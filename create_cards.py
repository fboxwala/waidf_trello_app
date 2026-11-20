import csv
import difflib
import os
import re
import time
from urllib.parse import unquote

import jinja2
import openpyxl
import requests
from bs4 import BeautifulSoup

# Optional: Google Cloud API key with Drive API enabled.
# If left empty, the script automatically reads the public Drive folder HTML views.
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")

DEFAULT_BOARD_NAME = "WAIDF Class of 2027 Selection Night"
LIST_NAME = "All Applicants"
TEMPLATE_FILE = "desc.txt"
BASE_URL = "https://api.trello.com/1"

# Populated at runtime from user input
AUTH_PARAMS = {}

# 0-based column indices for the raw Jotform Excel export
COL = {
    "fname": 4,
    "lname": 5,
    "pronoun": 7,
    "phone": 8,
    "email": 9,
    "county": 13,
    "resume": 16,
    "personal": 17,
    "lors": 18,
    "civically": 19,
    "equity": 20,
    "race": 21,
    "lgbt": 22,
    "disabled": 23,
    "immigrant_refugee": 24,
    "parent": 25,
    "societal_systems": 26,
    "poverty": 27,
    "education": 28,
    "veteran": 29,
    "union": 30,
    "union_name": 31,
    "reapplied": 32,
    "how_heard": 35,
    "submission_id": 39,
}

LABEL_COLORS = {
    "Outside King County": "pink_dark",
    "LGBTQIA+": "orange",
    "BIPOC": "purple",
    "No Higher Ed": "yellow",
    "Disabled": "sky",
    "Immigrant/Refugee": "green",
    "Caregiver": "lime",
    "Experience with Societal Systems": "red",
    "Experience with Poverty": "blue",
    "Veteran/Military": "black",
}

WHITE_ONLY_TOKENS = {"white", "caucasian"}


class TrelloAPIError(Exception):
    """Custom exception that preserves Trello's status code and response body."""

    def __init__(self, method: str, endpoint: str, status_code: int | None, message: str):
        self.method = method
        self.endpoint = endpoint
        self.status_code = status_code
        self.message = message
        super().__init__(f"Trello API error [{status_code}] on {method} {endpoint}: {message}")


def prompt_required(prompt_text: str, default: str = "") -> str:
    """Prompts the user for a value, optionally showing a default they can accept with Enter."""
    while True:
        suffix = f" [{default}]" if default else ""
        val = input(f"{prompt_text}{suffix}: ").strip()
        if not val and default:
            return default
        if val:
            return val
        print("This field is required. Please enter a value.")


def extract_drive_folder_id(drive_input: str) -> str:
    """Extracts the folder ID from a Google Drive folder URL, or returns the raw ID."""
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", drive_input)
    if match:
        return match.group(1)
    match = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", drive_input)
    if match:
        return match.group(1)
    return drive_input.strip()


def get_user_inputs() -> tuple[str, str, str, str, str]:
    """Collects filename, Trello board name, API key/token, and Google Drive folder ID from the user."""
    while True:
        raw_path = prompt_required("Excel/TSV filename (or drag-and-drop file here)")
        filepath = raw_path.strip().strip('"').strip("'")
        if os.path.isfile(filepath):
            break
        print(f"File not found: {filepath}. Please check the filename and try again.")

    board_name = prompt_required("Trello Board Name", default=DEFAULT_BOARD_NAME)

    env_key = os.getenv("TRELLO_API_KEY", "")
    api_key = prompt_required("Trello API Key", default=env_key)

    env_token = os.getenv("TRELLO_API_TOKEN", "")
    api_token = prompt_required("Trello API Token", default=env_token)

    drive_link = prompt_required("Public Google Drive folder link")
    folder_id = extract_drive_folder_id(drive_link)

    return filepath, board_name, api_key, api_token, folder_id


def trello_request(method: str, endpoint: str, params=None, retries: int = 3):
    """Wrapper around requests that injects auth, handles 429 rate limits, and surfaces Trello error bodies."""
    url = f"{BASE_URL}{endpoint}"
    req_params = AUTH_PARAMS.copy()
    if params:
        req_params.update(params)

    for attempt in range(retries):
        try:
            resp = requests.request(method, url, params=req_params, timeout=30)
        except requests.exceptions.RequestException as e:
            if attempt < retries - 1:
                wait = 2 * (attempt + 1)
                print(f"Network error on {method} {endpoint}: {e}. Retrying in {wait}s...")
                time.sleep(wait)
                continue
            raise TrelloAPIError(method, endpoint, None, f"Network error: {e}") from e

        if resp.status_code == 429:
            wait = 2 * (attempt + 1)
            print(f"Rate limited by Trello on {method} {endpoint}. Waiting {wait}s...")
            time.sleep(wait)
            continue

        if not resp.ok:
            raise TrelloAPIError(method, endpoint, resp.status_code, resp.text.strip())

        time.sleep(0.1)
        try:
            return resp.json()
        except ValueError as e:
            raise TrelloAPIError(
                method, endpoint, resp.status_code, f"Invalid JSON response: {resp.text[:200]}"
            ) from e

    raise TrelloAPIError(method, endpoint, 429, f"Exceeded {retries} retries due to rate limiting.")


def list_public_drive_folder(folder_id: str) -> list[dict]:
    """
    Lists immediate children (files and subfolders) of a public Google Drive folder.
    Returns a list of dicts: {"id": str, "name": str, "url": str, "is_folder": bool}.
    """
    items = []

    if GOOGLE_API_KEY:
        page_token = None
        while True:
            params = {
                "q": f"'{folder_id}' in parents and trashed = false",
                "fields": "nextPageToken, files(id, name, mimeType, webViewLink)",
                "pageSize": 1000,
                "supportsAllDrives": "true",
                "includeItemsFromAllDrives": "true",
                "key": GOOGLE_API_KEY,
            }
            if page_token:
                params["pageToken"] = page_token

            resp = requests.get(
                "https://www.googleapis.com/drive/v3/files",
                params=params,
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()

            for f in data.get("files", []):
                is_folder = f.get("mimeType") == "application/vnd.google-apps.folder"
                default_url = (
                    f"https://drive.google.com/drive/folders/{f['id']}"
                    if is_folder
                    else f"https://drive.google.com/file/d/{f['id']}/view?usp=drive_link"
                )
                items.append({
                    "id": f["id"],
                    "name": f["name"].strip(),
                    "url": f.get("webViewLink", default_url),
                    "is_folder": is_folder,
                })

            page_token = data.get("nextPageToken")
            if not page_token:
                break
        return items

    url = f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    for entry in soup.select(".flip-entry"):
        link_tag = entry.select_one("a[href]")
        title_tag = entry.select_one(".flip-entry-title")
        if not link_tag or not title_tag:
            continue

        href = link_tag["href"]
        name = title_tag.get_text(strip=True)
        folder_match = re.search(r"/folders/([a-zA-Z0-9_-]+)", href)
        file_match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", href)

        if folder_match:
            items.append({
                "id": folder_match.group(1),
                "name": name,
                "url": f"https://drive.google.com/drive/folders/{folder_match.group(1)}",
                "is_folder": True,
            })
        elif file_match:
            items.append({
                "id": file_match.group(1),
                "name": name,
                "url": f"https://drive.google.com/file/d/{file_match.group(1)}/view?usp=drive_link",
                "is_folder": False,
            })

    return items


def normalize_filename(name: str) -> str:
    """Normalize whitespace (including non-breaking spaces) and casing for filename matching."""
    cleaned = unquote(name).replace("\xa0", " ")
    return re.sub(r"\s+", " ", cleaned).strip().lower()


def resolve_drive_file_url(jotform_url: str, subfolder_files: dict[str, str]) -> tuple[str, str]:
    """
    Given a Jotform upload URL and a map of {normalized_filename: drive_url} from the
    candidate's Drive subfolder, returns (display_filename, drive_url).
    """
    raw_filename = unquote(jotform_url.rsplit("/", 1)[-1]).strip()
    norm = normalize_filename(raw_filename)

    if norm in subfolder_files:
        return raw_filename, subfolder_files[norm]

    close = difflib.get_close_matches(norm, list(subfolder_files.keys()), n=1, cutoff=0.75)
    if close:
        return raw_filename, subfolder_files[close[0]]

    return raw_filename, ""


def read_candidate_rows(filepath: str) -> list[list[str]]:
    """Reads rows (excluding the header row) from an .xlsx or .tsv file."""
    ext = os.path.splitext(filepath)[1].lower()
    rows = []

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append(["" if cell is None else str(cell).strip() for cell in row])
        wb.close()
    else:
        with open(filepath, mode="r", encoding="utf-8", newline="") as f:
            reader = csv.reader(f, delimiter="\t")
            next(reader, None)
            for row in reader:
                rows.append([cell.strip() for cell in row])

    return rows


def cell(row: list[str], key: str) -> str:
    idx = COL[key]
    return row[idx].strip() if idx < len(row) else ""


def get_board_and_list_ids(board_name: str):
    boards = trello_request("GET", "/members/me/boards", {"fields": "name,id"})
    board_id = next((b["id"] for b in boards if b["name"] == board_name), None)
    if not board_id:
        raise ValueError(f"Board not found: {board_name}")

    lists = trello_request("GET", f"/boards/{board_id}/lists", {"fields": "name,id"})
    list_id = next((lst["id"] for lst in lists if lst["name"] == LIST_NAME), None)
    if not list_id:
        raise ValueError(f"List not found on board '{board_name}': {LIST_NAME}")

    return board_id, list_id


def get_existing_cards(board_id: str) -> dict[str, dict]:
    """Fetch all open cards on the board, keyed by lowercase card name."""
    cards = trello_request(
        "GET",
        f"/boards/{board_id}/cards",
        {"fields": "id,name,desc,idLabels,idList"},
    )
    return {c["name"].strip().lower(): c for c in cards if c.get("name")}


def ensure_labels(board_id: str) -> dict[str, str]:
    """Fetch existing board labels and create any missing ones from LABEL_COLORS."""
    existing = trello_request("GET", f"/boards/{board_id}/labels", {"limit": 1000})
    existing_by_lower = {
        lbl["name"].strip().lower(): lbl["id"]
        for lbl in existing
        if lbl.get("name")
    }

    label_ids = {}
    for name, color in LABEL_COLORS.items():
        key = name.lower()
        if key in existing_by_lower:
            label_ids[name] = existing_by_lower[key]
        else:
            created = trello_request(
                "POST",
                f"/boards/{board_id}/labels",
                {"name": name, "color": color},
            )
            label_ids[name] = created["id"]
            print(f"Created label: {name}")

    return label_ids


def is_yes(value: str) -> bool:
    return (value or "").strip().lower() == "yes"


def is_bipoc(race_val: str) -> bool:
    """
    Returns True if the free-form race field includes any identity beyond
    white/caucasian (handling minor typos while preserving multiracial entries).
    """
    if not race_val or not race_val.strip():
        return False

    tokens = re.findall(r"[a-zA-Z]+", race_val.lower())
    if not tokens:
        return False

    connectors = {"and", "or"}
    non_white_tokens = []
    for tok in tokens:
        if tok in connectors:
            continue
        if difflib.get_close_matches(tok, WHITE_ONLY_TOKENS, n=1, cutoff=0.8):
            continue
        non_white_tokens.append(tok)

    return len(non_white_tokens) > 0


def get_candidate_label_ids(row: list[str], label_ids: dict[str, str]) -> list[str]:
    applied = []

    county = cell(row, "county")
    if county and "king" not in county.lower():
        applied.append(label_ids["Outside King County"])

    if is_yes(cell(row, "lgbt")):
        applied.append(label_ids["LGBTQIA+"])

    if is_bipoc(cell(row, "race")):
        applied.append(label_ids["BIPOC"])

    if cell(row, "education").lower() == "high school / ged":
        applied.append(label_ids["No Higher Ed"])

    if is_yes(cell(row, "disabled")):
        applied.append(label_ids["Disabled"])

    if is_yes(cell(row, "immigrant_refugee")):
        applied.append(label_ids["Immigrant/Refugee"])

    if is_yes(cell(row, "parent")):
        applied.append(label_ids["Caregiver"])

    if is_yes(cell(row, "societal_systems")):
        applied.append(label_ids["Experience with Societal Systems"])

    if is_yes(cell(row, "poverty")):
        applied.append(label_ids["Experience with Poverty"])

    if is_yes(cell(row, "veteran")):
        applied.append(label_ids["Veteran/Military"])

    return applied


def sync_checklist(
    card_id: str,
    existing_checklists: dict[str, dict],
    name: str,
    expected_items: list[str],
    preserve_manual_items: bool = False,
) -> bool:
    """Ensures a checklist exists on the card with the exact expected items."""
    changed = False
    clean_expected = [item.strip() for item in expected_items if item and item.strip()]
    key = name.strip().lower()

    if key not in existing_checklists:
        checklist = trello_request(
            "POST",
            f"/cards/{card_id}/checklists",
            {"name": name},
        )
        checklist_id = checklist["id"]
        changed = True
        for item in clean_expected:
            trello_request(
                "POST",
                f"/checklists/{checklist_id}/checkItems",
                {"name": item},
            )
        return changed

    checklist = existing_checklists[key]
    checklist_id = checklist["id"]

    if preserve_manual_items:
        return False

    existing_items = checklist.get("checkItems", [])
    existing_by_name = {}
    for item in existing_items:
        item_name = item["name"].strip()
        if item_name in existing_by_name or item_name not in clean_expected:
            trello_request("DELETE", f"/checklists/{checklist_id}/checkItems/{item['id']}")
            changed = True
        else:
            existing_by_name[item_name] = item["id"]

    for item_name in clean_expected:
        if item_name not in existing_by_name:
            trello_request(
                "POST",
                f"/checklists/{checklist_id}/checkItems",
                {"name": item_name},
            )
            changed = True

    return changed


def val_or_default(value: str, default: str = "N/A") -> str:
    cleaned = (value or "").strip()
    return cleaned if cleaned else default


def build_demographics_items(row: list[str]) -> list[str]:
    if is_yes(cell(row, "union")):
        union_val = val_or_default(cell(row, "union_name"), "Yes")
    else:
        union_val = "None"

    return [
        f"Race: {val_or_default(cell(row, 'race'))}",
        f"LGBTQ: {val_or_default(cell(row, 'lgbt'))}",
        f"Disabled: {val_or_default(cell(row, 'disabled'))}",
        f"Immigrant/Refugee: {val_or_default(cell(row, 'immigrant_refugee'))}",
        f"Caregiver/Parent: {val_or_default(cell(row, 'parent'))}",
        f"Societal Systems: {val_or_default(cell(row, 'societal_systems'))}",
        f"Poverty: {val_or_default(cell(row, 'poverty'))}",
        f"Veteran: {val_or_default(cell(row, 'veteran'))}",
        f"Union: {union_val}",
        f"Reapplied: {val_or_default(cell(row, 'reapplied'))}",
    ]


def build_lor_items(row: list[str], subfolder_files: dict[str, str], full_name: str) -> list[str]:
    raw_lors = cell(row, "lors")
    if not raw_lors:
        return []

    items = []
    for url in raw_lors.split():
        if not url.startswith("http"):
            continue
        filename, drive_url = resolve_drive_file_url(url, subfolder_files)
        if drive_url:
            items.append(f"{filename} - {drive_url}")
        else:
            print(f"Warning: No Drive match for LoR '{filename}' ({full_name})")
            items.append(filename)
    return items


def main():
    input_file, board_name, api_key, api_token, drive_root_folder_id = get_user_inputs()
    AUTH_PARAMS["key"] = api_key
    AUTH_PARAMS["token"] = api_token

    print(f"\nReading applications from {input_file}...")
    try:
        rows = read_candidate_rows(input_file)
        print(f"Loaded {len(rows)} row(s).")
    except Exception as e:
        print(f"Failed to read {input_file}: {e}")
        return

    print("Fetching Submission ID subfolders from Google Drive...")
    try:
        root_entries = list_public_drive_folder(drive_root_folder_id)
        subfolders_by_id = {
            entry["name"]: entry
            for entry in root_entries
            if entry["is_folder"]
        }
        print(f"Found {len(subfolders_by_id)} submission subfolder(s) in Drive.")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching Google Drive root folder: {e}")
        return

    try:
        board_id, list_id = get_board_and_list_ids(board_name)
        label_ids = ensure_labels(board_id)
        managed_label_id_set = set(label_ids.values())
        existing_cards = get_existing_cards(board_id)
        print(f"Found {len(existing_cards)} existing card(s) on board.")
    except (TrelloAPIError, ValueError) as e:
        print(f"Setup failed: {e}")
        return

    env = jinja2.Environment(loader=jinja2.FileSystemLoader("."))
    template = env.get_template(TEMPLATE_FILE)

    failed_cards = []

    for row in rows:
        full_name = f"{cell(row, 'fname')} {cell(row, 'lname')}".strip()
        if not full_name:
            continue

        submission_id = cell(row, "submission_id")
        subfolder = subfolders_by_id.get(submission_id)
        subfolder_files = {}
        folder_link = ""

        if subfolder:
            folder_link = subfolder["url"]
            try:
                children = list_public_drive_folder(subfolder["id"])
                subfolder_files = {
                    normalize_filename(c["name"]): c["url"]
                    for c in children
                    if not c["is_folder"]
                }
            except requests.exceptions.RequestException as e:
                print(f"Warning: Could not list Drive subfolder for {full_name} ({submission_id}): {e}")
        else:
            print(f"Warning: No Drive subfolder found for {full_name} (Submission ID: {submission_id})")

        app_link = subfolder_files.get(normalize_filename(f"{submission_id}.pdf"), folder_link)

        raw_personal = cell(row, "personal")
        if raw_personal.startswith("http"):
            _, personal_link = resolve_drive_file_url(raw_personal, subfolder_files)
            if not personal_link:
                print(f"Warning: No Drive match for Personal Statement ({full_name})")
        else:
            personal_link = raw_personal

        raw_resume = cell(row, "resume")
        if raw_resume.startswith("http"):
            _, resume_link = resolve_drive_file_url(raw_resume, subfolder_files)
            if not resume_link:
                print(f"Warning: No Drive match for Resume ({full_name})")
        else:
            resume_link = raw_resume

        try:
            desc = template.render(
                link=app_link,
                folder_link=folder_link,
                resume=resume_link,
                email=cell(row, "email"),
                pronouns=cell(row, "pronoun"),
                cellphone=cell(row, "phone"),
                personal=personal_link,
                civically=cell(row, "civically"),
                equity=cell(row, "equity"),
                how_heard=cell(row, "how_heard"),
            )

            expected_label_ids = get_candidate_label_ids(row, label_ids)
            card_key = full_name.lower()

            if card_key not in existing_cards:
                card = trello_request(
                    "POST",
                    "/cards",
                    {
                        "idList": list_id,
                        "name": full_name,
                        "desc": desc,
                        "idLabels": ",".join(expected_label_ids),
                    },
                )
                card_id = card["id"]
                existing_checklists = {}
                action = "Created"
            else:
                existing_card = existing_cards[card_key]
                card_id = existing_card["id"]
                card_updated = False

                current_label_ids = set(existing_card.get("idLabels", []))
                unmanaged_labels = current_label_ids - managed_label_id_set
                target_label_ids = unmanaged_labels | set(expected_label_ids)

                update_params = {}
                if existing_card.get("desc", "") != desc:
                    update_params["desc"] = desc
                if current_label_ids != target_label_ids:
                    update_params["idLabels"] = ",".join(sorted(target_label_ids))

                if update_params:
                    trello_request("PUT", f"/cards/{card_id}", update_params)
                    card_updated = True

                raw_checklists = trello_request("GET", f"/cards/{card_id}/checklists")
                existing_checklists = {
                    cl["name"].strip().lower(): cl for cl in raw_checklists
                }
                action = "Updated" if card_updated else "Up to date"

            cl_changed = False
            cl_changed |= sync_checklist(
                card_id,
                existing_checklists,
                "Letters of Recommendation",
                build_lor_items(row, subfolder_files, full_name),
            )
            cl_changed |= sync_checklist(
                card_id,
                existing_checklists,
                "Demographics",
                build_demographics_items(row),
            )
            cl_changed |= sync_checklist(
                card_id,
                existing_checklists,
                "Known Alum",
                [],
                preserve_manual_items=True,
            )

            if action == "Up to date" and cl_changed:
                action = "Updated"

            print(f"{action}: {full_name} ({len(expected_label_ids)} labels)")

        except TrelloAPIError as e:
            print(f"FAILED [{full_name}]: {e}")
            failed_cards.append((full_name, str(e)))
        except Exception as e:
            print(f"UNEXPECTED ERROR [{full_name}]: {e}")
            failed_cards.append((full_name, str(e)))

    if failed_cards:
        print(f"\nFinished with {len(failed_cards)} failure(s):")
        for name, err in failed_cards:
            print(f"  - {name}: {err}")
    else:
        print("\nAll cards synced successfully!")


if __name__ == "__main__":
    main()
