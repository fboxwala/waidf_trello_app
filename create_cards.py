import csv
import difflib
import os
import re
import time
from urllib.parse import unquote

import jinja2
import requests
from bs4 import BeautifulSoup

# Trello credentials (set via env vars or replace the fallback strings)
API_KEY = os.getenv("TRELLO_API_KEY", "YOUR_API_KEY")
API_TOKEN = os.getenv("TRELLO_API_TOKEN", "YOUR_API_TOKEN")

# Optional: if you have a Google Cloud API key with Drive API enabled, set it here.
# If left empty, the script automatically scrapes the public Drive folder listing instead.
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")

DRIVE_FOLDER_ID = "1QXvqTm3kit1wLXyMsSx-tt6XrecsC5hT"
BOARD_NAME = "WAIDF Class of 2027 Selection Night"
LIST_NAME = "All Applicants"
TSV_FILE = "cands.tsv"
TEMPLATE_FILE = "desc.txt"
BASE_URL = "https://api.trello.com/1"

AUTH_PARAMS = {
    "key": API_KEY,
    "token": API_TOKEN,
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


def fetch_drive_application_links(folder_id: str) -> dict[str, str]:
    """
    Returns a dict mapping Submission ID (filename without extension) -> Google Drive view URL.
    Uses Google Drive API v3 if GOOGLE_API_KEY is set; otherwise parses the public embedded folder view.
    """
    links = {}

    if GOOGLE_API_KEY:
        page_token = None
        while True:
            params = {
                "q": f"'{folder_id}' in parents and trashed = false",
                "fields": "nextPageToken, files(id, name, webViewLink)",
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

            for file_info in data.get("files", []):
                stem = os.path.splitext(file_info["name"])[0].strip()
                links[stem] = file_info.get(
                    "webViewLink",
                    f"https://drive.google.com/file/d/{file_info['id']}/view?usp=drive_link",
                )

            page_token = data.get("nextPageToken")
            if not page_token:
                break
        return links

    url = f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    for entry in soup.select(".flip-entry"):
        link_tag = entry.select_one("a[href]")
        title_tag = entry.select_one(".flip-entry-title")
        if not link_tag or not title_tag:
            continue

        stem = os.path.splitext(title_tag.get_text(strip=True))[0].strip()
        links[stem] = link_tag["href"]

    return links


def get_board_and_list_ids():
    boards = trello_request("GET", "/members/me/boards", {"fields": "name,id"})
    board_id = next((b["id"] for b in boards if b["name"] == BOARD_NAME), None)
    if not board_id:
        raise ValueError(f"Board not found: {BOARD_NAME}")

    lists = trello_request("GET", f"/boards/{board_id}/lists", {"fields": "name,id"})
    list_id = next((lst["id"] for lst in lists if lst["name"] == LIST_NAME), None)
    if not list_id:
        raise ValueError(f"List not found on board '{BOARD_NAME}': {LIST_NAME}")

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


def get_candidate_label_ids(row: dict, label_ids: dict[str, str]) -> list[str]:
    applied = []

    county = (row.get("county") or "").strip()
    if county and "king" not in county.lower():
        applied.append(label_ids["Outside King County"])

    if is_yes(row.get("lgbt")):
        applied.append(label_ids["LGBTQIA+"])

    if is_bipoc(row.get("race")):
        applied.append(label_ids["BIPOC"])

    if (row.get("education") or "").strip().lower() == "high school / ged":
        applied.append(label_ids["No Higher Ed"])

    if is_yes(row.get("disabled")):
        applied.append(label_ids["Disabled"])

    if is_yes(row.get("immigrant_refugee")):
        applied.append(label_ids["Immigrant/Refugee"])

    if is_yes(row.get("parent")):
        applied.append(label_ids["Caregiver"])

    if is_yes(row.get("societal_systems")):
        applied.append(label_ids["Experience with Societal Systems"])

    if is_yes(row.get("poverty")):
        applied.append(label_ids["Experience with Poverty"])

    if is_yes(row.get("veteran")):
        applied.append(label_ids["Veteran/Military"])

    return applied


def sync_checklist(
    card_id: str,
    existing_checklists: dict[str, dict],
    name: str,
    expected_items: list[str],
    preserve_manual_items: bool = False,
) -> bool:
    """
    Ensures a checklist exists on the card with the exact expected items.
    Returns True if any API changes were made.
    """
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

    # For checklists like "Known Alum" that are populated manually later, just ensure it exists.
    if preserve_manual_items:
        return False

    existing_items = checklist.get("checkItems", [])
    existing_by_name = {}
    for item in existing_items:
        item_name = item["name"].strip()
        # Remove duplicates if any exist from a prior partial run
        if item_name in existing_by_name or item_name not in clean_expected:
            trello_request("DELETE", f"/checklists/{checklist_id}/checkItems/{item['id']}")
            changed = True
        else:
            existing_by_name[item_name] = item["id"]

    # Add any missing expected items
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


def build_demographics_items(row: dict) -> list[str]:
    if is_yes(row.get("union")):
        union_val = val_or_default(row.get("Please specify which union"), "Yes")
    else:
        union_val = "None"

    return [
        f"Race: {val_or_default(row.get('race'))}",
        f"LGBTQ: {val_or_default(row.get('lgbt'))}",
        f"Disabled: {val_or_default(row.get('disabled'))}",
        f"Immigrant/Refugee: {val_or_default(row.get('immigrant_refugee'))}",
        f"Caregiver/Parent: {val_or_default(row.get('parent'))}",
        f"Societal Systems: {val_or_default(row.get('societal_systems'))}",
        f"Poverty: {val_or_default(row.get('poverty'))}",
        f"Veteran: {val_or_default(row.get('veteran'))}",
        f"Union: {union_val}",
        f"Reapplied: {val_or_default(row.get('Have you ever applied to WAIDF in the past?'))}",
    ]


def build_lor_items(row: dict) -> list[str]:
    raw_lors = (
        row.get("Up to two (2) letters of recommendation (in PDF format).") or ""
    ).strip()
    if not raw_lors:
        return []

    items = []
    for url in raw_lors.split():
        filename = unquote(url.rsplit("/", 1)[-1])
        items.append(f"{filename} - {url}")
    return items


def main():
    print("Fetching application links from Google Drive...")
    try:
        drive_links = fetch_drive_application_links(DRIVE_FOLDER_ID)
        print(f"Found {len(drive_links)} files in Drive folder.")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching Google Drive folder: {e}")
        return

    try:
        board_id, list_id = get_board_and_list_ids()
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

    with open(TSV_FILE, mode="r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f, delimiter="\t")

        for row in reader:
            full_name = f"{(row.get('fname') or '').strip()} {(row.get('lname') or '').strip()}".strip()
            if not full_name:
                continue

            submission_id = (row.get("Submission ID") or "").strip()
            app_link = drive_links.get(submission_id, "")
            if not app_link:
                print(f"Warning: No Drive PDF found for {full_name} (Submission ID: {submission_id})")

            try:
                desc = template.render(
                    link=app_link,
                    email=(row.get("email") or "").strip(),
                    pronouns=(row.get("pronoun") or "").strip(),
                    cellphone=(row.get("phone") or "").strip(),
                    personal=(row.get("personal") or "").strip(),
                    civically=(row.get("civically") or "").strip(),
                    equity=(row.get("equity") or "").strip(),
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

                    # Preserve any manually added labels that aren't managed by this script
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
                    card_id, existing_checklists, "Letters of Recommendation", build_lor_items(row)
                )
                cl_changed |= sync_checklist(
                    card_id, existing_checklists, "Demographics", build_demographics_items(row)
                )
                cl_changed |= sync_checklist(
                    card_id, existing_checklists, "Known Alum", [], preserve_manual_items=True
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
