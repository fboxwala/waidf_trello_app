# WAIDF Applicant-to-Trello Importer

This tool takes an Excel export straight from Jotform and automatically creates (or updates) Trello cards for every applicant—complete with their Google Drive document links, demographic labels, and checklists.

---

## How Re-Running Works (Safe to Run as Many Times as You Need)

This script is **idempotent**, which is a fancy way of saying you can run it over and over without breaking anything or creating duplicate cards.

Whenever new applications come in—or if an applicant's info/PDFs get updated, or if the script stops halfway due to a timeout—just download a fresh Excel file from Jotform and run the exact same steps again:

- **New applicants:** Automatically get brand-new cards in the `All Applicants` list.
- **Existing applicants:** Matched by name across the entire board (even if you've already moved their card out of `All Applicants` into another list). No duplicates are created.
- **Missing or updated info:** If a card is missing a Google Drive link, label, or checklist item—or if an answer changed in Jotform—the script updates just that piece.
- **Your manual work stays safe:** Any extra labels you've added by hand in Trello, any items you've typed into the `Known Alum` checklist, and any cards you've moved to other lists are left untouched.

---

## Before You Start (Checklist)

Make sure you have:
- Access to the **WAIDF Jotform** form, submissions table, and workflows
- Access to the **Google Drive folder** where Jotform saves applicant PDFs
- A **Trello account** with access to the WAIDF Workspace
- Access to the private **[WAIDF Trello API Credentials Google Doc](https://docs.google.com/document/d/1g7MCm6X1njoW1sfOXG8B4tScL8kON84IwDY-w9sGITs/edit?tab=t.0)**

---

## Part 1: Prepare Trello, Google Drive, and Jotform

### 1. Create the Trello Board
*(Skip this if the board already exists.)*
1. Log into [Trello](https://trello.com) and open the WAIDF Workspace.
2. Click **Create new board**.
3. Name the board (for example: `WAIDF Class of 2027 Selection Night`) and click **Create**.
4. Create a list named **`All Applicants`** (exact spelling and capitalization—the script looks for this list).
5. **Do not** manually create demographic labels; the script creates and color-codes all of them automatically.

### 2. Find & Check the Google Drive Folder Link
1. Open the WAIDF form in **Jotform** and go to **Settings → Workflows** (or open **My Workflows**).
2. Open the workflow called **`Make PDFs for Trello`**.
3. Find the Google Drive step in the workflow to get the link to the destination Google Drive folder (this is the root folder containing a subfolder for each applicant named by their numeric Submission ID).
4. Open that folder in Google Drive, click the folder name at the top → **Share** → **Share**.
5. Under **General access**, confirm it is set to **Anyone with the link** (as **Viewer**).
   > *Why?* The script reads the public folder listing to match each applicant's PDFs without requiring you to log into Google from the command line.
6. Click **Copy link** and keep it handy.

### 3. Download the Applications from Jotform
1. Open the WAIDF application table in Jotform.
2. Click **Download All** (top right) → **Download as Excel**.
3. Leave the file as `.xlsx`—**do not rename any columns** or edit the spreadsheet.

---

## Part 2: One-Time Computer Setup (Mac)

You only need to do this section **once** on your computer. All commands below are run in the **Terminal** app (press `Command + Space`, type `Terminal`, and press `Enter`).

> **Note:** When Terminal asks for your Mac password, nothing will appear on screen as you type—that's normal. Just type your password and press `Enter`.

### 1. Install Homebrew (a tool for installing developer software)
Copy and paste this into Terminal and press `Enter`:
```bash
/bin/bash -c "$(curl -fsSL [https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh](https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh))"
```
When it finishes, look at the last few lines under **"Next steps:"** in Terminal. If it gives you two commands to add Homebrew to your `PATH`, copy and run those too.

### 2. Install Git and Python
```bash
brew install git python
```

### 3. Download the Code
Choose where you want the folder to live (for example, your Desktop), then download the repository:
```bash
cd ~/Desktop
git clone [INSERT_PUBLIC_GIT_REPO_URL]
cd [INSERT_REPO_FOLDER_NAME]
```

### 4. Create a Virtual Environment & Install Required Packages
A virtual environment keeps this project's packages neatly isolated from the rest of your computer.
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```
You should now see `(venv)` at the beginning of your Terminal prompt.

---

## Part 3: Get Your Trello API Key & Token

1. Open the private **[WAIDF Trello API Credentials Google Doc]([INSERT_PRIVATE_GOOGLE_DOC_LINK])**.
2. Copy the **Trello API Key** listed in the document.
3. Click the **Generate Trello Token** link in the document, scroll to the bottom, and click **Allow**.
4. Copy the long token string (it starts with `ATTA...`). Keep it private—it acts like a password for your Trello account.

---

## Part 4: Run the Importer

Every time you want to import new applicants or sync updates from a fresh Jotform Excel download:

1. Open **Terminal** and go to your project folder:
   ```bash
   cd ~/Desktop/[INSERT_REPO_FOLDER_NAME]
   ```
2. Activate the virtual environment (you'll see `(venv)` appear):
   ```bash
   source venv/bin/activate
   ```
3. *(Optional)* Pull the latest version of the script in case anything was updated:
   ```bash
   git pull
   ```
4. Run the script:
   ```bash
   python3 create_cards.py
   ```
5. Answer the prompts as they appear:
   * **Excel/TSV filename:** Drag and drop the downloaded `.xlsx` file straight from Finder into the Terminal window, then press `Enter`.
   * **Trello Board Name:** Press `Enter` to use the default board name shown in brackets, or type your exact board name.
   * **Trello API Key:** Paste the API Key and press `Enter`.
   * **Trello API Token:** Paste your Token and press `Enter`.
   * **Public Google Drive folder link:** Paste the Google Drive folder link from the `Make PDFs for Trello` workflow and press `Enter`.

The script will print `Created`, `Updated`, or `Up to date` for each applicant. When it finishes, you'll see `All cards synced successfully!`.

---

## Troubleshooting

* **`ModuleNotFoundError: No module named 'openpyxl'` (or `requests`, `jinja2`, `bs4`)**
  You forgot to activate the virtual environment. Run `source venv/bin/activate` first, then run `python3 create_cards.py` again.
* **`List not found on board: All Applicants`**
  Check your Trello board and make sure you have a list named `All Applicants` (exact spelling and capitalization).
* **`Warning: No Drive subfolder found for [Name]`**
  Double-check that (1) the Google Drive folder from `Make PDFs for Trello` is set to **Anyone with the link can view**, and (2) the Jotform workflow has finished running for that submission. Once the subfolder appears in Drive, just re-run the script to fill in the missing links.
* **`Trello API error [401]: invalid token`**
  Your Trello token may have expired or had an extra space pasted in. Generate a fresh token from the link in the private Google Doc and run the script again.
* **The script stopped halfway due to a timeout or internet drop**
  Just run `python3 create_cards.py` again. Already-created cards will be skipped cleanly, and any incomplete cards will be finished.
