# Rebalancer

Reads each client's mutual-fund **holdings** plus the supplied **fund master** workbook and writes **`output/generate_transactions.xlsx`** with suggested switches and supporting sheets (Parameters, AllClientHoldings, Check tab, copy of the NAV data used).

Runs on **Windows 10/11** and **macOS**. No paid software.

---

## 1. Download this folder

1. Go to the project page on GitHub.
2. Click the green **Code** button → **Download ZIP**.
3. Unzip. Open the **`Rebalancer`** folder. All steps below assume this folder is the **working directory**.

(No GitHub account is needed to download a public repository.)

---

## 2. Install Python (one time per computer)

You need **Python 3.10 or newer**.

### Windows

1. Open https://www.python.org/downloads/windows/ → **Download Python 3.x.x** (64-bit).
2. Run the installer. **Tick "Add python.exe to PATH"** at the bottom of the first screen, then **Install Now**.
3. Open **Command Prompt** (press `Win`, type `cmd`, press Enter) and check:
   ```
   python --version
   ```
   You should see `Python 3.10.x` or higher.

### macOS

1. Open https://www.python.org/downloads/macos/ → **Download Python 3.x.x**.
2. Open the downloaded `.pkg` and follow the installer.
3. Open **Terminal** (press `Cmd+Space`, type `Terminal`, press Enter) and check:
   ```
   python3 --version
   ```
   You should see `Python 3.10.x` or higher.

(macOS users may also use `brew install python` if Homebrew is installed.)

---

## 3. Open a terminal and move into the `Rebalancer` folder

All following steps assume the terminal's working directory is the **`Rebalancer`** folder.

### Windows

**Open a terminal (Command Prompt):**

- Press `Win + R`, type `cmd`, press `Enter`. **Or** press the `Win` key, type `cmd`, press `Enter`.
- Alternative: in File Explorer, browse to the `Rebalancer` folder, then hold `Shift` and right-click the folder → **Open in Terminal** (Windows 11) or **Open command window here**. This skips the `cd` step.

**Change to the folder** (replace the path with where you unzipped):

```
cd /d "C:\Users\YOUR_USER\Downloads\MFD_Tools-main\Rebalancer"
```

Tips:
- The `/d` flag lets `cd` also switch drive letters (e.g. from `C:` to `D:`). Without it you'd also need `D:` first.
- Quotes are required only if the path has spaces; using them always is safe.

### macOS

**Open a terminal:**

- Press `Cmd + Space` (Spotlight), type `Terminal`, press `Return`.
- Alternative: in Finder, right-click the `Rebalancer` folder → **New Terminal at Folder**. This skips the `cd` step. If the option is missing, enable it in **System Settings → Keyboard → Keyboard Shortcuts → Services → Files and Folders → "New Terminal at Folder"**.

**Change to the folder** (replace the path with where you unzipped):

```
cd ~/Downloads/MFD_Tools-main/Rebalancer
```

Tips:
- `~` is your home folder, e.g. `/Users/yourname`.
- Quote paths with spaces: `cd "~/My Downloads/MFD_Tools-main/Rebalancer"`.

Verify you are in the right folder:

- Windows: `dir` (should list `build_mfd_pack.py`, `requirements.txt`, etc.)
- macOS: `ls` (same files visible).

---

## 4. One-time setup of this folder

### Windows

Double-click **`1_setup_venv.bat`** in the `Rebalancer` folder.

If double-click does nothing, open a terminal in the folder (section 3) and run:
```
1_setup_venv.bat
```

### macOS

In the terminal opened in section 3 (working directory = `Rebalancer`), run:
```
chmod +x 1_setup_venv.sh
./1_setup_venv.sh
```

The script creates a local Python environment (`.venv`), installs the libraries listed in `requirements.txt`, and creates `mfd_pack.ini` from the template.

---

## 5. Each time you use the tool

Open a terminal and `cd` into the `Rebalancer` folder (see section 3), then **activate the environment**:

- **Windows (Command Prompt):**
  ```
  .venv\Scripts\activate.bat
  ```
- **Windows (PowerShell):**
  ```
  .venv\Scripts\Activate.ps1
  ```
- **macOS:**
  ```
  source .venv/bin/activate
  ```

You will see `(.venv)` at the start of the prompt when active.

---

## 6. Add client data

Place **one holdings file per client** in **`data/clients/by_client/`**:

- File name (without extension) is the client id, e.g. `CLIENT_001.xlsx` → client id `CLIENT_001`.
- Accepted formats: `.xlsx`, `.xls`, `.csv`. The file must have a header row.
- The tool detects the ISIN column and the units column automatically.

Update the **fund master** when a refreshed copy is supplied: replace `data/latestNAV_Reports.xlsx` with the new file (same name, same location).

---

## 7. Configure `mfd_pack.ini`

Open **`mfd_pack.ini`** in Notepad (Windows) or TextEdit (macOS) and set:

- `clients_folder` — folder with client files (default: `data\clients\by_client`).
- `master` — fund master path (default: `data\latestNAV_Reports.xlsx`).
- `archetype` — `Averse`, `Moderate`, or `Aggressive`.
- `new_cash` — fresh money to invest in rupees, or `0`.

Lines beginning with `;` are off. Remove the `;` to enable optional features (see sections 9 and 10).

---

## 8. Run and view results

With the environment active, in the `Rebalancer` folder:

```
python build_mfd_pack.py
```

Open the result:

```
output\generate_transactions.xlsx
```

The **Transactions** sheet lists suggested actions per client. Other sheets (Parameters, AllClientHoldings, Check Tab, NAV copy) are for review and audit.

---

## 9. Optional: per-client age-based glide

For target equity/defensive that depends on each client's age:

1. With holding files in `data/clients/by_client/`, run:
   ```
   python build_mfd_pack.py --bootstrap-client-risk
   ```
2. Open `data/client_risk_pref.xlsx` → sheet **ClientAges** → fill **age** and **risk_preference** (Averse / Moderate / Aggressive) for each client.
3. In `mfd_pack.ini` remove the leading `;` on these lines:
   ```
   age_based = true
   client_risk_pref = data\client_risk_pref.xlsx
   ```
4. Run again:
   ```
   python build_mfd_pack.py
   ```

---

## 10. Optional: new-fund "sleeves"

If a FundChoices workbook has been supplied, in `mfd_pack.ini` remove the leading `;` from `allow_new_funds` and `fund_choices` and point the latter to that file. Then run `python build_mfd_pack.py`.

---

## Client data note

Keep client holdings and any client-specific files on your own computer or approved internal storage. Do not upload them to public locations.
