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

## 3. One-time setup of this folder

Open a terminal in the **`Rebalancer`** folder.

- **Windows tip:** in File Explorer, hold `Shift`, right-click the `Rebalancer` folder, choose **Open in Terminal** (or **Open command window here**).
- **macOS tip:** right-click the `Rebalancer` folder in Finder → **New Terminal at Folder** (enable in System Settings → Keyboard → Keyboard Shortcuts → Services if hidden).

### Windows

Double-click **`1_setup_venv.bat`** in the `Rebalancer` folder.

If double-click does nothing, open **Command Prompt** in the folder and run:
```
1_setup_venv.bat
```

### macOS

In Terminal, in the `Rebalancer` folder:
```
chmod +x 1_setup_venv.sh
./1_setup_venv.sh
```

The script creates a local Python environment (`.venv`), installs the libraries listed in `requirements.txt`, and creates `mfd_pack.ini` from the template.

---

## 4. Each time you use the tool

Open a terminal in the `Rebalancer` folder and **activate the environment**:

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

## 5. Add client data

Place **one holdings file per client** in **`data/clients/by_client/`**:

- File name (without extension) is the client id, e.g. `CLIENT_001.xlsx` → client id `CLIENT_001`.
- Accepted formats: `.xlsx`, `.xls`, `.csv`. The file must have a header row.
- The tool detects the ISIN column and the units column automatically.

Update the **fund master** when a refreshed copy is supplied: replace `data/latestNAV_Reports.xlsx` with the new file (same name, same location).

---

## 6. Configure `mfd_pack.ini`

Open **`mfd_pack.ini`** in Notepad (Windows) or TextEdit (macOS) and set:

- `clients_folder` — folder with client files (default: `data\clients\by_client`).
- `master` — fund master path (default: `data\latestNAV_Reports.xlsx`).
- `archetype` — `Averse`, `Moderate`, or `Aggressive`.
- `new_cash` — fresh money to invest in rupees, or `0`.

Lines beginning with `;` are off. Remove the `;` to enable optional features (see sections 8 and 9).

---

## 7. Run and view results

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

## 8. Optional: per-client age-based glide

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

## 9. Optional: new-fund "sleeves"

If a FundChoices workbook has been supplied, in `mfd_pack.ini` remove the leading `;` from `allow_new_funds` and `fund_choices` and point the latter to that file. Then run `python build_mfd_pack.py`.

---

## Client data note

Keep client holdings and any client-specific files on your own computer or approved internal storage. Do not upload them to public locations.
