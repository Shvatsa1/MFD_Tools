# Rebalancer — get transaction suggestions (Python)

This tool reads your clients’ **holdings** and the **fund master** (`data/latestNAV_Reports.xlsx`) and writes **`output/generate_transactions.xlsx`**: suggested switches and supporting sheets (Parameters, AllClientHoldings, Check tab, copy of the NAV data used).

Works on **Windows, macOS, and Linux**. You need **Python 3.10+**; no other paid software.

---

## Get this folder

- **If you have the repo on GitHub (public):** you do **not** need a GitHub account. Open the green **Code** button → **Download ZIP**, unzip, then open the **`Rebalancer`** folder.  
- **Optional:** with a `git` install you can `git clone` a public URL without logging in.  
- **Private** repos need access (e.g. a GitHub login that was invited) or a zip shared by the publisher.

The NAV workbook should already be in **`Rebalancer/data/latestNAV_Reports.xlsx`** in the usual distribution. If you update the pack later, replace that file (or the whole folder) the same way the publisher sends it.

---

## One-time setup (each laptop)

1. Install **Python 3.10+** from [python.org](https://www.python.org/downloads/) (Windows: check **Add python.exe to PATH**).
2. Open a **terminal** in this **`Rebalancer`** folder (this folder is your working directory for every step below).
3. Create a virtual environment and install dependencies:

   **Windows (simplest):** double-click or run in Command Prompt: **`1_setup_venv.bat`**  
   **macOS / Linux:** `chmod +x 1_setup_venv.sh && ./1_setup_venv.sh`  
   **Or manually:**  
   - `python -m venv .venv`  
   - **Activate**  
     - Windows cmd: `.venv\Scripts\activate.bat`  
     - Windows PowerShell: `.venv\Scripts\Activate.ps1`  
     - macOS/Linux: `source .venv/bin/activate`  
   - `pip install -r requirements.txt`  
   - `copy mfd_pack.example.ini mfd_pack.ini` (on macOS/Linux: `cp mfd_pack.example.ini mfd_pack.ini`)

After setup, every time you work in a **new** terminal, **activate the venv** again, then run the commands in the next section.

---

## Ideal flow: run a rebalance and open results

1. **Activate** the virtual environment (see above), current directory = this **`Rebalancer`** folder.
2. **Client holdings (Format B):** one file per client in **`data/clients/by_client/`**  
   - The **file name** (without extension) = that client’s **client id** (e.g. `Client_A.xlsx`).  
   - Supported: `.xlsx`, `.xls`, `.csv` with a header row; the tool finds the ISIN and units columns automatically.
3. **Edit `mfd_pack.ini`** in Notepad (or any editor): set  
   - `clients_folder` (default: `data\clients\by_client` on Windows)  
   - `master` = path to the NAV file (default: `data\latestNAV_Reports.xlsx`)  
   - `archetype` = `Averse` / `Moderate` / `Aggressive` (unless you use glide, below)  
   - `new_cash` = new money to invest in rupees, or `0`  
   - Uncomment other lines only when you use **glide** or **new-fund** options (see below).
4. **Run:**  
   `python build_mfd_pack.py`  
5. **Output:** open **`output/generate_transactions.xlsx`**. The **Transactions** sheet lists suggested actions; use **Check Tab** and other sheets as needed.

If something fails, read the messages in the terminal; often it is a missing file path or an empty `clients` folder.

---

## Optional: age-based “glide” (target mix depends on each client’s age)

1. Put all client holding files in **`data/clients/by_client/`** as above.  
2. Run: `python build_mfd_pack.py --bootstrap-client-risk`  
   This builds/updates **`data/client_risk_pref.xlsx`**.  
3. Open **`data/client_risk_pref.xlsx`** → sheet **ClientAges** → set **age** and **risk_preference** (Averse / Moderate / Aggressive) for each client.  
4. In **`mfd_pack.ini`**, uncomment:  
   `age_based = true`  
   and  
   `client_risk_pref = data\client_risk_pref.xlsx` (use `/` on macOS/Linux if you prefer)  
5. Run again: `python build_mfd_pack.py`

---

## Optional: new-fund “sleeves” (FundChoices)

Only if the publisher has enabled this and you have a FundChoices workbook. In **`mfd_pack.ini`**, uncomment **`allow_new_funds`** and **`fund_choices`**, and point to the file they gave you. Then `python build_mfd_pack.py`.

---

## Your clients’ data

Do **not** put client names, account details, or holdings in public issues, public uploads, or shared drives meant for **strangers**. Work on your own computer or approved internal storage.
