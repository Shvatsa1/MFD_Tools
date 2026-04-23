# MFD Rebalancer — run from source (no Windows `.exe`)

This folder is a **minimal Python-only** build of the rebalancer: the same `build_mfd_pack.py` entry point as the frozen pack, with **no** PyInstaller, **no** SmartScreen warnings, and it works on **macOS, Linux, and Windows**.

**Repository layout:** this tree lives under the **`MFD_Tools`** repo as **`Rebalancer/`**. Open a terminal **here** (inside `Rebalancer`) for all setup scripts and `python build_mfd_pack.py`.

---

## Distributor GitHub repo (MFD-facing only)

This folder is the MFD-facing app. It does **not** include the upstream “pack creator” (`tools/assemble_mfd_source_distribution.py`); that stays in the development repo.

| Topic | Guidance |
|--------|----------|
| **`latestNAV_Reports.xlsx` in git** | Yes. You can commit updated copies whenever you refresh NAV; MFDs `git pull` to update. The file is large and binary, so history grows over time—acceptable for a small repo, or use **Git LFS** if you prefer. |
| **Client holdings** | **Never** commit. Only the MFD’s machine; not this repo. |
| **Public vs private** | **Public** is usually fine for *code + AMFI-style NAV master* (no PII, no per-client data). If you are cautious about visibility or your terms of use, use a **private** repo and add only GitHub users you want (`Settings → Collaborators` or an org team). |
| **Refreshing the repo from development** | In the main project, run `python tools/assemble_mfd_source_distribution.py` with `-o` pointing at this `Rebalancer` folder (and `--include-nav` for the workbook), then commit and push **MFD_Tools**. |

---

## What you need (once per machine)

1. **Python 3.10+** from [python.org](https://www.python.org/downloads/) (Windows: tick **Add python.exe to PATH**).
2. Clone **MFD_Tools** and **`cd Rebalancer`**, or unzip a release that contains this folder.
3. **NAV master** `data/latestNAV_Reports.xlsx` (must include **Final** columns U–X). It is usually **committed** in this repo; after `git pull`, you get updates. If you only have a copy without it, get one from your distributor.

---

## Quickest path (create venv and install)

**Windows (double-click or run in cmd):** `1_setup_venv.bat`  
**macOS / Linux (Terminal):** `chmod +x 1_setup_venv.sh && ./1_setup_venv.sh`  
**Windows PowerShell:** `.\1_setup_venv.ps1`

Or manually:

```text
python3 -m venv .venv
```

Activate:

- **Windows (cmd):** `.venv\Scripts\activate.bat`
- **Windows (PowerShell):** `.venv\Scripts\Activate.ps1`
- **macOS / Linux:** `source .venv/bin/activate`

```text
pip install -r requirements.txt
copy mfd_pack.example.ini mfd_pack.ini
```

(On Unix use `cp` instead of `copy`.)

---

## Before each run

1. Put `latestNAV_Reports.xlsx` in **`data/`** (if not already there).
2. Put one holdings file per client in **`data/clients/by_client/`** (filename = client id), **or** set `clients_folder` in `mfd_pack.ini`.
3. Edit **`mfd_pack.ini`**: `master`, `output`, `archetype`, `new_cash`, and optionally glide (`age_based`, `client_risk_pref`) or new funds (`allow_new_funds`, `fund_choices`).

---

## Run the pack

```text
python build_mfd_pack.py
```

**Bootstrap** client list for the glide workbook (optional):

```text
python build_mfd_pack.py --bootstrap-client-risk
```

**Explicit paths** (no INI): see docstring at the top of `build_mfd_pack.py`.

Default output: `output/generate_transactions.xlsx`.

---

## Differences from the Windows `.exe` pack

| `.exe` pack | This source pack |
|-------------|------------------|
| One download, no Python | Python + venv + `pip install -r requirements.txt` |
| Windows only | **macOS, Linux, Windows** |
| Can trigger SmartScreen | No unsigned `.exe` |
| `MFRebalancerV1.1.exe` with cwd beside data | `python build_mfd_pack.py` with cwd in this folder |

Functionality of `build_mfd_pack.py` and `mfd_pack.ini` is the same as the V1.1 frozen app.

---

## Re-sync from upstream development (maintainers only)

In the **main** development repo (`MF_rebalancer`), from its root:

`python tools/assemble_mfd_source_distribution.py -o path/to/MFD_Tools/Rebalancer`  
(optional) `... --include-nav`  
Then commit and push the **MFD_Tools** repository.
