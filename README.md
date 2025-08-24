# Capitaline Consolidator

This tool consolidates multiple Capitaline exports (Excel/CSV) into one Excel file with three sheets:

1. **Div Adj Close Price** (NSE prioritized, fallback BSE)
2. **Daily Total Return (%)** (NSE prioritized, fallback BSE)
3. **Average Marketcap** (average of NSE and BSE)

---

## 📂 Folder Structure

```
project-folder/
│
├── main.py          # Main script
├── setup.bat        # Windows setup script (creates virtual env + installs deps)
├── README.md        # Instructions (this file)
├── logs/            # Logs will be written here
└── assets/          # Place all input Excel/CSV files here
```

* Put all your Capitaline Excel/CSV files inside the **assets/** folder.
* The script will automatically detect and process them.

---

## ⚙️ Setup Instructions

### 1. Install Python (if not already installed)

* Download and install Python 3.8+ from [python.org](https://www.python.org/downloads/windows/).
* During installation, ensure you check **"Add Python to PATH"**.

### 2. Run the setup script

* Double-click `setup.bat` in the project folder.
* This will:

  * Create a virtual environment in `.venv/`
  * Install required dependencies (`pandas`, `openpyxl`, `xlsxwriter`)

### 3. Activate the environment

Each time you open a new terminal/PowerShell session:

```powershell
call .venv\Scripts\activate
```

### 4. Run the script

```powershell
python main.py
```

### 5. Debug mode (optional)

To see detailed logs (per-column mapping, previews, etc.):

```powershell
python main.py --debug
```

---

## 📊 Output

* The script produces `consolidated_output.xlsx` in the project folder.
* Logs are written to `logs/consolidation.log`.

The Excel file will have 3 sheets:

* **Div Adj Close Price** → NSE prioritized closing price.
* **Daily Total Return (%)** → NSE prioritized returns.
* **Average Marketcap** → average of NSE and BSE marketcaps.

---

## 📝 Notes

* Ensure all Excel files are **closed** before running (to avoid `~$` lockfiles).
* The script expects real data headers to start at **row 2** (row 1 is metadata in Capitaline exports).
* Supported file formats: `.csv`, `.xlsx`, `.xls`.

---

## 🚀 Example Workflow

```powershell
cd D:\Coding Projects\Python\combinator
.\setup.bat                # run once to install deps
call .venv\Scripts\activate # activate venv
python main.py              # run normally
python main.py --debug      # run with detailed logs
```
