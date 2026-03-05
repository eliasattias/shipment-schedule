## Shipment Schedule App

This is a small Streamlit app that lets your team view the daily **Pending Orders** export and adjust the **Follow up** date for each row using a calendar UI. Changes are stored in a local SQLite database so they persist even when a new CSV is loaded.

### Folder structure

- `app.py` – main Streamlit app
- `requirements.txt` – Python dependencies
- `data/` – place your daily `Pending Orders *.csv` file(s) here
- `orders.db` – created automatically; stores user-edited follow-up dates

### Python & virtual environment setup

From PowerShell:

```powershell
cd "C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule"

# Create virtual environment (one time)
python -m venv .venv

# Activate virtual environment
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

Whenever you come back in a new terminal, just:

```powershell
cd "C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule"
.\.venv\Scripts\Activate.ps1
```

### Running the app

With the virtual environment activated:

```powershell
streamlit run app.py
```

Your browser will open to the local URL (usually `http://localhost:8501`). Press `Ctrl+C` in the terminal to stop the app.

### Daily workflow

1. Export/download the latest **Pending Orders** CSV from your source system.
2. Save or copy the file into:
   - `C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule\data`
   - The filename should match `Pending Orders *.csv` or `Pending Orders *.xlsx` (e.g. `Pending Orders  March 04 2026.xlsx`).
3. Make sure the Streamlit app is running (or restart it if needed).
4. Refresh/reload the browser page if it was already open.
5. Use the table to:
   - Filter by `Customer` (via the filter box).
   - Edit the `Follow up` column using the calendar picker.
6. Click **Save changes** to write all modified follow-up dates into `orders.db`.

### How tomorrow’s file works

The daily export usually doesn’t include Follow up (scheduled date). The app handles that by:

1. **Identifying the same orders** – Each order is matched by Customer + Created Date + Cases # + Sales (normalized so CSV and Excel formats match).
2. **Applying saved scheduled dates** – If an order was given a Follow up date before, that date is loaded from `orders.db` and applied to the same order in the new file.
3. **Adding new orders** – Rows that have never been saved appear with an empty Follow up; set the date in the table and click **Save changes** so it’s remembered for the next day.

So: drop the new file in `data/`, refresh the app—existing orders keep their scheduled date, new orders are there for you to set and save.

