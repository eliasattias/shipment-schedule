## Shipment Schedule App

This is a small Streamlit app that lets your team view the daily **Pending Orders** export and adjust the **Follow up** date for each row using a calendar UI. Changes are stored in a local SQLite database so they persist even when a new CSV is loaded.

### Folder structure

- `app.py` – main Streamlit app
- `requirements.txt` – Python dependencies
- `data/` – place your daily `Pending Orders *.csv` file(s) here
- `orders.db` – created automatically when using the local SQLite backend; stores user-edited follow-up dates


### Cloud database (PostgreSQL) 🔗

By default the app uses a local SQLite file (`orders.db`) which is fine for a
single user. To allow multiple team members or deployed instances of the app
to share updates you can point the app at a cloud database such as
PostgreSQL. Streamlit Cloud, Heroku, Supabase, ElephantSQL, and many other
services provide free or low-cost PostgreSQL instances.

#### Using Supabase (recommended)

Supabase is a developer-friendly hosted PostgreSQL service with a generous free
tier and an easy web UI. The steps below outline the minimal process:

1. **Create a Supabase account** at https://supabase.com and log in.
2. **Create a new project** and choose a name, password and region. You can
   start with the free tier.
3. Once the project is ready, go to **Settings → Database** and copy the
   connection string labeled **Connection string**. It will look like:
   `postgresql://postgres:[YOUR-PASSWORD]@db.lpwzixfqprgrhzeyplxq.supabase.co:5432/postgres`.
4. Within the same settings page, go to **API → Project URL** and note the
   base URL; you’ll need it for future Supabase client usage if you expand the
   app.
5. Add the `DATABASE_URL` to your Streamlit secrets or environment:
   ```toml
   # .streamlit/secrets.toml
   DATABASE_URL = "postgresql://postgres:[YOUR-PASSWORD]@db.lpwzixfqprgrhzeyplxq.supabase.co:5432/postgres"
   ```
   On Windows you can also `setx DATABASE_URL "postgresql://postgres:[YOUR-PASSWORD]@db.lpwzixfqprgrhzeyplxq.supabase.co:5432/postgres"`.

6. **Install dependencies** (if not already):
   ```powershell
   pip install -r requirements.txt
   ```
   `psycopg2-binary` is already included in `requirements.txt`.

7. **Run the app** locally or deploy to Streamlit Cloud. The first run will
   create the `followup_overrides` table in your Supabase database.

8. **Migrate existing overrides** (optional):
   ```bash
   sqlite3 orders.db "SELECT * FROM followup_overrides;" > overrides.csv
   psql "$DATABASE_URL" -c "\copy followup_overrides FROM 'overrides.csv' CSV;"
   ```
   (psql is included with PostgreSQL client tools; you can also upload via
   the Supabase SQL editor.)

Once configured, every upload or save will read/write the shared Supabase
table. Your teammates and other deployments will immediately see updates – no
more lost edits when the page reloads.

#### Generic PostgreSQL setup

If you prefer another provider, the same steps apply: obtain a
`postgresql://` URL, set `DATABASE_URL`, and ensure `psycopg2-binary` is
installed. The app automatically detects the presence of the variable and
switches backends.

#### Migration

If you already have overrides stored locally, you can migrate them by dumping
the SQLite table and importing into PostgreSQL. A simple command-line
sequence might look like:

```bash
# export SQLite rows as CSV
sqlite3 orders.db "SELECT * FROM followup_overrides;" > overrides.csv
# import into PostgreSQL (adjust flags as needed)
psql "$DATABASE_URL" -c "\copy followup_overrides FROM 'overrides.csv' CSV;"
```

Once the cloud database is in use, multiple users and deployed app instances
will see each other's changes. Uploading a new file still applies overrides by
matching `order_key` so existing rows aren’t overwritten but updated.

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

### Automated Email Data Retrieval

The app includes an automation script that can automatically fetch XLS attachments from your daily emails and update the data directory.

**Two automation options:**

1. **Local Windows Task Scheduler** (runs only when your computer is on)
2. **GitHub Actions** (runs in the cloud, even when your computer is off) ⭐ *Recommended*

---

### Option 1: Local Windows Task Scheduler (Computer must be on)

1. **Install additional dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

2. **Configure email settings in `.env`:**
   ```env
   # Email server settings (Gmail example - adjust for your email provider)
   IMAP_SERVER=imap.gmail.com
   IMAP_PORT=993
   EMAIL_USER=your-email@gmail.com
   EMAIL_PASSWORD=your-app-password  # Use app password, not regular password
   SMTP_SERVER=smtp.gmail.com
   SMTP_PORT=587

   # Search criteria (configured for your specific emails)
   SEARCH_SUBJECT=Sensi Medical Sales Open Order
   SEARCH_SENDER=customercare@optimalmax.com
   TARGET_FILENAME=searchresults.xlsx  # This will be saved as searchresults.xlsx

   # Notifications
   NOTIFY_EMAIL=your-email@gmail.com  # Email to send success/failure notifications

   # Git settings
   GIT_REMOTE=origin
   GIT_BRANCH=main
   ```

3. **For Gmail, enable 2-factor authentication and create an App Password:**
   - Go to Google Account settings
   - Security → 2-Step Verification → App passwords
   - Generate a password for "Mail"
   - Use this app password in `EMAIL_PASSWORD`

#### Manual Testing

Test the automation script:

```powershell
python email_automation.py
```

Check `email_automation.log` for results.

#### Daily Automation (Windows Task Scheduler)

1. **Create a batch file to run the script:**
   ```batch
   @echo off
   cd "C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule"
   call .venv\Scripts\activate.bat
   python email_automation.py
   ```

2. **Save as `run_automation.bat` in the project root**

3. **Schedule with Windows Task Scheduler:**
   - Search for "Task Scheduler" in Windows
   - Create Basic Task
   - Name: "Shipment Schedule Email Automation"
   - Trigger: Daily at 9:00 AM (or your preferred time)
   - Action: Start a program
   - Program/script: `C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule\run_automation.bat`
   - Start in: `C:\Users\Usuario\OneDrive\Projects\BIAutomations\shipment-schedule`

---

### Option 2: GitHub Actions (Cloud-based - runs even when your computer is off) ⭐

This is the recommended option since it runs automatically every day, including weekends, regardless of whether your computer is on.

#### Setup

1. **Push your code to GitHub:**
   ```bash
   git add .
   git commit -m "Add email automation"
   git push origin main
   ```

2. **Configure GitHub Secrets:**
   Go to your GitHub repository → Settings → Secrets and variables → Actions
   
   Add these secrets:
   ```
   EMAIL_USER          → your-email@gmail.com
   EMAIL_PASSWORD      → your-gmail-app-password (16-character)
   IMAP_SERVER         → imap.gmail.com
   IMAP_PORT           → 993
   SMTP_SERVER         → smtp.gmail.com
   SMTP_PORT           → 587
   SEARCH_SUBJECT      → Sensi Medical Sales Open Order
   SEARCH_SENDER       → customercare@optimalmax.com
   TARGET_FILENAME     → searchresults.xlsx
   NOTIFY_EMAIL        → your-email@gmail.com
   ```

3. **The workflow will automatically:**
   - Run daily at 9:00 AM UTC (adjustable in `.github/workflows/email-automation.yml`)
   - Download new email attachments
   - Update your repository with new data
   - Send you notification emails

4. **To get updates locally:**
   When you turn on your computer, pull the latest changes:
   ```bash
   git pull origin main
   ```

#### Adjusting the Schedule

To change when the automation runs, edit `.github/workflows/email-automation.yml`:

```yaml
schedule:
  # Run daily at 9:00 AM UTC
  - cron: '0 9 * * *'
  
  # Examples:
  # 2:00 PM UTC: '0 14 * * *'
  # 8:00 AM EST (UTC-5): '0 13 * * *'
  # Weekdays only: '0 9 * * 1-5'
```

#### Manual Testing

You can trigger the workflow manually:
- Go to your GitHub repository → Actions → "Daily Email Automation"
- Click "Run workflow"

---

#### How it works

1. Connects to your email via IMAP
2. Searches for emails from the last day matching your criteria
3. Downloads XLS/XLSX attachments from the most recent matching email
4. Validates the Excel file
5. Replaces the file in `data/`
6. Commits and pushes changes to git
7. Sends notification email with results

#### Live Console Refresh

Since this is a local Streamlit app, you'll need to:

1. **Pull latest changes** when the automation runs (if running from git)
2. **Restart Streamlit** or implement auto-reload

For auto-reload, you can modify the app to watch for file changes:

```python
# Add to app.py imports
import time
from streamlit.runtime.scriptrunner import add_script_run_ctx

# Add this function
def watch_for_changes():
    last_modified = os.path.getmtime(get_latest_file()) if get_latest_file() else 0
    while True:
        time.sleep(60)  # Check every minute
        current_modified = os.path.getmtime(get_latest_file()) if get_latest_file() else 0
        if current_modified > last_modified:
            st.rerun()
            last_modified = current_modified

# Call in main()
if __name__ == "__main__":
    # Start file watcher in background thread
    import threading
    watcher_thread = threading.Thread(target=watch_for_changes, daemon=True)
    watcher_thread.start()
    main()
```

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

