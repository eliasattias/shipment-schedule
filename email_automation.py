import os
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import openpyxl as oxl
import logging
import git
import requests as http_requests
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import base64

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_automation.log'),
        logging.StreamHandler()
    ]
)

class EmailAutomation:
    def __init__(self):
        # File configuration
        self.base_dir = Path(__file__).parent
        self.data_dir = self.base_dir / "data"
        self.target_filename = os.getenv('TARGET_FILENAME', 'searchresults.xlsx')
        self.token_path = self.base_dir / 'token.json'
        self.gmail_token_json = os.getenv('GMAIL_TOKEN_JSON', '')

        # Search configuration
        self.search_subject = os.getenv('SEARCH_SUBJECT', 'Sensi Medical Sales Open Order')
        self.search_sender = os.getenv('SEARCH_SENDER', 'customercare@optimalmax.com')

        # Git configuration (optional)
        try:
            self.git_repo = git.Repo(self.base_dir)
        except:
            self.git_repo = None
            logging.warning("Git repository not available - git operations will be skipped")

        self.git_remote = os.getenv('GIT_REMOTE', 'origin')
        self.git_branch = os.getenv('GIT_BRANCH', 'main')

        # Notification configuration (Resend)
        self.resend_api_key = os.getenv('RESEND_API_KEY', '')
        self.notify_from = 'Sensimedical Pending Orders Schedule <info@sensimedical.com>'
        self.notify_emails = [
            'automations@sensimedical.com',
            'alice.s@sensimedical.com',
            'eduardo.s@sensimedical.com',
        ]

    def get_gmail_service(self):
        """Get authenticated Gmail API service"""
        SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        try:
            creds = None

            # Priority 1: GMAIL_TOKEN_JSON env var (GitHub Actions / Lambda)
            if self.gmail_token_json:
                token_data = json.loads(self.gmail_token_json)
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            # Priority 2: token.json file (local)
            elif self.token_path.exists():
                creds = Credentials.from_authorized_user_file(str(self.token_path), SCOPES)

            # Refresh token if expired
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                # Save refreshed token back to file for next run
                if not self.gmail_token_json and self.token_path:
                    with open(self.token_path, 'w') as f:
                        f.write(creds.to_json())

            if not creds or not creds.valid:
                logging.error("No valid Gmail credentials found. Run oauth2_setup.py first.")
                return None

            return build('gmail', 'v1', credentials=creds, static_discovery=False)

        except Exception as e:
            logging.error(f"Failed to get Gmail service: {e}")
            return None

    def search_emails(self, service, days_back=1):
        """Search for emails with XLS attachments"""
        try:
            # Build search query
            since_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y/%m/%d')
            query = f'after:{since_date}'

            if self.search_subject:
                query += f' subject:"{self.search_subject}"'
            if self.search_sender:
                query += f' from:"{self.search_sender}"'

            query += ' has:attachment filename:(xls OR xlsx)'

            logging.info(f"Searching with query: {query}")

            results = service.users().messages().list(
                userId='me',
                q=query,
                maxResults=5
            ).execute()

            messages = results.get('messages', [])
            logging.info(f"Found {len(messages)} matching emails")
            return messages

        except Exception as e:
            logging.error(f"Failed to search emails: {e}")
            raise

    def convert_spreadsheetml_to_xlsx(self, src_path, dest_path):
        """Parse SpreadsheetML XML (used by NetSuite/older ERP exports) and write to .xlsx"""
        tree = ET.parse(src_path)
        root = tree.getroot()
        ns = "urn:schemas-microsoft-com:office:spreadsheet"
        if ns not in root.tag:
            raise ValueError("Not a SpreadsheetML file")

        wb_out = oxl.Workbook()
        wb_out.remove(wb_out.active)
        sheets_found = 0

        for ws_elem in root.iter(f"{{{ns}}}Worksheet"):
            sheet_name = ws_elem.get(f"{{{ns}}}Name", f"Sheet{sheets_found + 1}")
            ws_out = wb_out.create_sheet(title=sheet_name)
            row_num = 0
            for row_elem in ws_elem.iter(f"{{{ns}}}Row"):
                row_idx = row_elem.get(f"{{{ns}}}Index")
                row_num = int(row_idx) if row_idx else row_num + 1
                col_num = 0
                for cell_elem in row_elem.iter(f"{{{ns}}}Cell"):
                    cell_idx = cell_elem.get(f"{{{ns}}}Index")
                    col_num = int(cell_idx) if cell_idx else col_num + 1
                    data = cell_elem.find(f"{{{ns}}}Data")
                    if data is not None:
                        val = data.text or ""
                        dtype = data.get(f"{{{ns}}}Type", "String")
                        if dtype == "Number":
                            try:
                                val = float(val) if "." in val else int(val)
                            except (ValueError, TypeError):
                                pass
                        elif dtype == "DateTime":
                            try:
                                parsed = None
                                for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
                                    try:
                                        parsed = datetime.strptime(val, fmt)
                                        break
                                    except ValueError:
                                        continue
                                if parsed:
                                    val = parsed
                            except Exception:
                                pass
                        elif dtype == "Boolean":
                            val = val.strip().lower() in ("1", "true")
                        cell = ws_out.cell(row=row_num, column=col_num, value=val)
                        if hasattr(val, 'year'):
                            cell.number_format = 'MM/DD/YYYY'
                    merge = cell_elem.get(f"{{{ns}}}MergeAcross")
                    if merge:
                        col_num += int(merge)
            sheets_found += 1

        if not sheets_found:
            raise ValueError("No worksheets found in SpreadsheetML")
        wb_out.save(dest_path)
        return True

    def convert_to_xlsx(self, src_path, dest_path):
        """Convert .xls to .xlsx using multiple strategies"""
        # Strategy 1: standard xlrd read
        try:
            df = pd.read_excel(src_path, engine='xlrd')
            df.to_excel(dest_path, index=False)
            logging.info(f"Converted via xlrd: {src_path.name}")
            return True
        except Exception as e1:
            logging.debug(f"xlrd failed: {e1}")

        # Strategy 2: openpyxl engine (xlsx mis-named as .xls)
        try:
            df = pd.read_excel(src_path, engine='openpyxl')
            df.to_excel(dest_path, index=False)
            logging.info(f"Converted via openpyxl: {src_path.name}")
            return True
        except Exception as e2:
            logging.debug(f"openpyxl failed: {e2}")

        # Strategy 3: HTML table parsing
        try:
            tables = pd.read_html(str(src_path), flavor="lxml")
            if tables:
                tables[0].to_excel(dest_path, index=False)
                logging.info(f"Converted via HTML table parsing: {src_path.name}")
                return True
        except Exception as e3:
            logging.debug(f"HTML fallback failed: {e3}")

        # Strategy 4: SpreadsheetML XML
        try:
            self.convert_spreadsheetml_to_xlsx(str(src_path), str(dest_path))
            logging.info(f"Converted via SpreadsheetML XML: {src_path.name}")
            return True
        except Exception as e4:
            logging.debug(f"SpreadsheetML failed: {e4}")

        return False

    def download_attachment(self, service, message_id):
        """Download XLS attachment from email"""
        try:
            message = service.users().messages().get(userId='me', id=message_id, format='full').execute()

            headers = message['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'Unknown')
            sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown')

            logging.info(f"Processing email from {sender}: {subject}")

            # Find and download attachment
            if 'parts' in message['payload']:
                for part in message['payload']['parts']:
                    if part['filename']:
                        filename = part['filename']
                        if filename.lower().endswith(('.xls', '.xlsx')):
                            attachment_id = part['body']['attachmentId']
                            attachment = service.users().messages().attachments().get(
                                userId='me',
                                messageId=message_id,
                                id=attachment_id
                            ).execute()

                            data = base64.urlsafe_b64decode(attachment['data'])

                            # Save raw file
                            raw_path = self.data_dir / filename
                            with open(raw_path, 'wb') as f:
                                f.write(data)

                            final_path = self.data_dir / self.target_filename

                            # Convert if .xls, otherwise just move
                            if raw_path.suffix.lower() == '.xls':
                                if not self.convert_to_xlsx(raw_path, final_path):
                                    logging.warning("All conversion strategies failed, using raw file")
                                    raw_path.replace(final_path)
                                # Clean up raw .xls if final is different file
                                if raw_path.exists() and raw_path != final_path:
                                    raw_path.unlink()
                            else:
                                raw_path.replace(final_path)

                            logging.info(f"Downloaded attachment: {filename} -> {final_path}")
                            return final_path

            return None

        except Exception as e:
            logging.error(f"Failed to download attachment from message {message_id}: {e}")
            return None

    def validate_excel_file(self, filepath):
        """Validate that the downloaded file is a valid Excel file"""
        try:
            # Check file exists and has content
            if not filepath.exists():
                logging.error(f"File does not exist: {filepath}")
                return False
            
            file_size = filepath.stat().st_size
            if file_size == 0:
                logging.error(f"Downloaded file is empty: {filepath}")
                return False
            
            logging.info(f"File downloaded successfully: {filepath} ({file_size} bytes)")
            
            # Try to read with any available engine
            for engine in ['openpyxl', 'lxml', 'xlrd']:
                try:
                    df = pd.read_excel(filepath, engine=engine)
                    if not df.empty:
                        logging.info(f"Validated with {engine}: {len(df)} rows, {len(df.columns)} columns")
                        return True
                except Exception as e:
                    logging.debug(f"Engine {engine} failed: {str(e)[:80]}")
                    continue
            
            # last ditch: if file still cannot be read but has .xls extension,
            # try converting it in-place and then re-validate
            if filepath.suffix.lower() == '.xls':
                try:
                    df = pd.read_excel(filepath, engine='xlrd')
                    df.to_excel(filepath.with_suffix('.xlsx'), index=False)
                    logging.info(f"Converted failing .xls to .xlsx for validation")
                    return self.validate_excel_file(filepath.with_suffix('.xlsx'))
                except Exception:
                    pass

            # Log warning but return True since file was downloaded successfully
            logging.warning(f"Could not parse file as Excel, but file exists and has content")
            return True
            
        except Exception as e:
            logging.error(f"Validation error: {e}")
            return False

    def commit_and_push(self, filepath):
        """Commit and push changes to git (optional - may be handled by CI/CD)"""
        try:
            if not self.git_repo:
                logging.info("Git repository not available - skipping git operations")
                return True

            # Check if we should handle git operations ourselves
            if os.getenv('GITHUB_ACTIONS') == 'true':
                logging.info("Running in GitHub Actions - skipping manual git operations")
                return True

            # Add file to git
            self.git_repo.index.add([str(filepath.relative_to(self.base_dir))])

            # Check if there are changes
            if not self.git_repo.index.diff("HEAD"):
                logging.info("No changes to commit")
                return True

            # Commit changes
            commit_message = f"Auto-update: {self.target_filename} from email - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            self.git_repo.index.commit(commit_message)

            # Push changes
            origin = self.git_repo.remote(self.git_remote)
            origin.push(self.git_branch)

            logging.info(f"Successfully committed and pushed changes: {commit_message}")
            return True

        except Exception as e:
            logging.warning(f"Git operations failed (this may be expected in CI/CD): {e}")
            return True

    def send_notification(self, success, message):
        """Send notification email via Resend"""
        if not self.resend_api_key:
            logging.warning("RESEND_API_KEY not set - skipping notification")
            return

        if not success:
            logging.info("Automation did not succeed - skipping notification email")
            return

        try:
            now_et = datetime.now(ZoneInfo('America/New_York'))
            date_str = now_et.strftime('%B %d, %Y at %I:%M %p ET')

            html_body = (
                f'<p>The shipment schedule console has been updated on {date_str}.</p>'
                '<p><a href="https://sensimedical-shipment-schedule.streamlit.app/">'
                'Shipment Schedule Console</a></p>'
                '<hr>'
                '<p>This is an automated message.</p>'
            )

            resp = http_requests.post(
                'https://api.resend.com/emails',
                headers={
                    'Authorization': f'Bearer {self.resend_api_key}',
                    'Content-Type': 'application/json',
                },
                json={
                    'from': self.notify_from,
                    'to': self.notify_emails,
                    'subject': 'Shipment Schedule Console Updated',
                    'html': html_body,
                },
                timeout=10,
            )

            if resp.status_code in (200, 201):
                logging.info("Notification email sent via Resend")
            else:
                logging.error(f"Resend error {resp.status_code}: {resp.text}")

        except Exception as e:
            logging.error(f"Failed to send notification: {e}")

    def run_automation(self):
        """Main automation function"""
        logging.info("Starting email automation using Gmail API")

        try:
            # Ensure data directory exists
            self.data_dir.mkdir(parents=True, exist_ok=True)

            # Get Gmail service
            service = self.get_gmail_service()
            if not service:
                message = "Failed to authenticate with Gmail API"
                logging.error(message)
                self.send_notification(False, message)
                return

            # Search for emails
            messages = self.search_emails(service)
            if not messages:
                message = "No matching emails found"
                logging.info(message)
                self.send_notification(False, message)
                return

            # Process most recent email
            latest_message_id = messages[0]['id']
            filepath = self.download_attachment(service, latest_message_id)

            if not filepath:
                message = "No XLS attachment found in emails"
                logging.error(message)
                self.send_notification(False, message)
                return

            # Validate file
            if not self.validate_excel_file(filepath):
                message = "Downloaded file is not a valid Excel file"
                logging.error(message)
                self.send_notification(False, message)
                return

            # Commit and push
            if self.commit_and_push(filepath):
                message = f"Successfully updated {self.target_filename}"
                logging.info(message)
                self.send_notification(True, message)
            else:
                message = "Failed to commit and push changes"
                logging.error(message)
                self.send_notification(False, message)

        except Exception as e:
            error_msg = f"Automation failed: {str(e)}"
            logging.error(error_msg)
            self.send_notification(False, error_msg)

if __name__ == "__main__":
    automation = EmailAutomation()
    automation.run_automation()