#!/usr/bin/env python3
"""
OAuth2 Setup for Google Workspace Email Access
Alternative to app passwords for Google Workspace accounts
"""
import os
import json
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def setup_oauth2():
    """Set up OAuth2 credentials for Gmail API access"""

    creds = None
    token_path = Path(__file__).parent / 'token.json'
    credentials_path = Path(__file__).parent / 'credentials.json'

    # Check if token already exists
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    # If credentials are invalid or don't exist, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path.exists():
                print("credentials.json not found!")
                print("You need to:")
                print("1. Go to Google Cloud Console: https://console.cloud.google.com/")
                print("2. Create a new project or select existing")
                print("3. Enable Gmail API")
                print("4. Create OAuth2 credentials (Desktop application)")
                print("5. Download credentials.json and place in project root")
                return False

            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials as JSON
        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    print("OAuth2 setup complete!")
    print(f"Token saved to: {token_path}")
    return True

def test_gmail_api():
    """Test Gmail API access"""
    try:
        creds = None
        token_path = Path(__file__).parent / 'token.json'

        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

        service = build('gmail', 'v1', credentials=creds, static_discovery=False)

        # Test by getting user profile
        profile = service.users().getProfile(userId='me').execute()
        print(f"✅ Gmail API connected! Email: {profile['emailAddress']}")

        # Test searching for messages
        query = 'subject:"Sensi Medical Sales Open Order"'
        results = service.users().messages().list(userId='me', q=query, maxResults=5).execute()

        messages = results.get('messages', [])
        print(f"✅ Found {len(messages)} matching emails")

        return True

    except Exception as e:
        print(f"❌ Gmail API test failed: {e}")
        return False

if __name__ == "__main__":
    print("🔐 Gmail OAuth2 Setup")
    print("=" * 30)

    if setup_oauth2():
        print("\n🧪 Testing connection...")
        test_gmail_api()