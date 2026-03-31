"""
AWS Lambda entry point for shipment schedule email automation.
EventBridge invokes this daily at 8 AM ET.
Set config via Lambda env vars; Gmail token from Secrets Manager (GMAIL_TOKEN_SECRET_ARN).
"""

import json
import logging
import os

import boto3

logging.getLogger().setLevel(logging.INFO)


def _load_gmail_token_from_secrets():
    arn = os.getenv("GMAIL_TOKEN_SECRET_ARN")
    if not arn:
        return
    client = boto3.client("secretsmanager")
    resp = client.get_secret_value(SecretId=arn)
    secret = json.loads(resp["SecretString"])
    if "GMAIL_TOKEN_JSON" in secret:
        os.environ["GMAIL_TOKEN_JSON"] = secret["GMAIL_TOKEN_JSON"]
    else:
        os.environ["GMAIL_TOKEN_JSON"] = json.dumps(secret)


def handler(event, context):
    _load_gmail_token_from_secrets()
    from email_automation import EmailAutomation
    automation = EmailAutomation()
    automation.run_automation()
    return {"statusCode": 200, "body": "OK"}
