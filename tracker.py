import os
import json
import requests
from datetime import datetime
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"


def getenv(name):
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def get_token():
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token["access_token"]


def get_followers(handle):
    url = f"https://public.api.bsky.app/xrpc/app.bsky.actor.getProfile?actor={handle}"
    r = requests.get(url)
    r.raise_for_status()
    return r.json()["followersCount"]


def send_mail(token, sender, recipients, subject, html):
    to_recipients = [{"emailAddress": {"address": r.strip()}} for r in recipients.split(",")]

    url = f"{GRAPH}/users/{sender}/sendMail"

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html
            },
            "toRecipients": to_recipients
        }
    }

    requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload
    ).raise_for_status()


def main():
    accounts = json.loads(getenv("ACCOUNTS_JSON"))
    token = get_token()

    today = datetime.utcnow().strftime("%Y-%m-%d")

    rows = []
    html = "<h2>Bluesky Report</h2><table border='1'><tr><th>Account</th><th>Followers</th></tr>"

    for acc in accounts:
        handle = acc["handle"]
        followers = get_followers(handle)

        rows.append((handle, followers))
        html += f"<tr><td>{handle}</td><td>{followers}</td></tr>"

    html += "</table>"

    send_mail(
        token,
        getenv("SENDER_UPN"),
        getenv("RECIPIENTS"),
        f"Bluesky Report {today}",
        html
    )

    print(rows)


if __name__ == "__main__":
    main()
