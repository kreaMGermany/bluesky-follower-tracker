import os
import json
import requests
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from typing import Any, Dict, List, Optional, Tuple
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"
BERLIN = ZoneInfo("Europe/Berlin")

ALERT_DROP_THRESHOLD = -100


def getenv(name: str) -> str:
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Failed to get Graph token: {res}")
    return res["access_token"]


def graph_get(token: str, url: str) -> Dict[str, Any]:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()


def graph_post(token: str, url: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps(payload),
        timeout=60,
    )
    r.raise_for_status()
    return r.json() if r.text else {}


def parse_excel_date(x: Any) -> Optional[date]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        base = date(1899, 12, 30)
        return base + timedelta(days=int(float(x)))

    s = str(x).strip()
    if not s:
        return None
    if "T" in s:
        s = s[:10]

    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except:
        try:
            return datetime.strptime(s[:10], "%d.%m.%Y").date()
        except:
            return None


def fmt_ddmmyyyy(d: Optional[date]) -> str:
    return d.strftime("%d.%m.%Y") if d else "n/a"


def fmt_int(n: Optional[int]) -> str:
    if n is None:
        return "n/a"
    return f"{n:,}".replace(",", ".")


def fmt_delta(n: Optional[int]) -> str:
    if n is None:
        return "n/a"
    sign = "+" if n > 0 else ""
    return f"{sign}{n:,}".replace(",", ".")


def fmt_pct(p: Optional[float]) -> str:
    if p is None:
        return "n/a"
    sign = "+" if p > 0 else ""
    return f"{sign}{p:.2f}%"


def excel_safe_account(handle: str) -> str:
    return "'" + handle


def normalize_account_from_excel(x: Any) -> str:
    return str(x).lstrip("'").strip()


def build_mail_html(title: str, today: date, results: List[Dict[str, Any]]) -> str:
    rows = ""
    for r in results:
        rows += f"""
        <tr>
          <td>{r['account']}</td>
          <td>{fmt_int(r.get('followers'))}</td>
          <td>{fmt_delta(r.get('delta'))}</td>
          <td>{fmt_pct(r.get('delta_pct'))}</td>
        </tr>
        """

    return f"""
    <html>
      <body>
        <h2>{title} {fmt_ddmmyyyy(today)}</h2>
        <table border="1" cellpadding="6" cellspacing="0">
          <tr>
            <th>Account</th>
            <th>Followers</th>
            <th>Δ</th>
            <th>Δ %</th>
          </tr>
          {rows}
        </table>
      </body>
    </html>
    """


def send_mail(token: str, sender_upn: str, recipients: str, subject: str, html: str):
    to_recipients = [{"emailAddress": {"address": r.strip()}} for r in recipients.split(",")]
    url = f"{GRAPH}/users/{sender_upn}/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients,
        }
    }
    graph_post(token, url, payload)


def fetch_bluesky_followers(handle: str):
    url = f"https://public.api.bsky.app/xrpc/app.bsky.actor.getProfile?actor={handle}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    data = r.json()
    return data.get("handle"), int(data.get("followersCount", 0))


def main():
    tenant_id = getenv("TENANT_ID")
    client_id = getenv("CLIENT_ID")
    client_secret = getenv("CLIENT_SECRET")
    sender_upn = getenv("SENDER_UPN")
    recipients = getenv("RECIPIENTS")
    onedrive_path = getenv("ONEDRIVE_FILE_PATH")
    accounts = json.loads(getenv("ACCOUNTS_JSON"))

    today = datetime.now(BERLIN).date()
    token = graph_token(tenant_id, client_id, client_secret)

    file_meta = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/root:/{onedrive_path}")
    file_id = file_meta["id"]

    table = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables")["value"][0]["id"]

    rows_data = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables/{table}/rows?$top=5000")

    history = {}
    for r in rows_data["value"]:
        vals = r["values"][0]
        d = parse_excel_date(vals[0])
        acc = normalize_account_from_excel(vals[1])
        foll = int(vals[2])
        history.setdefault(acc, []).append((d, foll))

    results = []

    for acc in accounts:
        handle = acc["handle"]
        _, followers = fetch_bluesky_followers(handle)

        prev = [(d, f) for (d, f) in history.get(handle, []) if d < today]
        base_date, base_followers = (prev[-1] if prev else (None, None))

        delta = followers - base_followers if base_followers else None
        delta_pct = (delta / base_followers * 100) if base_followers else None

        results.append({
            "account": handle,
            "followers": followers,
            "delta": delta,
            "delta_pct": delta_pct,
            "base_date": base_date
        })

    html = build_mail_html("Bluesky Report", today, results)
    send_mail(token, sender_upn, recipients, f"Bluesky Report {fmt_ddmmyyyy(today)}", html)

    values = [[today.isoformat(), excel_safe_account(r["account"]), r["followers"]] for r in results]

    graph_post(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables/{table}/rows/add", {"values": values})


if __name__ == "__main__":
    main()
