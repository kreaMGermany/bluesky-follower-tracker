import os
import json
import requests
from datetime import datetime, date, timedelta
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"


def getenv(name):
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def graph_get(token, url):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()


def graph_post(token, url, payload):
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()
    return r.json() if r.text else {}


def get_token():
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        raise RuntimeError(f"Failed to get Graph token: {token}")
    return token["access_token"]


def get_followers(handle):
    url = f"https://public.api.bsky.app/xrpc/app.bsky.actor.getProfile?actor={handle}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    data = r.json()
    followers = data.get("followersCount")
    if followers is None:
        raise RuntimeError(f"followersCount missing for {handle}")
    return int(followers)


def send_mail(token, sender, recipients, subject, html):
    to_recipients = [
        {"emailAddress": {"address": r.strip()}}
        for r in recipients.split(",")
        if r.strip()
    ]

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html
            },
            "toRecipients": to_recipients
        },
        "saveToSentItems": True
    }

    graph_post(token, f"{GRAPH}/users/{sender}/sendMail", payload)


def parse_excel_date(value):
    if value is None:
        return None

    if isinstance(value, (int, float)):
        base = date(1899, 12, 30)
        return base + timedelta(days=int(float(value)))

    s = str(value).strip()
    if not s:
        return None

    if "T" in s:
        s = s[:10]

    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    return None


def fmt_int(n):
    if n is None:
        return "n/a"
    return f"{int(n):,}".replace(",", ".")


def fmt_delta(n):
    if n is None:
        return "n/a"
    sign = "+" if n > 0 else ""
    return f"{sign}{int(n):,}".replace(",", ".")


def fmt_pct(p):
    if p is None:
        return "n/a"
    sign = "+" if p > 0 else ""
    return f"{sign}{p:.2f}%"


def build_html(today, results):
    valid = [r for r in results if r["delta"] is not None]

    top_growth = sorted(
        [r for r in valid if r["delta"] > 0],
        key=lambda x: x["delta"],
        reverse=True
    )[:3]

    biggest_drop = sorted(
        [r for r in valid if r["delta"] < 0],
        key=lambda x: x["delta"]
    )[:3]

    def bullet(r):
        return f"<li><b>{r['account']}</b> ({fmt_int(r['followers'])}, {fmt_delta(r['delta'])}, {fmt_pct(r['delta_pct'])})</li>"

    rows = ""
    for r in results:
        delta_color = "#0a7f2e" if r["delta"] is not None and r["delta"] > 0 else "#b00020" if r["delta"] is not None and r["delta"] < 0 else "#111111"
        pct_color = "#0a7f2e" if r["delta_pct"] is not None and r["delta_pct"] > 0 else "#b00020" if r["delta_pct"] is not None and r["delta_pct"] < 0 else "#111111"

        base_date = r["base_date"].strftime("%d.%m.%Y") if r["base_date"] else "n/a"

        rows += f"""
        <tr>
          <td style="padding:10px 8px;border-bottom:1px solid #eaeaea;">
            <div style="font-weight:600;">{r['account']}</div>
            <div style="font-size:12px;color:#6b6b6b;">Basis: {base_date}</div>
          </td>
          <td style="padding:10px 8px;border-bottom:1px solid #eaeaea;text-align:right;">{fmt_int(r['followers'])}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #eaeaea;text-align:right;color:{delta_color};font-weight:700;">{fmt_delta(r['delta'])}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #eaeaea;text-align:right;color:{pct_color};font-weight:700;">{fmt_pct(r['delta_pct'])}</td>
        </tr>
        """

    html = f"""
    <html>
      <body style="font-family:Segoe UI, Arial, sans-serif; color:#111;">
        <h2 style="margin:0 0 8px 0;">Bluesky Report {today.strftime("%Y-%m-%d")}</h2>
        <div style="margin:0 0 18px 0;color:#444;">
          Delta = Vergleich zum letzten verfügbaren Eintrag vor heute
        </div>

        <h3 style="margin:18px 0 6px 0;">Top Growth</h3>
        <ul>{''.join([bullet(r) for r in top_growth]) if top_growth else "<li>n/a</li>"}</ul>

        <h3 style="margin:18px 0 6px 0;">Biggest Drop</h3>
        <ul>{''.join([bullet(r) for r in biggest_drop]) if biggest_drop else "<li>n/a</li>"}</ul>

        <h3 style="margin:18px 0 6px 0;">Alle Accounts</h3>
        <table style="border-collapse:collapse; width:780px; max-width:100%;">
          <thead>
            <tr>
              <th style="text-align:left;padding:10px 8px;border-bottom:2px solid #333;">Account</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Followers</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Δ</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Δ %</th>
            </tr>
          </thead>
          <tbody>
            {rows}
          </tbody>
        </table>
      </body>
    </html>
    """
    return html


def main():
    token = get_token()
    sender = getenv("SENDER_UPN")
    recipients = getenv("RECIPIENTS")
    path = getenv("ONEDRIVE_FILE_PATH")
    accounts = json.loads(getenv("ACCOUNTS_JSON"))

    today = datetime.utcnow().date()

    file_meta = graph_get(token, f"{GRAPH}/users/{sender}/drive/root:/{path}")
    file_id = file_meta["id"]

    tables = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables")
    if not tables.get("value"):
        raise RuntimeError("No Excel table found in workbook")
    table_id = tables["value"][0]["id"]

    rows_data = graph_get(
        token,
        f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables/{table_id}/rows?$top=5000"
    )

    history = {}

    for r in rows_data.get("value", []):
        vals = r.get("values", [[]])[0]

        if len(vals) < 3:
            continue

        d = parse_excel_date(vals[0])
        acc = str(vals[1]).replace("'", "").strip()
        foll_raw = str(vals[2]).strip()

        if not d or not acc or not foll_raw:
            continue

        try:
            foll = int(float(foll_raw))
        except Exception:
            continue

        history.setdefault(acc, []).append((d, foll))

    for acc in history:
        history[acc] = sorted(history[acc], key=lambda x: x[0])

    results = []

    for acc in accounts:
        handle = acc["handle"]
        followers = get_followers(handle)

        prev = [x for x in history.get(handle, []) if x[0] < today]
        if prev:
            base_date, base_followers = prev[-1]
            delta = followers - base_followers
            delta_pct = (delta / base_followers * 100) if base_followers > 0 else None
        else:
            base_date, base_followers = None, None
            delta, delta_pct = None, None

        results.append({
            "account": handle,
            "followers": followers,
            "delta": delta,
            "delta_pct": delta_pct,
            "base_date": base_date
        })

    html = build_html(today, results)

    send_mail(
        token,
        sender,
        recipients,
        f"Bluesky Report {today.strftime('%Y-%m-%d')}",
        html
    )

    values = [[today.isoformat(), "'" + r["account"], r["followers"]] for r in results]

    graph_post(
        token,
        f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables/{table_id}/rows/add",
        {"values": values}
    )

    print("DONE")


if __name__ == "__main__":
    main()
