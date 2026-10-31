import os
import smtplib
from datetime import datetime, timezone
from email.message import EmailMessage
from zoneinfo import ZoneInfo

import requests

MADAGASCAR_TZ = ZoneInfo("Indian/Antananarivo")


def required_env(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def main():
    supabase_url = required_env("SUPABASE_URL").rstrip("/")
    service_key = required_env("SUPABASE_SERVICE_ROLE_KEY")
    gmail_address = required_env("GMAIL_ADDRESS")
    gmail_password = required_env("GMAIL_APP_PASSWORD")
    notify_to = required_env("NOTIFY_TO")

    now_local = datetime.now(MADAGASCAR_TZ)
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    start_utc = start_local.astimezone(timezone.utc).isoformat()

    response = requests.get(
        f"{supabase_url}/rest/v1/app_events",
        # Supabase API keys (especially the new sb_secret_* keys) belong
        # in the apikey header. They are not JWTs, so do not send them as Bearer.
        headers={
            "apikey": service_key,
        },
        params={
            "select": "event_name,app_version,created_at",
            "created_at": f"gte.{start_utc}",
            "order": "created_at.desc",
            "limit": "5000",
        },
        timeout=30,
    )
    if not response.ok:
        # Show Supabase's diagnostic response without exposing request headers
        # or any secret key in the GitHub Actions log.
        try:
            detail = response.json()
        except ValueError:
            detail = response.text[:1500]
        raise RuntimeError(
            f"Supabase API error ({response.status_code}): {detail}"
        )

    events = response.json()

    counts = {}
    for event in events:
        name = event.get("event_name") or "unknown"
        counts[name] = counts.get(name, 0) + 1

    lines = [
        "CORRECTGLOSS — Rapport quotidien",
        f"Date : {now_local.strftime('%d/%m/%Y')}",
        "",
        f"Nombre total d'événements enregistrés : {len(events)}",
        "",
        "Détail par événement :",
    ]
    if counts:
        for name, count in sorted(counts.items()):
            lines.append(f"- {name} : {count}")
    else:
        lines.append("- Aucun événement enregistré aujourd'hui.")

    lines += [
        "",
        "Remarque : ce rapport compte les événements présents dans Supabase.",
        "Il ne représente pas nécessairement le nombre d'utilisateurs uniques.",
    ]

    message = EmailMessage()
    message["Subject"] = f"CORRECTGLOSS — Rapport du {now_local.strftime('%d/%m/%Y')}"
    message["From"] = gmail_address
    message["To"] = notify_to
    message.set_content("\n".join(lines))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(gmail_address, gmail_password)
        smtp.send_message(message)

    print(f"Daily digest sent to {notify_to}; {len(events)} events included.")


if __name__ == "__main__":
    main()
