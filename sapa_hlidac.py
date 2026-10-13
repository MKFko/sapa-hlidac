"""Hlídač volných termínů Sapa Trip (rezervační systém Reenio).

Kontroluje listopad 2026 a zvlášť 17. 11. Upozorní, když se objeví nový
český Info Trip s obědem / s brzkou večeří s volným místem. Kanály: Windows notifikace (vždy), ntfy.sh push na
mobil a e-mail (volitelně, nastavení v config.json).

Spuštění:  python sapa_hlidac.py          (jedna kontrola)
           python sapa_hlidac.py --test   (pošle zkušební upozornění)
"""
import http.cookiejar
import json
import os
import smtplib
import subprocess
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG_FILE = HERE / "config.json"
STATE_FILE = HERE / "state.json"
LOG_FILE = HERE / "log.txt"

BASE = "https://sapatrip.reenio.cz/cs"
BOOKING_URL = "https://www.sapatrip.cz/rezervace/"

DEFAULT_CONFIG = {
    "year": 2026,
    "month": 11,
    "target_day": 17,
    # Hlídané služby: 19944 = Info Trip s obědem, 27567 = Info trip s brzkou večeří (české)
    "service_ids": [19944, 27567],
    # ntfy.sh: nainstaluj appku ntfy na mobil a odebírej toto téma (vymysli si unikátní název)
    "ntfy_topic": "",
    # E-mail přes SMTP (např. Gmail: smtp.gmail.com, 587, heslo = "heslo aplikace")
    "email": {
        "enabled": False,
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "username": "",
        "password": "",
        "to": "",
    },
}


def log(msg):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line.encode("ascii", "replace").decode())
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config():
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=2, ensure_ascii=False), encoding="utf-8")
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))


def api_post(opener, token, path, params):
    req = urllib.request.Request(
        BASE + path,
        data=urllib.parse.urlencode(params).encode(),
        headers={
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "X-XSRF-TOKEN": token,
            "Accept": "application/json",
        },
    )
    return json.loads(opener.open(req, timeout=30).read())["data"]


def fetch_terms(year, month, service_ids):
    """Vrátí termíny hlídaných služeb v daném měsíci: [{day, time, service, free, max}]."""
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.addheaders = [("User-Agent", "Mozilla/5.0 (sapa-hlidac)")]
    opener.open(BASE + "/", timeout=30).read()
    token = urllib.parse.unquote(next(c.value for c in cj if c.name == "XSRF-TOKEN"))

    terms = []
    day = date(year, month, 1)
    while day.month == month:
        # API vrací nejvýš ~8 dní najednou, proto po týdnech
        end = min(day + timedelta(days=6), date(year + (month == 12), month % 12 + 1, 1) - timedelta(days=1))
        data = api_post(opener, token, "/api/Term/List",
                        {"date": day.isoformat(), "endDate": end.isoformat(), "viewMode": "day"})
        for e in data["events"]:
            service = e["eventResources"][0]["service"]
            if service["id"] not in service_ids or not e.get("isEnabledForCustomer", True):
                continue
            # start je v UTC; v listopadu platí v Praze UTC+1
            start = datetime.strptime(e["start"], "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=1)
            booked = sum(r["capacity"] for r in e["reservations"])
            terms.append({
                "id": e["id"],
                "day": start.date(),
                "time": start.strftime("%H:%M"),
                "service": service["name"].strip("🍜 ").strip(),
                "free": e["maxCapacity"] - booked,
                "max": e["maxCapacity"],
            })
        day = end + timedelta(days=1)
    return terms


def describe(t):
    return f"{t['day'].day}. {t['day'].month}. {t['time']} {t['service']} (volno {t['free']}/{t['max']})"


# ---------- upozornění ----------

def notify_windows(title, text):
    ps = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml(@'
<toast activationType="protocol" launch="{BOOKING_URL}" scenario="reminder">
  <visual><binding template="ToastGeneric"><text>{title}</text><text>{text}</text></binding></visual>
  <actions><action content="Otevřít rezervace" activationType="protocol" arguments="{BOOKING_URL}"/></actions>
</toast>
'@)
$app = '{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe'
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show([Windows.UI.Notifications.ToastNotification]::new($xml))
"""
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True, timeout=30)
    except Exception as e:
        log(f"Windows notifikace selhala: {e}")


def notify_ntfy(topic, title, text, priority=5):
    if not topic:
        return
    try:
        payload = {"topic": topic, "title": title, "message": text, "priority": priority,
                   "click": BOOKING_URL, "tags": ["ramen"]}
        req = urllib.request.Request(
            "https://ntfy.sh/",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        urllib.request.urlopen(req, timeout=30).read()
    except Exception as e:
        log(f"ntfy selhalo: {e}")


def notify_email(cfg, title, text):
    # Na GitHubu se údaje berou z tajných proměnných (Secrets)
    cfg = dict(cfg)
    for key in ("username", "password", "to"):
        if os.environ.get(f"EMAIL_{key.upper()}"):
            cfg[key] = os.environ[f"EMAIL_{key.upper()}"]
            cfg["enabled"] = True
    if not cfg.get("enabled"):
        return
    try:
        msg = EmailMessage()
        msg["Subject"] = title
        msg["From"] = cfg["username"]
        msg["To"] = cfg["to"]
        msg.set_content(f"{text}\n\nRezervace: {BOOKING_URL}")
        with smtplib.SMTP(cfg["smtp_host"], cfg["smtp_port"], timeout=30) as s:
            s.starttls()
            s.login(cfg["username"], cfg["password"])
            s.send_message(msg)
    except Exception as e:
        log(f"E-mail selhal: {e}")


def notify(config, title, text, priority=5):
    log(f"UPOZORNĚNÍ: {title} - {text}")
    if sys.platform == "win32":
        notify_windows(title, text)
    notify_ntfy(os.environ.get("NTFY_TOPIC") or config.get("ntfy_topic", ""), title, text, priority)
    notify_email(config.get("email", {}), title, text)


# ---------- hlavní logika ----------

def prague_now():
    """Aktuální pražský čas (letní čas: poslední neděle v březnu až poslední neděle v říjnu, 01:00 UTC)."""
    now = datetime.now(timezone.utc).replace(tzinfo=None)

    def last_sunday(month):
        d = date(now.year, month, 31)
        return datetime(d.year, d.month, d.day - (d.weekday() + 1) % 7, 1)

    summer = last_sunday(3) <= now < last_sunday(10)
    return now + timedelta(hours=2 if summer else 1)


def heartbeat_slot(now):
    """Kontrola „hlídač běží“ v 10:00 a 22:00 (s rezervou na zpoždění GitHubu)."""
    if 10 <= now.hour < 13:
        return f"{now:%Y-%m-%d} 10:00"
    if now.hour >= 22:
        return f"{now:%Y-%m-%d} 22:00"
    return None


def main():
    config = load_config()
    if "--test" in sys.argv:
        notify(config, "Sapa Trip – test", "Zkušební upozornění, hlídač funguje.")
        return

    year, month, target = config["year"], config["month"], config["target_day"]
    target_date = date(year, month, target)

    terms = [t for t in fetch_terms(year, month, set(config["service_ids"])) if t["free"] > 0]
    state = json.loads(STATE_FILE.read_text(encoding="utf-8")) if STATE_FILE.exists() else {}
    seen = set(state.get("seen_terms", []))
    new = [t for t in terms if t["id"] not in seen]

    log("Volné české infotripy {}/{}: {}".format(month, year, "; ".join(describe(t) for t in terms) or "žádné"))

    new_target = [t for t in new if t["day"] == target_date]
    if new_target:
        notify(config, f"🎉 Sapa Trip: infotrip {target}. {month}. je volný!",
               "; ".join(describe(t) for t in new_target) + ". Rychle rezervuj!")
    others = [t for t in new if t["day"] != target_date]
    if others:
        notify(config, "Sapa Trip: nový infotrip v listopadu", "; ".join(describe(t) for t in others))

    new_state = dict(state, seen_terms=sorted(seen | {t["id"] for t in terms}))

    now = prague_now()
    slot = heartbeat_slot(now)
    if slot and state.get("last_heartbeat") != slot:
        summary = "; ".join(describe(t) for t in terms) or "zatím žádný volný český infotrip"
        notify(config, "✅ Sapa hlídač běží",
               f"Kontrola {now:%d. %m. %H:%M}. Listopad: {summary}.", priority=2)
        new_state["last_heartbeat"] = slot

    # Stav se přepisuje jen při změně (na GitHubu se pak commituje jen tehdy)
    if new_state != state:
        STATE_FILE.write_text(json.dumps(new_state, indent=2), encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"CHYBA: {e!r}")
        sys.exit(1)
