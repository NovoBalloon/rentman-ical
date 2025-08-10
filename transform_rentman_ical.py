#!/usr/bin/env python3
# One-calendar Rentman ICS:
# - Keeps usage-time override when detectable
# - Prefixes titles with 🟢/🟡 and [Confirmed]/[Pending]
# - Sets STATUS and CATEGORIES accordingly

import re
import requests
from icalendar import Calendar, Event, vText
from datetime import datetime, date, time, timezone
import pytz

# --------- CONFIG ----------
RENTMAN_ICAL_URL = (
    "https://novolightingltd.sync.rentman.eu/ical.php?c=ap&i=33&t=hnAroMIDTw6eRlic8VZTa7tc1YVOgrCNTMC6tmD3gg4%3D%3A%3A%3Aef8c8108db57083c98a9fa98%3A%3A%3A"
    "52a43e6f375c6e78589ede917e39542c"
)
TZID = "America/Vancouver"
TZ = pytz.timezone(TZID)
OUTPUT_ICS = "rentman_usage_calendar.ics"

# --------- STATUS DETECTION ----------
CONFIRM_PAT = re.compile(r"\b(confirm(ed)?|best(ae|ä)tigt)\b", re.I)
OPTION_PAT  = re.compile(r"\b(option|concept|konzept|tentative|pending)\b", re.I)
CANCEL_PAT  = re.compile(r"\b(cancel(l)?(ed)?|storniert|abgesagt)\b", re.I)

def status_from_component(comp):
    """Return ('CONFIRMED'|'TENTATIVE'|'CANCELLED', 'Confirmed'|'Pending'|'Cancelled')."""
    s = comp.get("status")
    if s:
        sval = str(s).strip().upper()
        if "CANCEL" in sval:  return ("CANCELLED", "Cancelled")
        if "CONFIRM" in sval: return ("CONFIRMED", "Confirmed")
        if "TENTATIVE" in sval: return ("TENTATIVE", "Pending")

    cats = comp.get("categories")
    if cats:
        try:
            vals = [str(v) for v in (cats.cats if hasattr(cats, "cats") else [cats])]
        except Exception:
            vals = [str(cats)]
        joined = " ".join(vals)
        if CANCEL_PAT.search(joined):  return ("CANCELLED", "Cancelled")
        if CONFIRM_PAT.search(joined): return ("CONFIRMED", "Confirmed")
        if OPTION_PAT.search(joined):  return ("TENTATIVE", "Pending")

    text = (str(comp.get("summary","")) + "\n" + str(comp.get("description",""))).lower()
    if CANCEL_PAT.search(text):  return ("CANCELLED", "Cancelled")
    if CONFIRM_PAT.search(text): return ("CONFIRMED", "Confirmed")
    if OPTION_PAT.search(text):  return ("TENTATIVE", "Pending")
    return ("TENTATIVE", "Pending")

# --------- USAGE DETECTION ----------
USAGE_LINE = re.compile(r"usage[^:\n]*[:\-]?\s*(.+)", re.I)
DT_PAIR = re.compile(
    r"(\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]\d{1,2}:\d{2}(?::\d{2})?)\s*(?:→|to|-|–)\s*"
    r"(\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]\d{1,2}:\d{2}(?::\d{2})?)",
    re.I,
)
TIME_PAIR = re.compile(
    r"usage[^:\n]*[:\-]?\s*(\d{1,2}:\d{2})\s*(?:→|to|-|–)\s*(\d{1,2}:\d{2})",
    re.I,
)
DT_FORMATS = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M",
]

def parse_dt(s: str):
    s = s.strip().replace("Z", "+00:00")
    try:
        if ("T" in s and ("+" in s or s.endswith("Z"))):
            dtv = datetime.fromisoformat(s.replace("Z","+00:00"))
            return dtv if dtv.tzinfo else TZ.localize(dtv)
    except Exception:
        pass
    for fmt in DT_FORMATS:
        try:
            dtv = datetime.strptime(s, fmt)
            return TZ.localize(dtv)
        except Exception:
            continue
    return None

def usage_override(summary: str, desc: str, original_start, original_end):
    text = f"{summary}\n{desc}"
    for line in text.splitlines():
        m = USAGE_LINE.search(line)
        if not m: continue
        tail = m.group(1)
        p = DT_PAIR.search(tail)
        if p:
            s_dt = parse_dt(p.group(1)); e_dt = parse_dt(p.group(2))
            if s_dt and e_dt: return s_dt, e_dt
    for line in text.splitlines():
        p = TIME_PAIR.search(line)
        if p:
            s_t = datetime.strptime(p.group(1), "%H:%M").time()
            e_t = datetime.strptime(p.group(2), "%H:%M").time()
            sd = original_start.date() if isinstance(original_start, datetime) else original_start
            ed = original_end.date()   if isinstance(original_end, datetime) else original_end
            return TZ.localize(datetime.combine(sd, s_t)), TZ.localize(datetime.combine(ed, e_t))
    return None

def main():
    print("📥 Downloading source ICS…")
    resp = requests.get(RENTMAN_ICAL_URL, timeout=30)
    resp.raise_for_status()
    src = Calendar.from_ical(resp.content)

    out = Calendar()
    out.add("prodid", "-//Rentman Usage Calendar (One Feed)//")
    out.add("version", "2.0")
    out.add("method", "PUBLISH")
    out.add("x-wr-calname", "Rentman – Usage (One Feed)")
    out.add("x-wr-timezone", TZID)

    changed_usage = 0
    now = datetime.now(timezone.utc)

    for comp in src.walk():
        if comp.name != "VEVENT":
            continue

        summary = str(comp.get("summary", "Project"))
        desc    = str(comp.get("description", ""))
        loc     = str(comp.get("location", ""))

        dtstart = comp.get("dtstart").dt
        dtend   = comp.get("dtend").dt
        if isinstance(dtstart, date) and not isinstance(dtstart, datetime):
            dtstart = TZ.localize(datetime.combine(dtstart, time(0,0)))
        if isinstance(dtend, date) and not isinstance(dtend, datetime):
            dtend = TZ.localize(datetime.combine(dtend, time(0,0)))

        u = usage_override(summary, desc, dtstart, dtend)
        if u:
            dtstart, dtend = u
            changed_usage += 1

        ical_status, label = status_from_component(comp)

        # Emoji badges for universal visual “color”
        badge = "🟢" if ical_status == "CONFIRMED" else ("🟡" if ical_status == "TENTATIVE" else "⚫")
        new_summary = f"{badge} [{label}] {summary}"

        ev = Event()
        ev.add("uid", str(comp.get("uid", "")))
        ev.add("summary", new_summary)
        ev.add("dtstart", dtstart)
        ev.add("dtend", dtend)
        ev.add("dtstamp", now)
        ev.add("last-modified", now)
        ev.add("status", ical_status)
        ev.add("transp", "OPAQUE")
        # Categories can be used by Outlook rules; other clients ignore
        ev.add("categories", label)
        if loc:
            ev.add("location", vText(loc))
        if "Status:" in desc:
            ev.add("description", desc)
        else:
            ev.add("description", f"Status: {label}\n\n{desc}".strip())

        out.add_component(ev)

    with open(OUTPUT_ICS, "wb") as f:
        f.write(out.to_ical())
    print(f"✅ Wrote {OUTPUT_ICS} (usage overrides on {changed_usage} events)")

if __name__ == "__main__":
    main()

