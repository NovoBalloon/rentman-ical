#!/usr/bin/env python3
import os, re, sys, datetime as dt
import requests
from icalendar import Calendar, Event, vText

RENTMAN_ICAL_URL = os.getenv("RENTMAN_ICAL_URL")
OUTPUT_ICS       = os.getenv("OUTPUT_ICS", "rentman_usage_calendar.ics")

CONFIRM_PAT = re.compile(r"\b(confirm|confirmed)\b", re.I)
OPTION_PAT  = re.compile(r"\b(option|pending|tentative|concept)\b", re.I)

def label_status(text: str) -> str:
    if not text: return "Pending"
    if CONFIRM_PAT.search(text): return "Confirmed"
    if OPTION_PAT.search(text):  return "Pending"
    return "Pending"

def main():
    if not RENTMAN_ICAL_URL:
        print("Set RENTMAN_ICAL_URL to your Rentman Projects iCal URL.", file=sys.stderr)
        sys.exit(1)

    r = requests.get(RENTMAN_ICAL_URL, timeout=30)
    r.raise_for_status()
    src = Calendar.from_ical(r.content)

    out = Calendar()
    out.add("prodid", "-//Rentman Usage Calendar (Transformed)//")
    out.add("version", "2.0")
    out.add("method", "PUBLISH")
    out.add("x-wr-calname", "Rentman – Usage Periods")
    out.add("x-wr-timezone", "America/Vancouver")

    now = dt.datetime.now(dt.timezone.utc)

    for c in src.walk():
        if c.name != "VEVENT":
            continue
        summary = str(c.get("summary", "Project"))
        desc = str(c.get("description", ""))
        loc = str(c.get("location", ""))
        dtstart = c.get("dtstart").dt
        dtend   = c.get("dtend").dt

        label = label_status(summary + "\n" + desc)
        new_summary = f"[{label}] {summary}"

        ev = Event()
        ev.add("uid", str(c.get("uid", "")))
        ev.add("summary", new_summary)
        ev.add("dtstart", dtstart)
        ev.add("dtend", dtend)
        ev.add("dtstamp", now)
        ev.add("last-modified", now)
        ev.add("status", "CONFIRMED" if label == "Confirmed" else "TENTATIVE")
        ev.add("transp", "OPAQUE")
        if loc:
            ev.add("location", vText(loc))
        new_desc = desc if "Status:" in desc else f"Status: {label}\n\n{desc}".strip()
        ev.add("description", new_desc)

        out.add_component(ev)

    with open(OUTPUT_ICS, "wb") as f:
        f.write(out.to_ical())
    print(f"✅ Wrote {OUTPUT_ICS}")

if __name__ == "__main__":
    main()

