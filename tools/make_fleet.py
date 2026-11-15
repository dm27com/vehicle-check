"""Builds app/fleet.js from the Service Schedule workbook.
Usage: python3 make_fleet.py <workbook.xlsx> <output fleet.js>"""
import sys, json, re, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
fleet, seen = [], set()
def num(v):
    if isinstance(v, (int, float)): return int(v)
    if isinstance(v, str):
        m = re.match(r"\s*([\d,]+)\s*kms", v)
        if m: return int(m.group(1).replace(",", ""))
    return None
for ws in wb:
    site = (ws["A2"].value or ws.title.replace("Services", "")).strip().title().replace("Sb2", "SB2")
    rows = list(ws.iter_rows(values_only=True))
    for i, r in enumerate(rows):
        if len(r) < 13 or r[4] != "Scheduled Service Kms" or not r[2]: continue
        make = re.sub(r"\s{2,}.*", "", str(r[1] or "")).strip().replace("Peugoet", "Peugeot")
        reg = re.sub(r"\s+", " ", str(r[2])).strip().upper()
        interval = num(r[3])
        mv = re.search(r"Transfer+ed to\s+(\w+)", str(r[3] or ""), re.I)
        vsite = site
        if mv:  # e.g. "Transfered to Nacala 28 Oct 2026": list under the new site, interval from the schedule row
            vsite = mv.group(1).title()
            interval = num(r[5]) if isinstance(r[5], (int, float)) else None
        sched = [num(x) for x in r[5:13]]
        arrival = num(r[5]) if isinstance(r[5], str) else None  # "65,100 kms on arrival to site"
        if arrival: sched[0] = None
        actual = [num(x) for x in rows[i + 1][5:13]] if i + 1 < len(rows) else [None] * 8
        done = [k for k, a in enumerate(actual) if a]
        last_km = actual[done[-1]] if done else arrival
        nxt = None
        start = done[-1] + 1 if done else 0
        for k in range(start, 8):
            s = sched[k]
            if s and s != interval and (last_km is None or s > last_km):
                nxt = s; break
        if nxt is None and last_km and interval: nxt = last_km + interval
        key = reg.replace(" ", "")
        if key in seen: continue
        seen.add(key)
        fleet.append({"reg": reg, "make": make, "site": vsite, "interval": interval, "lastKm": last_km, "next": nxt})
with open(sys.argv[2], "w") as f:
    f.write("// Generated from " + sys.argv[1].split("/")[-1] + " — do not edit by hand; re-run tools/make_fleet.py\n")
    f.write("window.FLEET = " + json.dumps(fleet, ensure_ascii=False, indent=1) + ";\n")
for v in fleet: print(f"{v['site']:10} {v['reg']:12} {v['make']:18} every {v['interval']}  last {v['lastKm']}  next {v['next']}")
print(len(fleet), "vehicles")
