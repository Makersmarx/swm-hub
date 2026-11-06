"""Scrape 26-27 Indoor 1-MS Oberliga standings + fixtures from EZLeagues into league.json.
Usage: python3 -I scrape_oberliga.py <out.json>
"""
import re, sys, json, html, urllib.request, datetime

URL = "https://stjoekickers.ezleagues.ezfacility.com/leagues/481959/26-27-Indoor-1-MS-Oberliga.aspx"
MON = {m: i for i, m in enumerate(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], 1)}

def clean(c):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", c))).strip()

def tables(t):
    out = []
    for tb in re.findall(r"(?s)<table.*?</table>", t):
        rows = []
        for tr in re.findall(r"(?s)<tr.*?</tr>", tb):
            cells = [clean(c) for c in re.findall(r"(?s)<t[hd][^>]*>(.*?)</t[hd]>", tr)]
            if any(cells):
                rows.append(cells)
        out.append(rows)
    return out

def iso(d):  # "Wed-Nov 4" -> season runs Nov..Jan
    m = re.match(r"\w+-(\w{3}) (\d+)", d)
    mo, day = MON[m.group(1)], int(m.group(2))
    return f"{2026 if mo >= 8 else 2027}-{mo:02d}-{day:02d}"

def t24(s):
    m = re.match(r"(\d+):(\d+) (AM|PM)", s)
    if not m: return None
    h = int(m.group(1)) % 12 + (12 if m.group(3) == "PM" else 0)
    return f"{h:02d}:{m.group(2)}"

req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
page = urllib.request.urlopen(req, timeout=30).read().decode("utf8", "ignore")
st = fx = None
for rows in tables(page):
    if rows and "PTS" in rows[0]: st = rows
    elif rows and rows[0][:2] == ["Date", "Home"]: fx = rows
assert st and fx, "standings or schedule table not found"

hdr = st[0]
idx = {k: hdr.index(k) for k in ["GP","W","L","T","GF","GA","PTS","GD"]}
standings = [{"team": r[0], **{k.lower(): int(r[i]) for k, i in idx.items()}} for r in st[1:] if r[0]]
standings.sort(key=lambda x: (-x["pts"], -x["gd"], -x["gf"], x["team"]))

fixtures = []
for r in fx[1:]:
    date, home, mid, away, status = r[0], r[1], r[2], r[3], r[4]
    sc = re.match(r"(\d+)\s*-\s*(\d+)", mid)
    fixtures.append({"date": iso(date), "home": home, "away": away,
                     "hs": int(sc.group(1)) if sc else None, "as": int(sc.group(2)) if sc else None,
                     "time": t24(status), "status": "final" if sc else ("scheduled" if t24(status) else status)})

out = {"league": "26-27 Indoor 1-MS Oberliga", "source": URL,
       "scrapedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
       "standings": standings, "fixtures": fixtures}
json.dump(out, open(sys.argv[1], "w"), indent=1)
print(json.dumps(standings, indent=1)); print(len(fixtures), "fixtures")
