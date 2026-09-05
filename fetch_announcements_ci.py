import urllib.request, csv, json, io, re, time
from datetime import datetime

URL = ('https://docs.google.com/spreadsheets/d/e/'
       '2PACX-1vT9hMKavbas0ZlwI7Pb5vETPBiFiKslQNZImk_Cd0PeCZUTCP9QEtDTKyWmAb3mCMsUyCenu7DdNpUu'
       '/pub?gid=2124074003&single=true&output=csv')

# Google's publish endpoint occasionally responds slowly (~1/3 of scheduled runs timed out
# at 15s before this was added) — a longer timeout plus a couple of retries covers that
# without needing to wait for the next 5-min cron cycle to pick up missed data.
def fetch_url(url, tries=3, timeout=20):
    last_err = None
    for attempt in range(tries):
        try:
            return urllib.request.urlopen(url, timeout=timeout)
        except Exception as e:
            last_err = e
            if attempt < tries - 1:
                time.sleep(3)
    raise last_err

# Announcements response sheet columns (0-indexed):
#   0 Timestamp | 1 Email Address (unused, collection turned off) | 2 תאריך | 3 כותרת
#   4 תוכן | 5 קטגוריה | 6 עדיפות | 7 פעיל | 8 תוקף עד (if/when added)

def clean(s):
    return re.sub(r'[‎‏‪-‮⁦-⁩]', '', s).strip()

def col(row, i):
    return clean(row[i]) if len(row) > i and row[i].strip() else ''

# Tolerant of both the dropdown's own values (כן / גבוהה) and raw values that can end up
# here from pasting old Excel data directly (TRUE/FALSE, "1-גבוהה" style prefixes).
def is_active(v):
    return v.strip().upper() in ('כן', 'TRUE', '1')

def priority_num(v):
    if 'גבוהה' in v: return 1
    if 'בינונית' in v: return 2
    if 'נמוכה' in v: return 3
    return 2

def norm_date(v):
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', v)
    return f'{m.group(3)}/{m.group(2)}/{m.group(1)}' if m else v.split(' ')[0]

# Parses whichever date format ended up in this field (Form's own M/D/Y locale,
# or Y-M-D from a pasted Excel value) so we can tell if it's a future date.
def parse_date(v):
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(v.split(' ')[0], fmt).date()
        except ValueError:
            continue
    return None

r    = fetch_url(URL)
text = r.read().decode('utf-8')
rows = list(csv.reader(io.StringIO(text)))

today = datetime.now().date()
anns = []
for i, row in enumerate(rows[1:], 1):
    if not col(row, 3):  # no title — skip (works whether the row came from the
        continue         # Form, which auto-fills Timestamp/Email, or was typed directly)
    if not is_active(col(row, 7)):
        continue
    date_val = col(row, 2)
    parsed_date = parse_date(date_val) if date_val else None
    if parsed_date and parsed_date > today:
        continue  # scheduled for the future — not published yet
    anns.append({
        'id':       i,
        'date':     norm_date(col(row, 2)),
        'title':    col(row, 3),
        'content':  col(row, 4),
        'category': col(row, 5) or 'מידע',
        'priority': priority_num(col(row, 6)),
    })

with open('announcements.json', 'w', encoding='utf-8') as f:
    json.dump(anns, f, ensure_ascii=False, indent=2)
print(f'Wrote {len(anns)} announcements')
