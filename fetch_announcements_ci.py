import urllib.request, csv, json, io, re

URL = ('https://docs.google.com/spreadsheets/d/e/'
       '2PACX-1vT9hMKavbas0ZlwI7Pb5vETPBiFiKslQNZImk_Cd0PeCZUTCP9QEtDTKyWmAb3mCMsUyCenu7DdNpUu'
       '/pub?gid=2124074003&single=true&output=csv')

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

r    = urllib.request.urlopen(URL, timeout=15)
text = r.read().decode('utf-8')
rows = list(csv.reader(io.StringIO(text)))

anns = []
for i, row in enumerate(rows[1:], 1):
    if not col(row, 3):  # no title — skip (works whether the row came from the
        continue         # Form, which auto-fills Timestamp/Email, or was typed directly)
    if not is_active(col(row, 7)):
        continue
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
