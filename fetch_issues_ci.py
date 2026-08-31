import urllib.request, csv, json, io, re

URL = ('https://docs.google.com/spreadsheets/d/e/'
       '2PACX-1vTXlu7rgEen0MlaU_k0omizQ_9kUZgJ9M49cMiB7tumP__JEfjakPfiVpwN6bXUhQlYcrRGRhL9jQmJ'
       '/pub?output=csv&gid=386569253')

# Calls response sheet columns (0-indexed):
#   0 Timestamp | 1 שם מלא | 2 מספר בית | 3 מספר דירה | 4 תיאור התקלה
#   5 מיקום התקלה | 6 דחיפות | 7 העלאת תמונות/וידאו | 8 הערות
#   9 סטטוס (J) | 10 פעיל (K) | 11 תאריך עדכון סטטוס (L)

def clean(s):
    return re.sub(r'[‎‏‪-‮⁦-⁩]', '', s).strip()

def col(row, i):
    return clean(row[i]) if len(row) > i and row[i].strip() else ''

r    = urllib.request.urlopen(URL, timeout=15)
text = r.read().decode('utf-8')
rows = list(csv.reader(io.StringIO(text)))

issues = []
for i, row in enumerate(rows[1:], 1):
    if len(row) < 2 or not row[0].strip():
        continue
    if col(row, 10) == 'לא פעיל':
        continue
    b   = col(row, 2)
    a   = col(row, 3)
    apt = ('בניין ' + b + ' דירה ' + a).strip() if b or a else ''
    images_raw = col(row, 7)
    images = [u.strip() for u in images_raw.split(',') if u.strip()]
    issues.append({
        'id':          i,
        'date':        col(row, 0).split(' ')[0],
        'name':        col(row, 1),
        'apt':         apt,
        'location':    col(row, 5),
        'urgency':     col(row, 6),
        'desc':        col(row, 4),
        'status':      col(row, 9) or 'פתוח',
        'update_date': col(row, 11),
        'notes':       col(row, 8),
        'images':      images,
    })

with open('issues.json', 'w', encoding='utf-8') as f:
    json.dump(issues, f, ensure_ascii=False, indent=2)
print(f'Wrote {len(issues)} issues')
