import urllib.request, csv, json, io, re

URL = ('https://docs.google.com/spreadsheets/d/e/'
       '2PACX-1vTXlu7rgEen0MlaU_k0omizQ_9kUZgJ9M49cMiB7tumP__JEfjakPfiVpwN6bXUhQlYcrRGRhL9jQmJ'
       '/pub?output=csv&gid=386569253')

def clean(s):
    return re.sub(r'[‎‏‪-‮⁦-⁩]', '', s).strip()

r    = urllib.request.urlopen(URL, timeout=15)
text = r.read().decode('utf-8')
rows = list(csv.reader(io.StringIO(text)))

issues = []
for i, row in enumerate(rows[1:], 1):
    if len(row) < 2 or not row[0].strip():
        continue
    hidden = clean(row[12]) if len(row) > 12 else ''
    if hidden in ('כן', 'yes', '1'):
        continue
    b   = clean(row[2]) if len(row) > 2 else ''
    a   = clean(row[3]) if len(row) > 3 else ''
    apt = ('בניין ' + b + ' דירה ' + a).strip() if b or a else ''
    st  = clean(row[9]) if len(row) > 9 and row[9].strip() else 'פתוח'
    date_raw = clean(row[0]).split(' ')[0]
    issues.append({
        'id':          i,
        'date':        date_raw,
        'name':        clean(row[1]) if len(row) > 1 else '',
        'apt':         apt,
        'category':    clean(row[5]) if len(row) > 5 else '',
        'desc':        clean(row[4]) if len(row) > 4 else '',
        'status':      st,
        'update_date': clean(row[10]) if len(row) > 10 else '',
        'notes':       clean(row[11]) if len(row) > 11 else '',
    })

with open('issues.json', 'w', encoding='utf-8') as f:
    json.dump(issues, f, ensure_ascii=False, indent=2)
print(f'Wrote {len(issues)} issues')
