import urllib.request, csv, json, io

URL = ('https://docs.google.com/spreadsheets/d/e/'
       '2PACX-1vTXlu7rgEen0MlaU_k0omizQ_9kUZgJ9M49cMiB7tumP__JEfjakPfiVpwN6bXUhQlYcrRGRhL9jQmJ'
       '/pub?output=csv&gid=386569253')

r    = urllib.request.urlopen(URL, timeout=15)
text = r.read().decode('utf-8')
rows = list(csv.reader(io.StringIO(text)))

issues = []
for i, row in enumerate(rows[1:], 1):
    if len(row) < 2 or not row[0].strip():
        continue
    hidden = row[12].strip() if len(row) > 12 else ''
    if hidden in ('כן', 'yes', '1'):
        continue
    b   = row[2].strip() if len(row) > 2 else ''
    a   = row[3].strip() if len(row) > 3 else ''
    apt = ('בניין ' + b + ' דירה ' + a).strip() if b or a else ''
    st  = row[9].strip() if len(row) > 9 and row[9].strip() else 'פתוח'
    issues.append({
        'id':          i,
        'date':        row[0].strip(),
        'name':        row[1].strip() if len(row) > 1 else '',
        'apt':         apt,
        'category':    row[5].strip() if len(row) > 5 else '',
        'desc':        row[4].strip() if len(row) > 4 else '',
        'status':      st,
        'update_date': row[10].strip() if len(row) > 10 else '',
        'notes':       row[11].strip() if len(row) > 11 else '',
    })

with open('issues.json', 'w', encoding='utf-8') as f:
    json.dump(issues, f, ensure_ascii=False, indent=2)
print(f'Wrote {len(issues)} issues')
