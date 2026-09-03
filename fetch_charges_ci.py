import urllib.request, csv, io, json
from datetime import datetime

CHARGES_URL = ('https://docs.google.com/spreadsheets/d/e/'
                '2PACX-1vSjSHtIBvQQI72QQFjLoo3UezggmCZhfhixsg2a7ZZ9gcvHlU7JKFUyubhWpvAcgTVfSBGdaZ3D-7EJ'
                '/pub?gid=0&single=true&output=csv')
PAYMENTS_URL = ('https://docs.google.com/spreadsheets/d/e/'
                 '2PACX-1vSjSHtIBvQQI72QQFjLoo3UezggmCZhfhixsg2a7ZZ9gcvHlU7JKFUyubhWpvAcgTVfSBGdaZ3D-7EJ'
                 '/pub?gid=615900064&single=true&output=csv')

# Tenant count for the target calc — building has 14 units, changes rarely. The real
# dashboard's own Excel-based tenant count remains authoritative on the next local
# generator run; this is only used for the live progress-bar preview between runs.
TENANT_COUNT = 14

# גביות sheet columns (0-indexed): 0 charge_id | 1 name | 2 amount | 3 date | 4 active | 5 description
# תשלומים sheet columns (0-indexed): 0 charge_id | 1 tenant_name | 2 amount_paid | 3 updated_at

def fetch_csv(url):
    r = urllib.request.urlopen(url, timeout=15)
    text = r.read().decode('utf-8')
    return list(csv.reader(io.StringIO(text)))

charge = {}
rows = fetch_csv(CHARGES_URL)
for row in rows[1:]:
    if len(row) < 5 or not row[0].strip():
        continue
    if (row[4] or '').strip() != 'כן':
        continue
    try:
        amount = float(row[2])
    except (ValueError, IndexError):
        continue
    if amount <= 0:
        continue
    date_s = row[3].strip() if len(row) > 3 else ''
    if date_s:
        try:
            if datetime.now() < datetime.strptime(date_s, '%d/%m/%Y'):
                continue  # not started yet, same rule as the local generator
        except ValueError:
            pass
    charge = {
        'id': row[0].strip(),
        'name': row[1].strip() if len(row) > 1 else 'חיוב מיוחד',
        'amount': amount,
        'description': row[5].strip() if len(row) > 5 else '',
    }
    break  # first active row wins, same rule as the local generator

result = {}
if charge:
    collected = 0.0
    prows = fetch_csv(PAYMENTS_URL)
    for row in prows[1:]:
        if len(row) < 3 or not row[1].strip():
            continue
        if row[0].strip() != charge['id']:
            continue
        try:
            collected += float(row[2])
        except ValueError:
            continue
    target = charge['amount'] * TENANT_COUNT
    pct = min(round(collected / target * 100), 100) if target else 0
    result = {
        'id': charge['id'],
        'name': charge['name'],
        'description': charge['description'],
        'collected': collected,
        'target': target,
        'pct': pct,
    }

with open('charges.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)
print(f'Wrote charges.json: {result or "no active charge"}')
