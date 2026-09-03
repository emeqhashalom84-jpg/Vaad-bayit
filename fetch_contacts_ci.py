import urllib.request, csv, io, json

# Fill in once the restructured contacts sheet (see apps_script_admin.js's contacts
# section for the column layout) is published to web as CSV. Until then this script
# writes an empty contacts.json and the dashboard's tenant-card section stays empty
# (harmless — same "dormant until configured" pattern as the charges feature had).
CONTACTS_CSV_URL = 'PASTE_URL_HERE'

# Columns (0-indexed): 0 בניין | 1 דירה | 2 שם משפחה | 3 איש קשר 1 שם | 4 טלפון | 5 מייל |
#   6 איש קשר 2 שם | 7 טלפון | 8 מייל | 9 דירה שכורה (כן/לא) | 10 בעל הדירה שם | 11 טלפון | 12 מייל

def contact(row, name_i, phone_i, email_i):
    name = row[name_i].strip() if len(row) > name_i else ''
    if not name:
        return None
    return {
        'name': name,
        'phone': row[phone_i].strip() if len(row) > phone_i else '',
        'email': row[email_i].strip() if len(row) > email_i else '',
    }

out = []
if CONTACTS_CSV_URL and 'PASTE' not in CONTACTS_CSV_URL:
    r = urllib.request.urlopen(CONTACTS_CSV_URL, timeout=15)
    rows = list(csv.reader(io.StringIO(r.read().decode('utf-8'))))
    for row in rows[1:]:
        if len(row) < 3 or not row[0].strip():
            continue
        rented = (row[9].strip() if len(row) > 9 else '') == 'כן'
        owner = contact(row, 10, 11, 12) if rented else None
        out.append({
            'building':  row[0].strip(),
            'apt':       row[1].strip(),
            'lastName':  row[2].strip() if len(row) > 2 else '',
            'contact1':  contact(row, 3, 4, 5),
            'contact2':  contact(row, 6, 7, 8),
            'rented':    rented,
            'owner':     owner,
        })

with open('contacts.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False)
print(f'Wrote {len(out)} tenant cards')
