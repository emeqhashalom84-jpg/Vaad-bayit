import urllib.request, csv, io, json

# עדכון פרטים אישיים (Responses) — the single source of truth for tenant/contact data.
# Publish that sheet to web as CSV (File > Share > Publish to web > CSV, "Automatically
# republish when changes are made" checked) and paste the resulting URL here.
CONTACTS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSxiASsJ9hMm-5nd3mSEwRYk1-vuOLCU2kKGmeKx4PxBjT_nwA1WtP-uK61EjtcQ6oLXL7-vYg6I8uQ/pub?gid=1668586487&single=true&output=csv'

# Columns are found by matching HEADER TEXT, not fixed position — Google Forms keeps a
# response sheet's columns in the order questions were ORIGINALLY created, which does not
# follow later reordering in the form editor, so a fixed-position mapping breaks silently.
# Mirrors tenantColMap_/findCol_ in apps_script_admin.js exactly, so both stay in sync.
def find_col(headers, must_have, must_not_have=()):
    for i, h in enumerate(headers):
        h = h or ''
        if all(s in h for s in must_have) and not any(s in h for s in must_not_have):
            return i
    return -1

def build_col_map(headers):
    return {
        'building':   find_col(headers, ['בית']),
        'apt':        find_col(headers, ['מספר', 'דירה']),
        'lastName':   find_col(headers, ['משפחה']),
        'c1name':     find_col(headers, ['קשר', '1'], ['טלפון', 'מייל']),
        'c1phone':    find_col(headers, ['קשר', '1', 'טלפון']),
        'c1email':    find_col(headers, ['קשר', '1', 'מייל']),
        'c2name':     find_col(headers, ['קשר', '2'], ['טלפון', 'מייל']),
        'c2phone':    find_col(headers, ['קשר', '2', 'טלפון']),
        'c2email':    find_col(headers, ['קשר', '2', 'מייל']),
        'rented':     find_col(headers, ['שכורה']),
        'ownerName':  find_col(headers, ['בעל', 'שם']),
        'ownerPhone': find_col(headers, ['בעל', 'טלפון']),
        'ownerEmail': find_col(headers, ['בעל', 'מייל']),
        'status':     find_col(headers, ['סטטוס']),
    }

def get(row, col_map, key):
    i = col_map.get(key, -1)
    if i == -1 or i >= len(row):
        return ''
    return (row[i] or '').strip()

def contact(name, phone, email):
    if not name:
        return None
    return {'name': name, 'phone': phone, 'email': email}

# Standard household display name: contact1 [+ "ו"+contact2] + lastName, e.g.
# "אורן ואורלי אלקיים" (or just "אורן אלקיים" when there's no contact2).
def build_display_name(c1_name, c2_name, last_name):
    head = c1_name or ''
    if c2_name:
        head = f'{head} ו{c2_name}' if head else f'ו{c2_name}'
    if head:
        return f'{head} {last_name}' if last_name else head
    return last_name

out = []
if CONTACTS_CSV_URL and 'PASTE' not in CONTACTS_CSV_URL:
    r = urllib.request.urlopen(CONTACTS_CSV_URL, timeout=15)
    rows = list(csv.reader(io.StringIO(r.read().decode('utf-8'))))
    headers, data_rows = rows[0], rows[1:]
    col_map = build_col_map(headers)

    # Collapse every row to the LATEST non-dismissed submission per building+apt — a
    # household that updated its details more than once never shows as duplicate cards.
    # Rows marked 'בוטל' are excluded entirely. Mirrors getTenantCards() in
    # apps_script_admin.js exactly.
    winners, order = {}, []
    for row in data_rows:
        building, apt = get(row, col_map, 'building'), get(row, col_map, 'apt')
        if not building and not apt:
            continue
        if get(row, col_map, 'status') == 'בוטל':
            continue
        addr = f'{building}|{apt}'
        if addr not in winners:
            order.append(addr)
        winners[addr] = row  # later row always overwrites — latest wins

    for addr in order:
        row = winners[addr]
        rented = get(row, col_map, 'rented') == 'כן'
        last_name = get(row, col_map, 'lastName')
        c1_name, c2_name = get(row, col_map, 'c1name'), get(row, col_map, 'c2name')
        c1 = contact(c1_name, get(row, col_map, 'c1phone'), get(row, col_map, 'c1email'))
        c2 = contact(c2_name, get(row, col_map, 'c2phone'), get(row, col_map, 'c2email'))
        owner_name = get(row, col_map, 'ownerName')
        owner = contact(owner_name, get(row, col_map, 'ownerPhone'), get(row, col_map, 'ownerEmail')) if rented else None
        out.append({
            'building':    get(row, col_map, 'building'),
            'apt':         get(row, col_map, 'apt'),
            'lastName':    last_name,
            'displayName': build_display_name(c1_name, c2_name, last_name),
            'contact1':    c1,
            'contact2':    c2,
            'rented':      rented,
            'owner':       owner,
        })

with open('contacts.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False)
print(f'Wrote {len(out)} tenant cards')
