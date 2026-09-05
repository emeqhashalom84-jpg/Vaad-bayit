#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vaad_bayit_generator.py
ועד בית עמק השלום — Dashboard Generator
Reads Excel + Google Sheets → generates static HTML dashboard → Git push
"""

import configparser, openpyxl, requests, csv, io, os, sys, math, shutil, subprocess, logging, time, threading
from datetime import datetime
from pathlib import Path
from html import escape as he

# ── Watchdog (optional import) ──────────────────────────────────────────────
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    HAS_WATCHDOG = True
except ImportError:
    HAS_WATCHDOG = False

try:
    import schedule
    HAS_SCHEDULE = True
except ImportError:
    HAS_SCHEDULE = False

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            os.path.join(os.path.dirname(__file__), 'generator.log'),
            encoding='utf-8'
        )
    ]
)
log = logging.getLogger('vaad')

CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.ini')

MONTHS_HE    = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני',
                'יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר']
MONTHS_SHORT = ['ינו','פבר','מרץ','אפר','מאי','יוני','יולי','אוג','ספט','אוק','נוב','דצמ']

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
def load_config():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding='utf-8')
    return cfg

def cfg_bool(cfg, section, key, fallback=True):
    return cfg.get(section, key, fallback=str(fallback)).lower() in ('true','1','yes')

def cfg_int(cfg, section, key, fallback=0):
    try: return int(cfg.get(section, key, fallback=str(fallback)))
    except: return fallback

def cfg_float(cfg, section, key, fallback=0.0):
    try: return float(cfg.get(section, key, fallback=str(fallback)))
    except: return fallback

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL UTILITIES
# ─────────────────────────────────────────────────────────────────────────────
def _clean(v):
    if v is None: return None
    s = str(v).strip().replace('‫','').replace('‬','').replace('‏','').replace('‎','')
    return s if s not in ('None','') else None

def _num(v):
    if v is None: return None
    try: return float(str(v).replace('₪','').replace(',','').replace(' ','').strip())
    except: return None

def _sheet(wb, *names):
    for n in names:
        if n in wb.sheetnames: return wb[n]
        stripped = n.strip()
        for sn in wb.sheetnames:
            if sn.strip() == stripped: return wb[sn]
    return None

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL READER
# ─────────────────────────────────────────────────────────────────────────────
def read_excel(path):
    path = str(path)
    wb = openpyxl.load_workbook(path, data_only=True)

    # ── Dashboard — balance string only (income/expense come from sheets) ───
    ws_dash = _sheet(wb, 'Dashboard')
    balance_str = None
    if ws_dash:
        for row in ws_dash.iter_rows(values_only=True):
            for v in row:
                sv = _clean(v)
                if sv and '₪' in str(sv):
                    balance_str = sv

    # ── Tenant Payments ──────────────────────────────────────────────────────
    ws_inc = _sheet(wb, 'ריכוז הכנסות 2026 דיירים')
    tenants = []
    monthly_income = [0.0] * 12
    income_total_calc = 0.0
    _other_income     = 0.0
    collection_rate = None

    if ws_inc:
        _raw_rows = list(ws_inc.iter_rows())
        rows = [tuple(c.value for c in r) for r in _raw_rows]
        # collect cell comments (col indices 1-12 = months Jan-Dec)
        _comment_rows = [tuple(
            c.comment.text.strip() if c.comment and c.comment.text else None
            for c in r
        ) for r in _raw_rows]
        _inc_comments = {}
        for _ci, _cr in enumerate(rows):
            _nm = _clean(_cr[0]) if _cr else None
            if _nm and len(_nm) >= 2 and 'אחר' not in _nm and 'סה"כ' not in _nm:
                _inc_comments[_nm] = list(_comment_rows[_ci][1:13])
        for row in rows:
            name = _clean(row[0]) if row else None
            if not name or len(name) < 2: continue
            if 'אחר' in name:
                other_monthly = [_num(row[i]) if len(row) > i else None for i in range(1, 13)]
                _other_income += sum(v for v in other_monthly if v)
                continue

            # totals row — source of income_total_calc and monthly_income
            if 'סה"כ הכנסות' in name:
                for i in range(12):
                    v = _num(row[1+i]) if len(row) > 1+i else None
                    if v: monthly_income[i] = v
                # total paid is col N (index 13)
                t = _num(row[13]) if len(row) > 13 else None
                income_total_calc = t or sum(v for v in monthly_income if v)
                continue

            monthly = [_num(row[i]) if len(row) > i else None for i in range(1, 13)]
            total_paid   = _num(row[13]) if len(row) > 13 else None
            monthly_debt = _num(row[14]) if len(row) > 14 else None
            annual_debt  = _num(row[15]) if len(row) > 15 else None

            # skip header rows that have text (not numbers) in the totals column
            raw13 = row[13] if len(row) > 13 else None
            if isinstance(raw13, str) and _num(raw13) is None and str(raw13).strip():
                continue

            # skip metadata rows (collection rate decimal stored in monthly_debt)
            if all(v is None for v in monthly) and total_paid is None:
                if monthly_debt and 0 < monthly_debt < 1:
                    collection_rate = monthly_debt
                    continue

            # fallback: compute total_paid from monthly values if formula cache missing
            if total_paid is None:
                total_paid = sum(v for v in monthly if v) or 0.0

            tenants.append({
                'name': name,
                'monthly': monthly,
                'monthly_comments': _inc_comments.get(name, [None]*12),
                'total_paid': total_paid or 0.0,
                'monthly_debt': monthly_debt or 0.0,
                'annual_debt': annual_debt or 0.0,
            })

        # if income totals row had empty formula cache, recompute from tenants
        if not income_total_calc and tenants:
            income_total_calc = sum(t['total_paid'] for t in tenants) + _other_income
            for i in range(12):
                monthly_income[i] = sum(t['monthly'][i] or 0 for t in tenants)

        # collection rate from dedicated formula cell
        if collection_rate is None:
            for row in rows:
                if not row or len(row) < 15: continue
                v = _num(row[14])
                if v and 0 < v < 1 and _clean(row[0]) in (None, 'None', ''):
                    collection_rate = v
                    break

    # ── Expenses ─────────────────────────────────────────────────────────────
    ws_exp = _sheet(wb, 'ריכוז הוצאות 2026')
    expenses = []
    monthly_expenses = [0.0] * 12
    expense_total_actual = 0.0
    expense_categories = {}

    if ws_exp:
        rows_exp = list(ws_exp.iter_rows(values_only=True))
        current_cat = 'הוצאות צפויות'

        # Pass 1: build expense list + category totals + monthly expenses
        for row in rows_exp[1:]:
            if not row: continue
            # stop at second section (its header has text 'שנתי' in col[2])
            _c2r = row[2] if len(row) > 2 else None
            if isinstance(_c2r, str) and 'שנתי' in _c2r:
                break
            cat_cell = _clean(row[0])
            name_e   = _clean(row[1]) if len(row) > 1 else None

            if cat_cell and cat_cell not in ('None', 'סה"כ'):
                current_cat = cat_cell

            if not name_e or 'סה"כ' in (name_e or ''): continue

            monthly_vals  = [_num(row[5+i]) if len(row) > 5+i else None for i in range(12)]
            annual_est    = _num(row[4])  if len(row) > 4  else None
            annual_actual = _num(row[17]) if len(row) > 17 else None

            if current_cat not in expense_categories:
                expense_categories[current_cat] = 0.0
            for i, v in enumerate(monthly_vals):
                if v:
                    expense_categories[current_cat] += v
                    monthly_expenses[i] += v

            expenses.append({
                'category':      current_cat,
                'name':          name_e,
                'annual_est':    annual_est,
                'annual_actual': annual_actual,
                'monthly':       monthly_vals,
            })

        # Pass 2: find actual annual total — last סה"כ row col[2] > 10000
        # Structure: row[0]=None, row[1]='סה"כ', row[2]=19163.14
        for row in reversed(rows_exp):
            if not row or len(row) < 3: continue
            c1 = _clean(row[1]) if len(row) > 1 else None
            c2 = _num(row[2])   if len(row) > 2 else None
            if c1 == 'סה"כ' and c2 and c2 > 10000:
                expense_total_actual = c2
                break
        # fallback: row[0]='סה"כ', row[1]=value > 10000
        if not expense_total_actual:
            for row in reversed(rows_exp):
                if not row: continue
                c0 = _clean(row[0])
                c1 = _num(row[1]) if len(row) > 1 else None
                if c0 == 'סה"כ' and c1 and c1 > 10000:
                    expense_total_actual = c1
                    break

        # Pass 2b: read annual actuals from second section (col[2] has totals there)
        # Second section starts when col[2] is the text header 'הוצאה שנתית'
        _second_actuals = {}  # expense_name → annual_actual
        _in_second = False
        for row in rows_exp:
            if not row: continue
            c2_raw = row[2] if len(row) > 2 else None
            if isinstance(c2_raw, str) and 'שנתי' in c2_raw:
                _in_second = True
                continue
            if not _in_second: continue
            c1s = _clean(row[1]) if len(row) > 1 else None
            c2s = _num(row[2]) if len(row) > 2 else None
            if c1s and c1s != 'סה"כ' and c2s is not None:
                _second_actuals[c1s] = c2s

        # Update expense entries annual_actual from second section
        for _e in expenses:
            if _e['annual_actual'] is None and _e['name']:
                for _sn, _sv in _second_actuals.items():
                    if _e['name'] in _sn or _sn in _e['name']:
                        _e['annual_actual'] = _sv
                        break

        # Build named pie categories from second section actuals
        _PIE_MAP = [
            ('בזק',     ['בזק']),
            ('מעליות',  ['מעליות']),
            ('חשמל',    ['מונה חשמל', 'חשמל']),
            ('בנק',     ['עמלות בנק']),
            ('ביטוח',   ['ביטוח']),
            ('חריגות',  ['בדיקה', 'לא צפוי', 'חוב', 'ניקיון', 'אילן']),
        ]
        _named_cats = {}
        for _sn, _sv in _second_actuals.items():
            if not _sv or _sv <= 0: continue
            for _cname, _kws in _PIE_MAP:
                if any(_kw in _sn for _kw in _kws):
                    _named_cats[_cname] = _named_cats.get(_cname, 0) + _sv
                    break

        if _named_cats:
            expense_categories = _named_cats

    # ── Transactions ─────────────────────────────────────────────────────────
    ws_tr = _sheet(wb, 'תנועות חשבון')
    transactions = []
    latest_balance = None

    if ws_tr:
        rows_tr = list(ws_tr.iter_rows(values_only=True))
        for row in rows_tr[2:]:  # skip header + blank
            date_v = _clean(row[0]) if row else None
            if not date_v: continue
            name_v   = _clean(row[1]) if len(row) > 1 else ''
            action_v = _clean(row[2]) if len(row) > 2 else ''
            debit_v  = _num(row[3])  if len(row) > 3 else None
            credit_v = _num(row[4])  if len(row) > 4 else None
            bal_v    = _clean(row[5]) if len(row) > 5 else ''
            purp_v   = _clean(row[6]) if len(row) > 6 else ''
            if bal_v and latest_balance is None:
                latest_balance = _num(bal_v)
            transactions.append({
                'date':    date_v,
                'name':    name_v or '',
                'action':  action_v or '',
                'debit':   debit_v,
                'credit':  credit_v,
                'balance': bal_v or '',
                'purpose': purp_v or '',
            })

    # ── Budget ───────────────────────────────────────────────────────────────
    # 2026 annual budget plan (hardcoded — column J has live formulas not cached in file)
    _BUDGET_2026 = [
        ('בדיקת מעליות דו שנתי',  1412,  ['בדיקת מעליות'], None),
        ('עלות חודשית 2 מעליות',  17160, ['קונה'],          None),
        ('מעיינות העמקים',         400,   ['מעיינות'],       None),
        ('חשמל מדרגות',           2700,  ['מונה'],          None),
        ('ביטוח מבנה',            2700,  ['ביטוח'],         None),
        ('נקיון',                 9600,  ['ניקיון'],        None),
        ('נקיון גג',              1200,  ['גג'],             None),
        ('בזק מעליות',            600,   ['בזק'],           None),
        ('עמלות בנק',             240,   ['עמלות בנק'],     None),
        ('אחר - לא מתוכנן',       8000,  None,              'הוצאות לא צפויות'),
    ]
    budget = []
    budget_total = 35772
    for _nm, _tot, _kws, _cat in _BUDGET_2026:
        _act = 0.0
        for _e in expenses:
            if _cat:
                if _e.get('category') == _cat:
                    _act += _e.get('annual_actual') or 0
            elif _kws:
                if any(kw in (_e.get('name') or '') for kw in _kws):
                    _act += _e.get('annual_actual') or 0
        budget.append({'activity': _nm, 'total': _tot, 'actual': round(_act, 2)})
    budget_actual = expense_total_actual or sum(b['actual'] for b in budget)

    # ── Contacts ─────────────────────────────────────────────────────────────
    ws_con = _sheet(wb, 'פרטי תקשורת דיירים')
    contacts = []
    if ws_con:
        for row in list(ws_con.iter_rows(values_only=True))[1:]:
            name_c = _clean(row[0]) if row else None
            if not name_c: continue
            bld = row[1] if len(row) > 1 else None
            apt = row[2] if len(row) > 2 else None
            ph  = _clean(row[3]) if len(row) > 3 else None
            em  = _clean(row[4]) if len(row) > 4 else None
            contacts.append({'name': name_c, 'building': bld, 'apt': apt, 'phone': ph, 'email': em})

    # ── DATA sheet — pre-computed expense categories ─────────────────────────
    data_expense_cats = {}
    ws_data = _sheet(wb, 'DATA')
    if ws_data:
        rows_dat = list(ws_data.iter_rows(values_only=True))
        in_exp = False
        for row in rows_dat:
            if not row: continue
            c0 = _clean(row[0])
            c1 = _num(row[1]) if len(row) > 1 else None
            if c0 == 'קטגוריה': in_exp = True; continue
            if not in_exp: continue
            if c0 and c0 != 'סה"כ' and c1 is not None and c1 > 0:
                data_expense_cats[c0] = c1

    # ── Optional sheets ───────────────────────────────────────────────────────
    building_info = {}
    ws_bi = _sheet(wb, 'פרטי בניין')
    if ws_bi:
        for row in list(ws_bi.iter_rows(values_only=True))[1:]:
            if row and _clean(row[0]) and len(row) > 1:
                building_info[_clean(row[0])] = _clean(row[1]) or ''

    transaction_mapping = []
    ws_tm = _sheet(wb, 'מיפוי תנועות')
    if ws_tm:
        for row in list(ws_tm.iter_rows(values_only=True))[1:]:
            if row and _clean(row[0]):
                transaction_mapping.append({
                    'keyword':  _clean(row[0]),
                    'type':     _clean(row[1]) if len(row) > 1 else None,
                    'name':     _clean(row[2]) if len(row) > 2 else None,
                    'apt':      _clean(row[3]) if len(row) > 3 else None,
                    'day_from': _num(row[4])   if len(row) > 4 else None,
                    'day_to':   _num(row[5])   if len(row) > 5 else None,
                })

    wb.close()

    # ── Derive balance ────────────────────────────────────────────────────────
    balance = latest_balance
    if balance is None and balance_str:
        balance = _num(balance_str)
    if balance is None:
        balance = 0.0

    # ── Income total (always from ריכוז הכנסות) ──────────────────────────────
    income_total = income_total_calc or sum(v for v in monthly_income if v)

    # ── Expense total (actual annual from ריכוז הוצאות summary) ─────────────
    expense_total = expense_total_actual or sum(v for v in monthly_expenses if v)

    # ── Collection rate ───────────────────────────────────────────────────────
    if collection_rate is None and tenants:
        # Identify rate per tenant: corner (170) if min payment ≤175, else standard (210)
        total_exp_annual = 0.0
        total_paid_sum   = 0.0
        for t in tenants:
            paid_months = [v for v in t['monthly'] if v]
            if paid_months:
                rate = 170 if min(paid_months) <= 175 else 210
            else:
                rate = 210  # unpaying tenant — assume standard
            total_exp_annual += rate * 12
            total_paid_sum   += t['total_paid']
        collection_rate = total_paid_sum / total_exp_annual if total_exp_annual > 0 else 0.0

    # expense_categories may already be named_cats from Pass 2b; DATA sheet is fallback only
    final_expense_cats = expense_categories if expense_categories else data_expense_cats

    # Bar chart fallback: monthly formula cells are corrupted → distribute annual total
    # evenly across months that have income (gives a proportional view)
    if expense_total_actual > 0 and sum(monthly_expenses) == 0:
        active = [i for i in range(12) if monthly_income[i] > 0]
        base = active if active else list(range(12))
        pm = expense_total_actual / len(base)
        for i in base:
            monthly_expenses[i] = pm

    return {
        'balance':              balance,
        'income_total':         income_total,
        'expense_total':        expense_total,
        'collection_rate':      collection_rate or 0.0,
        'monthly_income':       monthly_income,
        'monthly_expenses':     monthly_expenses,
        'tenants':              tenants,
        'contacts':             contacts,
        'expenses':             expenses,
        'expense_categories':   final_expense_cats,
        'transactions':         transactions,
        'budget':               budget,
        'budget_total':         budget_total,
        'budget_actual':        budget_actual,
        'building_info':        building_info,
        'transaction_mapping':  transaction_mapping,
    }

# ─────────────────────────────────────────────────────────────────────────────
# GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────────────────────
import re as _re

def _sheet_to_csv_url(url):
    if not url or 'PASTE' in url: return None
    # Published URL: /d/e/2PACX-.../pubhtml  →  /d/e/2PACX-.../pub?output=csv&gid=X
    pub_m = _re.search(r'/d/e/([a-zA-Z0-9-_]+)/pub', url)
    if pub_m:
        pub_id = pub_m.group(1)
        gid_m = _re.search(r'gid=(\d+)', url)
        gid = gid_m.group(1) if gid_m else '0'
        return f'https://docs.google.com/spreadsheets/d/e/{pub_id}/pub?output=csv&gid={gid}'
    # Regular sheet URL: /d/SHEET_ID/
    m = _re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if not m: return None
    sid = m.group(1)
    gid_m = _re.search(r'gid=(\d+)', url)
    gid = gid_m.group(1) if gid_m else '0'
    return f'https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&gid={gid}'

def fetch_issues(issues_url, admin_url):
    import urllib3; urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    issues = []
    try:
        url = _sheet_to_csv_url(issues_url)
        if not url: return []
        r = requests.get(url, timeout=10, verify=False)
        r.encoding = 'utf-8'
        reader = csv.reader(io.StringIO(r.text))
        rows = list(reader)
        if not rows: return []
        for i, row in enumerate(rows[1:], start=1):
            if len(row) < 2 or not row[0].strip(): continue
            # Calls response sheet columns (0-indexed): 0 Timestamp | 1 שם מלא
            # 2 מספר בית | 3 מספר דירה | 4 תיאור התקלה | 5 מיקום התקלה | 6 דחיפות
            # 7 העלאת תמונות/וידאו | 8 הערות | 9 סטטוס (J) | 10 פעיל (K) | 11 תאריך עדכון סטטוס (L)
            _active = row[10].strip() if len(row) > 10 else ''
            if _active == 'לא פעיל': continue
            _building = row[2] if len(row) > 2 else ''
            _apt      = row[3] if len(row) > 3 else ''
            _apt_full = f'בניין {_building} דירה {_apt}'.strip() if _building or _apt else ''
            _status      = row[9].strip()  if len(row) > 9  and row[9].strip()  else 'פתוח'
            _update_date = row[11].strip() if len(row) > 11 and row[11].strip() else ''
            _images_raw  = row[7].strip()  if len(row) > 7  and row[7].strip()  else ''
            _images      = [u.strip() for u in _images_raw.split(',') if u.strip()]
            issues.append({
                'id':          i,
                'date':        row[0] if len(row) > 0 else '',
                'name':        row[1] if len(row) > 1 else '',
                'apt':         _apt_full,
                'location':    row[5] if len(row) > 5 else '',
                'urgency':     row[6] if len(row) > 6 else '',
                'desc':        row[4] if len(row) > 4 else '',
                'images':      _images,
                'status':      _status,
                'update_date': _update_date,
                'notes':       row[8] if len(row) > 8 else '',
            })
    except Exception as e:
        log.warning(f'Could not fetch issues sheet: {e}')
        return []

    # Merge admin responses
    try:
        url2 = _sheet_to_csv_url(admin_url)
        if url2:
            r2 = requests.get(url2, timeout=10, verify=False)
            r2.encoding = 'utf-8'
            reader2 = csv.reader(io.StringIO(r2.text))
            for row in list(reader2)[1:]:
                if len(row) < 2: continue
                try:
                    issue_id = int(row[0])
                    status   = row[1] if len(row) > 1 else ''
                    response = row[2] if len(row) > 2 else ''
                    for iss in issues:
                        if iss['id'] == issue_id:
                            if status: iss['status'] = status
                            if response: iss['response'] = response
                except: pass
    except Exception as e:
        log.warning(f'Could not fetch admin sheet: {e}')

    return issues

# Tolerant of both the dropdown's own values (כן / גבוהה) and raw values that can end up
# here from pasting old Excel data directly (TRUE/FALSE, "1-גבוהה" style prefixes).
def _ann_is_active(v):
    return v.strip().upper() in ('כן', 'TRUE', '1')

def _ann_priority(v):
    if 'גבוהה' in v: return 1
    if 'בינונית' in v: return 2
    if 'נמוכה' in v: return 3
    return 2

def _ann_norm_date(v):
    m = _re.match(r'^(\d{4})-(\d{2})-(\d{2})', v)
    return f'{m.group(3)}/{m.group(2)}/{m.group(1)}' if m else v.split(' ')[0]

# Parses whichever date format ended up in this field (Form's own M/D/Y locale,
# or Y-M-D from a pasted Excel value) so we can tell if it's a future date.
def _ann_parse_date(v):
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(v.split(' ')[0], fmt).date()
        except ValueError:
            continue
    return None

def fetch_charges(charges_url, payments_url):
    """Reads the current one-time charge + per-tenant payments from the admin-managed
    Google Sheet (replaces the old config.ini [one_time_charge] + inferred-payment heuristic).
    Charges sheet columns (0-indexed): 0 charge_id | 1 name | 2 amount | 3 date | 4 active | 5 description
    Payments sheet columns (0-indexed): 0 charge_id | 1 tenant_name | 2 amount_paid | 3 updated_at
    Returns (charge_dict, payments_dict). charge_dict is {} if no charge is currently active."""
    import urllib3; urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    charge = {}
    payments = {}
    try:
        curl = _sheet_to_csv_url(charges_url)
        if not curl: return {}, {}
        r = requests.get(curl, timeout=10, verify=False)
        r.encoding = 'utf-8'
        rows = list(csv.reader(io.StringIO(r.text)))
        for row in rows[1:]:
            if len(row) < 5 or not row[0].strip(): continue
            if (row[4] or '').strip() != 'כן': continue
            try:
                amount = float(row[2])
            except (ValueError, IndexError):
                continue
            date_s = row[3].strip() if len(row) > 3 else ''
            if date_s:
                try:
                    if datetime.now() < datetime.strptime(date_s, '%d/%m/%Y'): continue
                except ValueError:
                    pass
            if amount <= 0: continue
            charge = {
                'id':          row[0].strip(),
                'name':        row[1].strip() if len(row) > 1 else 'חיוב מיוחד',
                'amount':      amount,
                'description': row[5].strip() if len(row) > 5 else '',
            }
            break  # first active+ready row wins
        if charge:
            purl = _sheet_to_csv_url(payments_url)
            if purl:
                r2 = requests.get(purl, timeout=10, verify=False)
                r2.encoding = 'utf-8'
                prows = list(csv.reader(io.StringIO(r2.text)))
                for row in prows[1:]:
                    if len(row) < 3 or not row[1].strip(): continue
                    if row[0].strip() != charge['id']: continue
                    try:
                        payments[row[1].strip()] = float(row[2])
                    except ValueError:
                        continue
    except Exception as e:
        log.warning(f'Could not fetch charges sheet: {e}')
        return {}, {}
    return charge, payments

def fetch_announcements(announcements_url):
    import urllib3; urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    anns = []
    try:
        url = _sheet_to_csv_url(announcements_url)
        if not url: return []
        r = requests.get(url, timeout=10, verify=False)
        r.encoding = 'utf-8'
        rows = list(csv.reader(io.StringIO(r.text)))
        if not rows: return []
        # Response sheet columns (0-indexed): 0 Timestamp | 1 Email Address (unused)
        # 2 תאריך | 3 כותרת | 4 תוכן | 5 קטגוריה | 6 עדיפות | 7 פעיל
        today = datetime.now().date()
        for i, row in enumerate(rows[1:], start=1):
            title = row[3].strip() if len(row) > 3 else ''
            if not title: continue
            active = row[7] if len(row) > 7 else ''
            if not _ann_is_active(active): continue
            date_val = row[2].strip() if len(row) > 2 else ''
            parsed_date = _ann_parse_date(date_val) if date_val else None
            if parsed_date and parsed_date > today: continue  # scheduled for the future
            anns.append({
                'id':       i,
                'date':     _ann_norm_date(row[2]) if len(row) > 2 else '',
                'title':    title,
                'content':  row[4] if len(row) > 4 else '',
                'category': (row[5].strip() if len(row) > 5 and row[5].strip() else 'מידע'),
                'priority': _ann_priority(row[6] if len(row) > 6 else ''),
            })
    except Exception as e:
        log.warning(f'Could not fetch announcements sheet: {e}')
        return []
    return anns

# ─────────────────────────────────────────────────────────────────────────────
# SVG CHARTS
# ─────────────────────────────────────────────────────────────────────────────
def svg_bar_chart(monthly_income, monthly_expenses):
    W, H = 600, 380
    PL, PR, PT, PB = 72, 14, 32, 52
    cw = W - PL - PR
    ch = H - PT - PB

    all_vals = [v for v in monthly_income + monthly_expenses if v]
    inc_vals = [v for v in monthly_income if v and v > 0]
    if all_vals:
        raw_max = max(all_vals)
        # Don't let a single expense spike collapse all other bars:
        # cap scale at 2× max-income so normal months stay visible
        if inc_vals and raw_max > max(inc_vals) * 2.5:
            raw_max = max(inc_vals) * 2
        step = 1000 if raw_max <= 7000 else 2500 if raw_max <= 15000 else 5000
        max_v = math.ceil(raw_max / (5 * step)) * (5 * step)
    else:
        max_v = 5000

    def yp(v): return PT + ch - (v / max_v * ch)
    def hp(v): return v / max_v * ch

    bw = (cw / 12) * 0.38
    gap = (cw / 12) * 0.04

    lines = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" '
             f'style="width:100%;font-family:Segoe UI,Arial;direction:rtl">']

    # Grid lines + Y-axis labels
    for i in range(6):
        yv = max_v * i / 5
        y  = yp(yv)
        lines.append(f'<line x1="{PL}" y1="{y:.1f}" x2="{W-PR}" y2="{y:.1f}" '
                     f'stroke="#dde3ed" stroke-width="1" stroke-dasharray="3,3"/>')
        lines.append(f'<text x="{PL-5}" y="{y+4:.1f}" text-anchor="end" direction="ltr" '
                     f'font-size="10" fill="#94a3b8">₪{int(yv):,}</text>')

    # Bars + column value labels
    for i in range(12):
        xc = PL + (i + 0.5) * (cw / 12)
        xi = xc - bw - gap / 2
        xe = xc + gap / 2
        inc = monthly_income[i] or 0
        exp = monthly_expenses[i] or 0

        if inc > 0:
            bar_y = yp(inc)
            lines.append(f'<rect x="{xi:.1f}" y="{bar_y:.1f}" width="{bw:.1f}" '
                         f'height="{hp(inc):.1f}" fill="#22c55e" rx="2" opacity="0.88"/>')
            lbl = f'{int(round(inc)):,}'
            lines.append(f'<text x="{xi+bw/2:.1f}" y="{bar_y-4:.1f}" text-anchor="middle" '
                         f'font-size="10" font-weight="700" fill="#15803d">{lbl}</text>')

        if exp > 0:
            exp_draw = min(exp, max_v)  # cap to scale; label still shows real value
            bar_y = yp(exp_draw)
            lines.append(f'<rect x="{xe:.1f}" y="{bar_y:.1f}" width="{bw:.1f}" '
                         f'height="{hp(exp_draw):.1f}" fill="#f87171" rx="2" opacity="0.88"/>')
            lbl = f'{int(round(exp)):,}' + ('*' if exp > max_v else '')
            lines.append(f'<text x="{xe+bw/2:.1f}" y="{bar_y-4:.1f}" text-anchor="middle" '
                         f'font-size="10" font-weight="700" fill="#b91c1c">{lbl}</text>')

        lines.append(f'<text x="{xc:.1f}" y="{PT+ch+14}" text-anchor="middle" '
                     f'font-size="9" fill="#64748b">{MONTHS_SHORT[i]}</text>')

    # Legend — explicit LTR: [rect][gap][text], no RTL ambiguity
    ly = PT + ch + 32
    leg_x = W / 2 - 85
    lines += [
        f'<rect x="{leg_x:.0f}" y="{ly}" width="12" height="12" fill="#22c55e" rx="2"/>',
        f'<text x="{leg_x+17:.0f}" y="{ly+10}" text-anchor="start" direction="ltr" font-size="11" fill="#475569">הכנסות</text>',
        f'<rect x="{leg_x+90:.0f}" y="{ly}" width="12" height="12" fill="#f87171" rx="2"/>',
        f'<text x="{leg_x+107:.0f}" y="{ly+10}" text-anchor="start" direction="ltr" font-size="11" fill="#475569">הוצאות</text>',
    ]
    lines.append('</svg>')
    return '\n'.join(lines)


def svg_3d_pie_chart(expense_categories):
    """3D pie — Excel Style-8, outside callout labels, no legend.
    viewBox 640x608 (608/640=0.95=380/600*3/2) — equal height with bar at 3fr/2fr.
    Three layers: bottom face -> side face -> top face (painter's algorithm).
    Small bottom slices staggered to prevent overlap."""
    cats = {k: v for k, v in expense_categories.items() if v and v > 0}
    if not cats: return ''
    total = sum(cats.values())
    if total == 0: return ''

    COLORS = ['#1B7A8A','#ED7D31','#70AD47','#4BACC6','#C2247A','#1F3864','#C00000']

    W, H   = 640, 608
    cx, cy = 320, 190
    rx, ry = 160, 70
    depth  = 60

    def darken(hex_c, f=0.68):
        h = hex_c.lstrip('#')
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return f'#{int(r*f):02x}{int(g*f):02x}{int(b*f):02x}'

    def darken2(hex_c, f=0.52):
        h = hex_c.lstrip('#')
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return f'#{int(r*f):02x}{int(g*f):02x}{int(b*f):02x}'

    def tp(a): return cx + rx*math.cos(a), cy + ry*math.sin(a)
    def bp(a): x, y = tp(a); return x, y + depth

    lines = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" '
             f'style="width:100%;font-family:Segoe UI,Arial Hebrew,Arial">']

    slices = [(n, v, COLORS[i % len(COLORS)]) for i, (n, v) in enumerate(cats.items())]
    angle = -math.pi / 2
    sd = []
    for n, v, c in slices:
        da = v / total * 2 * math.pi
        ea = angle + da
        sd.append({'n': n, 'v': v, 'c': c, 'a1': angle, 'a2': ea,
                   'mid': angle + da / 2, 'da': da, 'pct': v / total * 100})
        angle = ea

    sorted_sd = sorted(sd, key=lambda s: math.sin(s['mid']))

    # Layer 1 — Bottom faces (painter's algorithm)
    for s in sorted_sd:
        a1, a2, da = s['a1'], s['a2'], s['da']
        lg = 1 if da > math.pi else 0
        x1b, y1b = bp(a1); x2b, y2b = bp(a2)
        d = (f'M {cx},{cy+depth} L {x1b:.1f},{y1b:.1f} '
             f'A {rx},{ry} 0 {lg} 1 {x2b:.1f},{y2b:.1f} Z')
        lines.append(f'<path d="{d}" fill="{darken2(s["c"])}" stroke="white" stroke-width="0.4"/>')

    # Layer 2 — Side faces (painter's algorithm)
    for s in sorted_sd:
        a1, a2, da = s['a1'], s['a2'], s['da']
        lg = 1 if da > math.pi else 0
        x1t, y1t = tp(a1); x2t, y2t = tp(a2)
        x1b, y1b = bp(a1); x2b, y2b = bp(a2)
        d = (f'M {x1t:.1f},{y1t:.1f} A {rx},{ry} 0 {lg} 1 {x2t:.1f},{y2t:.1f} '
             f'L {x2b:.1f},{y2b:.1f} A {rx},{ry} 0 {lg} 0 {x1b:.1f},{y1b:.1f} Z')
        lines.append(f'<path d="{d}" fill="{darken(s["c"])}" stroke="white" stroke-width="0.5"/>')

    # Layer 3 — Top faces
    for s in sd:
        a1, a2, da = s['a1'], s['a2'], s['da']
        lg = 1 if da > math.pi else 0
        x1, y1 = tp(a1); x2, y2 = tp(a2)
        d = f'M {cx},{cy} L {x1:.1f},{y1:.1f} A {rx},{ry} 0 {lg} 1 {x2:.1f},{y2:.1f} Z'
        lines.append(f'<path d="{d}" fill="{s["c"]}" stroke="white" stroke-width="1.5"/>')

    # ── Callout label positions ──────────────────────────────────────────────
    small_bottom_idx = 0
    pts = []
    for s in sd:
        mid = s['mid']
        cm, sm = math.cos(mid), math.sin(mid)
        pct = s['pct']

        # stagger small bottom slices: alternate line lengths to prevent overlap
        stagger = 0
        if pct < 4.0 and sm > 0.5:
            stagger = 70 if small_bottom_idx % 2 == 0 else 0
            small_bottom_idx += 1

        ex = cx + rx * cm
        ey = cy + ry * sm + (depth * 0.55 if sm > 0 else 0)

        if abs(cm) < 0.05:                    # truly vertical
            lx = cx + rx * 0.15 * cm
            ly = cy + (ry + depth + 75 + stagger) * sm
        elif abs(cm) < 0.20:                  # slightly off-vertical — use standard side formula
            if cm >= 0:
                lx = cx + (rx + 85) * cm
                ly = (cy - ry - 30) if sm < 0 else cy + (ry + depth + 55 + stagger) * sm
            else:
                extra = 80 + max(0, (4 - pct) * 14)
                lx = cx + (rx + 65 + extra) * cm
                if sm < 0:
                    ly = cy - ry - 30
                else:
                    ly = cy + (ry + depth + 55 + max(0, (4 - pct) * 18) + stagger) * sm
        elif cm >= 0:                          # right side
            lx = cx + (rx + 85) * cm
            ly = (cy - ry - 30) if sm < 0 else cy + (ry + depth + 55 + stagger) * sm
        else:                                  # left side
            extra = 80 + max(0, (4 - pct) * 14)
            lx = cx + (rx + 65 + extra) * cm
            if sm < 0:
                ly = cy - ry - 30
            else:
                ly = cy + (ry + depth + 55 + max(0, (4 - pct) * 18) + stagger) * sm

        lx = max(65, min(W - 65, lx))
        ly = max(18, min(H - 18, ly))
        pts.append({'s': s, 'ex': ex, 'ey': ey, 'lx': lx, 'ly': ly, 'cm': cm})

    # Collision avoidance — 12 passes, 24 px minimum gap
    for _ in range(12):
        for i in range(len(pts)):
            for j in range(i + 1, len(pts)):
                p1, p2 = pts[i], pts[j]
                if abs(p1['lx'] - p2['lx']) < 120 and abs(p1['ly'] - p2['ly']) < 24:
                    mid_y = (p1['ly'] + p2['ly']) / 2
                    p1['ly'] = max(18, min(H - 18, mid_y - 13))
                    p2['ly'] = max(18, min(H - 18, mid_y + 13))

    for p in pts:
        s = p['s']
        lx, ly, ex, ey = p['lx'], p['ly'], p['ex'], p['ey']

        lines.append(f'<line x1="{ex:.1f}" y1="{ey:.1f}" x2="{lx:.1f}" y2="{ly:.1f}" '
                     f'stroke="#94a3b8" stroke-width="1.0"/>')

        label = f'{he(s["n"])}, {int(round(s["v"])):,}'
        lines.append(f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
                     f'dominant-baseline="middle" '
                     f'font-size="14" font-weight="600" fill="#1e293b">{label}</text>')

    lines.append('</svg>')
    return '\n'.join(lines)

# ─────────────────────────────────────────────────────────────────────────────
# HTML HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt_ils(v, show_zero=True):
    if v is None: return '—'
    if v == 0 and not show_zero: return '—'
    return f'₪{v:,.0f}'

# Matches names across two independently-maintained lists (the charges/payments Sheet's
# free-typed tenant name vs. Excel's full household name, e.g. "אורן אלקיים" vs
# "אורן ואורלי אלקיים") without requiring one to be a literal substring of the other.
# Splits on whitespace, strips the Hebrew "and" prefix (ו) from each word, and matches
# if any word is shared — so a single first name or surname reliably links to the full
# household name.
def _name_match(a, b):
    def toks(s):
        return {w[1:] if w.startswith('ו') and len(w) > 1 else w for w in str(s).split()}
    return bool(toks(a) & toks(b))

def status_class(v, green, orange):
    if v >= green: return 'kpi-green'
    if v >= orange: return 'kpi-orange'
    return 'kpi-red'

_OC = ' onclick="dotTap(this,event)" ontouchend="dotTap(this,event)"'

def month_dot(val, rate, comment=None, approved_amount=None):
    """Return colored dot HTML for a monthly payment cell."""
    extra = f' — {comment}' if comment else ''
    if approved_amount is not None:
        if val and val >= approved_amount * 0.95:
            return f'<span class="dot dot-approved" title="שולם ₪{val:,.0f} — מאושר ועד{extra}"{_OC}>●</span>'
        elif val and val > 0:
            return f'<span class="dot dot-approved-partial" title="שולם חלקי ₪{val:,.0f} — מאושר ועד{extra}"{_OC}>◑</span>'
        else:
            return f'<span class="dot dot-approved-empty" title="טרם שולם — אושר ועד ₪{approved_amount:,.0f}{extra}"{_OC}>○</span>'
    if val is None or val == 0:
        return f'<span class="dot dot-empty" title="לא שולם{extra}"{_OC}>○</span>'
    if val >= rate * 0.95:
        return f'<span class="dot dot-paid" title="שולם ₪{val:,.0f}{extra}"{_OC}>●</span>'
    return f'<span class="dot dot-partial" title="שולם חלקי ₪{val:,.0f}{extra}"{_OC}>◑</span>'

# ─────────────────────────────────────────────────────────────────────────────
# HTML GENERATOR
# ─────────────────────────────────────────────────────────────────────────────
CSS = """
:root{
  --bg:#f0f4f8;--surface:#fff;--surface2:#f5f7fa;--border:#dde3ed;
  --text:#1e293b;--muted:#64748b;--primary:#1a3a5c;--accent:#2563eb;
  --green:#15803d;--orange:#b45309;--red:#b91c1c;
  --green-bg:#dcfce7;--orange-bg:#fef3c7;--red-bg:#fee2e2;--blue-bg:#dbeafe;
  --radius:10px;--font:'Segoe UI','Arial Hebrew',Arial,sans-serif;
}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#0d1521;--surface:#162032;--surface2:#1a2840;--border:#2d3f55;
  --text:#dde8f5;--muted:#7a92ab;--primary:#93c5fd;--accent:#60a5fa;
  --green-bg:#052e16;--orange-bg:#27180a;--red-bg:#2d0b0b;--blue-bg:#0c1e3d;
}}
:root[data-theme="dark"]{
  --bg:#0d1521;--surface:#162032;--surface2:#1a2840;--border:#2d3f55;
  --text:#dde8f5;--muted:#7a92ab;--primary:#93c5fd;--accent:#60a5fa;
  --green-bg:#052e16;--orange-bg:#27180a;--red-bg:#2d0b0b;--blue-bg:#0c1e3d;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth;height:100%}
body{background:var(--bg);color:var(--text);font-family:var(--font);font-size:14px;line-height:1.6;direction:rtl;min-height:100%;overflow-x:hidden}
.section{scroll-margin-top:92px}
a{color:var(--accent)}
/* Header */
.hdr{background:var(--primary);color:#fff;padding:14px 20px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;position:-webkit-sticky;position:sticky;top:0;z-index:100;width:100%}
.hdr-title{font-size:17px;font-weight:700;flex:1}
.hdr-meta{font-size:11px;opacity:.7}
.hdr-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:5px 12px;border-radius:20px;cursor:pointer;font-size:12px}
/* Nav */
.subnav{background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px;display:flex;gap:0;overflow-x:auto;position:-webkit-sticky;position:sticky;top:48px;z-index:99;width:100%}
.subnav a{padding:8px 14px;font-size:12px;color:var(--muted);text-decoration:none;white-space:nowrap;border-bottom:2px solid transparent}
.subnav a:hover{color:var(--text);border-bottom-color:var(--accent)}
/* Layout */
.wrap{max-width:1100px;margin:0 auto;padding:20px 16px 60px}
.section{margin-bottom:28px}
.section-title{font-size:15px;font-weight:700;color:var(--primary);padding-bottom:8px;border-bottom:2px solid var(--border);margin-bottom:14px;display:flex;align-items:center;gap:8px}
.section-title small{font-size:11px;font-weight:400;color:var(--muted);margin-right:auto}
/* KPI Tiles */
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:28px}
@media(max-width:700px){.kpis{grid-template-columns:repeat(2,1fr)}}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px;border-top:3px solid var(--border)}
.kpi.kpi-green{border-top-color:#22c55e}
.kpi.kpi-orange{border-top-color:#f59e0b}
.kpi.kpi-red{border-top-color:#ef4444}
.kpi.kpi-blue{border-top-color:#3b82f6}
.kpi-label{font-size:11px;color:var(--muted);margin-bottom:6px}
.kpi-val{font-size:20px;font-weight:700;color:var(--text);font-variant-numeric:tabular-nums}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:4px}
.kpi-bar{height:4px;background:var(--border);border-radius:2px;margin-top:8px;overflow:hidden}
.kpi-bar-fill{height:100%;border-radius:2px;transition:width .5s}
/* Announcements */
.ann-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px}
.ann-card{background:var(--surface);border:1px solid var(--border);border-right:4px solid;border-radius:var(--radius);padding:14px}
.ann-card.urgent{border-right-color:#ef4444}
.ann-card.maintenance{border-right-color:#f59e0b}
.ann-card.financial{border-right-color:#8b5cf6}
.ann-card.meeting{border-right-color:#22c55e}
.ann-card.safety{border-right-color:#1e40af}
.ann-card.info{border-right-color:#3b82f6}
.ann-cat{font-size:10px;font-weight:600;text-transform:uppercase;margin-bottom:4px;color:var(--muted)}
.ann-title{font-weight:600;margin-bottom:4px}
.ann-content{font-size:12px;color:var(--muted)}
.ann-date{font-size:10px;color:var(--muted);margin-top:6px}
/* Charts */
.chart-row{display:grid;grid-template-columns:3fr 2fr;gap:20px;align-items:start}
@media(max-width:800px){.chart-row{grid-template-columns:1fr}}
.chart-box{background:var(--surface);border-radius:var(--radius);padding:10px;display:flex;flex-direction:column}
.chart-box svg{width:100%;height:auto;display:block}
.chart-title{font-size:13px;font-weight:700;color:var(--primary);text-align:center;margin-bottom:8px}
/* Tenant table */
.tbl-wrap{overflow-x:auto}
.tr-scroll{max-height:420px;overflow-y:auto;border-radius:6px}
.tr-scroll thead th{position:sticky;top:0;z-index:1;background:var(--surface2)}
.th-sort{cursor:pointer;user-select:none}
.th-sort:hover{color:var(--primary)}
.tr-search{width:220px;padding:5px 10px;border:1px solid var(--border);border-radius:6px;background:var(--surface);color:var(--text);font-size:13px;direction:rtl;margin-bottom:8px}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:var(--surface2);padding:7px 8px;font-size:11px;color:var(--muted);font-weight:600;border-bottom:2px solid var(--border);text-align:right;white-space:nowrap}
td{padding:6px 8px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--surface2)}
.dot{font-size:14px;cursor:pointer;display:inline-block;padding:6px 4px;margin:-6px -4px;touch-action:manipulation}
#dot-popup{position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:8px 12px;font-size:13px;color:var(--text);max-width:260px;box-shadow:0 4px 16px rgba(0,0,0,.18);pointer-events:none;opacity:0;transition:opacity .15s;line-height:1.5;direction:rtl;text-align:right}
#dot-popup.show{opacity:1}
.dot-paid{color:#22c55e}.dot-partial{color:#f59e0b}.dot-empty{color:#e2e8f0}
.dot-approved{color:#2563eb}.dot-approved-partial{color:#60a5fa}.dot-approved-empty{color:#93c5fd}
/* Status badges */
.badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
.badge-red{background:var(--red-bg);color:var(--red)}
.badge-orange{background:var(--orange-bg);color:var(--orange)}
.badge-green{background:var(--green-bg);color:var(--green)}
.badge-credit{background:#d1fae5;color:#065f46}
.badge-blue{background:var(--blue-bg);color:var(--accent)}
/* Progress bars */
.prog-item{margin-bottom:10px}
.prog-label{display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px}
.prog-bar{height:8px;background:var(--border);border-radius:4px;overflow:hidden}
.prog-fill{height:100%;border-radius:4px}
/* Issues */
.issues-table td:first-child{font-variant-numeric:tabular-nums;color:var(--muted);font-size:11px}
/* Contact grid */
.contact-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px}
.contact-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px}
.contact-name{font-weight:600;margin-bottom:4px}
.contact-detail{font-size:12px;color:var(--muted)}
/* Links section */
.links-grid{display:flex;flex-wrap:wrap;gap:12px;margin-top:8px}
.link-btn{display:inline-flex;align-items:center;gap:8px;padding:12px 20px;border-radius:var(--radius);font-size:14px;font-weight:600;text-decoration:none;transition:opacity .15s}
.link-btn:hover{opacity:.85}
.link-btn-red{background:#fee2e2;color:#b91c1c}
.link-btn-blue{background:#dbeafe;color:#1d4ed8}
/* FAB */
.fab{position:fixed;bottom:24px;left:20px;background:var(--accent);color:#fff;border:none;border-radius:50px;padding:12px 18px;font-size:14px;font-weight:600;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.2);z-index:100;text-decoration:none;display:flex;align-items:center;gap:6px}
.fab:hover{opacity:.9}
/* Alert banner */
.alert-banner{background:var(--red-bg);border:1px solid #fca5a5;border-radius:var(--radius);padding:12px 16px;margin-bottom:16px;font-size:13px;color:var(--red)}
"""

def generate_html(data, issues, anns, cfg, updated_at, charge=None, charge_payments=None):
    building_name = cfg.get('building','name', fallback='ועד בית')
    reserve_target = cfg_float(cfg, 'thresholds', 'reserve_target', 8000)
    col_green  = cfg_float(cfg, 'thresholds', 'collection_green',  85)
    col_orange = cfg_float(cfg, 'thresholds', 'collection_orange', 70)
    bal_warn   = cfg_float(cfg, 'thresholds', 'balance_warning', 2000)
    warn_debt    = cfg_float(cfg, 'thresholds', 'warning_debt',    100)
    crit_debt    = cfg_float(cfg, 'thresholds', 'critical_debt',   1000)
    min_credit   = cfg_float(cfg, 'thresholds', 'min_credit_display', 50)
    refresh_min= cfg_int(cfg, 'display', 'auto_refresh_min', 5)
    tenant_form= cfg.get('forms','tenant_form_url', fallback='')
    fault_form = cfg.get('forms','fault_form_url',  fallback='')
    admin_form = cfg.get('forms','admin_form_url',  fallback='')
    open_warn  = cfg_int(cfg, 'thresholds', 'open_issues_warn', 3)

    # Approved payments: one-time committee-approved reduced payment per tenant per month.
    # config key format: "name:month_number = amount"  (month 1=Jan … 12=Dec)
    _approved_payments = {}  # partial_name -> {month_idx_0based: amount}
    if cfg.has_section('approved_payments'):
        for name, val in cfg.items('approved_payments'):
            try:
                # format: name = month:amount  (ConfigParser treats ':' as delimiter so month goes in value)
                parts = val.strip().split(':')
                if len(parts) == 2:
                    ap_month = int(parts[0].strip()) - 1  # 0-based
                    ap_amt   = float(parts[1].strip())
                    _approved_payments.setdefault(name.strip(), {})[ap_month] = ap_amt
            except:
                pass

    # Approved notes: optional reason text for tooltip on the blue badge.
    _approved_notes = {}  # partial_name -> note string
    if cfg.has_section('approved_notes'):
        for name, note in cfg.items('approved_notes'):
            _approved_notes[name.strip()] = note.strip()

    _debt_carryover = {}  # partial_name -> carryover amount from prior year
    if cfg.has_section('debt_carryover'):
        for name, val in cfg.items('debt_carryover'):
            try: _debt_carryover[name.strip()] = float(val)
            except: pass

    # One-time special charge (e.g. roof repair) — same amount for every tenant, tracked
    # separately from regular dues (per-tenant dot + aggregate progress + its own debt component).
    # Managed via the admin GUI + Google Sheet now (see fetch_charges) — real per-tenant payment
    # amounts, not inferred from dues surplus. Empty dict when no charge is currently active.
    _otc = charge or {}
    _otc_payments = charge_payments or {}

    show_ann  = cfg_bool(cfg,'display','show_announcements')
    show_bud  = cfg_bool(cfg,'display','show_budget')
    show_exp  = cfg_bool(cfg,'display','show_expense_chart')
    show_res  = cfg_bool(cfg,'display','show_reserve_kpi')
    show_iss  = cfg_bool(cfg,'display','show_issues')
    show_tr   = cfg_bool(cfg,'display','show_transactions')

    bal    = data['balance']
    inc    = data['income_total']
    exp    = data['expense_total']
    mi     = data['monthly_income']
    me     = data['monthly_expenses']
    tenants= data['tenants']

    # Recompute collection rate: YTD expected, capped per tenant (advance payers don't inflate rate)
    _now_month   = datetime.now().month
    _std_rate    = cfg_int(cfg, 'rates', 'standard', 210)
    _cor_rate    = cfg_int(cfg, 'rates', 'corner',   170)
    _cor_names   = [n.strip() for n in cfg.get('rates','corner_tenants',fallback='').split(',') if n.strip()]
    _MONTH_HE    = ['','ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר']
    _col_month_name = _MONTH_HE[_now_month] if 1 <= _now_month <= 12 else ''
    if tenants:
        _exp_ytd  = 0.0
        _eff_paid = 0.0
        for t in tenants:
            _is_corner = any(cn in t['name'] for cn in _cor_names)
            _base = _cor_rate if _is_corner else _std_rate
            _appr = {}
            for _apn, _apm in _approved_payments.items():
                if _apn in t['name']:
                    _appr = _apm
                    break
            _t_ytd = sum(_appr.get(i, _base) for i in range(_now_month))
            _exp_ytd  += _t_ytd
            _eff_paid += min(t['total_paid'], _t_ytd)
        col_pct = round(_eff_paid / _exp_ytd * 100) if _exp_ytd else 0
    else:
        col_pct = round(data['collection_rate'] * 100)
        _col_month_name = ''
    trans  = data['transactions']
    cats   = data['expense_categories']
    binfo  = data['building_info']

    # open issues count
    open_count = sum(1 for i in issues if i.get('status','פתוח') in ('פתוח','בטיפול'))

    # ── KPI section ──────────────────────────────────────────────────────────
    bal_cls = 'kpi-red' if bal < bal_warn else 'kpi-green'
    col_cls = status_class(col_pct, col_green, col_orange)
    res_pct = min(bal / reserve_target * 100, 100) if reserve_target else 0
    res_cls = 'kpi-green' if res_pct >= 100 else ('kpi-orange' if res_pct >= 50 else 'kpi-red')
    iss_cls = 'kpi-red' if open_count >= open_warn else ('kpi-orange' if open_count > 0 else 'kpi-green')

    kpi_html = f"""
<div class="kpis">
  <div class="kpi {bal_cls}">
    <div class="kpi-label">יתרת חשבון</div>
    <div class="kpi-val">{fmt_ils(bal)}</div>
    <div class="kpi-sub">{'⚠ יתרה נמוכה' if bal < bal_warn else 'תקין'}</div>
  </div>
  <div class="kpi kpi-blue">
    <div class="kpi-label">הכנסות שנה</div>
    <div class="kpi-val">{fmt_ils(inc)}</div>
    <div class="kpi-sub">מתחילת 2026</div>
  </div>
  <div class="kpi {'kpi-green' if exp < inc else 'kpi-red'}">
    <div class="kpi-label">הוצאות שנה</div>
    <div class="kpi-val">{fmt_ils(exp)}</div>
    <div class="kpi-sub">מתחילת 2026</div>
  </div>
  <div class="kpi {col_cls}">
    <div class="kpi-label">אחוז גבייה</div>
    <div class="kpi-val">{col_pct:.0f}%</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(col_pct,100):.0f}%;background:{'#22c55e' if col_pct>=col_green else ('#f59e0b' if col_pct>=col_orange else '#ef4444')}"></div></div>
    <div class="kpi-sub">מתוך צפוי עד {_col_month_name} 2026</div>
  </div>"""

    if show_res:
        kpi_html += f"""
  <div class="kpi {res_cls}">
    <div class="kpi-label">קרן רזרבה</div>
    <div class="kpi-val">{fmt_ils(min(bal, reserve_target))}</div>
    <div class="kpi-sub">יעד: {fmt_ils(reserve_target)}</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{res_pct:.0f}%;background:{'#22c55e' if res_pct>=100 else '#f59e0b'}"></div></div>
  </div>"""

    kpi_html += '\n</div>'

    # ── Announcements (static pre-render + background JS refresh) ─────────────
    def _ann_cat_cls(cat):
        return ('urgent'      if 'דחוף'    in cat else
                'maintenance' if 'תחזוקה'  in cat else
                'financial'   if 'כספי'    in cat else
                'meeting'     if 'כינוסים' in cat else
                'safety'      if 'בטיחות'  in cat else 'info')

    def _ann_cards_html(items):
        cards = ''
        for a in sorted(items, key=lambda x: x.get('priority', 9)):
            cat = a.get('category') or 'מידע'
            cards += (f'<div class="ann-card {_ann_cat_cls(cat)}">'
                      f'<div class="ann-cat">{he(cat)}</div>'
                      f'<div class="ann-title">{he(a.get("title",""))}</div>'
                      f'<div class="ann-content">{he(a.get("content",""))}</div>'
                      f'<div class="ann-date">{he(a.get("date",""))}</div></div>')
        return cards

    ann_html = ''
    if show_ann:
        if anns:
            _ann_static = f'<div class="ann-grid">{_ann_cards_html(anns)}</div>'
        else:
            _ann_static = '<p style="color:var(--muted);font-size:13px;padding:8px 0">אין הודעות פעילות</p>'
        ann_html = f"""
<div class="section" id="announcements">
  <div class="section-title">📢 הודעות לדיירים</div>
  <div id="announcements-body">{_ann_static}</div>
</div>
<script>
(function(){{
  function esc(s){{return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}}
  function catCls(c){{return c.indexOf('דחוף')>-1?'urgent':(c.indexOf('תחזוקה')>-1?'maintenance':(c.indexOf('כספי')>-1?'financial':(c.indexOf('כינוסים')>-1?'meeting':(c.indexOf('בטיחות')>-1?'safety':'info'))));}}
  function buildHTML(anns){{
    if(!anns.length)return'<p style="color:var(--muted);font-size:13px;padding:8px 0">אין הודעות פעילות</p>';
    var sorted=anns.slice().sort(function(a,b){{return (a.priority||9)-(b.priority||9);}});
    var h='<div class="ann-grid">';
    for(var j=0;j<sorted.length;j++){{
      var a=sorted[j],cat=a.category||'מידע';
      h+='<div class="ann-card '+catCls(cat)+'"><div class="ann-cat">'+esc(cat)+'</div>'+
         '<div class="ann-title">'+esc(a.title)+'</div><div class="ann-content">'+esc(a.content)+'</div>'+
         '<div class="ann-date">'+esc(a.date)+'</div></div>';
    }}
    return h+'</div>';
  }}
  fetch('announcements.json?_='+Date.now()).then(function(r){{return r.json();}}).then(function(anns){{
    document.getElementById('announcements-body').innerHTML=buildHTML(anns);
  }}).catch(function(){{}});  // silent fail — static content stays
}})();
</script>"""

    # ── Charts ────────────────────────────────────────────────────────────────
    bar_svg   = svg_bar_chart(mi, me)
    pie3d_svg = svg_3d_pie_chart(cats) if show_exp else ''
    chart_html = f"""
<div class="section" id="charts">
  <div class="section-title">📊 הכנסות והוצאות 2026</div>
  <div class="chart-row">
    <div class="chart-box">
      <div class="chart-title">הכנסות מול הוצאות חודשי</div>
      {bar_svg}
    </div>
    {'<div class="chart-box"><div class="chart-title">הוצאות</div>' + pie3d_svg + '</div>' if pie3d_svg else ''}
  </div>
</div>"""

    # ── Budget tracker ────────────────────────────────────────────────────────
    bud_html = ''
    if show_bud and data['budget']:
        # Total summary bar — I14 vs J14 directly from Excel
        total_budget = data.get('budget_total') or 0
        total_actual = data.get('budget_actual') or 0
        tot_pct = min(total_actual / total_budget * 100, 100) if total_budget > 0 else 0
        tot_color = '#22c55e' if tot_pct < 80 else ('#f59e0b' if tot_pct < 100 else '#ef4444')
        over = total_actual > total_budget
        over_label = f' <span style="color:#ef4444;font-weight:700">({fmt_ils(total_actual - total_budget)} חריגה)</span>' if over else ''
        summary_bar = f"""
<div class="prog-item" style="margin-bottom:16px;padding-bottom:14px;border-bottom:2px solid var(--border)">
  <div class="prog-label" style="font-size:13px;font-weight:700">
    <span>סה"כ תקציב שנתי{over_label}</span>
    <span style="color:var(--muted)">{fmt_ils(total_actual)} / {fmt_ils(total_budget)}</span>
  </div>
  <div class="prog-bar" style="height:12px"><div class="prog-fill" style="width:{tot_pct:.0f}%;background:{tot_color}"></div></div>
</div>"""

        items = summary_bar
        for b in data['budget']:
            est    = b['total']  or 0
            actual = b['actual'] or 0
            if not est and not actual:
                continue
            pct = min(actual / est * 100, 100) if est > 0 else 0
            color = '#22c55e' if pct < 80 else ('#f59e0b' if pct < 100 else '#ef4444')
            pct_label = f'{pct:.0f}%'
            items += f"""
<div class="prog-item">
  <div class="prog-label">
    <span>{he(b['activity'])}</span>
    <span style="color:var(--muted)">{fmt_ils(actual)} / {fmt_ils(est)} &nbsp;<span style="color:{color};font-weight:600">{pct_label}</span></span>
  </div>
  <div class="prog-bar"><div class="prog-fill" style="width:{pct:.0f}%;background:{color}"></div></div>
</div>"""
        bud_html = f"""
<div class="section" id="budget">
  <div class="section-title">📋 מעקב תקציב</div>
  {items}
</div>"""

    # ── Tenant payments ───────────────────────────────────────────────────────
    # Recompute monthly_debt: formula cells corrupted → all read as ₪0.
    # Formula: sum of monthly expected amounts − total paid.
    # For months with an approved_payment entry use that amount; else use standard/corner rate.
    _now_month = datetime.now().month
    for t in tenants:
        _is_corner = any(cn in t['name'] for cn in _cor_names)
        _base_rate = _cor_rate if _is_corner else _std_rate

        # Find per-month approved payments for this tenant
        _appr = {}
        _appr_note = ''
        for ap_name, ap_months in _approved_payments.items():
            if ap_name in t['name']:
                _appr = ap_months
                break
        for an_name, an_note in _approved_notes.items():
            if an_name in t['name']:
                _appr_note = an_note
                break
        t['_approved_months'] = _appr   # {month_idx: amount}
        t['_approved_note']   = _appr_note

        # Expected total = sum of required amount per elapsed month
        _expected_normal = sum(
            _appr[i] if i in _appr else _base_rate
            for i in range(_now_month)
        )
        # Regular dues debt/credit — unaffected by the one-time charge, which is tracked
        # and paid via a completely separate pool (real per-tenant amounts from the admin sheet).
        _raw = _expected_normal - t['total_paid']
        _regular_debt   = max(0.0,  _raw)
        _regular_credit = max(0.0, -_raw)

        _otc_amount = _otc['amount'] if _otc else 0.0
        _otc_paid_raw = next((v for k, v in _otc_payments.items() if _name_match(k, t['name'])), 0.0)
        _otc_paid = min(max(0.0, _otc_paid_raw), _otc_amount) if _otc_amount else 0.0
        _otc_debt = max(0.0, _otc_amount - _otc_paid) if _otc_amount else 0.0
        t['_otc_paid'] = _otc_paid

        # Combined single debt badge = regular dues owed + any unpaid portion of the charge
        t['monthly_debt']   = _regular_debt + _otc_debt
        t['monthly_credit'] = _regular_credit

        # Carryover debt from prior year (stored in config.ini [debt_carryover])
        _carry = next((v for k, v in _debt_carryover.items() if k in t['name']), 0.0)
        t['_carryover'] = _carry

    # Build contact lookup for apartment info
    contact_map = {c['name'].strip(): c for c in data['contacts']}

    header_months = ''.join(f'<th>{m}</th>' for m in MONTHS_SHORT)
    rows_html = ''
    for t in tenants:
        name = t['name'].strip()
        debt   = t['monthly_debt']
        credit = t.get('monthly_credit', 0.0)
        carry  = t.get('_carryover', 0.0)
        annual_debt = t['annual_debt']

        if credit >= min_credit:
            debt_cls   = 'badge-credit'
            debt_label = f'מקדמה {fmt_ils(credit)}'
        elif carry > 0:
            total_debt_display = debt + carry
            debt_cls   = ('badge-red'    if total_debt_display > crit_debt else
                          'badge-orange' if total_debt_display >= warn_debt else 'badge-green')
            debt_label = f'{fmt_ils(debt)} + {fmt_ils(carry)} העברה'
        elif debt > 0:
            debt_cls   = ('badge-red'    if debt > crit_debt else
                          'badge-orange' if debt >= warn_debt else 'badge-green')
            debt_label = fmt_ils(debt)
        else:
            debt_cls   = 'badge-green'
            debt_label = 'ללא חוב'

        # find contact for this tenant
        con = contact_map.get(name, {})
        for cname, cdata in contact_map.items():
            if name in cname or cname in name:
                con = cdata
                break
        apt_info = f'בניין {con.get("building","")} דירה {con.get("apt","")}' if con.get('building') else ''

        # determine rate for this tenant
        rate = 170 if any(v == 170 for v in t['monthly'] if v) else 210
        _cmnts       = t.get('monthly_comments') or [None]*12
        _appr_months = t.get('_approved_months', {})
        dots = ''.join(
            month_dot(t['monthly'][i], rate,
                      _cmnts[i] if i < len(_cmnts) else None,
                      _appr_months.get(i))
            for i in range(12)
        )

        _otc_cell = ''
        if _otc:
            _op = t.get('_otc_paid', 0.0)
            if _op >= _otc['amount'] - 0.01:
                _dc, _dl = '#22c55e', 'שולם'
            elif _op > 0:
                _dc, _dl = '#f59e0b', 'חלקי'
            else:
                _dc, _dl = '#cbd5e1', 'לא שולם'
            _otc_title = f'{he(_otc["name"])}: {fmt_ils(_op)} / {fmt_ils(_otc["amount"])} — {_dl}'
            _otc_cell = f'<td><span style="color:{_dc};font-size:18px" title="{_otc_title}">●</span></td>'

        rows_html += f"""
<tr>
  <td><strong>{he(name)}</strong><br><small style="color:var(--muted)">{he(apt_info)}</small></td>
  {dots.replace('<span', '<td><span').replace('</span>', '</span></td>')}
  {_otc_cell}
  <td style="font-variant-numeric:tabular-nums">{fmt_ils(t['total_paid'])}</td>
  <td><span class="badge {debt_cls}" title="{he(debt_label)}">{he(debt_label)}</span></td>
</tr>"""

    # ── Totals row ────────────────────────────────────────────────────────────
    _total_paid    = sum(t['total_paid'] for t in tenants)
    _total_debt26  = sum(t['monthly_debt'] for t in tenants)
    _total_carry   = sum(t.get('_carryover', 0.0) for t in tenants)
    _grand_debt    = _total_debt26 + _total_carry
    _carry_note    = f' (כולל {fmt_ils(_total_carry)} העברה מ-2025)' if _total_carry > 0 else ''
    _total_cls     = 'badge-red' if _grand_debt > crit_debt else ('badge-orange' if _grand_debt >= warn_debt else 'badge-green')

    _otc_header = ''
    _otc_totals_cell = ''
    _otc_progress_html = ''
    if _otc:
        _otc_header = f'<th>{he(_otc["name"])}</th>'
        _otc_target    = _otc['amount'] * len(tenants)
        _otc_collected = sum(t.get('_otc_paid', 0.0) for t in tenants)
        _otc_pct = min(_otc_collected / _otc_target * 100, 100) if _otc_target else 0
        _otc_color = '#22c55e' if _otc_pct >= 100 else ('#f59e0b' if _otc_pct >= 50 else '#ef4444')
        _otc_totals_cell = f'<td><span class="badge" style="background:{_otc_color}22;color:{_otc_color}">{_otc_pct:.0f}%</span></td>'
        _otc_desc = f' — {he(_otc["description"])}' if _otc.get('description') else ''
        _otc_progress_html = f"""
<div class="prog-item" style="margin-top:14px">
  <div class="prog-label">
    <span>📢 {he(_otc['name'])}{_otc_desc}</span>
    <span style="color:var(--muted)">{fmt_ils(_otc_collected)} / {fmt_ils(_otc_target)} &nbsp;<span style="color:{_otc_color};font-weight:600">{_otc_pct:.0f}%</span></span>
  </div>
  <div class="prog-bar"><div class="prog-fill" style="width:{_otc_pct:.0f}%;background:{_otc_color}"></div></div>
</div>"""

    totals_row = f"""
<tr style="font-weight:600;border-top:2px solid var(--border);background:var(--surface2)">
  <td>סה"כ</td>
  {'<td></td>' * 12}
  {_otc_totals_cell}
  <td style="font-variant-numeric:tabular-nums">{fmt_ils(_total_paid)}</td>
  <td><span class="badge {_total_cls}" title="חוב 2026: {fmt_ils(_total_debt26)}{_carry_note}">{fmt_ils(_grand_debt)}{' ↑' if _total_carry > 0 else ''}</span></td>
</tr>"""

    tenant_html = f"""
<div class="section" id="tenants">
  <div class="section-title">🏠 תשלומי דיירים 2026</div>
  <div class="tbl-wrap">
  <table>
    <thead>
      <tr>
        <th>שם דייר</th>
        {header_months}
        {_otc_header}
        <th>סה"כ שולם</th>
        <th>חוב חודשי</th>
      </tr>
    </thead>
    <tbody>{rows_html}{totals_row}</tbody>
  </table>
  </div>
  <div id="otc-progress">{_otc_progress_html}</div>
  <div style="font-size:12px;color:var(--muted);margin-top:10px;display:flex;flex-wrap:wrap;gap:18px;align-items:center">
    <span style="display:flex;gap:12px;align-items:center">
      <span><span style="color:#22c55e;font-size:15px">●</span> שולם</span>
      <span><span style="color:#f59e0b;font-size:15px">◑</span> חלקי</span>
      <span><span style="color:#cbd5e1;font-size:15px">○</span> לא שולם</span>
      <span><span style="color:#2563eb;font-size:15px">●</span> זיכוי ועד</span>
    </span>
    <span style="color:#94a3b8">|</span>
    <span style="display:flex;gap:6px;align-items:center">
      <span style="font-size:11px;color:var(--muted)">חוב:</span>
      <span class="badge badge-credit" style="font-size:10px">מקדמה</span>
      <span class="badge badge-green" style="font-size:10px">ללא חוב</span>
      <span class="badge badge-orange" style="font-size:10px">1–3 חודשים</span>
      <span class="badge badge-red" style="font-size:10px">4+ חודשים</span>
    </span>
  </div>
</div>"""

    # ── Transactions ──────────────────────────────────────────────────────────
    tr_html = ''
    if show_tr:
        tr_rows = ''
        for t in trans:
            amount  = ''
            amt_val = 0
            color   = ''
            if t['credit']:
                amount  = f'+{fmt_ils(t["credit"])}'
                color   = '#22c55e'
                amt_val = t['credit']
            elif t['debit']:
                amount  = f'-{fmt_ils(t["debit"])}'
                color   = '#ef4444'
                amt_val = -(t['debit'])
            try:
                bal_val = float(str(t['balance']).replace(',','').replace('₪','').replace(' ',''))
            except Exception:
                bal_val = 0
            clr_style = f';color:{color}' if color else ''
            tr_rows += f"""
<tr>
  <td data-val="{he(t['date'])}">{he(t['date'])}</td>
  <td>{he(t['name'])}</td>
  <td>{he(t['action'])}</td>
  <td data-val="{amt_val}" style="font-variant-numeric:tabular-nums{clr_style}">{amount}</td>
  <td data-val="{bal_val}" style="font-variant-numeric:tabular-nums">{he(t['balance'])}</td>
  <td style="font-size:11px;color:var(--muted)">{he(t['purpose'])}</td>
</tr>"""
        tr_html = f"""
<div class="section" id="transactions">
  <div class="section-title">🏦 תנועות בנק
    <small>{len(trans)} תנועות</small>
  </div>
  <input type="text" class="tr-search" id="tr-search" placeholder="חיפוש לפי שם / תיאור..." oninput="filterTrans()">
  <div class="tbl-wrap tr-scroll">
  <table id="tr-table">
    <thead><tr>
      <th class="th-sort" onclick="sortTrans(0)">תאריך ↕</th>
      <th>שם</th>
      <th>פעולה</th>
      <th class="th-sort" onclick="sortTrans(3)">סכום ↕</th>
      <th class="th-sort" onclick="sortTrans(4)">יתרה ↕</th>
      <th>עבור</th>
    </tr></thead>
    <tbody id="tr-body">{tr_rows}</tbody>
  </table>
  </div>
</div>"""

    # ── Issues (static pre-render + background JS refresh) ───────────────────
    def _urg_cls(u):
        if u in ('גבוהה','דחוף','חריגה'): return 'badge-red'
        if u in ('בינונית',): return 'badge-orange'
        if u in ('נמוכה',): return 'badge-green'
        return 'badge-blue'

    def _img_id(u):
        m = _re.search(r'/d/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)', u)
        return (m.group(1) or m.group(2)) if m else None

    def _iss_img_html(urls):
        if not urls: return ''
        out = ''
        for u in urls[:4]:
            fid = _img_id(u)
            if fid:
                thumb = f'https://drive.google.com/thumbnail?id={fid}&sz=w200'
                out += f'<a href="{he(u)}" target="_blank"><img src="{thumb}" loading="lazy" style="width:44px;height:44px;object-fit:cover;border-radius:6px;margin:2px" alt="תמונה"></a>'
            else:
                out += f'<a href="{he(u)}" target="_blank" style="font-size:11px">📎 קובץ</a>'
        return out

    iss_html = ''
    if show_iss:
        _fl = he(fault_form or '#')
        # Build static content (always shown immediately)
        if not issues:
            _iss_static = f'<p style="color:var(--muted);font-size:13px;padding:8px 0">אין קריאות פתוחות · <a href="{_fl}" target="_blank">הגש קריאה חדשה</a></p>'
        else:
            _iss_rows = ''
            for iss in sorted(issues, key=lambda x: x.get('id',0), reverse=True)[:20]:
                stat = iss.get('status','פתוח')
                sc = 'badge-green' if stat == 'סגור' else ('badge-orange' if stat == 'בטיפול' else 'badge-red')
                urg = iss.get('urgency','')
                _upd = he(iss.get("update_date",""))
                _nts = he(iss.get("notes",""))
                _img = _iss_img_html(iss.get('images') or [])
                _iss_rows += (f'<tr><td>#{iss["id"]}</td><td>{he(iss.get("date","")[:10])}</td><td>{he(iss.get("name",""))}</td>'
                              f'<td dir="ltr" style="text-align:right">{he(iss.get("apt",""))}</td>'
                              f'<td>{he(iss.get("location",""))}</td>'
                              f'<td><span class="badge {_urg_cls(urg)}">{he(urg)}</span></td>'
                              f'<td>{he(iss.get("desc",""))}</td>'
                              f'<td><span class="badge {sc}">{he(stat)}</span></td>'
                              f'<td style="font-size:11px;color:var(--muted)">{_upd}</td>'
                              f'<td style="font-size:11px;color:var(--muted)">{_nts}</td>'
                              f'<td>{_img}</td></tr>')
            _iss_static = f'<div class="tbl-wrap"><table class="issues-table"><thead><tr><th>#</th><th>תאריך</th><th>שם</th><th>דירה</th><th>מיקום</th><th>דחיפות</th><th>תיאור</th><th>סטטוס</th><th>תאריך עדכון</th><th>הערות</th><th>תמונה</th></tr></thead><tbody>{_iss_rows}</tbody></table></div>'
        iss_html = f"""
<div class="section" id="issues">
  <div class="section-title">🔧 קריאות שירות
    <small id="issues-count">פתוחות: {open_count}</small>
    <a href="{_fl}" target="_blank" class="hdr-btn" style="font-size:12px;padding:4px 10px">+ הגש קריאה</a>
  </div>
  <div id="issues-body">{_iss_static}</div>
</div>
<script>
(function(){{
  var FF='{_fl}';
  function esc(s){{return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}}
  function urgCls(u){{return u==='גבוהה'||u==='דחוף'||u==='חריגה'?'badge-red':(u==='בינונית'?'badge-orange':(u==='נמוכה'?'badge-green':'badge-blue'));}}
  function imgId(u){{var m=u.match(/\\/d\\/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)/);return m?(m[1]||m[2]):null;}}
  function imgHtml(urls){{
    if(!urls||!urls.length)return'';
    var out='';
    for(var k=0;k<Math.min(urls.length,4);k++){{
      var u=urls[k],fid=imgId(u);
      if(fid){{out+='<a href="'+esc(u)+'" target="_blank"><img src="https://drive.google.com/thumbnail?id='+fid+'&sz=w200" loading="lazy" style="width:44px;height:44px;object-fit:cover;border-radius:6px;margin:2px" alt="תמונה"></a>';}}
      else{{out+='<a href="'+esc(u)+'" target="_blank" style="font-size:11px">📎 קובץ</a>';}}
    }}
    return out;
  }}
  function buildHTML(issues){{
    var open=issues.filter(function(x){{return x.status!=='סגור';}}).length;
    if(!issues.length)return['<p style="color:var(--muted);font-size:13px;padding:8px 0">אין קריאות פתוחות · <a href="'+FF+'" target="_blank">הגש קריאה חדשה</a></p>',0];
    var h='<div class="tbl-wrap"><table class="issues-table"><thead><tr><th>#</th><th>תאריך</th><th>שם</th><th>דירה</th><th>מיקום</th><th>דחיפות</th><th>תיאור</th><th>סטטוס</th><th>תאריך עדכון</th><th>הערות</th><th>תמונה</th></tr></thead><tbody>';
    var sorted=issues.slice().sort(function(a,b){{return b.id-a.id;}});
    for(var j=0;j<Math.min(sorted.length,20);j++){{
      var is=sorted[j],sc=is.status==='סגור'?'badge-green':(is.status==='בטיפול'?'badge-orange':'badge-red');
      h+='<tr><td>#'+is.id+'</td><td>'+esc(is.date)+'</td><td>'+esc(is.name)+'</td><td dir="ltr" style="text-align:right">'+esc(is.apt)+'</td>'+
         '<td>'+esc(is.location)+'</td><td><span class="badge '+urgCls(is.urgency)+'">'+esc(is.urgency)+'</span></td>'+
         '<td>'+esc(is.desc)+'</td><td><span class="badge '+sc+'">'+esc(is.status)+'</span></td>'+
         '<td style="font-size:11px;color:var(--muted)">'+esc(is.update_date)+'</td>'+
         '<td style="font-size:11px;color:var(--muted)">'+esc(is.notes)+'</td>'+
         '<td>'+imgHtml(is.images)+'</td></tr>';
    }}
    return[h+'</tbody></table></div>',open];
  }}
  // Fetch from issues.json (same domain — no CORS, always fast)
  fetch('issues.json?_='+Date.now()).then(function(r){{return r.json();}}).then(function(issues){{
    var res=buildHTML(issues);
    document.getElementById('issues-body').innerHTML=res[0];
    document.getElementById('issues-count').textContent='פתוחות: '+res[1];
  }}).catch(function(){{}});  // silent fail — static content stays

  // Live-refresh the one-time-charge aggregate progress bar (Sheet-only data, no
  // Excel dependency) — the per-tenant dot column inside the table above still
  // needs a local generator run, since that merges in Excel dues data.
  fetch('charges.json?_='+Date.now()).then(function(r){{return r.json();}}).then(function(c){{
    var el=document.getElementById('otc-progress');
    if(!el) return;
    if(!c||!c.id){{el.innerHTML='';return;}}
    var color=c.pct>=100?'#22c55e':(c.pct>=50?'#f59e0b':'#ef4444');
    var desc=c.description?' — '+esc(c.description):'';
    el.innerHTML='<div class="prog-item" style="margin-top:14px">'+
      '<div class="prog-label"><span>📢 '+esc(c.name)+desc+'</span>'+
      '<span style="color:var(--muted)">₪'+Math.round(c.collected).toLocaleString()+' / ₪'+Math.round(c.target).toLocaleString()+' &nbsp;<span style="color:'+color+';font-weight:600">'+c.pct+'%</span></span></div>'+
      '<div class="prog-bar"><div class="prog-fill" style="width:'+c.pct+'%;background:'+color+'"></div></div></div>';
  }}).catch(function(){{}});  // silent fail — static content stays

  // Live-refresh per-tenant contact cards (כרטיס דייר) — fully Sheet-based now, no
  // Excel dependency, so this always reflects the admin panel's tenant-cards tab.
  fetch('contacts.json?_='+Date.now()).then(function(r){{return r.json();}}).then(function(list){{
    var el=document.getElementById('tenant-contacts-live');
    if(!el) return;
    function line(label,c){{
      if(!c||!c.name) return '';
      var rows='';
      if(c.phone) rows+='<div style="direction:ltr;text-align:right">📞 <a href="tel:'+esc(c.phone.replace(/[^0-9+]/g,''))+'">'+esc(c.phone)+'</a></div>';
      if(c.email) rows+='<div style="direction:ltr;text-align:right">✉️ <a href="mailto:'+esc(c.email)+'">'+esc(c.email)+'</a></div>';
      var head=label?(label+': '+esc(c.name)):esc(c.name);
      return '<div class="contact-detail"><strong>'+head+'</strong>'+rows+'</div>';
    }}
    el.innerHTML=list.map(function(t){{
      var h='<div class="contact-card"><div class="contact-name">'+esc(t.displayName||t.lastName)+'</div>'+
        '<div class="contact-detail" style="color:var(--muted)">בניין '+esc(t.building)+' דירה '+esc(t.apt)+'</div>';
      h+=line('',t.contact1);
      h+=line('',t.contact2);
      if(t.rented) h+=line('בעל הדירה',t.owner);
      return h+'</div>';
    }}).join('');
  }}).catch(function(){{}});  // silent fail — static content stays
}})();
</script>"""

    # ── Contacts ──────────────────────────────────────────────────────────────
    # Bank/payment info card stays Excel-derived (building-level, not per-tenant).
    # Per-tenant "כרטיס דייר" cards are rendered live from contacts.json (see JS below) —
    # migrated off the old Excel-based per-tenant contacts list, which required a local
    # generator run to reflect any change and could drift from the Sheet-based admin panel.
    bank_section = ''
    for k, v in binfo.items():
        if v:
            bank_section += f'<div class="contact-detail"><strong>{he(k)}:</strong> {he(v)}</div>'
    bank_card = f'<div class="contact-card"><div class="contact-name">💳 פרטי תשלום</div>{bank_section}</div>' if bank_section else ''
    con_html = f"""
<div class="section" id="contacts">
  <div class="section-title">📞 אנשי קשר</div>
  <div class="contact-grid">{bank_card}<div id="tenant-contacts-live" style="display:contents"></div></div>
</div>"""

    # ── Alert for unmatched transactions ──────────────────────────────────────
    alert_html = ''

    # ── Admin hidden link ─────────────────────────────────────────────────────
    admin_link = ''
    if admin_form and 'PASTE' not in admin_form:
        admin_link = f'<a href="{he(admin_form)}" target="_blank" style="color:inherit;text-decoration:none" title="כניסת ועד">🔒</a>'

    # ── Links section ─────────────────────────────────────────────────────────
    links_html = ''
    _link_items = []
    if fault_form  and 'PASTE' not in fault_form:
        _link_items.append(f'<a href="{he(fault_form)}"  target="_blank" class="link-btn link-btn-red">🔧 פתיחת קריאת שירות</a>')
    if tenant_form and 'PASTE' not in tenant_form:
        _link_items.append(f'<a href="{he(tenant_form)}" target="_blank" class="link-btn link-btn-blue">✏️ עדכון פרטים אישיים</a>')
    if _link_items:
        links_html = f"""
<div class="section" id="links">
  <div class="section-title">🔗 קישורים שימושיים</div>
  <div class="links-grid">{''.join(_link_items)}</div>
</div>"""

    # ── FAB button ────────────────────────────────────────────────────────────
    fab_html = ''
    if fault_form and 'PASTE' not in fault_form:
        fab_html = f'<a href="{he(fault_form)}" target="_blank" class="fab">🔧 דווח על תקלה</a>'

    # ── Subnav ────────────────────────────────────────────────────────────────
    nav_links = [
        ('#kpis','סיכום'),('#charts','גרפים'),('#tenants','תשלומים'),
        ('#transactions','תנועות'),('#issues','קריאות'),('#contacts','קשר'),
        ('#links','קישורים'),
    ]
    if anns: nav_links.insert(1, ('#announcements','הודעות'))
    if not _link_items: nav_links = [l for l in nav_links if l[0] != '#links']
    nav_html = ''.join(f'<a href="{href}">{label}</a>' for href, label in nav_links)

    # ── Assemble ──────────────────────────────────────────────────────────────
    refresh_seconds = refresh_min * 60
    html = f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="{refresh_seconds}">
<title>{he(building_name)} — דשבורד {cfg.get('display','default_year',fallback='2026')}</title>
<style>{CSS}</style>
</head>
<body>

<header class="hdr">
  <div class="hdr-title">{he(building_name)}</div>
  <div class="hdr-meta">עודכן: {he(updated_at)} {admin_link}</div>
  <button class="hdr-btn" onclick="toggleTheme()">◐</button>
</header>

<nav class="subnav">{nav_html}</nav>

{alert_html}

<div class="wrap">

<div id="kpis">{kpi_html}</div>

{ann_html}
{chart_html}
{bud_html}
{tenant_html}
{tr_html}
{iss_html}
{con_html}
{links_html}

</div>

{fab_html}

<div id="dot-popup"></div>
<script>
function dotTap(el,e){{
  e.preventDefault();e.stopPropagation();
  var popup=document.getElementById('dot-popup');
  var msg=el.getAttribute('title');
  if(!msg)return;
  popup.textContent=msg;
  var t=e.changedTouches&&e.changedTouches[0];
  var x=t?t.clientX:e.clientX, y=t?t.clientY:e.clientY;
  var vw=window.innerWidth;
  var pw=Math.min(260,vw-20);
  popup.style.maxWidth=pw+'px';
  var left=x-pw/2;
  var top=y-56;
  if(left<8)left=8;
  if(left+pw>vw-8)left=vw-pw-8;
  if(top<8)top=y+16;
  popup.style.left=left+'px';
  popup.style.top=top+'px';
  popup.classList.add('show');
  clearTimeout(window._dotHide);
  window._dotHide=setTimeout(function(){{popup.classList.remove('show');}},3500);
}}
document.addEventListener('click',function(e){{
  if(!e.target.closest('.dot'))document.getElementById('dot-popup').classList.remove('show');
}});
function toggleTheme(){{
  var r=document.documentElement;
  r.setAttribute('data-theme',r.getAttribute('data-theme')==='dark'?'light':'dark');
}}
function filterTrans(){{
  var q=document.getElementById('tr-search').value.trim().toLowerCase();
  document.querySelectorAll('#tr-body tr').forEach(function(r){{
    r.style.display=(!q||r.textContent.toLowerCase().includes(q))?'':'none';
  }});
}}
var _tc=-1,_ta=true;
function sortTrans(c){{
  var tb=document.getElementById('tr-body');
  if(!tb)return;
  var rows=Array.from(tb.querySelectorAll('tr'));
  if(_tc===c){{_ta=!_ta;}}else{{_tc=c;_ta=true;}}
  rows.sort(function(a,b){{
    var av=a.cells[c].getAttribute('data-val')||a.cells[c].textContent.trim();
    var bv=b.cells[c].getAttribute('data-val')||b.cells[c].textContent.trim();
    var an=parseFloat(av),bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn))return _ta?an-bn:bn-an;
    return _ta?av.localeCompare(bv,'he'):bv.localeCompare(av,'he');
  }});
  rows.forEach(function(r){{tb.appendChild(r);}});
}}
</script>
<script>
// Reports real content height to a parent window (e.g. the Admin app's דשבורד tab, which
// embeds this page in an iframe) so the iframe can be sized exactly to fit — otherwise the
// iframe gets its own internal scrollbar on top of the outer page's scrollbar ("page in
// page"). Harmless no-op when this page isn't embedded (postMessage just goes nowhere).
(function(){{
  // html{{height:100%}}/body{{min-height:100%}} make document.documentElement.scrollHeight
  // self-referential once embedded: the page stretches to fill whatever height the iframe
  // currently has, then reports THAT inflated size back, locking in the iframe's starting
  // height forever regardless of real content. Measuring the actual bottom-most content
  // element instead sidesteps that entirely.
  function postHeight(){{
    try{{
      var h=0;
      Array.prototype.forEach.call(document.body.children,function(el){{
        if(el.tagName==='SCRIPT'||el.tagName==='STYLE') return;
        var bottom=el.offsetTop+el.offsetHeight;
        if(bottom>h) h=bottom;
      }});
      parent.postMessage({{vaadDashboardHeight: h||document.documentElement.scrollHeight}}, '*');
    }}catch(e){{}}
  }}
  if (window.ResizeObserver) new ResizeObserver(postHeight).observe(document.documentElement);
  window.addEventListener('load', postHeight);
  postHeight();
}})();
</script>
</body>
</html>"""
    return html

# ─────────────────────────────────────────────────────────────────────────────
# GIT PUSH
# ─────────────────────────────────────────────────────────────────────────────
def git_push(html_path, repo_dir, auto_push):
    if not auto_push: return
    try:
        html_path = str(html_path)
        repo_dir  = str(repo_dir)
        # copy as index.html for GitHub Pages
        index_path = os.path.join(repo_dir, 'index.html')
        if html_path != index_path:
            shutil.copy2(html_path, index_path)
        subprocess.run(['git','-C',repo_dir,'add','index.html'], check=True, capture_output=True)
        msg = f'Update dashboard {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        result = subprocess.run(['git','-C',repo_dir,'commit','-m',msg], capture_output=True)
        if result.returncode == 0:
            subprocess.run(['git','-C',repo_dir,'push'], check=True, capture_output=True)
            log.info('Git push OK')
        else:
            log.info('Git: nothing to commit')
    except subprocess.CalledProcessError as e:
        log.error(f'Git push failed: {e}')
    except Exception as e:
        log.error(f'Git error: {e}')

# ─────────────────────────────────────────────────────────────────────────────
# BACKUP
# ─────────────────────────────────────────────────────────────────────────────
def backup_excel(excel_path, backup_folder):
    try:
        Path(backup_folder).mkdir(parents=True, exist_ok=True)
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = Path(excel_path).stem + f'_{ts}.xlsx'
        shutil.copy2(excel_path, Path(backup_folder) / name)
        # keep only last 10 backups
        backups = sorted(Path(backup_folder).glob('*.xlsx'))
        for old in backups[:-10]:
            old.unlink()
    except Exception as e:
        log.warning(f'Backup failed: {e}')

# ─────────────────────────────────────────────────────────────────────────────
# MAIN RUN
# ─────────────────────────────────────────────────────────────────────────────
def run_once():
    cfg = load_config()
    excel_path    = cfg.get('paths','excel_file')
    output_html   = cfg.get('paths','output_html')
    backup_folder = cfg.get('paths','backup_folder', fallback='')
    repo_dir      = cfg.get('paths','repo_dir', fallback=os.path.dirname(__file__))
    auto_push     = cfg_bool(cfg,'github','auto_push', False)
    issues_url    = cfg.get('google','issues_sheet_url', fallback='')
    admin_url     = cfg.get('google','admin_sheet_url',  fallback='')
    announcements_url = cfg.get('google','announcements_sheet_url', fallback='')
    charges_url   = cfg.get('google','charges_sheet_url', fallback='')
    charge_payments_url = cfg.get('google','charge_payments_sheet_url', fallback='')

    log.info('Reading Excel...')
    try:
        if backup_folder:
            backup_excel(excel_path, backup_folder)
        data = read_excel(excel_path)
    except Exception as e:
        log.error(f'Excel read failed: {e}')
        return

    log.info(f'Balance: {data["balance"]:,.0f}  Income: {data["income_total"]:,.0f}  '
             f'Expenses: {data["expense_total"]:,.0f}  Collection: {data["collection_rate"]*100:.0f}%  '
             f'Tenants: {len(data["tenants"])}')

    log.info('Fetching Google Sheets...')
    issues = fetch_issues(issues_url, admin_url)
    log.info(f'Issues: {len(issues)}')
    anns = fetch_announcements(announcements_url)
    log.info(f'Announcements: {len(anns)}')
    charge, charge_payments = fetch_charges(charges_url, charge_payments_url)
    log.info(f'Active charge: {charge.get("name") if charge else "none"}')

    updated_at = datetime.now().strftime('%d/%m/%Y %H:%M')
    log.info('Generating HTML...')
    html = generate_html(data, issues, anns, cfg, updated_at, charge, charge_payments)

    Path(output_html).parent.mkdir(parents=True, exist_ok=True)
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html)
    log.info(f'HTML saved → {output_html}  ({len(html):,} bytes)')

    git_push(output_html, repo_dir, auto_push)

# ─────────────────────────────────────────────────────────────────────────────
# WATCHDOG
# ─────────────────────────────────────────────────────────────────────────────
if HAS_WATCHDOG:
    class ExcelHandler(FileSystemEventHandler):
        def __init__(self, excel_path):
            self.excel_path = os.path.abspath(excel_path)
            self._last = 0

        def on_modified(self, event):
            if event.is_directory: return
            if os.path.abspath(event.src_path) != self.excel_path: return
            now = time.time()
            if now - self._last < 5: return  # debounce 5s
            self._last = now
            log.info(f'Excel modified — regenerating...')
            threading.Thread(target=run_once, daemon=True).start()

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    cfg = load_config()
    excel_path = cfg.get('paths','excel_file')
    refresh_min = cfg_int(cfg,'display','auto_refresh_min', 5)

    log.info('=== ועד בית Dashboard Generator ===')
    log.info(f'Excel: {excel_path}')
    log.info(f'Refresh interval: {refresh_min} min')

    # First run immediately
    run_once()

    # Watchdog for Excel file changes
    if HAS_WATCHDOG:
        observer = Observer()
        handler  = ExcelHandler(excel_path)
        watch_dir = str(Path(excel_path).parent)
        observer.schedule(handler, watch_dir, recursive=False)
        observer.start()
        log.info(f'Watchdog monitoring: {watch_dir}')
    else:
        log.warning('watchdog not installed — file monitoring disabled')

    # Scheduler for Google Sheets (every N minutes)
    if HAS_SCHEDULE:
        schedule.every(refresh_min).minutes.do(run_once)
        log.info(f'Scheduler: every {refresh_min} min')
        try:
            while True:
                schedule.run_pending()
                time.sleep(30)
        except KeyboardInterrupt:
            log.info('Stopped by user')
            if HAS_WATCHDOG:
                observer.stop()
                observer.join()
    else:
        log.warning('schedule not installed — running once only')


if __name__ == '__main__':
    main()
