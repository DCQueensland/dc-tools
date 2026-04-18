#!/usr/bin/env python3
"""Consolidate DC graduate rolls into two HubSpot-ready CSVs.

Sources:
  ~/Dropbox/2026/William/Database All/4 - Graduates/
    Updated Graduates for Farmer 2014-2019.xls
    Updated Graduates for Farmer 01092019 to 31082024.xls
    William Farmer Past Graduates.xlsx

Outputs:
  Master Graduates - All - <DATE>.csv           (one row per person)
  Master Graduates - Contactable - <DATE>.csv   (subset with email AND phone/mobile)
  Master Graduates - Summary - <DATE>.txt       (counts)

Rules:
  - Territory: QLD (4xxx), SA (5xxx), WA (6xxx). Drop all others.
  - Exclude Bradnam's Windows and Doors, National Glass/Aluminium, Triscol (separate campaign).
  - Dedupe by email, fallback to firstname+lastname+company.
"""
import re, csv, datetime, collections
from pathlib import Path
import xlrd, openpyxl

SRC = Path('/Users/williamfarmer/Dropbox/2026/William/Database All/4 - Graduates')
OUT = Path('/Users/williamfarmer/Dropbox/2026/William/Database All')
TODAY = datetime.date.today().isoformat()

EXCLUDE_COMPANY_SUBSTRINGS = [
    'bradnam', 'national glass', 'national aluminium', 'national aluminum', 'triscol',
]

PROGRAM_MAP = {
    'DCC': 'DCC', 'ADCC': 'Advanced DCC', 'UYLP': 'UYLP', 'LTM': 'LTM',
    'LTR': 'LTR', 'HIP': 'HIP', 'WRS': 'WRS', 'WGYH': 'WGYH', 'WGHY': 'WGYH',
    'SALES': 'Sales', 'HPT': 'High Performance Teams', 'WCCS': 'World Class Customer Service',
    'WCS': 'World Class Customer Service',
    'LEAD': 'Leadership', 'LEADERSHIP': 'Leadership',
    'PRES': 'HIP', 'PRESENTATIONS': 'HIP',
    'TEAMS': 'High Performance Teams',
    'CUSTOMER SERVICE': 'World Class Customer Service',
    'GEN NEXT': 'Gen Next', 'GENNEXT': 'Gen Next', 'DCCGN': 'Gen Next',
    'SB': 'Schools', 'SBDC': 'Schools', 'SCDCC': 'Schools', 'BSBDC': 'Schools',
    'ECHR': 'ECHR',
    'DYLP': 'UYLP',
    'DC': 'DCC',
    'DCCVIATRIS': 'DCC', 'TWDCC': 'DCC',
}

PROGCODE_RE = re.compile(r'^(\d{2})([A-Z]{2,4})([A-Z_]+?)(?:_([A-Za-z0-9]+))?(?:_?(\d+))?$')

def parse_programcode(pc: str):
    """Parse codes like 18QLDDCC02, 17QLDDCC_BRADNAMS_01, 22QLDSALES_BENESTAR, 15SBDC04."""
    if not pc: return {}
    pc = str(pc).strip()
    m = re.match(r'^(\d{2})', pc)
    year = None
    if m:
        yy = int(m.group(1))
        year = 2000 + yy if yy < 80 else 1900 + yy
        rest = pc[2:]
    else:
        rest = pc
    state = None
    for s in ('QLD', 'NSW', 'VIC', 'ACT', 'TAS', 'NT', 'WA', 'SA', 'SB'):
        if rest.startswith(s):
            state = s
            rest = rest[len(s):]
            break
    rest = rest.lstrip('_')
    # Program is the alphabetic prefix of rest, up to digit or underscore
    prog_m = re.match(r'^([A-Za-z]+)', rest)
    raw_prog = prog_m.group(1) if prog_m else ''
    remainder = rest[len(raw_prog):].lstrip('_')
    program = PROGRAM_MAP.get(raw_prog.upper(), raw_prog.upper())
    # Detect inhouse vs public: if remainder contains letters it's a client, else trailing digits = public class number
    inhouse_client = None
    if remainder:
        client_m = re.match(r'^([A-Za-z][A-Za-z0-9]*)', remainder)
        if client_m:
            inhouse_client = client_m.group(1)
    delivery = 'Inhouse' if inhouse_client else 'Public'
    return {
        'year': year, 'state_guess': state, 'program': program,
        'program_raw': raw_prog, 'delivery': delivery,
        'inhouse_client': inhouse_client, 'program_code': pc,
    }

def state_from_zip(z):
    if not z: return None
    s = str(z).strip()
    digits = ''.join(c for c in s if c.isdigit())
    if not digits: return None
    first = digits[0] if len(digits) >= 4 else digits[0]
    return {'2':'NSW','3':'VIC','4':'QLD','5':'SA','6':'WA','7':'TAS','0':'NT_ACT'}.get(first)

def normalise_state(s):
    if not s: return None
    s = str(s).strip().upper()
    mapping = {
        'QLD':'QLD','QUEENSLAND':'QLD','QLD ':'QLD',
        'NSW':'NSW','NEW SOUTH WALES':'NSW',
        'VIC':'VIC','VICTORIA':'VIC',
        'WA':'WA','WESTERN AUSTRALIA':'WA',
        'SA':'SA','SOUTH AUSTRALIA':'SA',
        'ACT':'ACT','TAS':'TAS','TASMANIA':'TAS','NT':'NT',
    }
    return mapping.get(s)

def excluded_company(c):
    if not c: return False
    cl = str(c).lower()
    return any(sub in cl for sub in EXCLUDE_COMPANY_SUBSTRINGS)

def norm_phone(p):
    if not p: return ''
    s = str(p).strip()
    if s in ('', 'nan', 'None', '0'): return ''
    return s

def norm_email(e):
    if not e: return ''
    s = str(e).strip().lower()
    return s if '@' in s else ''

def parse_xls_date(v, datemode):
    if not v: return None
    try:
        if isinstance(v, float):
            tup = xlrd.xldate_as_tuple(v, datemode)
            return datetime.date(tup[0], tup[1], tup[2])
    except Exception:
        pass
    if isinstance(v, str):
        for fmt in ('%Y-%m-%d','%d/%m/%Y','%m/%d/%Y','%d-%b-%y','%d/%m/%y'):
            try: return datetime.datetime.strptime(v.strip(), fmt).date()
            except Exception: pass
    return None

def parse_any_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    if isinstance(v, str):
        for fmt in ('%Y-%m-%d','%d/%m/%Y','%m/%d/%Y','%d-%b-%y','%d/%m/%y'):
            try: return datetime.datetime.strptime(v.strip(), fmt).date()
            except Exception: pass
    return None

records = []

# --- XLS files ---
for fname in ['Updated Graduates for Farmer 2014-2019.xls',
              'Updated Graduates for Farmer 01092019 to 31082024.xls']:
    wb = xlrd.open_workbook(SRC / fname)
    ws = wb.sheet_by_index(0)
    dm = wb.datemode
    for i in range(1, ws.nrows):
        row = ws.row_values(i)
        pc_info = parse_programcode(row[10])
        sess_date = parse_xls_date(row[11], dm)
        year = pc_info.get('year') or (sess_date.year if sess_date else None)
        zip_state = state_from_zip(row[6])
        prog_state = pc_info.get('state_guess')
        # SB (schools) rows: trust zipcode
        if prog_state == 'SB': prog_state = None
        state = zip_state or prog_state
        records.append({
            'first': str(row[1]).strip(),
            'last': str(row[2]).strip(),
            'nickname': str(row[3]).strip(),
            'title': str(row[4]).strip(),
            'company': str(row[5]).strip(),
            'zipcode': str(row[6]).strip() if row[6] else '',
            'phone': norm_phone(row[7]),
            'mobile': norm_phone(row[8]),
            'email': norm_email(row[9]),
            'program_code': row[10],
            'program': pc_info.get('program') or '',
            'delivery': pc_info.get('delivery') or '',
            'inhouse_client': pc_info.get('inhouse_client') or '',
            'year': year,
            'session_date': sess_date.isoformat() if sess_date else '',
            'state': state,
            'source': fname,
        })

# --- XLSX file ---
wb = openpyxl.load_workbook(SRC / 'William Farmer Past Graduates.xlsx', read_only=True, data_only=True)
ws = wb.active
headers = None
for idx, row in enumerate(ws.iter_rows(values_only=True)):
    if idx == 0:
        headers = row
        continue
    d = dict(zip(headers, row))
    grad_date = parse_any_date(d.get('Graduation Date'))
    pg = (d.get('Product Group Code') or '').strip()
    program = PROGRAM_MAP.get(pg.upper(), pg.title() if pg else '')
    state = normalise_state(d.get('Mailing State/Province'))
    records.append({
        'first': str(d.get('First Name') or '').strip(),
        'last': str(d.get('Last Name') or '').strip(),
        'nickname': '',
        'title': str(d.get('Title') or '').strip(),
        'company': str(d.get('Account Name') or '').strip(),
        'zipcode': '',
        'phone': norm_phone(d.get('Phone')),
        'mobile': norm_phone(d.get('Mobile')),
        'email': norm_email(d.get('Primary Email')),
        'program_code': '',
        'program': program,
        'delivery': '',
        'inhouse_client': '',
        'year': grad_date.year if grad_date else None,
        'session_date': grad_date.isoformat() if grad_date else '',
        'state': state,
        'source': 'William Farmer Past Graduates.xlsx',
    })

print(f'Loaded {len(records)} raw graduate rows')

# --- Filter: territory + excluded companies ---
KEEP_STATES = {'QLD', 'SA', 'WA'}
filtered = []
dropped_territory = dropped_excluded = dropped_unknown_state = 0
for r in records:
    if excluded_company(r['company']):
        dropped_excluded += 1; continue
    st = r['state']
    if st in KEEP_STATES:
        filtered.append(r)
    elif st is None:
        dropped_unknown_state += 1
    else:
        dropped_territory += 1

print(f'After territory/exclusion filter: {len(filtered)} rows')
print(f'  Dropped (excluded companies): {dropped_excluded}')
print(f'  Dropped (other territory):    {dropped_territory}')
print(f'  Dropped (unknown state):      {dropped_unknown_state}')

# --- Dedupe into people; aggregate programs ---
def person_key(r):
    if r['email']: return ('e', r['email'])
    return ('n', (r['first'] or '').lower(), (r['last'] or '').lower(), (r['company'] or '').lower())

people = {}
for r in filtered:
    k = person_key(r)
    if k not in people:
        people[k] = {
            'first': r['first'], 'last': r['last'], 'nickname': r['nickname'],
            'title': r['title'], 'company': r['company'],
            'email': r['email'], 'phone': r['phone'], 'mobile': r['mobile'],
            'zipcode': r['zipcode'], 'state': r['state'],
            'attendances': [],
        }
    p = people[k]
    # Fill in missing contact details from later rows
    if not p['email'] and r['email']: p['email'] = r['email']
    if not p['phone'] and r['phone']: p['phone'] = r['phone']
    if not p['mobile'] and r['mobile']: p['mobile'] = r['mobile']
    if not p['title'] and r['title']: p['title'] = r['title']
    if not p['company'] and r['company']: p['company'] = r['company']
    if not p['zipcode'] and r['zipcode']: p['zipcode'] = r['zipcode']
    if not p['state'] and r['state']: p['state'] = r['state']
    p['attendances'].append(r)

print(f'Unique people after dedupe: {len(people)}')

def fmt_attendance(a):
    parts = []
    if a['program']: parts.append(a['program'])
    if a['year']: parts.append(str(a['year']))
    if a['delivery']:
        d = a['delivery']
        if a['inhouse_client']: d = f"{d} - {a['inhouse_client']}"
        parts.append(f"({d})")
    return ' '.join(parts) if parts else a['program_code']

# --- Build output rows ---
HEADERS = [
    'First Name','Last Name','Nickname','Title','Company','State','Postcode',
    'Email','Phone','Mobile',
    'Programs Attended','First Year','Last Year','Total Attendances',
    'Has DCC','Has UYLP','Has LTM','Has HIP','Has WRS','Has WGYH','Has Sales','Has Leadership',
]

def build_row(p):
    atts = p['attendances']
    years = sorted({a['year'] for a in atts if a['year']})
    programs_set = sorted({a['program'] for a in atts if a['program']})
    prog_strs = sorted({fmt_attendance(a) for a in atts})
    return {
        'First Name': p['first'],
        'Last Name': p['last'],
        'Nickname': p['nickname'],
        'Title': p['title'],
        'Company': p['company'],
        'State': p['state'] or '',
        'Postcode': p['zipcode'],
        'Email': p['email'],
        'Phone': p['phone'],
        'Mobile': p['mobile'],
        'Programs Attended': '; '.join(prog_strs),
        'First Year': years[0] if years else '',
        'Last Year': years[-1] if years else '',
        'Total Attendances': len(atts),
        'Has DCC': 'Y' if 'DCC' in programs_set or 'Advanced DCC' in programs_set else '',
        'Has UYLP': 'Y' if 'UYLP' in programs_set else '',
        'Has LTM': 'Y' if 'LTM' in programs_set or 'Leadership' in programs_set else '',
        'Has HIP': 'Y' if 'HIP' in programs_set else '',
        'Has WRS': 'Y' if 'WRS' in programs_set else '',
        'Has WGYH': 'Y' if 'WGYH' in programs_set else '',
        'Has Sales': 'Y' if 'Sales' in programs_set else '',
        'Has Leadership': 'Y' if 'Leadership' in programs_set else '',
    }

master_rows = [build_row(p) for p in people.values()]

def has_contact(row):
    has_email = bool(row['Email'])
    has_phone = bool(row['Phone']) or bool(row['Mobile'])
    return has_email and has_phone

contactable_rows = [r for r in master_rows if has_contact(r)]

master_rows.sort(key=lambda r: (-(int(r['Last Year']) if r['Last Year'] else 0), r['Company'] or '', r['Last Name'] or ''))
contactable_rows.sort(key=lambda r: (-(int(r['Last Year']) if r['Last Year'] else 0), r['Company'] or '', r['Last Name'] or ''))

master_path = OUT / f'4 - Graduates' / f'Master Graduates - All - {TODAY}.csv'
contactable_path = OUT / f'4 - Graduates' / f'Master Graduates - Contactable - {TODAY}.csv'
summary_path = OUT / f'4 - Graduates' / f'Master Graduates - Summary - {TODAY}.txt'

def write_csv(path, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow(r)

write_csv(master_path, master_rows)
write_csv(contactable_path, contactable_rows)

# --- Summary ---
year_state = collections.Counter((a['year'], a['state']) for p in people.values() for a in p['attendances'] if a['state'] in KEEP_STATES)
prog_counts = collections.Counter(a['program'] for p in people.values() for a in p['attendances'] if a['program'])
state_counts_master = collections.Counter(r['State'] for r in master_rows)
state_counts_contactable = collections.Counter(r['State'] for r in contactable_rows)

with open(summary_path, 'w') as f:
    f.write(f"DC Graduates Database Build — {TODAY}\n")
    f.write(f"="*60 + "\n\n")
    f.write(f"Raw graduate rows loaded:          {len(records)}\n")
    f.write(f"  Dropped — excluded companies:    {dropped_excluded}\n")
    f.write(f"  Dropped — other territory:       {dropped_territory}\n")
    f.write(f"  Dropped — unknown state:         {dropped_unknown_state}\n")
    f.write(f"Kept attendance records:           {len(filtered)}\n\n")
    f.write(f"Unique people (master list):       {len(master_rows)}\n")
    f.write(f"  By state:\n")
    for s, c in sorted(state_counts_master.items()):
        f.write(f"    {s or '(blank)':10s}  {c}\n")
    f.write(f"\nContactable (has email AND phone/mobile): {len(contactable_rows)}\n")
    f.write(f"  By state:\n")
    for s, c in sorted(state_counts_contactable.items()):
        f.write(f"    {s or '(blank)':10s}  {c}\n")
    f.write(f"\nAttendances by program:\n")
    for p, c in sorted(prog_counts.items(), key=lambda kv: -kv[1]):
        f.write(f"  {p:30s}  {c}\n")
    f.write(f"\nAttendances by year:\n")
    year_counts = collections.Counter(a['year'] for p in people.values() for a in p['attendances'] if a['year'] and a['state'] in KEEP_STATES)
    for y, c in sorted(year_counts.items()):
        f.write(f"  {y}: {c}\n")
    f.write(f"\nOutputs:\n  {master_path}\n  {contactable_path}\n")

print(f'Wrote {master_path}')
print(f'Wrote {contactable_path}')
print(f'Wrote {summary_path}')
print(open(summary_path).read())
