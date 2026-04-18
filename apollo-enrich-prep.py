#!/usr/bin/env python3
"""Preprocess the Contactable Graduates CSV into Apollo-ready batches of 10."""
import csv, json, os
from pathlib import Path

SRC = Path('/Users/williamfarmer/Dropbox/2026/William/Database All/Graduates 2016 - 2026/Graduates - Contactable 2250 People (Email + Phone) - 18 Apr 2026.csv')
BATCH_DIR = Path('/tmp/apollo_batches')
BATCH_DIR.mkdir(exist_ok=True)
# Clean old batches
for f in BATCH_DIR.glob('*.json'):
    f.unlink()

rows = []
with open(SRC, newline='', encoding='utf-8') as f:
    for i, r in enumerate(csv.DictReader(f)):
        r['_row_id'] = f'g{i:04d}'
        rows.append(r)

print(f'Loaded {len(rows)} graduates.')

# Strip DC staff — multiple domains + internal placeholder company
DC_EMAIL_DOMAINS = ('@dalecarnegie.com', '@dalecarnegieqld.com', '@dalecarnegie.com.au')
def is_dc_staff(r):
    email = (r.get('Email') or '').lower()
    company = (r.get('Company') or '').lower()
    if any(d in email for d in DC_EMAIL_DOMAINS): return True
    if company.startswith('40626a-farmer'): return True
    if 'dale carnegie' in company: return True
    return False

dc_staff = [r for r in rows if is_dc_staff(r)]
rows = [r for r in rows if not is_dc_staff(r)]
print(f'Dropped {len(dc_staff)} DC staff/internal. Processing {len(rows)} graduates.')

# Save full row index so we can rejoin later
with open('/tmp/apollo_source_rows.json', 'w') as f:
    json.dump(rows, f)

# Batch into 10s — Apollo bulk match max
batches = []
for i in range(0, len(rows), 10):
    chunk = rows[i:i+10]
    batch = []
    for r in chunk:
        batch.append({
            'id': r['_row_id'],
            'first_name': r.get('First Name', ''),
            'last_name': r.get('Last Name', ''),
            'organization_name': r.get('Company', ''),
            'email': r.get('Email', ''),
        })
    batches.append(batch)

for idx, batch in enumerate(batches):
    with open(BATCH_DIR / f'batch_{idx:03d}.json', 'w') as f:
        json.dump(batch, f)

print(f'Wrote {len(batches)} batches to {BATCH_DIR}')
print(f'First batch preview: {json.dumps(batches[0][:2], indent=2)}')

# Truncate checkpoint file
with open('/tmp/apollo_enriched.jsonl', 'w') as f:
    pass
print('Initialised /tmp/apollo_enriched.jsonl')
