#!/usr/bin/env python3
"""Tier the Apollo-enriched graduate records into Active Match, Moved,
No Match, and Priority Target segments, and produce CSVs + summary.

Reads:
  /tmp/apollo_source_rows.json       (original records, all columns, _row_id)
  /tmp/apollo_enriched.jsonl         (one JSON record per row: _req_id, matched, ...)
  /tmp/apollo_responses/batch_*.json (for credits_consumed totals)

Writes CSVs + summary into
  /Users/williamfarmer/Dropbox/2026/William/Database All/Graduates 2016 - 2026/
"""
import csv
import glob
import json
import os
import re
from collections import Counter

SOURCE_ROWS = "/tmp/apollo_source_rows.json"
ENRICHED_JSONL = "/tmp/apollo_enriched.jsonl"
RESPONSES_GLOB = "/tmp/apollo_responses/batch_*.json"

OUTPUT_DIR = (
    "/Users/williamfarmer/Dropbox/2026/William/Database All/"
    "Graduates 2016 - 2026"
)
DATE_STR = "18 Apr 2026"

ACTIVE_CSV = f"Apollo Enriched - Active Match - {DATE_STR}.csv"
MOVED_CSV = f"Apollo Enriched - Moved New Employer - {DATE_STR}.csv"
NO_MATCH_CSV = f"Apollo Enriched - No Match - {DATE_STR}.csv"
PRIORITY_CSV = f"Apollo Enriched - PRIORITY Owner and Ops Targets - {DATE_STR}.csv"
SUMMARY_TXT = f"Apollo Enriched - Summary - {DATE_STR}.txt"

PRIORITY_KEYWORDS = [
    "owner",
    "ceo",
    "founder",
    "president",
    "managing director",
    "director",
    "general manager",
    "gm",
    "operations manager",
    "ops manager",
    "head of operations",
    "chief",
    "partner",
    "principal",
    "coo",
]

_NORM_STOP = [
    "pty. ltd.",
    "pty ltd",
    "pty",
    "ltd.",
    "ltd",
    "inc.",
    "inc",
    " & co",
    " and co",
    " group",
]


def _normalise_company(name):
    if not name:
        return ""
    s = name.lower().strip()
    # Strip company-type suffixes.
    for stop in _NORM_STOP:
        s = s.replace(stop, " ")
    # Strip punctuation.
    s = re.sub(r"[^\w\s&]+", " ", s)
    # Collapse spaces.
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_priority(title):
    if not title:
        return False
    t = title.lower()
    for kw in PRIORITY_KEYWORDS:
        if kw == "gm":
            # Word-bound "gm" only.
            if re.search(r"\bgm\b", t):
                return True
        else:
            if kw in t:
                return True
    return False


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(SOURCE_ROWS) as f:
        rows = json.load(f)

    enriched_by_id = {}
    with open(ENRICHED_JSONL) as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            r = json.loads(line)
            enriched_by_id[r["_req_id"]] = r

    # Sum credits across all saved batch responses.
    total_credits = 0.0
    for path in glob.glob(RESPONSES_GLOB):
        try:
            with open(path) as f:
                resp = json.load(f)
            c = resp.get("credits")
            if isinstance(c, (int, float)):
                total_credits += float(c)
        except Exception:
            continue

    # Columns.
    original_cols = list(rows[0].keys()) if rows else []
    # Drop _row_id from original_cols export position; keep internal only.
    # Task says: all original CSV columns + Apollo_* + Match_Status +
    # Priority_Target. We'll exclude _row_id to keep it clean.
    export_original_cols = [c for c in original_cols if c != "_row_id"]
    apollo_cols = [
        "Apollo_Matched",
        "Apollo_Current_Company",
        "Apollo_Current_Title",
        "Apollo_Verified_Email",
        "Apollo_Personal_Emails",
        "Apollo_LinkedIn_URL",
        "Apollo_City",
        "Apollo_State",
        "Apollo_Seniority",
        "Match_Status",
        "Priority_Target",
    ]
    final_cols = export_original_cols + apollo_cols

    active_rows = []
    moved_rows = []
    no_match_rows = []

    priority_active = 0
    priority_moved = 0
    priority_by_state = Counter()
    new_employer_counter = Counter()
    processed_count = 0
    matched_count = 0

    for row in rows:
        rid = row.get("_row_id")
        enr = enriched_by_id.get(rid)
        out = {c: row.get(c, "") for c in export_original_cols}

        if enr is None:
            # Not enriched yet.
            out["Apollo_Matched"] = ""
            out["Apollo_Current_Company"] = ""
            out["Apollo_Current_Title"] = ""
            out["Apollo_Verified_Email"] = ""
            out["Apollo_Personal_Emails"] = ""
            out["Apollo_LinkedIn_URL"] = ""
            out["Apollo_City"] = ""
            out["Apollo_State"] = ""
            out["Apollo_Seniority"] = ""
            out["Match_Status"] = "Not Processed"
            out["Priority_Target"] = ""
            no_match_rows.append(out)
            continue

        processed_count += 1
        matched = bool(enr.get("matched"))
        current_org = enr.get("current_organization_name") or ""
        current_title = enr.get("current_title") or ""
        if matched:
            matched_count += 1

        norm_orig = _normalise_company(row.get("Company", ""))
        norm_new = _normalise_company(current_org)

        # Decide bucket.
        if not matched or not current_org:
            status = "No Match"
        elif norm_new and norm_orig and norm_new == norm_orig:
            status = "Active Match"
        elif norm_new and norm_new != norm_orig:
            status = "Moved - New Employer"
        else:
            status = "No Match"

        priority = _is_priority(current_title)
        out["Apollo_Matched"] = "Y" if matched else "N"
        out["Apollo_Current_Company"] = current_org
        out["Apollo_Current_Title"] = current_title
        out["Apollo_Verified_Email"] = enr.get("verified_email") or ""
        out["Apollo_Personal_Emails"] = enr.get("personal_emails") or ""
        out["Apollo_LinkedIn_URL"] = enr.get("linkedin_url") or ""
        out["Apollo_City"] = enr.get("city") or ""
        out["Apollo_State"] = enr.get("state") or ""
        out["Apollo_Seniority"] = enr.get("seniority") or ""
        out["Match_Status"] = status
        out["Priority_Target"] = "Y" if priority else ""

        if status == "Active Match":
            active_rows.append(out)
            if priority:
                priority_active += 1
        elif status == "Moved - New Employer":
            moved_rows.append(out)
            new_employer_counter[current_org] += 1
            if priority:
                priority_moved += 1
        else:
            no_match_rows.append(out)

        if priority and status in ("Active Match", "Moved - New Employer"):
            orig_state = (row.get("State") or "").strip().upper()
            if orig_state in ("QLD", "SA", "WA"):
                priority_by_state[orig_state] += 1
            else:
                priority_by_state[orig_state or "OTHER"] += 1

    def _sort_key(r):
        return ((r.get("Company") or "").lower(), (r.get("Last Name") or "").lower())

    active_rows.sort(key=_sort_key)
    moved_rows.sort(key=_sort_key)
    no_match_rows.sort(key=_sort_key)

    priority_rows = [
        r for r in (active_rows + moved_rows) if r.get("Priority_Target") == "Y"
    ]
    priority_rows.sort(key=_sort_key)

    def _write_csv(path, data, cols):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=cols)
            writer.writeheader()
            for r in data:
                writer.writerow({c: r.get(c, "") for c in cols})

    _write_csv(os.path.join(OUTPUT_DIR, ACTIVE_CSV), active_rows, final_cols)
    _write_csv(os.path.join(OUTPUT_DIR, MOVED_CSV), moved_rows, final_cols)
    _write_csv(os.path.join(OUTPUT_DIR, NO_MATCH_CSV), no_match_rows, final_cols)
    _write_csv(os.path.join(OUTPUT_DIR, PRIORITY_CSV), priority_rows, final_cols)

    top_employers = new_employer_counter.most_common(20)

    with open(os.path.join(OUTPUT_DIR, SUMMARY_TXT), "w", encoding="utf-8") as f:
        f.write("Apollo Enrichment Summary - 18 Apr 2026\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"Total input rows:              {len(rows)}\n")
        f.write(f"Total processed (enriched):    {processed_count}\n")
        f.write(f"Total Apollo matches:          {matched_count}\n")
        f.write(f"Total credits consumed:        {total_credits:.2f}\n\n")
        f.write("Buckets:\n")
        f.write(f"  Active Match:                {len(active_rows)}\n")
        f.write(f"  Moved - New Employer:        {len(moved_rows)}\n")
        f.write(f"  No Match / Not Processed:    {len(no_match_rows)}\n\n")
        f.write(f"Priority Targets (Active+Moved): {len(priority_rows)}\n")
        f.write(f"  Active Match priority:       {priority_active}\n")
        f.write(f"  Moved priority:              {priority_moved}\n\n")
        f.write("Priority Targets by state (original graduate state):\n")
        for st in ("QLD", "SA", "WA"):
            f.write(f"  {st}: {priority_by_state.get(st, 0)}\n")
        other = sum(
            v for k, v in priority_by_state.items() if k not in ("QLD", "SA", "WA")
        )
        if other:
            f.write(f"  Other: {other}\n")
        f.write("\n")
        f.write("Top 20 new employers grads have moved to:\n")
        for emp, n in top_employers:
            f.write(f"  {n:>3}  {emp}\n")

    print("Wrote:")
    for name in (ACTIVE_CSV, MOVED_CSV, NO_MATCH_CSV, PRIORITY_CSV, SUMMARY_TXT):
        print(f"  {os.path.join(OUTPUT_DIR, name)}")
    print()
    print(f"Active Match:          {len(active_rows)}")
    print(f"Moved - New Employer:  {len(moved_rows)}")
    print(f"No Match:              {len(no_match_rows)}")
    print(f"Priority Targets:      {len(priority_rows)}  (Active {priority_active} / Moved {priority_moved})")
    print(f"Total credits:         {total_credits:.2f}")


if __name__ == "__main__":
    main()
