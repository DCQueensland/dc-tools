#!/usr/bin/env python3
"""Extract per-person enrichment records from an Apollo bulk_match response.

Usage:
    python3 apollo-extract.py /path/to/apollo_response.json

Appends one JSON line per person (including unmatched) to
/tmp/apollo_enriched.jsonl. The input response filename must be of the
form batch_NNN.json so the corresponding request batch can be found at
/tmp/apollo_batches/batch_NNN.json.
"""
import json
import os
import sys

OUTPUT_PATH = "/tmp/apollo_enriched.jsonl"
BATCHES_DIR = "/tmp/apollo_batches"


def _org_name_from_match(m):
    """Best-effort current organisation name from a match object."""
    if not m:
        return ""
    org = m.get("organization") or {}
    if isinstance(org, dict):
        name = org.get("name")
        if name:
            return name
    # Fall back to employment_history (latest with current=true).
    history = m.get("employment_history") or []
    if isinstance(history, list):
        for entry in history:
            if not isinstance(entry, dict):
                continue
            if entry.get("current") is True and entry.get("organization_name"):
                return entry["organization_name"]
        # If no "current" flag, try the first entry with an org name.
        for entry in history:
            if isinstance(entry, dict) and entry.get("organization_name"):
                return entry["organization_name"]
    return ""


def extract(response_path):
    fname = os.path.basename(response_path)
    # Expect batch_NNN.json.
    stem = os.path.splitext(fname)[0]
    batch_path = os.path.join(BATCHES_DIR, f"{stem}.json")
    if not os.path.exists(batch_path):
        raise FileNotFoundError(f"Could not find batch file at {batch_path}")

    with open(response_path) as f:
        resp = json.load(f)
    with open(batch_path) as f:
        batch = json.load(f)

    matches = resp.get("matches") or []
    # Prefer _requested_ids from the saved response if present; otherwise use
    # ids from the batch itself (these are the _row_id values we passed in).
    requested_ids = resp.get("_requested_ids") or [b.get("id") for b in batch]

    rows_out = []
    # The matches list is aligned to the requested ids by position.
    for i, req_id in enumerate(requested_ids):
        match = matches[i] if i < len(matches) else None
        # The API may also attach _req_id on the match object; prefer that
        # when present because it is the authoritative link.
        effective_req_id = req_id
        if match and match.get("_req_id"):
            effective_req_id = match.get("_req_id")

        if not match:
            rec = {
                "_req_id": effective_req_id,
                "matched": False,
                "current_organization_name": "",
                "current_title": "",
                "verified_email": "",
                "personal_emails": "",
                "linkedin_url": "",
                "city": "",
                "state": "",
                "country": "",
                "seniority": "",
            }
        else:
            personal = match.get("personal_emails") or []
            if not isinstance(personal, list):
                personal = []
            rec = {
                "_req_id": effective_req_id,
                "matched": True,
                "current_organization_name": _org_name_from_match(match) or "",
                "current_title": match.get("title") or "",
                "verified_email": match.get("email") or "",
                "personal_emails": ";".join([str(e) for e in personal if e]),
                "linkedin_url": match.get("linkedin_url") or "",
                "city": match.get("city") or "",
                "state": match.get("state") or "",
                "country": match.get("country") or "",
                "seniority": match.get("seniority") or "",
            }
        rows_out.append(rec)

    with open(OUTPUT_PATH, "a") as f:
        for rec in rows_out:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    return len(rows_out), sum(1 for r in rows_out if r["matched"])


def main():
    if len(sys.argv) < 2:
        print("Usage: apollo-extract.py <response_json>", file=sys.stderr)
        sys.exit(1)
    total, matched = extract(sys.argv[1])
    print(f"Wrote {total} records ({matched} matched) -> {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
