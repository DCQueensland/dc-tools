#!/usr/bin/env python3
"""Unwrap a persisted Claude tool-result file and write the inner Apollo
bulk_match response JSON to the given output path.

Usage:
    python3 apollo-save-persisted.py <persisted_tool_result.json> <output_path>

The persisted file format from Claude is either:
  [ { "type": "text", "text": "<stringified JSON>" }, ... ]
or already a plain JSON object. This helper handles both.
"""
import json
import sys


def main():
    if len(sys.argv) < 3:
        print("Usage: apollo-save-persisted.py <persisted> <output>", file=sys.stderr)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    with open(src) as f:
        raw = json.load(f)
    if isinstance(raw, list) and raw and isinstance(raw[0], dict) and "text" in raw[0]:
        inner = json.loads(raw[0]["text"])
    elif isinstance(raw, dict) and raw.get("type") == "text" and "text" in raw:
        inner = json.loads(raw["text"])
    else:
        inner = raw
    with open(dst, "w") as f:
        json.dump(inner, f, indent=2)
    matches = inner.get("unique_enriched_records")
    missing = inner.get("missing_records")
    credits = inner.get("credits_consumed")
    print(f"saved -> {dst}  matched={matches} missing={missing} credits={credits}")


if __name__ == "__main__":
    main()
