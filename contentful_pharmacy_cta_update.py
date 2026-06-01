#!/usr/bin/env python3
"""
Pharmacy CTA Bulk Update — Contentful Automation
Reads a hospital mapping from Excel, scans Contentful, and previews URL changes.
"""

import os
import re
import sys
import requests
import warnings
import openpyxl

warnings.filterwarnings("ignore")

# ── Configuration ──────────────────────────────────────────────────────────────

SPACE_ID = os.environ.get("CONTENTFUL_SPACE_ID", "YOUR_SPACE_ID")
ENV      = os.environ.get("CONTENTFUL_ENVIRONMENT", "master")
TOKEN    = os.environ.get("CONTENTFUL_MANAGEMENT_TOKEN", "YOUR_MANAGEMENT_TOKEN")
BASE_URL = f"https://api.contentful.com/spaces/{SPACE_ID}/environments/{ENV}"

EXCEL_PATH = os.environ.get("PHARMACY_EXCEL", "pharmacy_updates.xlsx")

# ── Display helpers ────────────────────────────────────────────────────────────

LINE  = "─" * 64
DLINE = "═" * 64

def header(text):
    print(f"\n{DLINE}")
    print(f"  {text}")
    print(DLINE)

def step(n, text):
    print(f"\n[Step {n}] {text}")
    print(LINE)

# ── Excel parsing ──────────────────────────────────────────────────────────────

def load_excel(path):
    """Return list of {yext_id, hospital_name, old_url, new_url} from Excel."""
    try:
        wb = openpyxl.load_workbook(path)
    except FileNotFoundError:
        sys.exit(f"ERROR: Excel file not found: {path}")

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("ERROR: Excel file is empty.")

    # Expect header row: yext_id, hospital_name, old_url, new_url
    headers = [str(h).strip().lower() for h in rows[0]]
    required = {"yext_id", "hospital_name", "old_url", "new_url"}
    missing = required - set(headers)
    if missing:
        sys.exit(f"ERROR: Excel missing columns: {missing}")

    idx = {h: headers.index(h) for h in required}
    records = []
    for row in rows[1:]:
        if not any(row):
            continue
        records.append({
            "yext_id":       str(row[idx["yext_id"]]).strip(),
            "hospital_name": str(row[idx["hospital_name"]]).strip(),
            "old_url":       str(row[idx["old_url"]]).strip(),
            "new_url":       str(row[idx["new_url"]]).strip(),
        })
    return records

# ── Contentful helpers ─────────────────────────────────────────────────────────

def auth_headers():
    return {"Authorization": f"Bearer {TOKEN}"}


def fetch_online_store_entries():
    """Fetch all link entries with dataAnalyticsValue='Online Store'."""
    entries, skip, limit = [], 0, 1000
    while True:
        resp = requests.get(
            f"{BASE_URL}/entries",
            headers=auth_headers(),
            params={
                "content_type": "link",
                "fields.dataAnalyticsValue": "Online Store",
                "limit": limit,
                "skip": skip,
            },
        )
        resp.raise_for_status()
        data  = resp.json()
        batch = data.get("items", [])
        entries.extend(batch)
        skip += len(batch)
        if skip >= data.get("total", 0) or not batch:
            break
    return entries


def get_field(entry, field_name, locale="en-US"):
    return entry.get("fields", {}).get(field_name, {}).get(locale)


def extract_yext_id(title):
    """Pull 4-digit hospital ID from title, e.g. 'LINK - Hospital Name 1370 - ...'"""
    if not title:
        return None
    m = re.search(r"\b(\d{4})\b", title)
    return m.group(1) if m else None


def update_entry(entry_id, version, fields):
    resp = requests.put(
        f"{BASE_URL}/entries/{entry_id}",
        headers={**auth_headers(),
                 "Content-Type": "application/vnd.contentful.management.v1+json",
                 "X-Contentful-Version": str(version)},
        json={"fields": fields},
    )
    resp.raise_for_status()
    return resp.json()


def publish_entry(entry_id, version):
    resp = requests.put(
        f"{BASE_URL}/entries/{entry_id}/published",
        headers={**auth_headers(),
                 "Content-Type": "application/vnd.contentful.management.v1+json",
                 "X-Contentful-Version": str(version)},
    )
    resp.raise_for_status()
    return resp.json()

# ── Main workflow ──────────────────────────────────────────────────────────────

def main():
    if not SPACE_ID or not TOKEN:
        sys.exit("ERROR: set CONTENTFUL_SPACE_ID and CONTENTFUL_MANAGEMENT_TOKEN")

    header("Pharmacy CTA Bulk Update — Contentful Automation")
    print(f"  Space       : {SPACE_ID}")
    print(f"  Environment : {ENV}")
    print(f"  Excel file  : {EXCEL_PATH}")

    # ── Step 1: Parse Excel ────────────────────────────────────────────────────
    step(1, "Parsing Excel upload")
    mapping = load_excel(EXCEL_PATH)
    print(f"  Loaded {len(mapping)} hospital(s) from '{EXCEL_PATH}':\n")
    for r in mapping:
        print(f"  • [{r['yext_id']}] {r['hospital_name']}")
        print(f"      OLD: {r['old_url']}")
        print(f"      NEW: {r['new_url']}")

    # Build lookup: yext_id → record
    lookup = {r["yext_id"]: r for r in mapping}

    # ── Step 2: Scan Contentful ────────────────────────────────────────────────
    step(2, "Scanning Contentful for 'Online Store' link entries")
    print("  Fetching entries… ", end="", flush=True)
    entries = fetch_online_store_entries()
    print(f"found {len(entries)} entries.")

    # ── Step 3: Match entries to Excel hospitals ───────────────────────────────
    step(3, f"Matching {len(entries)} Contentful entries against {len(mapping)} hospitals from Excel")

    matches   = []
    no_yext   = 0
    no_match  = 0
    url_wrong = 0

    for entry in entries:
        entry_id = entry["sys"]["id"]
        version  = entry["sys"]["version"]
        title    = get_field(entry, "title")
        yext_id  = extract_yext_id(title)

        if not yext_id:
            no_yext += 1
            continue

        if yext_id not in lookup:
            no_match += 1
            continue

        record      = lookup[yext_id]
        current_url = get_field(entry, "url")

        if current_url != record["old_url"]:
            url_wrong += 1
            print(f"  SKIP  [{yext_id}] {record['hospital_name']} — URL already updated")
            continue

        print(f"  MATCH [{yext_id}] {record['hospital_name']}")
        matches.append((entry_id, version, entry.get("fields", {}), record, current_url))

    print(f"\n  {len(entries)} scanned  |  {len(matches)} matched  |  {url_wrong} already updated  |  {no_yext + no_match} skipped")

    # ── Step 4: Preview changes ────────────────────────────────────────────────
    step(4, f"Pending changes — {len(matches)} entry(ies) ready to update")

    if not matches:
        print("  Nothing to update.")
    else:
        # Table header
        col_yext  = 10
        col_name  = 44
        col_url   = 52
        print(f"\n  {'YEXT ID':<{col_yext}}  {'HOSPITAL':<{col_name}}  {'OLD URL':<{col_url}}  NEW URL")
        print(f"  {'─'*col_yext}  {'─'*col_name}  {'─'*col_url}  {'─'*col_url}")
        for _, _, _, record, current_url in matches:
            print(f"  {record['yext_id']:<{col_yext}}  {record['hospital_name']:<{col_name}}  {current_url:<{col_url}}  {record['new_url']}")

    # ── Step 5: Confirm and apply ──────────────────────────────────────────────
    step(5, "Apply changes?")

    if not matches:
        print("  No changes to apply. Exiting.")
        return

    print(f"\n  {len(matches)} entry(ies) will be updated in Contentful.")
    print("  This action cannot be undone automatically.\n")
    answer = input("  Apply changes now? [y/N] › ").strip().lower()

    if answer != "y":
        print("\n  Aborted. No changes were made to Contentful.")
        return

    print()
    saved, save_failed = [], []
    for entry_id, version, fields, record, current_url in matches:
        updated_fields = dict(fields)
        updated_fields["url"] = {"en-US": record["new_url"]}
        try:
            updated = update_entry(entry_id, version, updated_fields)
            new_version = updated["sys"]["version"]
            print(f"  SAVED  [{record['yext_id']}] {record['hospital_name']}")
            print(f"         {current_url}")
            print(f"      → {record['new_url']}")
            saved.append((entry_id, new_version, record))
        except requests.HTTPError as exc:
            print(f"  FAILED [{record['yext_id']}] {entry_id} — {exc}")
            save_failed.append(record)

    print(f"\n  {len(saved)} saved, {len(save_failed)} failed.")

    # ── Step 6: Publish ────────────────────────────────────────────────────────
    step(6, "Publish changes to live site?")

    if not saved:
        print("  Nothing to publish.")
        print(f"\n{DLINE}")
        print(f"  Done. 0 published.")
        print(DLINE)
        return

    print(f"\n  {len(saved)} entry(ies) saved as draft in Contentful.")
    print("  Publishing will make the new URLs live immediately.\n")
    answer = input("  Publish to live site now? [y/N] › ").strip().lower()

    if answer != "y":
        print("\n  Entries saved as draft. Publish manually in Contentful when ready.")
        print(f"\n{DLINE}")
        print(f"  Done. {len(saved)} saved (unpublished), {len(save_failed)} failed.")
        print(DLINE)
        return

    print()
    pub_success, pub_failed = 0, 0
    for entry_id, version, record in saved:
        try:
            publish_entry(entry_id, version)
            print(f"  PUBLISHED  [{record['yext_id']}] {record['hospital_name']}")
            pub_success += 1
        except requests.HTTPError as exc:
            print(f"  PUB FAILED [{record['yext_id']}] {entry_id} — {exc}")
            pub_failed += 1

    print(f"\n{DLINE}")
    print(f"  Done. {pub_success} published live, {pub_failed} failed.")
    print(DLINE)


if __name__ == "__main__":
    main()
