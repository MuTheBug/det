#!/usr/bin/env python3
"""
Merge sample.xlsx and chat_extracted.xlsx into a single SQLite database.
Removes duplicates based on name matching.
"""

import openpyxl
import sqlite3
import os
import re

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "combined_database.db")
SAMPLE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample.xlsx")
CHAT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chat_extracted.xlsx")


def normalize_name(name):
    """Normalize Arabic name for deduplication comparison."""
    if not name:
        return ""
    name = str(name).strip()
    # Remove extra whitespace
    name = re.sub(r'\s+', ' ', name)
    # Remove diacritics/tashkeel
    name = re.sub(r'[\u0610-\u061A\u064B-\u065F\u0670]', '', name)
    return name


def read_sample_sheet(wb, sheet_name, category):
    """Read a sheet from sample.xlsx and return list of record dicts."""
    ws = wb[sheet_name]
    records = []

    for row in range(4, ws.max_row + 1):  # Data starts at row 4 (row 1=title, 2=header, 3=sub-header)
        name = ws.cell(row, 2).value
        if not name or str(name).strip() == "":
            continue

        if category == "ناجي":
            # Survivors sheet has different column layout (13 cols)
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": ws.cell(row, 6).value,
                "release_death_disappearance_date": ws.cell(row, 7).value,
                "detention_authority": ws.cell(row, 8).value,
                "children_male": ws.cell(row, 9).value,
                "children_female": ws.cell(row, 10).value,
                "contact_number": ws.cell(row, 11).value,
                "current_address": ws.cell(row, 12).value,
                "source": "sample.xlsx",
            }
        elif category == "شهيد":
            # Martyrs sheet
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": None,
                "release_death_disappearance_date": ws.cell(row, 6).value,  # death date
                "detention_authority": ws.cell(row, 7).value,  # place of death
                "children_male": ws.cell(row, 8).value,
                "children_female": ws.cell(row, 9).value,
                "contact_number": ws.cell(row, 10).value,
                "current_address": ws.cell(row, 11).value,
                "source": "sample.xlsx",
            }
        elif category == "مختفي":
            # Forcibly disappeared sheet
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": None,
                "release_death_disappearance_date": ws.cell(row, 6).value,  # disappearance date
                "detention_authority": ws.cell(row, 7).value,
                "children_male": ws.cell(row, 8).value,
                "children_female": ws.cell(row, 9).value,
                "contact_number": ws.cell(row, 10).value,
                "current_address": ws.cell(row, 11).value,
                "source": "sample.xlsx",
            }

        records.append(record)

    return records


def read_chat_sheet(wb, sheet_name, category):
    """Read a sheet from chat_extracted.xlsx and return list of record dicts."""
    ws = wb[sheet_name]
    records = []

    for row in range(3, ws.max_row + 1):  # Data starts at row 3 (row 1=title, 2=header)
        name = ws.cell(row, 2).value
        if not name or str(name).strip() == "":
            continue

        if category == "ناجي":
            # 12 columns
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": ws.cell(row, 6).value,
                "release_death_disappearance_date": ws.cell(row, 7).value,
                "detention_authority": ws.cell(row, 8).value,
                "children_male": ws.cell(row, 9).value,
                "children_female": ws.cell(row, 10).value,
                "contact_number": ws.cell(row, 11).value,
                "current_address": ws.cell(row, 12).value,
                "source": "chat_extracted.xlsx",
            }
        elif category == "شهيد":
            # 11 columns
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": None,
                "release_death_disappearance_date": ws.cell(row, 6).value,
                "detention_authority": ws.cell(row, 7).value,
                "children_male": ws.cell(row, 8).value,
                "children_female": ws.cell(row, 9).value,
                "contact_number": ws.cell(row, 10).value,
                "current_address": ws.cell(row, 11).value,
                "source": "chat_extracted.xlsx",
            }
        elif category == "مختفي":
            # 11 columns
            record = {
                "name": str(name).strip(),
                "category": category,
                "mother_name": ws.cell(row, 3).value,
                "wife_name": ws.cell(row, 4).value,
                "birthplace_date": ws.cell(row, 5).value,
                "arrest_date": None,
                "release_death_disappearance_date": ws.cell(row, 6).value,
                "detention_authority": ws.cell(row, 7).value,
                "children_male": ws.cell(row, 8).value,
                "children_female": ws.cell(row, 9).value,
                "contact_number": ws.cell(row, 10).value,
                "current_address": ws.cell(row, 11).value,
                "source": "chat_extracted.xlsx",
            }

        records.append(record)

    return records


def clean_value(val):
    """Clean a cell value for database storage."""
    if val is None:
        return None
    val = str(val).strip()
    if val in ("", "لا يوجد", "-", "0", "عازب"):
        return val  # Keep as-is, these are meaningful
    return val


def create_database(records):
    """Create SQLite database and insert deduplicated records."""
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Create main table
    cursor.execute("""
        CREATE TABLE persons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT NOT NULL,
            mother_name TEXT,
            wife_name TEXT,
            birthplace_date TEXT,
            arrest_date TEXT,
            release_death_disappearance_date TEXT,
            detention_authority TEXT,
            children_male TEXT,
            children_female TEXT,
            contact_number TEXT,
            current_address TEXT,
            source TEXT
        )
    """)

    # Create index on normalized name + category for fast lookups
    cursor.execute("""
        CREATE INDEX idx_name_category ON persons(name, category)
    """)

    # Deduplicate: use normalized name + category as the key
    seen = {}
    duplicates = 0
    inserted = 0

    for rec in records:
        key = (normalize_name(rec["name"]), rec["category"])

        if key in seen:
            duplicates += 1
            # If existing record has less info, prefer the one with more data
            existing = seen[key]
            filled_existing = sum(1 for v in existing.values() if v and str(v).strip())
            filled_new = sum(1 for v in rec.values() if v and str(v).strip())
            if filled_new > filled_existing:
                seen[key] = rec  # Replace with more complete record
            continue

        seen[key] = rec

    # Insert deduplicated records
    for rec in seen.values():
        cursor.execute("""
            INSERT INTO persons (
                name, category, mother_name, wife_name, birthplace_date,
                arrest_date, release_death_disappearance_date, detention_authority,
                children_male, children_female, contact_number, current_address, source
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            clean_value(rec["name"]),
            rec["category"],
            clean_value(rec["mother_name"]),
            clean_value(rec["wife_name"]),
            clean_value(rec["birthplace_date"]),
            clean_value(rec["arrest_date"]),
            clean_value(rec["release_death_disappearance_date"]),
            clean_value(rec["detention_authority"]),
            clean_value(rec["children_male"]),
            clean_value(rec["children_female"]),
            clean_value(rec["contact_number"]),
            clean_value(rec["current_address"]),
            rec["source"],
        ))
        inserted += 1

    conn.commit()

    # Print summary
    print(f"{'='*50}")
    print(f"  Database created: {DB_PATH}")
    print(f"{'='*50}")
    print(f"  Total records processed: {len(records)}")
    print(f"  Duplicates removed:      {duplicates}")
    print(f"  Records inserted:        {inserted}")
    print(f"{'='*50}")

    # Per-category breakdown
    cursor.execute("SELECT category, COUNT(*) FROM persons GROUP BY category")
    print("\n  Breakdown by category:")
    for cat, count in cursor.fetchall():
        print(f"    {cat}: {count}")

    # Per-source breakdown
    cursor.execute("SELECT source, COUNT(*) FROM persons GROUP BY source")
    print("\n  Breakdown by source:")
    for src, count in cursor.fetchall():
        print(f"    {src}: {count}")

    conn.close()
    print(f"\nDone! Database saved to: {DB_PATH}")


def main():
    all_records = []

    # Read sample.xlsx (3 sheets)
    print("Reading sample.xlsx...")
    wb_sample = openpyxl.load_workbook(SAMPLE_PATH)
    all_records.extend(read_sample_sheet(wb_sample, "قائمة الناجيين ", "ناجي"))
    all_records.extend(read_sample_sheet(wb_sample, "قائمة الشهداء ", "شهيد"))
    all_records.extend(read_sample_sheet(wb_sample, "قائمة المختفيين قسراً", "مختفي"))
    print(f"  -> {len(all_records)} records from sample.xlsx")

    # Read chat_extracted.xlsx (3 category sheets, skip القائمة الشاملة since it's a combined view)
    print("Reading chat_extracted.xlsx...")
    wb_chat = openpyxl.load_workbook(CHAT_PATH)
    before = len(all_records)
    all_records.extend(read_chat_sheet(wb_chat, "قائمة الناجيين", "ناجي"))
    all_records.extend(read_chat_sheet(wb_chat, "قائمة الشهداء", "شهيد"))
    all_records.extend(read_chat_sheet(wb_chat, "قائمة المختفيين قسراً", "مختفي"))
    print(f"  -> {len(all_records) - before} records from chat_extracted.xlsx")

    print(f"\nTotal records before deduplication: {len(all_records)}")
    print("Creating database with deduplication...")

    create_database(all_records)


if __name__ == "__main__":
    main()
