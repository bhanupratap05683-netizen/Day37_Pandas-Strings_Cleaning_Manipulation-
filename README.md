# Day 37 — String Cleaning & Manipulation

**Roadmap Phase:** 2 (Advanced Excel + pandas intro)
**Date:** May 2026
**Topic:** strip · replace · split · case conversion · regex basics · pandas .str accessor

---

## What This Project Does

Loads two messy real-world datasets (financial transactions + employee records),
applies Python and pandas string-cleaning techniques, and exports a fully
standardised Excel report with two clean sheets.

---

## Files

| File | Purpose |
|------|---------|
| `Day37_String_Cleaning_Input.xlsx` | Raw messy data (practice input) |
| `Day37_String_Cleaning.py` | All cleaning logic — copy-paste ready |
| `Day37_String_Cleaning_Output.xlsx` | Cleaned and formatted output |

---

## Skills Demonstrated

- `str.strip()`, `lstrip()`, `rstrip()` — whitespace removal
- `str.upper()`, `lower()`, `title()`, `capitalize()` — case normalisation
- `str.replace()` — substring substitution
- `str.split()` — string to list conversion
- `str.startswith()`, `endswith()`, `in` — substring checks
- `re.sub()`, `re.search()`, `re.findall()` — regex pattern matching
- `pandas .str` accessor — vectorised column-level string ops
- `str.contains(regex=True)` — conditional flagging

---

## Key Cleaning Operations

| Column | Problem | Fix Applied |
|--------|---------|-------------|
| Company_Name | Extra spaces, random casing | `.strip().title()` |
| Transaction_Type | Mixed case, padding | `.strip().upper()` |
| Amount_Raw | `Rs. 45,000.50` format | regex prefix strip + comma remove |
| Email_Raw | Internal spaces, mixed case | `.strip().lower()` + regex |
| Phone_Raw | `+91-`, `-`, `.` variations | `re.sub([^0-9])` → last 10 digits |

---

## Portfolio Connection

String cleaning is the first step in every real data pipeline.
This script is a reusable template for the **Day 50 Sales Analyzer**
and the **Day 78 Financial Dashboard** — both require clean inputs
before any analysis can run.
