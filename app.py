"""
app.py  –  Single-resume debug runner
Edit RESUME_PATH and run:  python app.py
"""

from parser import (
    extract_text,
    extract_name,
    extract_email,
    extract_mobile,
    extract_education_section,
    extract_experience_section,
    extract_companies_from_experience,
    extract_experience_records,
    extract_highest_qualification_details,
    extract_latest_company,
    extract_total_experience,
)

# ── SELECT RESUME ─────────────────────────────
RESUME_PATH = "uploads/Dhananjayresume_581638820403511.pdf"
# RESUME_PATH = "uploads/Onkar_Raut_Resume.pdf"

# ── EXTRACT TEXT ──────────────────────────────
print(f"\nReading: {RESUME_PATH}")
text = extract_text(RESUME_PATH)

# ── EDUCATION ─────────────────────────────────
print("\n" + "=" * 65)
print("EDUCATION SECTION")
print("=" * 65)
print(extract_education_section(text))

# ── QUALIFICATION DETAILS ─────────────────────
qual = extract_highest_qualification_details(text)

# ── EXPERIENCE ────────────────────────────────
exp = extract_experience_section(text)

print("\n" + "=" * 65)
print("EXPERIENCE SECTION")
print("=" * 65)
print(exp)

print("\n" + "=" * 65)
print("COMPANIES FOUND  (spaCy ORG – for audit)")
print("=" * 65)
for c in extract_companies_from_experience(exp) or ["None found"]:
    print(" •", c)

print("\n" + "=" * 65)
print("EXPERIENCE RECORDS  (date-anchored)")
print("=" * 65)
records = extract_experience_records(exp)
if records:
    for r in records:
        status = "CURRENT" if r["is_current"] else f"until {r['end_year']}/{r['end_month']:02d}"
        print(f"  {r['company']:<40}  {r['start_year']}/{r['start_month']:02d} → {status}")
else:
    print("  None found")

# ── FINAL OUTPUT ──────────────────────────────
print("\n" + "=" * 65)
print("EXTRACTED DATA")
print("=" * 65)
print(f"  Name             : {extract_name(text)}")
print(f"  Email            : {extract_email(text)}")
print(f"  Mobile           : {extract_mobile(text)}")
print(f"  Qualification    : {qual['qualification']}")
print(f"  Passing Year     : {qual['passing_year']}")
print(f"  Latest Company   : {extract_latest_company(text)}")
print(f"  Total Experience : {extract_total_experience(text)}")
print("=" * 65)

