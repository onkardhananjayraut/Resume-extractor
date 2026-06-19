"""
bulk_processor.py
─────────────────
Processes up to 50 resumes concurrently and writes results to Excel.

Usage:
    python bulk_processor.py                          # default ./uploads folder
    python bulk_processor.py --folder /path/to/pdfs  # custom folder
    python bulk_processor.py --workers 10             # tune thread count
    python bulk_processor.py --limit 30               # process only 30
"""

import os
import sys
import time
import argparse
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

from parser import (
    extract_text,
    extract_name,
    extract_email,
    extract_mobile,
    extract_highest_qualification_details,
    extract_latest_company,
    extract_total_experience,
)

# ════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════

DEFAULT_UPLOAD_FOLDER = "uploads"
DEFAULT_OUTPUT_PATH   = "output/candidates.xlsx"
MAX_WORKERS           = 8
MAX_RESUMES           = 50

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ════════════════════════════════════════════════
# PER-FILE WORKER
# ════════════════════════════════════════════════

def process_resume(pdf_path: str) -> dict:
    """
    Extract all fields from one resume.
    Any exception is caught and stored in the Error column so
    one bad PDF never aborts the whole batch.
    """
    filename = os.path.basename(pdf_path)
    try:
        text = extract_text(pdf_path)

        # Call expensive PDF-level helpers once and reuse
        qual = extract_highest_qualification_details(text)

        return {
            "Resume":           filename,
            "Name":             extract_name(text),
            "Email":            extract_email(text),
            "Mobile":           extract_mobile(text),
            "Qualification":    qual["qualification"],
            "Passing Year":     qual["passing_year"],
            "Latest Company":   extract_latest_company(text),
            "Total Experience": extract_total_experience(text),
            "Error":            "",
        }

    except Exception as exc:
        log.error("✗ Failed: %s  →  %s", filename, exc)
        return {
            "Resume":           filename,
            "Name":             "Error",
            "Email":            "Error",
            "Mobile":           "Error",
            "Qualification":    "Error",
            "Passing Year":     "Error",
            "Latest Company":   "Error",
            "Total Experience": "Error",
            "Error":            str(exc),
        }


# ════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════

def main(folder: str, output: str, workers: int, limit: int) -> None:

    # ── Collect PDFs ──────────────────────────────────────────────
    if not os.path.isdir(folder):
        log.error("Folder not found: %s", folder)
        sys.exit(1)

    resume_files = sorted(
    os.path.join(folder, f)
    for f in os.listdir(folder)
    if f.lower().endswith((".pdf", ".docx"))
)
    

    if not resume_files:
        log.warning("No resume files found in '%s'.", folder)
        return

    if len(resume_files) > limit:
        log.warning("Found %d resume files – processing first %d (cap=%d).",
                    len(resume_files), limit, limit)
        resume_files = resume_files[:limit]

    log.info("Processing %d resume(s) with %d worker(s) …",
             len(resume_files), workers)
    t0 = time.time()

    # ── Parallel extraction ────────────────────────────────────────
    results   = [None] * len(resume_files)
    completed = 0

    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {pool.submit(process_resume, p): i
                      for i, p in enumerate(resume_files)}

        for fut in as_completed(future_map):
            idx           = future_map[fut]
            results[idx]  = fut.result()
            completed    += 1
            status = "✓" if results[idx]["Error"] == "" else "✗"
            log.info("[%d/%d] %s %s", completed, len(resume_files),
                     status, results[idx]["Resume"])

    elapsed = time.time() - t0
    log.info("Done in %.1f s", elapsed)

    # ── Build DataFrame ────────────────────────────────────────────
    columns = [
        "Resume", "Name", "Email", "Mobile",
        "Qualification", "Passing Year",
        "Latest Company", "Total Experience", "Error",
    ]
    df = pd.DataFrame(results, columns=columns)

    # ── Console summary ────────────────────────────────────────────
    ok     = (df["Error"] == "").sum()
    failed = len(df) - ok

    print("\n" + "=" * 80)
    print(f"  SUMMARY   Total: {len(df)}   OK: {ok}   Failed: {failed}"
          f"   Time: {elapsed:.1f}s")
    print("=" * 80)
    print(df[[
        "Resume", "Name", "Qualification", "Passing Year",
        "Latest Company", "Total Experience"
    ]].to_string(index=False))
    print("=" * 80 + "\n")

    # ── Write Excel ────────────────────────────────────────────────
    os.makedirs(os.path.dirname(output) or ".", exist_ok=True)
    df.to_excel(output, index=False)
    log.info("Excel saved → %s", output)

    if failed:
        log.warning("%d file(s) had errors – see 'Error' column in Excel.", failed)


# ════════════════════════════════════════════════
# CLI
# ════════════════════════════════════════════════

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Bulk resume parser (up to 50 PDFs).")
    ap.add_argument("--folder",  default=DEFAULT_UPLOAD_FOLDER)
    ap.add_argument("--output",  default=DEFAULT_OUTPUT_PATH)
    ap.add_argument("--workers", type=int, default=MAX_WORKERS,
                    help="Parallel worker threads (default 8)")
    ap.add_argument("--limit",   type=int, default=MAX_RESUMES,
                    help="Max resumes to process (default 50)")
    args = ap.parse_args()
    main(args.folder, args.output, args.workers, args.limit)

