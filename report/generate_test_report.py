#!/usr/bin/env python3
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
"""
generate_test_report.py
=======================
Runs the AI Companion integration test suite, parses the TRX output, and
prints a Markdown table suitable for pasting into an academic report.

Usage
-----
  cd "C:\\Users\\yyurc\\Desktop\\IBM PROJECT\\AICompanion"
  python report/generate_test_report.py

  # Skip backend-dependent tests (no Python server needed):
  python report/generate_test_report.py --no-backend

  # Include all tests (requires: uvicorn server:app --port 8000):
  python report/generate_test_report.py --with-backend
"""

import subprocess
import sys
import os
import xml.etree.ElementTree as ET
from datetime import datetime

TRX_PATH   = os.path.join("TestResults", "integration_results.trx")
PROJ_PATH  = os.path.join("tests", "AICompanion.IntegrationTests")
NS         = {"t": "http://microsoft.com/schemas/VisualStudio/TeamTest/2010"}


def run_tests(include_backend: bool) -> int:
    filter_arg = [] if include_backend else ["--filter", "Category!=RequiresBackend"]
    cmd = [
        "dotnet", "test", PROJ_PATH,
        "--logger", f"trx;LogFileName=integration_results.trx",
        "--verbosity", "normal",
        *filter_arg,
    ]
    print(f"[RUN] {' '.join(cmd)}\n")
    result = subprocess.run(cmd, capture_output=False)
    return result.returncode


def find_trx() -> str | None:
    """dotnet test puts TRX in TestResults/ relative to the project dir."""
    candidate = os.path.join(PROJ_PATH, "TestResults", "integration_results.trx")
    if os.path.exists(candidate):
        return candidate
    # also try repo-root TestResults/
    if os.path.exists(TRX_PATH):
        return TRX_PATH
    return None


def parse_trx(path: str) -> list[dict]:
    tree = ET.parse(path)
    root = tree.getroot()
    rows = []
    for r in root.findall(".//t:UnitTestResult", NS):
        name     = r.get("testName", "")
        outcome  = r.get("outcome", "Unknown")
        duration = r.get("duration", "—")
        # Strip namespace prefix from test name for readability
        short_name = name.split(".")[-1] if "." in name else name
        rows.append({"name": short_name, "outcome": outcome, "duration": duration})
    return rows


def print_markdown(rows: list[dict], return_code: int):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    passed  = sum(1 for r in rows if r["outcome"] == "Passed")
    failed  = sum(1 for r in rows if r["outcome"] == "Failed")
    skipped = sum(1 for r in rows if r["outcome"] not in ("Passed", "Failed"))
    total   = len(rows)

    print(f"\n## Integration Test Results  ({now})\n")
    print(f"**{passed}/{total} passed** | {failed} failed | {skipped} skipped\n")
    print("| # | Test Name | Result | Duration |")
    print("|---|-----------|:------:|----------|")
    for i, r in enumerate(rows, 1):
        icon = "✅" if r["outcome"] == "Passed" else ("❌" if r["outcome"] == "Failed" else "⏭️")
        # Trim long durations to ms
        dur = r["duration"].split(".")[0] if "." in r["duration"] else r["duration"]
        print(f"| {i} | `{r['name']}` | {icon} {r['outcome']} | {dur} |")

    print(f"\n**Overall: {'✅ PASS' if return_code == 0 else '❌ FAIL'}**")


def main():
    include_backend = "--with-backend" in sys.argv
    no_backend      = "--no-backend" in sys.argv or not include_backend

    print("=" * 60)
    print("AI Companion — Integration Test Report Generator")
    print("=" * 60)
    if no_backend:
        print("[INFO] Skipping RequiresBackend tests (pass --with-backend to include)\n")

    rc = run_tests(include_backend=include_backend)

    trx = find_trx()
    if trx is None:
        print("[ERROR] TRX file not found. Did the tests run successfully?")
        sys.exit(1)

    print(f"[PARSE] Reading results from: {trx}\n")
    rows = parse_trx(trx)
    print_markdown(rows, rc)

    sys.exit(rc)


if __name__ == "__main__":
    main()
