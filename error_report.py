import csv

LOG_FILE = r"Z:\BILLING 2025-2026\Automation Logs\job_move_log.csv"
OUT_FILE = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\unified file manager\error_analysis.txt"

with open(LOG_FILE, "r", encoding="utf-8") as f:
    rows = list(csv.reader(f))

moved = [r for r in rows[1:] if len(r) > 5 and r[5] == "MOVED"]
skipped = [r for r in rows[1:] if len(r) > 5 and r[5] == "SKIPPED"]
errors = [r for r in rows[1:] if len(r) > 5 and r[5] == "ERROR"]

with open(OUT_FILE, "w", encoding="utf-8") as out:
    out.write(f"Total log entries: {len(rows)-1}\n")
    out.write(f"MOVED: {len(moved)}\n")
    out.write(f"SKIPPED: {len(skipped)}\n")
    out.write(f"ERROR: {len(errors)}\n")
    out.write("\n===== ALL ERRORS =====\n\n")
    for row in errors:
        out.write(f"Time:    {row[0]}\n")
        out.write(f"Job:     {row[1]}\n")
        out.write(f"Import:  {row[2]}\n")
        out.write(f"Folder:  {row[3]}\n")
        out.write(f"Trigger: {row[4]}\n")
        out.write(f"Comment: {row[6]}\n")
        out.write("---\n")

    # Unique skip reasons
    skip_reasons = set()
    for row in skipped:
        skip_reasons.add(row[6])
    out.write("\n===== UNIQUE SKIP REASONS =====\n")
    for r in skip_reasons:
        out.write(f"  - {r}\n")

print("Report written to:", OUT_FILE)
