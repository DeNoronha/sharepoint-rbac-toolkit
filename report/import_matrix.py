#!/usr/bin/env python3
"""Takes the matrix tab from a report that the owner edited and compares it with the
current matrix file. Shows every difference; writes nothing unless you pass --write.

    python3 report/import_matrix.py --config config/contoso "returned report.xlsx"
    python3 report/import_matrix.py --config config/contoso "returned report.xlsx" --write

With --write the matrix file is replaced and the previous version is kept next to it as
matrix_<timestamp>.bak.csv in the output folder.

A new column (target) also needs an entry under "targets" in config.json; a new row (role)
needs a group, which 04-Sync-GroupNesting.ps1 -CreateMissingGroups creates.

Needs: pip install openpyxl
"""
import argparse
import datetime
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sprbac import LEVELS, NO_MARKS, Config, fail, write_matrix  # noqa: E402

try:
    import openpyxl
except ImportError:
    fail("Python package openpyxl is missing. Install it with: python3 -m pip install openpyxl")


def main():
    ap = argparse.ArgumentParser(description="Import the matrix tab of an edited report.")
    ap.add_argument("--config", required=True)
    ap.add_argument("xlsx", help="the report the owner sent back")
    ap.add_argument("--sheet", help="name of the matrix tab (default: report.matrixSheet)")
    ap.add_argument("--write", action="store_true", help="replace the matrix file")
    args = ap.parse_args()

    cfg = Config(args.config)
    sheet = args.sheet or cfg.report.get("matrixSheet", "Access Matrix")
    if not os.path.isfile(args.xlsx):
        fail(f"File not found: {args.xlsx}")
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    if sheet not in wb.sheetnames:
        fail(f"No tab '{sheet}' in {args.xlsx}. Tabs: {', '.join(wb.sheetnames)}")
    rows = [["" if v is None else str(v).strip() for v in r] for r in wb[sheet].iter_rows(values_only=True)]
    rows = [r for r in rows if any(r)]
    header = rows[0]
    while header and not header[-1]:
        header.pop()
    new_targets = header[1:]
    new_roles, new_cells = [], {}
    for r in rows[1:]:
        role = r[0]
        if not role:
            continue
        if role in new_cells:
            fail(f"Role '{role}' appears twice in the tab.")
        new_roles.append(role)
        new_cells[role] = {}
        for i, t in enumerate(new_targets, start=1):
            v = r[i] if i < len(r) else ""
            if v.lower() in NO_MARKS:
                continue
            new_cells[role][t] = next((l for l in LEVELS if l.lower() == v.lower()), "*")

    old_roles, old_targets, old_cells = cfg.roles, cfg.matrix_targets, cfg.cells
    print(f"Current matrix : {len(old_roles)} roles x {len(old_targets)} columns ({os.path.basename(cfg.matrix_path)})")
    print(f"Returned tab   : {len(new_roles)} roles x {len(new_targets)} columns ({os.path.basename(args.xlsx)} / {sheet})\n")

    changes = 0
    for r in [r for r in new_roles if r not in old_roles]:
        print(f"  NEW ROLE       {r}: {', '.join(new_cells[r]) or '(no marks)'}"); changes += 1
    for r in [r for r in old_roles if r not in new_roles]:
        print(f"  ROLE REMOVED   {r} (had: {', '.join(old_cells[r]) or 'no marks'})"); changes += 1
    for t in [t for t in new_targets if t not in old_targets]:
        print(f"  NEW COLUMN     {t}: {', '.join(r for r in new_roles if t in new_cells[r]) or '(no marks)'}"); changes += 1
    for t in [t for t in old_targets if t not in new_targets]:
        print(f"  COLUMN REMOVED {t}"); changes += 1
    for r in [r for r in new_roles if r in old_roles]:
        for t in [t for t in new_targets if t in old_targets]:
            a, b = old_cells[r].get(t), new_cells[r].get(t)
            if a == b:
                continue
            changes += 1
            show = lambda c: "default level" if c == "*" else c  # noqa: E731
            if a is None:
                print(f"  ACCESS ADDED   {r} -> {t}" + ("" if b == "*" else f" ({b})"))
            elif b is None:
                print(f"  ACCESS REMOVED {r} -> {t}")
            else:
                print(f"  LEVEL CHANGED  {r} -> {t}: {show(a)} -> {show(b)}")
    if new_roles != [r for r in old_roles if r in new_roles] + [r for r in new_roles if r not in old_roles] and not changes:
        print("  (only the order of the rows differs)")

    if not changes:
        print("No differences: the returned matrix equals the current one.")
        return
    print(f"\n{changes} difference(s).")

    unknown_targets = [t for t in new_targets if t not in cfg.target_names]
    if unknown_targets:
        print(f"NOTE: add these columns under 'targets' in config.json before running the scripts: {', '.join(unknown_targets)}")

    if not args.write:
        print("Nothing was written. Run again with --write to replace the matrix file.")
        return
    os.makedirs(cfg.output_dir, exist_ok=True)
    backup = os.path.join(cfg.output_dir, f"matrix_{datetime.datetime.now():%Y%m%d_%H%M%S}.bak.csv")
    shutil.copy2(cfg.matrix_path, backup)
    with open(cfg.matrix_path, encoding="utf-8-sig") as f:
        first = f.readline()
    role_label = (first.split(";")[0] if first.count(";") >= first.count(",") else first.split(",")[0]).strip() or "Role"
    write_matrix(cfg.matrix_path, role_label, new_roles, new_targets, new_cells)
    print(f"Matrix file replaced: {cfg.matrix_path}\nPrevious version: {backup}")
    print("Next: 04-Sync-GroupNesting.ps1 (dry run), 01-Export-EntraSnapshot.ps1, 03-Sync-Permissions.ps1 (dry run).")


if __name__ == "__main__":
    main()
