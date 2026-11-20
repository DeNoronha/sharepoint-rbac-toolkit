#!/usr/bin/env python3
"""Turns the list of individually shared files and folders into a workbook the owner can
decide in, and turns the decided workbook back into a file for the removal script.

    python3 report/itemshares.py build --config config/contoso
        latest itemshares_*.csv (06-Audit-ItemShares.ps1)  ->  workbook with a Keep/Remove dropdown

    python3 report/itemshares.py decisions --config config/contoso "decided workbook.xlsx"
        decided workbook  ->  itemshares-decisions_<timestamp>.csv for 07-Remove-ItemShares.ps1

The technical columns (site, list id, item id, login name) travel along in hidden columns,
so the removal script can find each share again. Do not delete or reorder rows' hidden cells.

Needs: pip install xlsxwriter openpyxl
"""
import argparse
import csv
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sprbac import Config, fail, read_csv  # noqa: E402

VISIBLE = [("Target", 16), ("Library", 16), ("Type", 8), ("Path", 70), ("SharedWith", 26), ("Account", 44),
           ("How", 22), ("Access", 12), ("AlreadyHasAccessViaGroup", 40), ("Decision", 22)]
HEADS = {"Path": "File or folder", "SharedWith": "Shared with", "AlreadyHasAccessViaGroup": "Has access through a group anyway?",
         "Decision": "Decision (Keep / Remove)"}
HIDDEN = ["SiteUrl", "ListId", "ItemId", "LoginName", "LinkGroup"]


def build(cfg, args):
    try:
        import xlsxwriter
    except ImportError:
        fail("Python package xlsxwriter is missing. Install it with: python3 -m pip install xlsxwriter")
    src = cfg.latest("itemshares")
    rows = read_csv(src)
    out_dir = args.out_dir or cfg.report.get("outputDir", "reports")
    if not os.path.isabs(out_dir):
        out_dir = os.path.normpath(os.path.join(cfg.dir if not args.out_dir else os.getcwd(), out_dir))
    os.makedirs(out_dir, exist_ok=True)
    base = cfg.report.get("itemSharesFileName", f"Individually-Shared-Items {cfg.name}")
    out = os.path.join(out_dir, f"{base} {datetime.date.today():%Y%m%d}.xlsx")
    wb = xlsxwriter.Workbook(out, {"strings_to_numbers": False, "strings_to_urls": False, "strings_to_formulas": False, "in_memory": True})
    ws = wb.add_worksheet("Shared items")
    f_head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "top"})
    f_cell = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    f_dec = wb.add_format({"text_wrap": True, "valign": "top", "border": 1, "bg_color": "#FFF2CC"})
    cols = [c for c, _ in VISIBLE] + HIDDEN
    for i, c in enumerate(cols):
        ws.write(0, i, HEADS.get(c, c), f_head)
    for i, (_, w) in enumerate(VISIBLE):
        ws.set_column(i, i, w)
    ws.set_column(len(VISIBLE), len(cols) - 1, 12, None, {"hidden": True})
    dec = cols.index("Decision")
    for ri, r in enumerate(rows, start=1):
        for ci, c in enumerate(cols):
            ws.write(ri, ci, r.get(c, ""), f_dec if ci == dec else f_cell)
    ws.freeze_panes(1, 0)
    if rows:
        ws.autofilter(0, 0, len(rows), len(VISIBLE) - 1)
        ws.data_validation(1, dec, len(rows), dec, {"validate": "list", "source": ["Keep", "Remove"]})
    wb.close()
    print(f"Using {os.path.basename(src)}\nSaved: {out} ({len(rows)} rows)")
    for r in rows:
        if r.get("AlreadyHasAccessViaGroup", "").startswith("No"):
            print(f"  loses access if removed: {r['Library']} | ...{r['Path'][-50:]} | {r['SharedWith']}")


def decisions(cfg, args):
    try:
        import openpyxl
    except ImportError:
        fail("Python package openpyxl is missing. Install it with: python3 -m pip install openpyxl")
    if not os.path.isfile(args.xlsx):
        fail(f"File not found: {args.xlsx}")
    ws = openpyxl.load_workbook(args.xlsx, data_only=True).worksheets[0]
    data = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    back = {v: k for k, v in HEADS.items()}
    header = [back.get(h, h) for h in data[0]]
    needed = ["Decision"] + HIDDEN
    missing = [c for c in needed if c not in header]
    if missing:
        fail(f"The workbook has no column(s): {', '.join(missing)}. Use a workbook made by 'itemshares.py build'.")
    rows = [dict(zip(header, r)) for r in data[1:] if any(r)]
    bad = sorted({r["Decision"] for r in rows if r["Decision"] not in ("", "Keep", "Remove")})
    if bad:
        fail(f"Unknown decision value(s): {', '.join(bad)}. Use Keep or Remove.")
    os.makedirs(cfg.output_dir, exist_ok=True)
    out = os.path.join(cfg.output_dir, f"itemshares-decisions_{datetime.datetime.now():%Y%m%d_%H%M%S}.csv")
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=header, quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(rows)
    count = lambda v: sum(1 for r in rows if r["Decision"] == v)  # noqa: E731
    print(f"{len(rows)} rows: {count('Remove')} Remove, {count('Keep')} Keep, {count('')} undecided")
    print(f"Saved: {out}\nNext: 07-Remove-ItemShares.ps1 (dry run first).")


def main():
    ap = argparse.ArgumentParser(description="Workbook for individually shared items.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build"); b.add_argument("--config", required=True); b.add_argument("--out-dir")
    d = sub.add_parser("decisions"); d.add_argument("--config", required=True); d.add_argument("xlsx")
    args = ap.parse_args()
    cfg = Config(args.config)
    (build if args.cmd == "build" else decisions)(cfg, args)


if __name__ == "__main__":
    main()
