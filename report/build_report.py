#!/usr/bin/env python3
"""Builds the Excel access report from the latest Entra snapshot and the access matrix.

    python3 report/build_report.py --config config/contoso

Reads   <config>/output/entra-groups_*.csv and entra-users_*.csv (01-Export-EntraSnapshot.ps1),
        the matrix file and, when configured, the report notes file.
Writes  <report.outputDir>/<report.fileName> <YYYYMMDD>.xlsx

Tabs
    <matrix sheet>       the access matrix: the only tab the owner edits
    User → Role Groups   per person: job title, role groups, and what those give access to
    Groups → Members     per group: the members as they are in Entra now
    Users → Groups       per person: every authorization group
    Summary              counts, plus the rows under "summary" in the notes file
    No Access            people who deliberately have no access (from the notes file)
    Notes                how-to and changelog (from the notes file)

Needs: pip install xlsxwriter
"""
import argparse
import datetime
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sprbac import Config, fail, read_csv, short_name  # noqa: E402

try:
    import xlsxwriter
except ImportError:
    fail("Python package xlsxwriter is missing. Install it with: python3 -m pip install xlsxwriter")


def main():
    ap = argparse.ArgumentParser(description="Build the Excel access report.")
    ap.add_argument("--config", required=True, help="config.json or the folder that contains it")
    ap.add_argument("--out-dir", help="write the report here instead of report.outputDir")
    ap.add_argument("--date", help="date for the file name, YYYYMMDD (default: today)")
    args = ap.parse_args()

    cfg = Config(args.config)
    rep = cfg.report
    today = datetime.date.today()
    stamp = args.date or today.strftime("%Y%m%d")
    refreshed = today.strftime("%Y-%m-%d")
    sheet_matrix = rep.get("matrixSheet", "Access Matrix")
    target_label = rep.get("targetLabel", "Library" if cfg.model == "direct" else "Department")
    consolidate = bool(rep.get("consolidateByName", False))
    include_disabled = bool(rep.get("includeDisabled", False))

    groups_csv = cfg.latest("entra-groups")
    users_csv = cfg.latest("entra-users")
    print(f"Using:\n  {os.path.basename(groups_csv)}\n  {os.path.basename(users_csv)}\n  {os.path.basename(cfg.matrix_path)}")

    notes = {"summary": [], "notes": [], "noAccess": []}
    notes_file = rep.get("notesFile")
    if notes_file:
        p = notes_file if os.path.isabs(notes_file) else os.path.join(cfg.dir, notes_file)
        if not os.path.isfile(p):
            fail(f"report.notesFile not found: {p}")
        with open(p, encoding="utf-8-sig") as f:
            notes.update(json.load(f))

    # ── Data ──────────────────────────────────────────────────────────
    prefix = cfg.prefix
    in_scope = lambda g: g.startswith(prefix) or g in cfg.exception_groups  # noqa: E731
    rows = [r for r in read_csv(groups_csv) if in_scope(r["GroupName"])]
    if not include_disabled:
        rows = [r for r in rows if not (r["MemberType"] == "User" and r["AccountEnabled"] == "False")]
    all_groups = sorted({r["GroupName"] for r in read_csv(groups_csv) if in_scope(r["GroupName"])})
    with_rows = {r["GroupName"] for r in rows if r["MemberName"] != "(empty)"}
    descriptions = {r["GroupName"]: r["GroupDescription"] for r in read_csv(groups_csv)}
    group_rows = [r for r in rows if r["MemberName"] != "(empty)"]
    group_rows += [{"GroupName": g, "GroupDescription": descriptions.get(g, ""), "MemberName": "(empty)", "MemberUPN": "", "MemberType": "", "AccountEnabled": ""}
                   for g in all_groups if g not in with_rows]
    group_rows.sort(key=lambda r: (r["GroupName"], r["MemberName"]))

    titles = {u["UserPrincipalName"].lower(): (u.get("JobTitle") or "").strip() for u in read_csv(users_csv)}
    role_of = {prefix + r: r for r in cfg.roles}

    people = {}   # key -> {name, upns, enabled, groups}
    for r in rows:
        if r["MemberType"] != "User" or not r["GroupName"].startswith(prefix):
            continue
        name = short_name(r["MemberName"])
        key = name.lower() if consolidate else r["MemberUPN"].lower()
        p = people.setdefault(key, {"name": r["MemberName"], "short": name, "upns": set(), "enabled": r["AccountEnabled"], "groups": set()})
        p["upns"].add(r["MemberUPN"])
        p["groups"].add(r["GroupName"])

    user_rows = []
    for p in people.values():
        roles = sorted(role_of[g] for g in p["groups"] if g in role_of)
        if not roles:
            continue
        reach = sorted({t for r in roles for t in cfg.targets_for(r)})
        title = next((titles[u.lower()] for u in sorted(p["upns"]) if titles.get(u.lower())), "")
        user_rows.append([p["short"], title, "; ".join(sorted(p["upns"])), "; ".join(roles), "; ".join(reach), len(reach) + len(roles)])
    user_rows.sort(key=lambda r: r[0].lower())

    # ── Workbook ──────────────────────────────────────────────────────
    out_dir = args.out_dir or rep.get("outputDir", "reports")
    if not os.path.isabs(out_dir):
        out_dir = os.path.normpath(os.path.join(cfg.dir if not args.out_dir else os.getcwd(), out_dir))
    os.makedirs(out_dir, exist_ok=True)
    base = rep.get("fileName", f"SharePoint-Access-Report {cfg.name}")
    out = os.path.join(out_dir, f"{base} {stamp}.xlsx")

    # strings_to_urls off: Excel shows a "repair" prompt on workbooks with many URL-like strings
    wb = xlsxwriter.Workbook(out, {"strings_to_numbers": False, "strings_to_urls": False, "strings_to_formulas": False, "in_memory": True})
    f_head = wb.add_format({"bold": True, "bg_color": "#2F5496", "font_color": "#FFFFFF", "font_name": "Calibri", "font_size": 11, "align": "left", "valign": "vcenter"})
    f_src = wb.add_format({"italic": True, "font_color": "#595959", "bg_color": "#E7E6E6", "font_name": "Calibri", "font_size": 9, "valign": "vcenter"})
    f_check = wb.add_format({"bold": True, "font_color": "#2F5496", "bg_color": "#D9E1F2", "font_name": "Calibri", "font_size": 11, "align": "center"})
    f_center = wb.add_format({"align": "center"})
    f_cell = wb.add_format({"font_name": "Calibri", "font_size": 11})

    def header(ws, heads, widths):
        for i, h in enumerate(heads):
            ws.write(0, i, h, f_head)
        for i, w in enumerate(widths):
            ws.set_column(i, i, w)
        ws.set_row(0, 22)
        ws.freeze_panes(1, 0)

    def source(ws, ncols, text):
        ws.merge_range(1, 0, 1, ncols - 1, text, f_src)

    # 1. Matrix
    ws = wb.add_worksheet(sheet_matrix)
    header(ws, [f"Role Group ↓ / {target_label} →"] + cfg.target_names, [32] + [14] * len(cfg.target_names))
    ws.freeze_panes(1, 1)
    for ri, role in enumerate(cfg.roles, start=1):
        ws.write(ri, 0, role, f_cell)
        for ci, t in enumerate(cfg.target_names, start=1):
            cell = cfg.cells[role].get(t)
            if cell:
                ws.write(ri, ci, "✓" if cell == "*" else cell, f_check)
            else:
                ws.write(ri, ci, "", f_center)

    # 2. User → Role Groups
    ws = wb.add_worksheet("User → Role Groups")
    header(ws, ["User Name", "Job Title", "UPN", "Role Group(s)", f"Inherited {target_label}s" if not target_label.endswith("y") else f"Inherited {target_label[:-1]}ies", "Total Access"], [28, 28, 44, 50, 80, 12])
    source(ws, 6, f"Source: Entra snapshot {os.path.basename(groups_csv)} (refreshed {refreshed}) + {sheet_matrix} tab.")
    for i, r in enumerate(user_rows, start=2):
        ws.write_row(i, 0, r)

    # 3. Groups → Members
    ws = wb.add_worksheet("Groups → Members")
    header(ws, ["Group Name", "Description", "Member Name", "Member UPN", "Member Type", "Account Enabled"], [38, 50, 38, 44, 15, 18])
    source(ws, 6, f"Source: Entra snapshot {os.path.basename(groups_csv)} (refreshed {refreshed})")
    for i, r in enumerate(group_rows, start=2):
        ws.write_row(i, 0, [r["GroupName"], r["GroupDescription"], r["MemberName"], r["MemberUPN"], r["MemberType"], r["AccountEnabled"]])

    # 4. Users → Groups
    ws = wb.add_worksheet("Users → Groups")
    header(ws, ["User Name", "UPN(s)" if consolidate else "UPN", "Account Enabled", "SP Group Count", "SP Groups"], [32, 44, 15, 14, 100])
    source(ws, 5, f"Source: Entra snapshot {os.path.basename(groups_csv)} (refreshed {refreshed}). Only groups starting with '{prefix.strip()}' are shown.")
    for i, p in enumerate(sorted(people.values(), key=lambda p: p["name"].lower()), start=2):
        ws.write_row(i, 0, [p["short"] if consolidate else p["name"], "; ".join(sorted(p["upns"])), p["enabled"], len(p["groups"]), "; ".join(sorted(p["groups"]))])

    # 5. Summary
    ws = wb.add_worksheet("Summary")
    header(ws, ["Metric", "Value"], [44, 90])
    prefix_groups = [g for g in all_groups if g.startswith(prefix)]
    empty_roles = sorted(role_of[g] for g in prefix_groups if g in role_of and g not in with_rows)
    missing_roles = sorted(r for r in cfg.roles if prefix + r not in all_groups)
    summary = [
        ("Report Date", refreshed),
        ("", ""),
        (f"Active '{prefix.strip()}' groups", len(prefix_groups)),
        ("Role groups (rows in the matrix)", len(cfg.roles)),
        (f"{target_label}s (columns in the matrix)" if not target_label.endswith("y") else f"{target_label[:-1]}ies (columns in the matrix)", len(cfg.target_names)),
        ("Unique users (in role groups)", len(user_rows)),
        ("Empty role groups", ", ".join(empty_roles) if empty_roles else "(none)"),
    ]
    if missing_roles:
        summary.append(("Roles without a group in Entra", ", ".join(missing_roles)))
    if notes.get("summary"):
        summary.append(("", ""))
        summary += [tuple(r) for r in notes["summary"]]
    for i, (k, v) in enumerate(summary, start=1):
        ws.write_row(i, 0, [k, v])

    # 6. No Access
    if notes.get("noAccess"):
        ws = wb.add_worksheet("No Access")
        header(ws, ["User Name", "UPN", "Job Title", "Reason"], [30, 46, 40, 60])
        for i, r in enumerate(notes["noAccess"], start=1):
            ws.write_row(i, 0, [r.get("name", ""), r.get("upn", ""), r.get("jobTitle", ""), r.get("reason", "")])

    # 7. Notes
    if notes.get("notes"):
        ws = wb.add_worksheet("Notes")
        header(ws, ["Principle", "Description"], [40, 110])
        for i, r in enumerate(notes["notes"], start=1):
            ws.write_row(i, 0, list(r))

    wb.close()

    # Move older versions of this report out of the way
    archive = rep.get("archiveDir")
    if archive and not args.out_dir:
        adir = archive if os.path.isabs(archive) else os.path.join(out_dir, archive)
        for name in sorted(os.listdir(out_dir)):
            if name.startswith(base + " ") and name.endswith(".xlsx") and name != os.path.basename(out) and name[len(base) + 1:-5].isdigit():
                os.makedirs(adir, exist_ok=True)
                shutil.move(os.path.join(out_dir, name), os.path.join(adir, name))
                print(f"Archived: {name}")

    print(f"Saved: {out}")
    print(f"  {len(user_rows)} users in role groups, {len(prefix_groups)} groups, {len(cfg.roles)} roles x {len(cfg.target_names)} {target_label.lower()} columns")
    if empty_roles:
        print(f"  Empty role groups: {', '.join(empty_roles)}")
    if missing_roles:
        print(f"  WARNING: roles without a group in Entra: {', '.join(missing_roles)}")


if __name__ == "__main__":
    main()
