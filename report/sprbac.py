"""Shared helpers for the report scripts: read the same config.json and matrix.csv as the
PowerShell scripts, and find the latest snapshot files."""
import csv
import glob
import json
import os
import re
import sys

LEVELS = ["Read", "Contribute", "Edit", "Design", "Full Control"]
NO_MARKS = {"", "-", "0", "no", "n", "false"}


def fail(msg):
    sys.exit(f"ERROR: {msg}")


class Config:
    def __init__(self, path):
        if os.path.isdir(path):
            path = os.path.join(path, "config.json")
        if not os.path.isfile(path):
            fail(f"Config file not found: {path}")
        self.path = os.path.abspath(path)
        self.dir = os.path.dirname(self.path)
        with open(self.path, encoding="utf-8-sig") as f:
            self.raw = json.load(f)
        r = self.raw
        self.name = r["name"]
        self.model = r["model"].lower()
        self.prefix = r["groupPrefix"]
        self.default_level = r.get("defaultPermission", "Contribute")
        self.role_levels = r.get("rolePermissions", {}) or {}
        self.everyone_group = r.get("everyoneGroup", "") or ""
        self.exception_groups = list(r.get("exceptionGroups", []) or [])
        self.admin_groups = [a if isinstance(a, str) else a["group"] for a in (r.get("adminGroups", []) or [])]
        self.report = r.get("report", {}) or {}
        self.output_dir = self._abs(r.get("outputDir", "output"))
        self.matrix_path = self._abs(r.get("matrixFile", "matrix.csv"))
        self.targets = r["targets"]                       # ordered: name -> settings
        self.target_names = list(self.targets.keys())
        self.roles, self.matrix_targets, self.cells = read_matrix(self.matrix_path)

    def _abs(self, p):
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(self.dir, p))

    def access_group(self, target):
        return (self.targets.get(target) or {}).get("accessGroup") or self.prefix + target

    def roles_for(self, target):
        return [r for r in self.roles if target in self.cells[r]]

    def targets_for(self, role):
        return [t for t in self.target_names if t in self.cells.get(role, {})]

    def level(self, role, target):
        cell = self.cells[role].get(target)
        if cell and cell != "*":
            return cell
        return self.role_levels.get(role) or (self.targets.get(target) or {}).get("permission") or self.default_level

    def latest(self, name, required=True):
        files = sorted(glob.glob(os.path.join(self.output_dir, f"{name}_*.csv")))
        if not files:
            if required:
                fail(f"No {name}_*.csv in {self.output_dir}. Run the PowerShell script that writes it first.")
            return None
        return files[-1]


def read_matrix(path):
    """First column = role, other columns = targets. Returns (roles, targets, cells) where
    cells[role][target] is a permission level name or '*' (the default level)."""
    if not os.path.isfile(path):
        fail(f"Matrix file not found: {path}")
    with open(path, encoding="utf-8-sig", newline="") as f:
        text = f.read()
    first = text.splitlines()[0] if text.strip() else ""
    delim = ";" if first.count(";") > first.count(",") else ","
    rows = [r for r in csv.reader(text.splitlines(), delimiter=delim) if any(c.strip() for c in r)]
    if not rows:
        fail(f"Matrix file is empty: {path}")
    header = [c.strip() for c in rows[0]]
    while header and not header[-1]:
        header.pop()
    targets = header[1:]
    roles, cells = [], {}
    for r in rows[1:]:
        role = r[0].strip()
        if not role:
            continue
        roles.append(role)
        cells[role] = {}
        for i, t in enumerate(targets, start=1):
            v = r[i].strip() if i < len(r) else ""
            if v.lower() in NO_MARKS:
                continue
            level = next((l for l in LEVELS if l.lower() == v.lower()), None)
            cells[role][t] = level or "*"
    return roles, targets, cells


def write_matrix(path, role_label, roles, targets, cells):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([role_label] + targets)
        for role in roles:
            w.writerow([role] + [("x" if cells[role].get(t) == "*" else cells[role].get(t, "")) for t in targets])


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def short_name(display_name):
    """'Jane Doe | Contoso' -> 'Jane Doe' (some tenants append the company to the display name)."""
    return re.split(r"\s\|\s", display_name or "")[0].strip()
