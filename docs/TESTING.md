# Testing

What has been verified, and how, so you know how far to trust each part before you run it on your own tenant.

## Without a tenant

`tests/Test-Offline.ps1` needs no sign-in and no network. It checks:

- every script parses, has a help block, and every changing script has the `-Apply` switch (and every read-only script does not)
- config and matrix parsing for both models, including permission levels per cell, per role and per target
- validation: wrong model, a matrix column without a target, an unknown role, a client id that is not a GUID
- classification of SharePoint principals: Entra groups, Microsoft 365 groups, users, guests, Everyone claims, sharing links, SharePoint groups, system entries
- permission levels by type, so a site in another language gives the same result
- the report: built from a made-up snapshot, names shortened, disabled accounts left out, all tabs present
- the matrix round trip: the matrix tab of a built report imports back without differences

Run it after every change to the scripts:

```sh
pwsh -NoProfile -File ./tests/Test-Offline.ps1
```

## Against a live tenant

Verified on a production Microsoft 365 tenant that uses both models, by comparing the output with independent earlier exports and audits of the same tenant.

| Part | Status | How it was verified |
|---|---|---|
| `01-Export-EntraSnapshot.ps1` | Verified | 51 groups, 184 memberships: identical, row for row, to an export made earlier the same day with a different script. |
| `02-Audit-Permissions.ps1`, library targets | Verified on 15 of 16 libraries | No differences reported, matching an earlier independent audit. The sixteenth library, which has a guest with direct access, exposed a bug that is fixed; the re-run is pending. |
| `report/build_report.py` | Verified | Both reports compared cell by cell with the reports from the previous generation of scripts: every data tab identical. |
| `report/import_matrix.py` | Verified | The matrix tabs of both production reports import without differences. |
| `00-Test-Prerequisites.ps1` | Verified | Run against both production configs and both examples. |
| `02-Audit-Permissions.ps1`, site targets | Not yet verified | |
| `03-Sync-Permissions.ps1` | Not yet verified | |
| `04-Sync-GroupNesting.ps1` | Not yet verified | |
| `05-Set-GroupMembership.ps1` | Not yet verified | |
| `06-Audit-ItemShares.ps1` | Not yet verified | |
| `07-Remove-ItemShares.ps1` | Not yet verified | |
| `08-Remove-UserAccess.ps1` | Not yet verified | |
| `09-Grant-FolderAccess.ps1` | Not yet verified | |
| `10-Update-GroupTitles.ps1` | Not yet verified | |
| `report/itemshares.py` | Not yet verified | |

"Not yet verified" means the script was written from proven building blocks of its predecessor but has not itself been run against a tenant. Treat its first dry run on your tenant as a test.

## Test your own setup

Before you apply anything:

1. `00-Test-Prerequisites.ps1 -ConfigPath <config>` and `tests/Test-Offline.ps1`.
2. `01-Export-EntraSnapshot.ps1`: open the two CSV files and check that the groups and members are what you expect.
3. `02-Audit-Permissions.ps1`: pick one library or site you know well and check its rows against the permission page in SharePoint.
4. The dry run of `03-Sync-Permissions.ps1 -Target <that one>`: read every line. If a line surprises you, find out why before you go on.
5. Apply to that one target, run the audit again, and ask someone who should have access, and someone who should not, to try.
