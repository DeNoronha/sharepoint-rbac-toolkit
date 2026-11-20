# SharePoint RBAC Toolkit

Role-based access for SharePoint Online, driven by one matrix in a spreadsheet.

You decide per role which libraries or sites it may reach. The scripts make Entra ID and SharePoint match that decision, tell you where reality differs, and produce an Excel report the business owner can read and edit.

## The problem

SharePoint permissions drift. Someone shares a folder with a colleague, a site owner adds three people by hand, a group is renamed and shows up twice. After a year nobody can answer "who can open the Finance library, and why?" without clicking through dozens of permission pages.

Doing it properly by hand is slow and error-prone: every library has its own permission page, every person has to be looked up, and there is no overview.

## The idea

1. **People are only ever members of role groups.** A role group belongs to a job: `SP - UK Finance Team`, `SP - UK Site Managers`. These are ordinary Entra ID security groups.
2. **A matrix says which role reaches what.** Rows are roles, columns are libraries or sites, a mark means access. The matrix is a CSV file; the business owner sees and edits it as a tab in an Excel workbook.
3. **Nothing else gets access.** No people directly on a library, no sharing links on libraries, no leftover groups. The scripts report all of these and can remove them.

A new colleague gets access by joining a role group. A leaver loses it by leaving the group. A change in who may see what is a mark in the matrix.

## Two models

| | Direct | Nested |
|---|---|---|
| Typical use | One site with many libraries | One site per department |
| Matrix columns | Libraries | Sites |
| How access is granted | Each role group sits directly on the library, at its own permission level | Each site has one access group in its Members group; role groups are nested inside that access group |
| Config | `"model": "direct"` | `"model": "nested"` |

Both models can live side by side in one tenant, each with its own config folder.

## What is in the box

| Script | What it does | Changes anything? |
|---|---|---|
| `00-Test-Prerequisites.ps1` | Checks PowerShell, modules, Python and your config. | No |
| `01-Export-EntraSnapshot.ps1` | Reads the groups, members and users from Entra ID. Every other script builds on this snapshot. | No |
| `02-Audit-Permissions.ps1` | Compares SharePoint with the matrix and lists every difference. | No |
| `03-Sync-Permissions.ps1` | Makes SharePoint match the matrix: adds missing groups, fixes levels, removes what does not belong. | Yes, with `-Apply` |
| `04-Sync-GroupNesting.ps1` | Makes the groups in Entra match the matrix: creates missing groups, fixes nesting. | Yes, with `-Apply` |
| `05-Set-GroupMembership.ps1` | Adds people to, or removes them from, a role group. | Yes, with `-Apply` |
| `06-Audit-ItemShares.ps1` | Lists files and folders that were shared with individual people. | No |
| `07-Remove-ItemShares.ps1` | Removes the shares the owner marked "Remove". | Yes, with `-Apply` |
| `08-Remove-UserAccess.ps1` | Removes all access of someone who left. | Yes, with `-Apply` |
| `09-Grant-FolderAccess.ps1` | Gives one group access to one folder: a documented exception. | Yes, with `-Apply` |
| `10-Update-GroupTitles.ps1` | Fixes group names SharePoint still shows under an old name. | Yes, with `-Apply` |
| `report/build_report.py` | Builds the Excel report: matrix, people per role, members per group. | Writes the workbook |
| `report/import_matrix.py` | Reads the matrix tab the owner edited and shows what changed. | With `--write` |
| `report/itemshares.py` | Workbook with a Keep/Remove dropdown for individually shared items. | Writes the workbook |

## Safety by design

- **Every script that can change something is a dry run unless you add `-Apply`.** A dry run shows exactly what would happen and ends with `DRY RUN`.
- **Read-only scripts say so** in their first lines of output and have no `-Apply` switch.
- **Groups are added before anything is removed**, so nobody is locked out halfway.
- **Nothing is removed on incomplete information.** If an expected group cannot be found, that library or group is skipped and reported.
- **Site owners and visitors are never removed.** They are reported for a person to judge.
- **No passwords, keys or secrets anywhere.** Every sign-in is interactive, with your own account.
- **Every run writes a CSV log** to the `output` folder of your config.

## Quick start

Read [docs/PREREQUISITES.md](docs/PREREQUISITES.md) first. It takes fifteen minutes and saves an afternoon.

```sh
# 1. Check this computer
pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1

# 2. Make your own config from an example
cp -r examples/contoso-direct config/mycompany
#    edit config/mycompany/config.json and matrix.csv  (see docs/CONFIGURATION.md)
pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1 -ConfigPath ./config/mycompany

# 3. Read the current state. Both steps are read-only.
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath ./config/mycompany
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1    -ConfigPath ./config/mycompany

# 4. Build the report for the business owner
python3 report/build_report.py --config config/mycompany
```

From there: [docs/WORKFLOWS.md](docs/WORKFLOWS.md) describes each everyday task step by step.

Always start a script in a new process with `pwsh -NoProfile -File`. The PnP and Microsoft Graph modules cannot be loaded into the same PowerShell session.

## Working with an AI assistant

The toolkit works well with a coding assistant that can read this folder and run commands, such as Claude Code or GitHub Copilot in agent mode. `CLAUDE.md` and `.github/copilot-instructions.md` give the assistant its rules; [docs/AI-ASSISTANT.md](docs/AI-ASSISTANT.md) has ready-made prompts for the common requests.

A chat window without access to this folder cannot read the scripts and will guess. Use it to explain the documentation, not to drive the scripts.

## What it does not do

- It does not manage Microsoft 365 group or Teams membership. It works with Entra ID security groups.
- It does not decide who should have access. The matrix is a business decision; the toolkit enforces it.
- It manages permissions on sites and libraries. Individual files and folders are audited and cleaned up (scripts 06 and 07), not governed by the matrix.
- For a site, only the Members and Owners groups are managed. Custom permission structures on a site are reported, not rebuilt.
- It has no unattended mode. Certificate-based sign-in for scheduled runs is not included.

## How well is it tested?

The scripts were developed against a production tenant with both models in use. `tests/Test-Offline.ps1` checks the config, matrix and report logic without any sign-in. See [docs/TESTING.md](docs/TESTING.md) for what was verified against a live tenant and what was not.

Run the read-only scripts and the dry runs on your own tenant, and read their output, before you apply anything.

## Documentation

- [Prerequisites](docs/PREREQUISITES.md): what you need, and how to set it up
- [Configuration](docs/CONFIGURATION.md): `config.json`, the matrix, the report notes
- [Workflows](docs/WORKFLOWS.md): the everyday tasks, step by step
- [AI assistant](docs/AI-ASSISTANT.md): prompts for Claude Code and GitHub Copilot
- [Troubleshooting](docs/TROUBLESHOOTING.md): known pitfalls
- [Testing](docs/TESTING.md): what has been verified

## License

MIT. See [LICENSE](LICENSE). Use at your own risk: these scripts change permissions in your tenant when you tell them to.
