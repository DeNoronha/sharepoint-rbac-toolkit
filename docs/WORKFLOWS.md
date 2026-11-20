# Workflows

The everyday tasks, step by step. `CFG` stands for your config folder, for example `./config/mycompany`.

Three habits apply to every task:

1. **Dry run first.** Run the script without `-Apply`, read what it would do, then run it again with `-Apply`.
2. **Read back.** After a change, run the read-only script again and look at the actual state before you tell anyone it is done.
3. **Refresh the snapshot** after every change in Entra. The SharePoint scripts and the report read the snapshot, not Entra itself.

Start each command in a new process: `pwsh -NoProfile -File ...`.

## Who may ask for what

Decide this once and write it down. A model that works:

| Request | Who may ask |
|---|---|
| Add or remove a person, change someone's role | The business owner of that environment |
| Change the matrix: marks, a new role, a new library or site | The business owner, by sending back the edited report |
| An exception outside the matrix (one folder, an external person) | Only with explicit approval of whoever owns the model |

A request from anyone else, such as an employee asking for access to a library, is not carried out. It goes to the business owner, who decides whether it is a role change or a matrix change.

## First-time setup

```sh
pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1 -ConfigPath CFG
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1   -ConfigPath CFG                              # which groups are missing?
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1   -ConfigPath CFG -CreateMissingGroups -Apply   # create them, set nesting
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1    -ConfigPath CFG                              # where does SharePoint differ?
```

Put people in the role groups (`05-Set-GroupMembership.ps1`) **before** you sync SharePoint. Otherwise the sync removes people's direct access while their group is still empty.

Then sync one target, check it, and only then the rest:

```sh
pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1 -ConfigPath CFG -Target Finance            # dry run
pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1 -ConfigPath CFG -Target Finance -Apply
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1 -ConfigPath CFG -Target Finance           # confirm
```

Two switches make a careful first sync possible:

- `-AddOnly` adds and fixes, and removes nothing. Use it to put the groups in place while the old access still works.
- `-KeepDirectUsers` leaves people with direct access alone, so you can clean those up later, one conversation at a time.

## Add a person to a role

```sh
pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath CFG -Role "Finance Team" -Add jane@contoso.com
pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath CFG -Role "Finance Team" -Add jane@contoso.com -Apply
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG
python3 report/build_report.py --config CFG
```

Nothing changes in SharePoint: the group already has the access. The person can use it within about an hour; signing out and in again speeds it up.

## Someone changes role

Add to the new role group, remove from the old one. Two runs of script 05, one with `-Add` and one with `-Remove`. Then the snapshot and the report.

## Someone leaves

```sh
pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath CFG -User jane@contoso.com -Step Groups
pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath CFG -User jane@contoso.com -Step Groups -Apply
pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath CFG -User jane@contoso.com -Step Sites
pwsh -NoProfile -File ./scripts/08-Remove-UserAccess.ps1 -ConfigPath CFG -User jane@contoso.com -Step Sites -Apply
```

`Groups` takes the account out of every authorization group. `Sites` removes it from each site in the config, which drops everything the person held there directly: SharePoint groups, direct permissions, sharing links.

Then the snapshot and the report.

## The matrix changes

The business owner edits the matrix tab of the report and sends the workbook back.

```sh
python3 report/import_matrix.py --config CFG "returned report.xlsx"            # shows every difference, writes nothing
python3 report/import_matrix.py --config CFG "returned report.xlsx" --write    # replaces matrix.csv, keeps a backup
```

Check the list of differences with the owner if anything in it surprises you. Then:

**Direct model**

```sh
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1    -ConfigPath CFG                     # new roles need a group
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG
pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1     -ConfigPath CFG                     # dry run
pwsh -NoProfile -File ./scripts/03-Sync-Permissions.ps1     -ConfigPath CFG -Apply
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1    -ConfigPath CFG                     # must report no differences
```

**Nested model**

```sh
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1    -ConfigPath CFG                     # dry run: nesting to add and remove
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1    -ConfigPath CFG -Apply
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG
```

In the nested model a changed mark is a change in Entra only. SharePoint needs a sync only when a site is new.

Finish with the report, and add a line to the changelog in `report-notes.json`.

## A new role

Add a row to the matrix with its marks. `04-Sync-GroupNesting.ps1 -CreateMissingGroups -Apply` creates the group. Add the people with script 05. Then continue as under "The matrix changes".

## A new library or site

Add a column to the matrix and an entry under `targets` in `config.json`. Create the library or site in SharePoint yourself; the toolkit does not create those. Then continue as under "The matrix changes". A new library gets its own permissions on the first sync.

## Someone needs one folder only

First ask whether an existing role would do: then it is a matrix or membership question. If not, it is an exception. Get approval, and keep to the rule that access goes to a group:

1. Create a security group for the exception and add the people (Entra admin center, or `05-Set-GroupMembership.ps1 -Group "<name>"`).
2. Add the group name to `exceptionGroups` in the config and take a snapshot.
3. Grant the group the folder:

```sh
pwsh -NoProfile -File ./scripts/09-Grant-FolderAccess.ps1 -ConfigPath CFG -Site /sites/Operations -Folder "/sites/Operations/Shared Documents/Board meeting" -Group "SP - Board Meeting Guests"
```

4. Write it down in the report notes: who asked, why, which group. From now on, permission changes on the library no longer reach this folder.

## Clean up individually shared files and folders

```sh
pwsh -NoProfile -File ./scripts/06-Audit-ItemShares.ps1 -ConfigPath CFG -Target Finance,Board
python3 report/itemshares.py build --config CFG
```

Send the workbook to the business owner. Per row it says whether the person has access through a group anyway. The owner picks Keep or Remove.

```sh
python3 report/itemshares.py decisions --config CFG "decided workbook.xlsx"
pwsh -NoProfile -File ./scripts/07-Remove-ItemShares.ps1 -ConfigPath CFG            # dry run
pwsh -NoProfile -File ./scripts/07-Remove-ItemShares.ps1 -ConfigPath CFG -Apply
pwsh -NoProfile -File ./scripts/06-Audit-ItemShares.ps1 -ConfigPath CFG -Target Finance,Board   # confirm what is left
```

Rows without a decision are left alone.

## A group was renamed in Entra

Access keeps working, but SharePoint shows the old name.

```sh
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG
pwsh -NoProfile -File ./scripts/10-Update-GroupTitles.ps1   -ConfigPath CFG
pwsh -NoProfile -File ./scripts/10-Update-GroupTitles.ps1   -ConfigPath CFG -Apply
```

If the rename changes a role name, change the row in the matrix too: the group name is always prefix plus role.

## Periodic review

Once a quarter, or before an audit:

```sh
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath CFG    # also lists disabled accounts still in groups
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1    -ConfigPath CFG
pwsh -NoProfile -File ./scripts/04-Sync-GroupNesting.ps1    -ConfigPath CFG
pwsh -NoProfile -File ./scripts/06-Audit-ItemShares.ps1     -ConfigPath CFG
python3 report/build_report.py --config CFG
```

All read-only or dry run. Send the report to the business owner and ask them to confirm the matrix and the people per role.
