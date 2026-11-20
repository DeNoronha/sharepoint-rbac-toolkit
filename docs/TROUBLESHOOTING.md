# Troubleshooting

## A script hangs right after "Signing in"

It is waiting for you. The sign-in opened a window in your default browser, possibly behind other windows or in another desktop space. Complete the sign-in there.

- If your organization requires re-authentication or MFA after some time, the window comes back even when earlier runs were silent.
- Stop the script before you start it again. Every attempt opens another window.
- No browser on this machine: set `"pnpAuth": "DeviceLogin"` in the config.

## "No groups found" or the wrong account is shown

You are signed in to another tenant. The scripts force the tenant derived from `tenantUrl` (`https://contoso.sharepoint.com` gives `contoso.onmicrosoft.com`). If your tenant's name is different, set `tenantId` in the config to the tenant id or a verified domain.

To clear a remembered sign-in:

```powershell
Disconnect-MgGraph                          # Microsoft Graph
Disconnect-PnPOnline -ClearPersistedLogin   # SharePoint
```

## "Could not load file or assembly" or a script returns nothing

PnP.PowerShell and Microsoft.Graph were loaded into the same PowerShell session. Their assemblies conflict; the result is a load error or, worse, an empty result without an error.

Start every script in a fresh process: `pwsh -NoProfile -File ./scripts/<script>.ps1`. The `-NoProfile` matters when your profile imports one of the modules.

## "Attempted to perform an unauthorized operation" on one site

Your account is not a site collection administrator of that site. The SharePoint Administrator role does not include access to sites; add yourself in the SharePoint admin center (Active sites, the site, Membership, Site admins).

## MISSING_GROUP in the audit

The matrix asks for a group that is not in the snapshot.

- New role: run `04-Sync-GroupNesting.ps1 -CreateMissingGroups -Apply`, then take a snapshot.
- The group exists but is named differently: the name must be exactly `groupPrefix` plus the role name in the matrix, spaces included.
- The snapshot is old: run `01-Export-EntraSnapshot.ps1`.

`03-Sync-Permissions.ps1` skips a target with a missing group on purpose. Removing access while the picture is incomplete would be guessing.

## "Group not resolvable yet"

A group created a few minutes ago is not known to SharePoint yet. The scripts retry for a few minutes. If it still fails, wait ten minutes and run the script again; it continues where it stopped.

## The same group appears under two names

SharePoint keeps the display name a group had when it was first used on a site. After a rename in Entra the site shows the old name, and the audit shows the current one. It is one group. Run `10-Update-GroupTitles.ps1`.

The scripts always match groups on object id, never on the name SharePoint shows.

## A person still has access after being removed from a group

- Group membership changes take up to an hour to reach SharePoint. Signing out and in again speeds it up.
- They may have access another way: a second role group, a direct share on a file or folder (`06-Audit-ItemShares.ps1`), or membership of the site's Owners or Visitors group (the REVIEW rows of `02-Audit-Permissions.ps1`).

## A person has no access after being added

- Wait an hour, or have them sign out and in.
- Check they were added to the role group and not to an access group.
- Nested model: check that the role group is nested in the access group (`04-Sync-GroupNesting.ps1`, dry run) and that the access group is in the site's Members group (`02-Audit-Permissions.ps1`).

## Several addresses or targets with `pwsh -File`

`pwsh -File` passes arguments as plain text. Give lists as one comma-separated value without spaces:

```sh
pwsh -NoProfile -File ./scripts/05-Set-GroupMembership.ps1 -ConfigPath CFG -Role "Finance Team" -Add jane@contoso.com,joe@contoso.com
```

Quote names that contain spaces: `-Target "Human Resources",Finance`.

## The audit reports REVIEW rows

These are things the toolkit does not manage and will not change:

- members of a site's Owners group other than the admin groups
- anyone in a site's Visitors group
- access granted on the site itself, outside the Owners, Members and Visitors groups

Judge each one. Remove by hand what should not be there, or list what should stay under `keepPrincipals` in the config so the audit reports it as KEEP.

## Excel offers to repair the workbook

The report is written with `strings_to_urls` off because Excel shows a repair prompt on workbooks with many URL-like strings. If you still see the prompt, check that you are opening the file the script just wrote and not an older copy, and that `xlsxwriter` is current (`python3 -m pip install -U xlsxwriter`).

## "ModuleNotFoundError: No module named 'xlsxwriter'"

The package is installed for another Python than the one you ran. Install it for this one (`python3 -m pip install xlsxwriter openpyxl`) or call the interpreter that has it.

## A library has thousands of items and script 06 is slow

Every item with its own permissions costs one request, and listing the items of a large library takes time by itself. Run it per target (`-Target`) and let it finish; it is read-only and can be repeated safely.
