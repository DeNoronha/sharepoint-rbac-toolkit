# Prerequisites

Everything you need before the first run, in the order you set it up. `00-Test-Prerequisites.ps1` checks parts 1 and 5 for you.

## Checklist

- [ ] PowerShell 7.4 or newer
- [ ] PowerShell modules `PnP.PowerShell` (3.x) and `Microsoft.Graph.Authentication` (2.x)
- [ ] Python 3.8 or newer with `xlsxwriter` and `openpyxl` (only for the Excel report)
- [ ] An app registration in Entra ID for the PnP sign-in, and its client id
- [ ] An account with the right roles (table below), and admin consent for the Graph permissions
- [ ] That account is site collection administrator on every site in your config
- [ ] Security groups that follow one naming convention, with a common prefix
- [ ] A config folder with `config.json` and `matrix.csv`

## 1. On your computer

Works on Windows, macOS and Linux.

**PowerShell 7.4 or newer.** Windows PowerShell 5.1 does not work. Install from <https://aka.ms/powershell> and check with `pwsh --version`.

**Two PowerShell modules:**

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

You only need `Microsoft.Graph.Authentication`, not the full `Microsoft.Graph` module. The toolkit talks to Graph through plain REST calls.

**Python 3 with two packages**, only for the Excel report and the workbook helpers:

```sh
python3 -m pip install xlsxwriter openpyxl
```

If you have several Python installations, the packages must be installed for the one you run. `00-Test-Prerequisites.ps1 -Python <path>` checks a specific interpreter.

**The two modules cannot share a PowerShell session.** PnP.PowerShell and Microsoft.Graph ship conflicting assemblies. Start every script in a new process:

```sh
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1 -ConfigPath ./config/mycompany
```

Do not dot-source the scripts into a session where the other module is loaded. The scripts refuse to start when they detect that.

## 2. An app registration for PnP

PnP.PowerShell needs your own app registration in Entra ID for interactive sign-in. You create it once per tenant:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint RBAC Toolkit" -Tenant contoso.onmicrosoft.com -SharePointDelegatePermissions "AllSites.FullControl"
```

- Run it as someone who may register applications and grant admin consent (Global Administrator, Application Administrator or Cloud Application Administrator).
- The command prints an **AzureAppId/ClientId**. Put it in `config.json` as `clientId`.
- `AllSites.FullControl` is a *delegated* permission: the app can do what the signed-in person can do, never more. There is no app secret and no certificate.
- If you only ever run the read-only scripts, `AllSites.Read` is enough for them. The scripts that change permissions need `AllSites.FullControl`.

Check the result in the Entra admin center under App registrations: the app should list the SharePoint delegated permission, with admin consent granted.

## 3. Your account

Every sign-in is interactive with your own account. What that account needs:

| What | Why | Needed by |
|---|---|---|
| Read groups and users in Entra ID | Take the snapshot | 01 |
| **Groups Administrator** (or owner of every group involved) | Create security groups, change members and nesting | 04, 05, 08 (step Groups) |
| **Site collection administrator on every site in the config** | Read and change permissions on the site, its libraries and items | 02, 03, 06, 07, 08 (step Sites), 09, 10 |
| Consent for the app registration from part 2 | Sign in through PnP | All SharePoint scripts |

Two things that surprise people:

- **The SharePoint Administrator role does not give you access to a site's content or permissions.** It lets you make yourself site collection administrator, in the SharePoint admin center under Active sites, the site, Membership, Site admins. Do that for each site in the config. Without it the scripts report "unauthorized" for that site.
- **Read-only is enough to start.** Scripts 01, 02 and 06 change nothing. You can audit a tenant with read rights and decide later who applies changes.

## 4. Consent for Microsoft Graph

The Entra scripts sign in through the Microsoft Graph Command Line Tools enterprise application and ask for these delegated permissions:

| Permission | Used by |
|---|---|
| `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All` | 01, and the dry runs of 04 |
| `GroupMember.ReadWrite.All` | 05, 08 |
| `Group.ReadWrite.All` | 04 with `-Apply` (creating groups, nesting) |

The first time, Entra asks for consent. An administrator has to approve these permissions once for the organization; after that the sign-in is silent for everyone who is allowed to use it.

The sign-in is forced to the tenant of your `tenantUrl` (for `https://contoso.sharepoint.com` that is `contoso.onmicrosoft.com`). If you are signed in to several tenants in your browser, pick the account of the right one. If your tenant name differs from its SharePoint host name, set `tenantId` in the config.

## 5. Groups and naming

The toolkit recognizes its groups by name:

- **Role groups** are named `<groupPrefix><role>`, where the role is the text in the first column of the matrix. With prefix `SP - UK ` and role `Finance Team` the group is `SP - UK Finance Team`.
- **Access groups** (nested model only) are named `<groupPrefix><column name>`, unless the config sets another name for a target.
- **Group names must be unique** in the tenant. The scripts stop when two groups share a name.
- Use **security groups**, not Microsoft 365 groups. Security groups can be nested and carry no mailbox, Team or site of their own.

You do not have to create the groups by hand. `04-Sync-GroupNesting.ps1 -CreateMissingGroups` creates every group the matrix asks for.

Pick a prefix that nothing else in the tenant uses. Everything that starts with the prefix is treated as an authorization group: it appears in the report, and a leaver is removed from it.

## 6. Your config

Copy one of the examples and edit it:

```sh
cp -r examples/contoso-direct config/mycompany
```

See [CONFIGURATION.md](CONFIGURATION.md) for every setting. The `config` folder is in `.gitignore` on purpose: your configuration names your sites, your groups and your people.

## 7. Verify

```sh
pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1 -ConfigPath ./config/mycompany
pwsh -NoProfile -File ./tests/Test-Offline.ps1
```

Then the two read-only runs that prove sign-in and permissions work:

```sh
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath ./config/mycompany
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1    -ConfigPath ./config/mycompany
```

If both finish without errors, everything is in place.

## Sign-in behaviour

- A sign-in opens a window in your default browser. It can open behind other windows. A script that seems to hang right after "Signing in" is almost always waiting for that window.
- The SharePoint sign-in is remembered in the credential store of your operating system (`"persistLogin": true`, the default), so later runs are silent. To forget it: `Disconnect-PnPOnline -ClearPersistedLogin`. Set `"persistLogin": false` on a shared computer.
- On a machine without a browser, set `"pnpAuth": "DeviceLogin"` in the config.
- If your organization requires re-authentication after a while, the browser window comes back. Sign in again; nothing is lost.
