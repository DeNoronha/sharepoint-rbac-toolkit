# Configuration

One folder per environment. A folder holds three files you maintain and one folder the scripts fill:

```
config/mycompany/
  config.json          settings: tenant, model, groups, targets
  matrix.csv           who reaches what
  report-notes.json    optional: texts for the Excel report
  output/              snapshots and logs, written by the scripts
```

Every script takes `-ConfigPath` (PowerShell) or `--config` (Python): the folder or the `config.json` inside it.

If you have two populations with different rules, for example two countries, give each its own folder. They can share a tenant.

## matrix.csv

The first column holds the roles, the other columns are the targets. A target is a library (direct model) or a site (nested model).

```csv
Role;Procedures;Customers;Finance;Board
General Manager;x;x;x;x
Finance Team;x;x;x;
Customer Service;x;x;;
Auditor;Read;Read;Read;
```

- **An empty cell means no access.** So do `-`, `0`, `no` and `false`.
- **Any other mark means access** at the default permission level: `x`, `yes`, `1`, a tick.
- **A permission level name sets that level** for that one cell: `Read`, `Contribute`, `Edit`, `Design`, `Full Control`.
- Separator `;` or `,`: both work, the script looks at the header line.
- Every column must have an entry under `targets` in `config.json`.
- The role name plus `groupPrefix` is the group name. Role `Finance Team` with prefix `SP - UK ` is the group `SP - UK Finance Team`.

Edit the file in a text editor or in Excel. When a business owner edits the matrix tab of the report instead, `report/import_matrix.py` brings the changes back into this file.

## config.json

### Required

| Setting | Meaning |
|---|---|
| `name` | Short name of this environment. Used in output and in the default report name. |
| `tenantUrl` | Root of your SharePoint tenant, without a path: `https://contoso.sharepoint.com`. |
| `clientId` | Client id of the app registration for the PnP sign-in. See [PREREQUISITES.md](PREREQUISITES.md). |
| `model` | `direct` or `nested`. See below. |
| `groupPrefix` | What every role group name starts with, including the trailing space: `"SP - UK "`. |
| `targets` | One entry per matrix column. See below. |

### Optional

| Setting | Default | Meaning |
|---|---|---|
| `matrixFile` | `matrix.csv` | Path of the matrix, relative to the config folder. |
| `defaultSite` | none | Site for targets that do not name one. Handy when all libraries are on one site. |
| `defaultPermission` | `Contribute` | Level for a plain mark in the matrix. |
| `rolePermissions` | none | Level per role, overriding the default: `{ "General Manager": "Edit" }`. |
| `adminGroups` | none | Groups that get access to every target, outside the matrix. Default level `Full Control`. On a site they belong in the Owners group. |
| `everyoneGroup` | none | Nested model: the group that contains every access group. |
| `exceptionGroups` | none | Groups outside the prefix that you still want in the snapshot and the report, such as a group for a folder exception. |
| `keepPrincipals` | none | Names or login names (wildcards allowed) that the audit and sync must leave alone: `["* Owners"]`. |
| `extraSites` | none | Sites that are not a target but must be covered when a leaver is removed and when group names are refreshed. |
| `snapshotPrefixes` | `groupPrefix` | Which groups the snapshot includes. Set it wider than `groupPrefix` when several configs share groups, for example `["SP -"]`. |
| `userFilter` | none | Which users count as "in scope" for the user list: `{ "domains": ["contoso.com"], "companies": ["Contoso"] }`. Without it, only people who are in a group are listed. |
| `groupDescription` | generic text | Description for groups the toolkit creates. `{name}` is replaced by the role or target name. |
| `tenantId` | derived | Tenant for sign-in. Derived from `tenantUrl` (`contoso.onmicrosoft.com`); set it when that is not right. |
| `pnpAuth` | `Interactive` | `Interactive` (browser) or `DeviceLogin` (code on another device). |
| `persistLogin` | `true` | Remember the SharePoint sign-in in the operating system's credential store. |
| `outputDir` | `output` | Where snapshots and logs go, relative to the config folder. |
| `report` | none | Settings for the Excel report. See below. |

### targets

The key is the column name in the matrix.

**Direct model**: a target is a library.

```json
"targets": {
  "Finance": { "library": "Finance" },
  "Board": { "library": "Board", "aliases": ["Board Documents"], "site": "/sites/Management" },
  "Everyone": { "library": "Everyone", "fixedGroups": ["SP - UK Everyone"] }
}
```

| Setting | Meaning |
|---|---|
| `library` | Title of the library. |
| `aliases` | Other names to try: the internal name or an old title. |
| `site` | Site of the library, when it differs from `defaultSite`. Full URL or the part after the tenant. |
| `permission` | Default level for this target, overriding `defaultPermission`. |
| `fixedGroups` | Groups that always get access to this target, whatever the matrix says. A name, or `{ "group": "...", "permission": "Read" }`. |

**Nested model**: a target is a site.

```json
"targets": {
  "Finance": { "site": "/sites/Finance" },
  "Sales": { "site": "/sites/Commercial", "accessGroup": "SP - NL Sales Department" },
  "Archive": {}
}
```

| Setting | Meaning |
|---|---|
| `site` | The department site. A target without a site still gets its access group and nesting in Entra; nothing is checked in SharePoint for it. |
| `accessGroup` | Name of the access group. Default: `groupPrefix` plus the column name. |

### The two models

**`direct`.** For each library, the role groups with a mark are granted the library directly, at their level. The library gets its own permissions; inheritance from the site is broken. Everything else on the library is reported and can be removed: other groups, people, sharing links, SharePoint groups.

**`nested`.** Each site has one access group, which sits in the site's Members group. The role groups with a mark in that column are nested inside the access group. Changing who reaches a site is then a change in Entra only (`04-Sync-GroupNesting.ps1`); SharePoint is not touched. The admin groups sit in the site's Owners group.

Permission levels per role do not apply to the nested model: everyone who reaches the site does so through the Members group and gets what that group has, normally Edit.

### report

| Setting | Default | Meaning |
|---|---|---|
| `fileName` | `SharePoint-Access-Report <name>` | Base name of the workbook. The date is appended. |
| `outputDir` | `reports` | Where the workbook goes, relative to the config folder. |
| `archiveDir` | none | When set, older versions of the report are moved here, relative to `outputDir`. |
| `matrixSheet` | `Access Matrix` | Name of the matrix tab. |
| `targetLabel` | `Library` or `Department` | What a matrix column is called in headings. |
| `consolidateByName` | `false` | Merge several accounts of the same person into one row, matched on display name. |
| `includeDisabled` | `false` | Also list disabled accounts that are still in a group. |
| `notesFile` | none | JSON file with texts for the report. |
| `itemSharesFileName` | `Individually-Shared-Items <name>` | Base name of the workbook from `itemshares.py build`. |

## report-notes.json

Optional texts for the report. Three lists:

```json
{
  "summary": [ ["PERMISSION MODEL", ""], ["General Manager", "Edit on every library"] ],
  "notes":   [ ["How to use this workbook", "Edit only the matrix tab ..."], ["2026-01-15", "First version."] ],
  "noAccess": [ { "name": "Break Glass", "upn": "breakglass@contoso.com", "jobTitle": "", "reason": "Emergency account." } ]
}
```

- `summary` rows are added to the Summary tab below the counts.
- `notes` rows become the Notes tab: the place for instructions to the owner and a changelog.
- `noAccess` becomes the No Access tab: people who deliberately have no access, with the reason. It stops them from being flagged as forgotten at every review.

Keep the changelog here. It is the one place the business owner sees.
