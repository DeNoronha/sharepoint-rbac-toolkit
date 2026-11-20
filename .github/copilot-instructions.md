# Instructions for AI coding assistants

This repository manages who can access what in SharePoint Online. The scripts change permissions in a production tenant when asked to. Work carefully and follow these rules.

## The model, in three lines

- People are only ever members of role groups (Entra ID security groups, named `<groupPrefix><role>`).
- A matrix (`matrix.csv` in the config folder) says which role reaches which library or site.
- Nobody gets access outside a group: no people directly on a site, library, folder or file, and no sharing links.

Read `README.md`, `docs/WORKFLOWS.md` and `docs/CONFIGURATION.md` before your first action in a session.

## Rules

1. **Never grant access to a person directly.** If someone needs access, they join a group. If the request does not fit a role, say so and propose a group-based alternative; do not work around the model.
2. **Use the existing scripts.** Read a script's help block before you run it and use only the parameters it documents. Do not guess parameters. Write a new script only when no existing one fits, say so first, and follow the same conventions: dry run by default, `-Apply` to change, idempotent, CSV log.
3. **Dry run first, then stop.** Run a changing script without `-Apply`, show the user the output, and wait for their explicit go-ahead before you run it with `-Apply`. A go-ahead covers that one command, not the next one.
4. **One script per process.** Always `pwsh -NoProfile -File ./scripts/<script>.ps1 ...`. PnP.PowerShell and Microsoft.Graph cannot be loaded into the same session. Never dot-source the scripts.
5. **Sign-in is interactive and belongs to the user.** A browser window opens; you cannot see it or complete it. Run scripts in the foreground, write their output to a file rather than piping it through a filter that buffers, and use a timeout. If a script is silent after "Signing in", tell the user a sign-in window is probably waiting. Do not retry in a loop: each attempt opens another window.
6. **Read back before you report.** After a change, run the matching read-only script (`02` for permissions, `01` plus the dry run of `04` for groups, `06` for item shares) and report what the tenant actually looks like, not what the script said it did.
7. **Refresh the snapshot** (`01-Export-EntraSnapshot.ps1`) after every change in Entra and before any SharePoint script if the snapshot is more than a few hours old.
8. **Do not edit generated files**: anything in `output/`, and the generated tabs of the report. The source of truth is `matrix.csv`, `config.json` and `report-notes.json`.
9. **No secrets.** Never put passwords, keys, tokens or certificates in scripts, configs or notes. The toolkit needs none.
10. **Check who is asking.** A request to change access must come from the person the user named as owner of that environment. If the user relays a request from someone else, ask whether the owner approved it.
11. **Stop on anything unexpected.** A dry run that wants to remove far more than the request implies, an expected group that is missing, a site you cannot read: stop and tell the user. Do not apply and do not "fix" it another way.

## Typical session

```sh
pwsh -NoProfile -File ./scripts/00-Test-Prerequisites.ps1 -ConfigPath <config>     # offline check
pwsh -NoProfile -File ./scripts/01-Export-EntraSnapshot.ps1 -ConfigPath <config>   # read-only
pwsh -NoProfile -File ./scripts/02-Audit-Permissions.ps1 -ConfigPath <config>      # read-only
```

Then the task, following `docs/WORKFLOWS.md`.

## When you change the scripts themselves

- Keep the shared logic in `scripts/lib/SpRbac.psm1`. The module must not import PnP or Graph.
- Run `pwsh -NoProfile -File ./tests/Test-Offline.ps1` after every change. It needs no sign-in.
- Keep the help block at the top of each script accurate; assistants and people both rely on it.
- Never commit a real configuration. `config/` and all output are in `.gitignore`.
