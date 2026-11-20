# Working with an AI assistant

The toolkit is built to be driven by a coding assistant: every script explains itself in its help block, never changes anything without `-Apply`, and says clearly in its output whether it was a dry run.

## Which assistant

You need one that can **read the files in this folder and run commands in a terminal**:

- **Claude Code**: open the repository folder and start a session. It reads `CLAUDE.md` by itself.
- **GitHub Copilot** in agent mode (VS Code): open the folder. It reads `.github/copilot-instructions.md`. Other agents read `AGENTS.md`.

A chat window without access to the folder (Copilot in Microsoft 365 or Teams, a chat in the browser) cannot read the scripts. It will invent script names and parameters. Use such a chat to have the documentation explained, never to write commands you then paste into a terminal.

## What stays with you

- **Signing in.** Every sign-in opens a browser window. The assistant cannot see or complete it. If a run goes quiet right after "Signing in", look for the window.
- **Saying yes.** The assistant shows you a dry run and waits. You decide whether it applies.
- **Knowing who asked.** The assistant cannot check whether a request really comes from the business owner.

## Start prompt

Paste this at the start of a session. With Claude Code or Copilot the rules are already in the instruction files; the prompt adds your own context.

```text
You are helping me manage SharePoint access with the scripts in this folder.
Read CLAUDE.md, README.md and docs/WORKFLOWS.md first.

My config folder is ./config/<name>. The business owner who may ask for changes is <name>.

Rules for this session:
- Access only goes through groups. Never put a person directly on a site, library, folder or file.
- Use the existing scripts and only the parameters in their help. Tell me before you write anything new.
- Run each script in its own process: pwsh -NoProfile -File ...
- Always do a dry run first, show me the output, and wait for my OK before you use -Apply.
- Sign-in opens a browser window that only I can complete. Tell me when you are waiting for one. Do not retry in a loop.
- After a change, run the read-only check and tell me what the tenant actually looks like.
```

## Prompts for common requests

Replace the parts in angle brackets.

**Check that everything is ready**

```text
Run 00-Test-Prerequisites.ps1 for my config and tests/Test-Offline.ps1. Explain every warning and tell me what I have to fix myself.
```

**Where do we stand?**

```text
Take a fresh snapshot (01) and run the audit (02) for my config. Summarize per library or site what differs from the matrix. Group the differences by kind: missing groups, groups that should not be there, people with direct access, items to review. Change nothing.
```

**Add a person**

```text
<owner> asks to add <name> (<sign-in address>) to the role <role>. Check that the role is a row in the matrix. Do the dry run with script 05 and show me the output.
```

**Someone changes role**

```text
<name> (<sign-in address>) moves from <old role> to <new role>. Dry run both changes with script 05 and tell me which libraries or sites they gain and lose according to the matrix.
```

**Someone has left**

```text
<name> (<sign-in address>) has left. Do the dry run of script 08 for step Groups and for step Sites and give me the list of what would be removed. Apply nothing yet.
```

**The owner sent back the matrix**

```text
Attached is the report <owner> sent back: <path>. Run import_matrix.py without --write and list the differences in plain language: who gains access to what, who loses it. Do not write anything.
```

```text
The differences are right. Write the matrix, then do the dry runs that the workflow for my model describes, and show me what would change in Entra and in SharePoint.
```

**Rebuild the report**

```text
Take a fresh snapshot and build the report. Compare it with the previous version tab by tab and list every difference, including ones that have nothing to do with today's change. Add a changelog line to report-notes.json for <what changed> and rebuild.
```

**Individually shared files**

```text
Run script 06 for <targets> and build the workbook with itemshares.py. Tell me how many shares give access the person does not have through a group, and for which people.
```

```text
<owner> has filled in the decisions: <path>. Convert them with 'itemshares.py decisions' and do the dry run of script 07. Show me what would be removed.
```

**A request for one folder**

```text
<requester> asks access to the folder <path> on <site> for <names>. First find out who can reach that folder now and through which group, and whether the folder has its own permissions. Tell me which of the named people already have access, and propose how the others could get it through a group. Run nothing that changes anything.
```

**When something fails**

```text
The script gave this error: <error>. Check docs/TROUBLESHOOTING.md for a known cause before you propose a fix. Do not change a script to make the error go away without telling me what the cause is.
```

## Reviewing what the assistant did

Every run leaves a CSV in `<config>/output/`: what was planned, what was applied, and any error. If you want to know what happened in a session, read those files; they are the record, the chat is not.
