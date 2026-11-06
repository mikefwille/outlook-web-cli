---
name: outlook-web
description: Read Outlook mail and calendars, manage drafts and contacts, and perform requested mailbox actions through the local Outlook web-session CLI.
---

# Outlook on the web

Use `outlook-web-pp-cli` for the user's Outlook mailbox. Keep the CLI and
`outlook-web-pp-mcp` together on PATH. This v1 is a standalone source project;
it is not available from the Printing Press catalog.

## Session and account

The current account-bound session pointer is
`~/.local/share/outlook-web-dev/private/current-outlook.json`. Keep the private
directory and credential contents private. Normal reads use the saved session
and do not launch a browser.

```bash
outlook-web-pp-cli auth status --json --no-learn
outlook-web-pp-cli doctor --live --json --no-learn
outlook-web-pp-cli accounts list --json --no-learn
```

Status is local; live doctor makes one read-only Outlook request. `accounts use`
previews selection of an already registered account and needs `--execute` to
change it. If the session needs recovery, `auth recover` previews the action;
`auth recover --execute` uses the separately installed `m365-session` helper and
may start the saved browser. Report an interactive login or missing helper when required. A
separate directory credential can expire independently.

## Read and search

```bash
outlook-web-pp-cli which "search mail" --json --no-learn
outlook-web-pp-cli mail list --limit 25 --json --no-learn
outlook-web-pp-cli mail search "invoice" --limit 25 --json --no-learn
outlook-web-pp-cli mail get "$ITEM_ID" --json --no-learn
```

The mail commands are live reads. Choose an item ID from a suitable item read;
server search returns conversations and can include Deleted Items. Check
pagination and partial-result indicators. `search` and `sql` instead use the
local metadata snapshot; report its freshness and incomplete coverage.
`sync folder` previews reconciliation, and `--execute` writes a private local
snapshot without changing Outlook.

## Drafts and requested changes

```bash
outlook-web-pp-cli drafts create --subject "Example" --text "Example body" --json --no-learn
outlook-web-pp-cli drafts get "$DRAFT_ID" --json --no-learn
outlook-web-pp-cli drafts send "$DRAFT_ID" --json --no-learn
```

Create and send preview by default. Creating with `--execute` saves a server
draft; it does not send. The send preview shows the saved body, all recipients,
attachments and change key. When that delivery is authorized by the user's
task or session instructions, execute the previewed send with `--execute` and
its fresh `--change-key`. Sending can notify recipients; server acceptance does
not prove delivery. Reconcile an unknown write outcome before retrying.

Other supported writes preview by default and require `--execute`; inspect
scope, version checks and notification effects first. The CLI also supports
folders, personal contacts and lists, selected mail actions, bounded calendar
appointments, and self-only availability. Some version checks are non-atomic.

## V1 limits

Meeting invitations, meeting accept/decline/tentative/respond/cancel/forward,
Junk/Not junk classification, and directory photo are deferred. Calendar
occurrence subject edits are limited to observed timed Mountain-time shapes;
broader recurrence, other zones and attendee notifications are unverified.
Do not infer that a missing command means Outlook itself lacks the feature.

## MCP

Use the local `outlook-web-pp-mcp` server with its sibling CLI.
It exposes a reviewed subset of the command tree. Set
`OUTLOOK_WEB_MCP_READ_ONLY=1` in the server environment to remove write tools;
execution flags must be JSON booleans. Local file paths and session credentials
are not exposed as MCP tool arguments. Starting the binary does not register it
with an MCP host.
