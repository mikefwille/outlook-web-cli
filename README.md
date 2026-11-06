# Outlook Web CLI

Read and manage Outlook mail, drafts, personal calendars and contacts through a saved Outlook on the web session. The CLI and MCP server share the same account-bound client.

Created by [@mikefwille](https://github.com/mikefwille).

## Install

After the repository is published, install the two binaries with Go 1.26.10 or newer:

```sh
go install github.com/mikefwille/outlook-web-cli/cmd/outlook-web-pp-cli@latest
go install github.com/mikefwille/outlook-web-cli/cmd/outlook-web-pp-mcp@latest
```

Alternatively, run `make build-all` in a checkout and add `bin/` to your PATH. Keep the CLI and MCP binaries together. Agent instructions are in [SKILL.md](SKILL.md).

## Session setup

The CLI needs a compatible, account-bound Outlook web session and private operation templates. This repository contains no sessions, cookies, credentials, or captured requests. It does not yet provide a complete one-command setup for a new account.

The separate shared `m365-session` helper can recover an already configured session and must be on PATH or beside the CLI. The default session pointer is `~/.local/share/outlook-web-dev/private/current-outlook.json`; keep the whole private directory private. The helper and its browser profile are not included in this repository. Ordinary supported reads use the saved session without starting a browser.

```sh
outlook-web-pp-cli auth status --json --no-learn
outlook-web-pp-cli doctor --live --json --no-learn
```

`auth status` checks the local session; live doctor makes a read-only Outlook request. `auth recover` previews recovery, while `auth recover --execute` invokes the shared helper and may start its saved browser profile. Directory relationship reads use a separate credential that expires independently.

## Quick start

```sh
outlook-web-pp-cli which "search mail" --json --no-learn
outlook-web-pp-cli mail list --limit 25 --json --no-learn
outlook-web-pp-cli mail search "invoice" --limit 25 --json --no-learn
outlook-web-pp-cli mail get "$ITEM_ID" --json --no-learn
```

Choose an item ID from an appropriate item read. Mail search returns conversations and may include Deleted Items; check pagination and partial-result indicators. Local `search` and `sql` use a reconciled metadata snapshot, so check its freshness and coverage.

Drafts and supported writes preview by default. Saving a draft requires `--execute` and does not send it. Sending is a separate action: review the saved body, all recipients, attachments and previewed change key, then execute only the delivery the user authorized. Reconcile an uncertain write before retrying.

## MCP

Run `outlook-web-pp-mcp` as a stdio server with its sibling CLI available. Configure an MCP host with the following entry; remove the environment setting if the host should expose reviewed write tools:

```json
{
  "mcpServers": {
    "outlook-web": {
      "command": "outlook-web-pp-mcp",
      "env": {"OUTLOOK_WEB_MCP_READ_ONLY": "1"}
    }
  }
}
```

Starting the server does not register it with a host. Write execution flags must be JSON booleans. Requests with duplicate or incorrectly cased JSON-RPC envelope/tool members are rejected before tool dispatch; the upstream notification path is outside that guard.

## V1 scope

V1 covers ordinary mail reads, drafts and selected message actions, folders, personal contacts and lists, bounded personal calendar appointments, self-only availability, and a private metadata cache. Meeting response/cancellation/forwarding, Junk/Not junk classification, and directory photo are deferred. Occurrence subject edits are limited to observed timed Mountain-time shapes. The missing CLI paths do not imply that Outlook lacks those features.

This build has been exercised on macOS. Other platforms have not been verified end to end; cache ownership checks fail closed on unsupported platforms.

## Development

```sh
make build-all
make test
go vet ./...
```

Tests use synthetic data and may bind local loopback ports. Keep credentials and captures outside the checkout.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
