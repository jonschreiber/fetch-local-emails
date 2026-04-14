# Summarize Local Emails

This project extracts recent emails directly from Thunderbird's local mail
storage on macOS and writes them to either:

- a structured Markdown file, or
- JSON on stdout

It also includes an MCP server so Codex can call the extractor directly.

It is intentionally simple, and doesn't depend on cloud/app access to your organization's email server. It runs locally:

- no API keys
- no OAuth
- no cloud access
- local Thunderbird mailbox access only

## What It Does

The script reads a Thunderbird mbox file from your local profile, filters
messages by date, optional sender filter, and optional subject filter, and extracts:

- sender name
- sender email address
- received date
- decoded subject
- plain-text body

For multipart emails it prefers `text/plain`. If only `text/html` is present,
it converts the HTML into Markdown so structure like headings, lists, links,
line breaks, and tables is retained more cleanly. Long bodies are truncated.

## Requirements

- macOS
- `python3`
- Thunderbird installed
- Thunderbird must have synced the mailbox at least once

This project is set up so you can run it with `uv`. The runtime dependencies are
small: `mcp` for the server wrapper and `rapidfuzz` for fuzzy subject matching.

## How Profile Detection Works

By default the script looks for the first Thunderbird profile matching:

```text
~/Library/Thunderbird/Profiles/*/ImapMail/outlook.office365.com/
```

If auto-detect fails, you can point it at a profile manually with `--profile`.

## How To Run

From this project directory:

```bash
uv run python3 extract_emails.py
```

That will:

- auto-detect the Thunderbird profile
- read the `INBOX` mbox file
- extract emails from the last 7 days
- write Markdown to `~/emails_last_week.md`

## Common Examples

Extract the last 7 days to Markdown:

```bash
uv run python3 extract_emails.py
```

Extract the last 3 days as JSON:

```bash
uv run python3 extract_emails.py --days 3 --json
```

Extract JSON with both Markdown-preserving and plain-text body fields:

```bash
uv run python3 extract_emails.py --days 3 --json --body-mode both
```

Extract a different folder:

```bash
uv run python3 extract_emails.py --folder Sent --output ~/sent.md
```

Filter by sender with a partial email match:

```bash
uv run python3 extract_emails.py --sender @zoom.us
```

Filter by subject using case-insensitive contains matching plus fuzzy fallback:

```bash
uv run python3 extract_emails.py --subject "Meeting Assets"
```

Combine sender and subject filters:

```bash
uv run python3 extract_emails.py --sender @zoom.us --subject "Meeting Assets"
```

Use a specific Thunderbird profile:

```bash
uv run python3 extract_emails.py --profile /path/to/Thunderbird/Profile
```

## Running Tests

The project includes `pytest` tests against a Thunderbird-style mailbox fixture.

Run them with:

```bash
uv run pytest
```

## MCP Server

The project also includes:

- [thunderbird_email_mcp.py](/Users/jonschre/Code/tools/summarize-local-emails/thunderbird_email_mcp.py)

Run it locally with:

```bash
uv run python3 thunderbird_email_mcp.py
```

It exposes one MCP tool:

- `fetch_thunderbird_local_emails`

Tool inputs:

- `days`
- `folder`
- `max_body`
- `profile`
- `format`
  `markdown` or `json`
- `sender`
  Optional partial sender filter such as `@zoom.us`
- `subject`
  Optional subject filter using contains matching plus fuzzy fallback
- `body_mode`
  Optional body output mode: `rendered` or `both`
- `write_to_file`
  only used with `format = "markdown"`
- `output_file`
  only used when `write_to_file = true`

In `markdown` mode, the tool returns a JSON object containing metadata plus a
`markdown` field with the rendered email output.

In `json` mode, the tool returns a JSON object containing metadata plus an
`emails` array.

### Codex Config

Add this to `~/.codex/config.toml`:

```toml
[mcp_servers.thunderbird_local_emails]
command = "uv"
args = [
  "--directory",
  "/Users/jonschre/Code/tools/summarize-local-emails",
  "run",
  "python3",
  "thunderbird_email_mcp.py",
]
startup_timeout_sec = 120
tool_timeout_sec = 300
```

## Command-Line Options

- `--days`
  Number of days back to include. Default: `7`
- `--folder`
  Thunderbird mbox filename to read. Default: `INBOX`
- `--output`
  Markdown output file path. Default: `~/emails_last_week.md`
- `--max-body`
  Maximum body length per email. Default: `1000`
- `--profile`
  Thunderbird profile path override
- `--sender`
  Optional case-insensitive sender filter with partial matches
- `--subject`
  Optional subject filter with case-insensitive contains matching and fuzzy fallback
- `--body-mode`
  Body output mode: `rendered` keeps `body`, while `both` also adds `body_markdown` and `body_text`
- `--json`
  Print a JSON array to stdout instead of writing Markdown

## Output Format

Markdown output looks like this:

```md
---
**From:** Name <email@example.com>  
**Date:** 2025-01-15 09:30  
**Subject:** Weekly sync notes

body text here...
```

With `--json`, the script writes one JSON array of objects with:

- `from_name`
- `from_email`
- `date`
- `subject`
- `body`

With `--body-mode both`, each JSON object also includes:

- `body_markdown`
- `body_text`

## Troubleshooting

If the Thunderbird profile is not found:

- check `~/Library/Thunderbird/Profiles/`
- confirm the profile contains `ImapMail/YOUR-MAIL-SERVER-HERE/`
- rerun with `--profile /path/to/profile`

If the mailbox file is missing:

- open Thunderbird
- click the target folder so Thunderbird syncs it locally
- run the script again

If the mailbox is empty or no recent emails match:

- make sure Thunderbird has synced that folder
- try a larger date window with `--days`

If you get permission errors:

- confirm your user can read the Thunderbird profile files

## Files

- [extract_emails.py](/Users/jonschre/Code/tools/summarize-local-emails/extract_emails.py)
  Main script
- [thunderbird_email_mcp.py](/Users/jonschre/Code/tools/summarize-local-emails/thunderbird_email_mcp.py)
  MCP server wrapper for Codex
- [pyproject.toml](/Users/jonschre/Code/tools/summarize-local-emails/pyproject.toml)
  Minimal `uv` project config