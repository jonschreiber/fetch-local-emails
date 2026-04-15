#!/usr/bin/env python3
"""
Extract recent emails directly from Thunderbird's local mbox storage on macOS.

This script finds your Thunderbird profile, opens a local mbox file such as
`INBOX`, extracts recent messages, optionally filters them by sender and
subject, and writes them as either Markdown or JSON. It works whether
Thunderbird is open or closed.

Prerequisites:
- `python3`
- Thunderbird installed and signed in
- Thunderbird must have synced the target mailbox at least once

Usage examples:
    python3 extract_emails.py
    python3 extract_emails.py --list-accounts
    python3 extract_emails.py --account mail.example.invalid
    python3 extract_emails.py --days 3 --json
    python3 extract_emails.py --folder Sent --output ~/sent.md
    python3 extract_emails.py --sender @zoom.us --subject "Meeting assets"

If you prefer to run it with `uv` without creating a package environment:
    uv run python3 extract_emails.py

How to find your Thunderbird profile manually if auto-detect fails:
- Open `~/Library/Thunderbird/Profiles/`
- Look for a profile directory such as `xxxx.default-release`
- Inside that directory, confirm you have:
  `ImapMail/<account>/INBOX` or `Mail/<account>/INBOX`
- Pass the profile directory with `--profile /path/to/profile`

Note:
- Thunderbird must have synced the mailbox at least once, and in some cases you
  may need to click the folder in Thunderbird first so the local mbox file is
  created or refreshed.
"""

from __future__ import annotations

import argparse
import json
import mailbox
import re
import sys
from dataclasses import dataclass, replace
from datetime import datetime, timedelta
from email.header import decode_header
from email.message import Message
from email.utils import parseaddr, parsedate_to_datetime
from html import unescape
from pathlib import Path
from typing import Final

from bs4 import BeautifulSoup
from markdownify import MarkdownConverter
from rapidfuzz import fuzz

DAYS_BACK = 7
MAIL_FOLDER = "INBOX"
OUTPUT_FILE = "~/emails_last_week.md"
MAX_BODY_LENGTH = 1000
THUNDERBIRD_PROFILE = ""
THUNDERBIRD_ACCOUNT = ""
SENDER_FILTER = ""
SUBJECT_FILTER = ""
BODY_MODE = "rendered"

SUBJECT_FUZZY_THRESHOLD: Final[int] = 80
HTML_STRIP_TAGS: Final[list[str]] = ["script", "style"]
VALID_BODY_MODES: Final[set[str]] = {"rendered", "both"}
MAIL_STORAGE_DIRECTORIES: Final[tuple[str, ...]] = ("ImapMail", "Mail")
EXPECTED_PROFILE_HINT: Final[str] = "~/Library/Thunderbird/Profiles/*/{ImapMail,Mail}/<account>/"
DISCOVERY_COMMAND_HINT: Final[str] = "uv run python3 extract_emails.py --list-accounts"


@dataclass(frozen=True)
class Config:
    """Runtime configuration for email extraction."""

    days_back: int = DAYS_BACK
    mail_folder: str = MAIL_FOLDER
    output_file: str = OUTPUT_FILE
    max_body_length: int = MAX_BODY_LENGTH
    thunderbird_profile: str = THUNDERBIRD_PROFILE
    thunderbird_account: str = THUNDERBIRD_ACCOUNT
    sender_filter: str = SENDER_FILTER
    subject_filter: str = SUBJECT_FILTER
    body_mode: str = BODY_MODE
    as_json: bool = False
    list_accounts: bool = False


@dataclass
class ExtractionStats:
    """Holds summary information about the extraction process."""

    total_in_mbox: int = 0
    matched_in_range: int = 0
    matched_after_filters: int = 0
    skipped_malformed: int = 0


@dataclass(frozen=True)
class ExtractionResult:
    """The extracted emails plus metadata about the mailbox that was used."""

    profile_path: Path
    account_path: Path
    mbox_path: Path
    emails: list[dict[str, str]]
    stats: ExtractionStats


@dataclass(frozen=True)
class MailAccount:
    """A discoverable Thunderbird mail account directory."""

    profile_path: Path
    account_path: Path
    storage_root: str
    account_name: str


RUNTIME_PROFILE_OVERRIDE = THUNDERBIRD_PROFILE
RUNTIME_ACCOUNT_OVERRIDE = THUNDERBIRD_ACCOUNT
LAST_EXTRACTION_STATS = ExtractionStats()


def eprint(message: str) -> None:
    """Print a message to stderr."""

    print(message, file=sys.stderr)


def is_mail_account_path(path: Path) -> bool:
    """Return true when the path points at a Thunderbird account directory."""

    return path.is_dir() and path.parent.name in MAIL_STORAGE_DIRECTORIES


def discover_mail_accounts(profile: Path) -> list[MailAccount]:
    """Return the discoverable Thunderbird mail accounts for a profile."""

    profile_path = profile.expanduser()
    account_paths: list[Path] = []

    if is_mail_account_path(profile_path):
        base_profile = profile_path.parent.parent
        account_paths = [profile_path]
    elif profile_path.name in MAIL_STORAGE_DIRECTORIES and profile_path.is_dir():
        base_profile = profile_path.parent
        account_paths = sorted(path for path in profile_path.iterdir() if path.is_dir())
    else:
        base_profile = profile_path
        for storage_root in MAIL_STORAGE_DIRECTORIES:
            storage_path = profile_path / storage_root
            if not storage_path.exists():
                continue
            account_paths.extend(sorted(path for path in storage_path.iterdir() if path.is_dir()))

    return [
        MailAccount(
            profile_path=base_profile,
            account_path=account_path,
            storage_root=account_path.parent.name,
            account_name=account_path.name,
        )
        for account_path in account_paths
    ]


def render_account_discovery(profile: Path, accounts: list[MailAccount]) -> str:
    """Render a plain-text list of discoverable Thunderbird accounts."""

    lines = [
        f"Thunderbird profile: {profile}",
        "",
        "Available mail accounts:",
    ]
    for account in accounts:
        lines.append(f"- {account.account_name} ({account.storage_root})")
    lines.extend(
        [
            "",
            "Use one with:",
            "  uv run python3 extract_emails.py --account <account-name>",
        ]
    )
    return "\n".join(lines) + "\n"


def resolve_mail_account(profile: Path, requested_account: str = "") -> MailAccount:
    """Resolve the Thunderbird account directory to use for extraction."""

    accounts = discover_mail_accounts(profile)
    if not accounts:
        raise FileNotFoundError(
            f"Thunderbird mail accounts were not found under {profile}. Expected a path like "
            f"{EXPECTED_PROFILE_HINT} Run `{DISCOVERY_COMMAND_HINT}` after syncing a mailbox."
        )

    normalized_requested = requested_account.strip().casefold()
    if not normalized_requested:
        if len(accounts) == 1:
            return accounts[0]
        available_accounts = ", ".join(
            f"{account.account_name} ({account.storage_root})" for account in accounts
        )
        raise ValueError(
            f"Multiple Thunderbird mail accounts were found under {accounts[0].profile_path}. "
            f"Run `{DISCOVERY_COMMAND_HINT}` and rerun with `--account <name>`. "
            f"Available accounts: {available_accounts}"
        )

    matches = [
        account for account in accounts if account.account_name.strip().casefold() == normalized_requested
    ]
    if len(matches) == 1:
        return matches[0]

    available_accounts = ", ".join(
        f"{account.account_name} ({account.storage_root})" for account in accounts
    )
    raise ValueError(
        f"Thunderbird account '{requested_account}' was not found under {accounts[0].profile_path}. "
        f"Available accounts: {available_accounts}. Run `{DISCOVERY_COMMAND_HINT}` to inspect available accounts."
    )


def find_thunderbird_profile() -> Path:
    """Locate the Thunderbird profile directory on macOS."""

    if RUNTIME_PROFILE_OVERRIDE:
        override_path = Path(RUNTIME_PROFILE_OVERRIDE).expanduser()
        if not override_path.exists():
            raise FileNotFoundError(
                f"Thunderbird profile override was not found: {override_path}. "
                f"Expected a profile or account path like {EXPECTED_PROFILE_HINT}"
            )
        if is_mail_account_path(override_path):
            return override_path.parent.parent
        if override_path.name in MAIL_STORAGE_DIRECTORIES and override_path.is_dir():
            return override_path.parent
        if discover_mail_accounts(override_path):
            return override_path
        raise FileNotFoundError(
            f"Thunderbird profile override is not a profile or account path: {override_path}. "
            f"Expected a path like {EXPECTED_PROFILE_HINT}"
        )

    profiles_root = Path.home() / "Library" / "Thunderbird" / "Profiles"
    candidates = sorted(
        profile_path
        for profile_path in profiles_root.glob("*")
        if profile_path.is_dir() and discover_mail_accounts(profile_path)
    )
    if not candidates:
        raise FileNotFoundError(
            "Thunderbird profile not found. Expected a path like "
            f"{EXPECTED_PROFILE_HINT} Run `{DISCOVERY_COMMAND_HINT}` after syncing a mailbox."
        )
    return candidates[0]


def find_mbox_path(profile: Path, folder: str) -> Path:
    """Resolve the Thunderbird mbox path for the selected folder."""

    account = resolve_mail_account(profile.expanduser(), RUNTIME_ACCOUNT_OVERRIDE)
    mbox_path = account.account_path / folder
    if not mbox_path.exists():
        raise FileNotFoundError(
            f"The Thunderbird mail file {mbox_path} was not found. "
            "Open Thunderbird, click the target folder so it syncs locally, then try again."
        )
    return mbox_path


def decode_header_value(raw: str | None) -> str:
    """Decode a possibly MIME-encoded header into plain Unicode text."""

    if not raw:
        return ""

    decoded_parts: list[str] = []
    for part, charset in decode_header(raw):
        if isinstance(part, bytes):
            charsets = [charset, "utf-8", "latin-1"]
            for candidate in charsets:
                if not candidate:
                    continue
                try:
                    decoded_parts.append(part.decode(candidate))
                    break
                except (LookupError, UnicodeDecodeError):
                    continue
            else:
                decoded_parts.append(part.decode("utf-8", errors="replace"))
        else:
            decoded_parts.append(part)
    return re.sub(r"\s+", " ", "".join(decoded_parts)).strip()


def decode_bytes(payload: bytes, preferred_charset: str | None) -> str:
    """Decode raw message bytes using the preferred charset, utf-8, then latin-1."""

    for charset in (preferred_charset, "utf-8", "latin-1"):
        if not charset:
            continue
        try:
            return payload.decode(charset)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="replace")


class EmailMarkdownConverter(MarkdownConverter):
    """Markdownify converter with email-oriented handling for links and tables."""

    def convert_a(self, el, text, parent_tags):  # type: ignore[override]
        href = (el.get("href") or "").strip()
        if text.strip():
            return super().convert_a(el, text, parent_tags)

        label = (
            (el.get("aria-label") or "").strip()
            or (el.get("title") or "").strip()
            or href
        )
        if not label:
            return ""
        if not href or label == href:
            return f"<{label}>"
        return f"[{label}]({href})"


def html_to_markdown_body(html_string: str) -> str:
    """Convert HTML into readable Markdown while preserving common structure."""

    markdown = EmailMarkdownConverter(
        heading_style="ATX",
        bullets="-",
        newline_style="BACKSLASH",
        table_infer_header=True,
        strip=HTML_STRIP_TAGS,
        wrap=False,
    ).convert(html_string)
    markdown = unescape(markdown)
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")
    markdown = re.sub(r"[ \t]+\n", "\n", markdown)
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    return markdown.strip()


def html_to_plain_text(html_string: str) -> str:
    """Flatten HTML into readable plain text for search-friendly output."""

    soup = BeautifulSoup(html_string, "html.parser")

    for tag_name in HTML_STRIP_TAGS:
        for tag in soup.find_all(tag_name):
            tag.decompose()

    for br in soup.find_all("br"):
        br.replace_with("\n")

    for img in soup.find_all("img"):
        replacement = (img.get("alt") or "").strip() or (img.get("src") or "").strip()
        img.replace_with(replacement)

    for anchor in soup.find_all("a"):
        href = (anchor.get("href") or "").strip()
        label = anchor.get_text(" ", strip=True)
        title = (anchor.get("title") or "").strip()
        if label and href and label != href:
            replacement = f"{label} ({href})"
        elif title and href and title != href:
            replacement = f"{title} ({href})"
        else:
            replacement = label or title or href
        anchor.replace_with(replacement)

    for table in soup.find_all("table"):
        row_lines: list[str] = []
        for row in table.find_all("tr"):
            cells = [cell.get_text(" ", strip=True) for cell in row.find_all(["th", "td"])]
            if cells:
                row_lines.append(" | ".join(cells))
        if row_lines:
            row_text = "\n".join(row_lines)
            table.replace_with(f"\n\n{row_text}\n\n")
        else:
            table.replace_with("")

    for item in soup.find_all("li"):
        item.replace_with(f"- {item.get_text(' ', strip=True)}\n")

    for block in soup.find_all(
        [
            "p",
            "div",
            "section",
            "article",
            "header",
            "footer",
            "blockquote",
            "pre",
            "ul",
            "ol",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
        ]
    ):
        block.append("\n\n")

    text = unescape(soup.get_text())
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def strip_html(html_string: str) -> str:
    """Backward-compatible wrapper for HTML-to-Markdown conversion."""

    return html_to_markdown_body(html_string)


def normalize_body(text: str, max_length: int) -> str:
    """Normalize whitespace and enforce a maximum output length."""

    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = "\n".join(line.rstrip() if line.strip() else "" for line in normalized.splitlines())
    normalized = re.sub(r"\n{3,}", "\n\n", normalized).strip()
    if len(normalized) > max_length:
        return normalized[:max_length].rstrip()
    return normalized


def get_plain_text_body(message: Message, max_length: int) -> str:
    """Return the default rendered body variant for a message."""

    rendered_body, _ = get_body_variants(message, max_length)
    return rendered_body


def get_body_variants(message: Message, max_length: int) -> tuple[str, str]:
    """Extract Markdown-oriented and plain-text body variants from a message."""

    plain_text: str | None = None
    html_text: str | None = None

    if message.is_multipart():
        for part in message.walk():
            if part.is_multipart():
                continue
            disposition = part.get_content_disposition()
            if disposition == "attachment":
                continue

            content_type = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload is None:
                raw_payload = part.get_payload()
                if isinstance(raw_payload, str):
                    decoded = raw_payload
                else:
                    continue
            else:
                decoded = decode_bytes(payload, part.get_content_charset())

            if content_type == "text/plain" and plain_text is None:
                plain_text = decoded
            elif content_type == "text/html" and html_text is None:
                html_text = decoded
    else:
        payload = message.get_payload(decode=True)
        if payload is None:
            raw_payload = message.get_payload()
            decoded = raw_payload if isinstance(raw_payload, str) else ""
        else:
            decoded = decode_bytes(payload, message.get_content_charset())

        if message.get_content_type() == "text/html":
            html_text = decoded
        else:
            plain_text = decoded

    body_markdown = plain_text if plain_text is not None else html_to_markdown_body(html_text or "")
    body_text = plain_text if plain_text is not None else html_to_plain_text(html_text or "")
    return normalize_body(body_markdown, max_length), normalize_body(body_text, max_length)


def parse_message_datetime(message: Message) -> datetime | None:
    """Return the best available message timestamp as a local naive datetime."""

    candidates: list[str] = []
    for header_name in ("Delivery-date", "Date"):
        header_value = message.get(header_name)
        if header_value:
            candidates.append(decode_header_value(header_value))

    for received_value in message.get_all("Received", []):
        if ";" in received_value:
            candidates.append(received_value.rsplit(";", 1)[-1].strip())

    for candidate in candidates:
        try:
            parsed = parsedate_to_datetime(candidate)
        except (TypeError, ValueError, IndexError, OverflowError):
            continue
        if parsed is None:
            continue
        if parsed.tzinfo is not None:
            return parsed.astimezone().replace(tzinfo=None)
        return parsed
    return None


def extract_sender(message: Message) -> tuple[str, str]:
    """Extract and decode sender name and email address."""

    decoded_from = decode_header_value(message.get("From"))
    from_name, from_email = parseaddr(decoded_from)
    return decode_header_value(from_name), from_email.strip()


def format_message_date(message_datetime: datetime) -> str:
    """Format a message timestamp for output."""

    return message_datetime.strftime("%Y-%m-%d %H:%M")


def normalize_search_value(value: str) -> str:
    """Normalize search text for case-insensitive matching."""

    return re.sub(r"\s+", " ", value).strip().casefold()


def sender_matches(from_name: str, from_email: str, sender_filter: str) -> bool:
    """Return true when the sender matches a case-insensitive partial filter."""

    normalized_filter = normalize_search_value(sender_filter)
    if not normalized_filter:
        return True

    sender_values = [
        from_name,
        from_email,
        f"{from_name} <{from_email}>".strip(),
    ]
    return any(
        normalized_filter in normalize_search_value(candidate)
        for candidate in sender_values
        if candidate.strip()
    )


def subject_matches(subject: str, subject_filter: str) -> bool:
    """Return true when the subject matches by substring or fuzzy similarity."""

    normalized_filter = normalize_search_value(subject_filter)
    if not normalized_filter:
        return True

    normalized_subject = normalize_search_value(subject)
    if normalized_filter in normalized_subject:
        return True

    if len(normalized_filter) < 4:
        return False

    return fuzz.partial_ratio(normalized_filter, normalized_subject) >= SUBJECT_FUZZY_THRESHOLD


def extract_emails(
    mbox_path: Path,
    days_back: int,
    max_body: int,
    sender_filter: str = "",
    subject_filter: str = "",
    body_mode: str = BODY_MODE,
) -> list[dict[str, str]]:
    """Stream messages from an mbox file and return recent emails."""

    global LAST_EXTRACTION_STATS
    LAST_EXTRACTION_STATS = ExtractionStats()

    cutoff = datetime.now() - timedelta(days=days_back)
    extracted: list[dict[str, str]] = []

    try:
        mbox = mailbox.mbox(str(mbox_path), create=False)
    except PermissionError as exc:
        raise PermissionError(
            f"Permission denied when opening {mbox_path}. Check file permissions."
        ) from exc

    try:
        for index, message in enumerate(mbox, start=1):
            LAST_EXTRACTION_STATS.total_in_mbox += 1
            try:
                message_datetime = parse_message_datetime(message)
                if message_datetime is None:
                    raise ValueError("missing or invalid Date header")
                if message_datetime < cutoff:
                    continue

                from_name, from_email = extract_sender(message)
                subject = decode_header_value(message.get("Subject")) or "(No Subject)"
                LAST_EXTRACTION_STATS.matched_in_range += 1

                if not sender_matches(from_name, from_email, sender_filter):
                    continue
                if not subject_matches(subject, subject_filter):
                    continue

                body_markdown, body_text = get_body_variants(message, max_body)

                email_record = {
                    "from_name": from_name,
                    "from_email": from_email,
                    "date": format_message_date(message_datetime),
                    "subject": subject,
                    "body": body_markdown,
                }
                if body_mode == "both":
                    email_record["body_markdown"] = body_markdown
                    email_record["body_text"] = body_text

                extracted.append(email_record)
            except Exception as exc:  # noqa: BLE001
                LAST_EXTRACTION_STATS.skipped_malformed += 1
                eprint(f"Warning: skipped malformed email #{index}: {exc}")
                continue
    finally:
        mbox.close()

    LAST_EXTRACTION_STATS.matched_after_filters = len(extracted)
    return extracted


def write_markdown(emails: list[dict[str, str]], output_path: Path) -> None:
    """Write extracted emails to a Markdown file."""

    output_path = output_path.expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        handle.write(render_markdown(emails))


def write_json(emails: list[dict[str, str]]) -> None:
    """Write extracted emails to stdout as JSON."""

    json.dump(emails, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")


def parse_args() -> Config:
    """Parse CLI arguments into a runtime config."""

    parser = argparse.ArgumentParser(
        description="Extract recent Thunderbird emails from a local mbox file."
    )
    parser.add_argument("--days", type=int, default=DAYS_BACK, help="Days back to include. Default: 7.")
    parser.add_argument("--folder", default=MAIL_FOLDER, help='Mailbox filename. Default: "INBOX".')
    parser.add_argument("--output", default=OUTPUT_FILE, help="Markdown output path. Default: ~/emails_last_week.md")
    parser.add_argument("--max-body", type=int, default=MAX_BODY_LENGTH, help="Maximum body length. Default: 1000.")
    parser.add_argument(
        "--account",
        default=THUNDERBIRD_ACCOUNT,
        help="Thunderbird account directory name inside ImapMail/ or Mail/. Use --list-accounts to discover values.",
    )
    parser.add_argument(
        "--sender",
        default=SENDER_FILTER,
        help='Optional case-insensitive sender filter. Partial matches are supported, for example "@zoom.us".',
    )
    parser.add_argument(
        "--subject",
        default=SUBJECT_FILTER,
        help='Optional subject filter. Case-insensitive contains matches are supported, with fuzzy fallback for close matches like "Meeting Assets".',
    )
    parser.add_argument(
        "--body-mode",
        choices=sorted(VALID_BODY_MODES),
        default=BODY_MODE,
        help='Body output mode. "rendered" keeps the current body field, while "both" also adds body_markdown and body_text. Default: rendered.',
    )
    parser.add_argument(
        "--profile",
        default=THUNDERBIRD_PROFILE,
        help="Thunderbird profile path override. Default: auto-detect.",
    )
    parser.add_argument(
        "--list-accounts",
        action="store_true",
        help="List discoverable Thunderbird mail accounts for the selected profile and exit.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Write JSON to stdout instead of Markdown.",
    )
    args = parser.parse_args()

    return Config(
        days_back=args.days,
        mail_folder=args.folder,
        output_file=args.output,
        max_body_length=args.max_body,
        thunderbird_profile=args.profile,
        thunderbird_account=args.account,
        sender_filter=args.sender,
        subject_filter=args.subject,
        body_mode=args.body_mode,
        as_json=args.json,
        list_accounts=args.list_accounts,
    )


def validate_config(config: Config) -> Config:
    """Validate runtime config values."""

    if config.days_back < 1:
        raise ValueError("--days must be at least 1.")
    if config.max_body_length < 1:
        raise ValueError("--max-body must be at least 1.")
    mail_folder = config.mail_folder.strip()
    if not mail_folder:
        raise ValueError("--folder must not be empty.")
    output_file = config.output_file.strip()
    if not output_file:
        raise ValueError("--output must not be empty.")
    body_mode = config.body_mode.strip().lower()
    if body_mode not in VALID_BODY_MODES:
        raise ValueError(f"--body-mode must be one of: {', '.join(sorted(VALID_BODY_MODES))}.")
    return replace(
        config,
        mail_folder=mail_folder,
        output_file=output_file,
        thunderbird_profile=config.thunderbird_profile.strip(),
        thunderbird_account=config.thunderbird_account.strip(),
        sender_filter=config.sender_filter.strip(),
        subject_filter=config.subject_filter.strip(),
        body_mode=body_mode,
        list_accounts=config.list_accounts,
    )


def render_markdown(emails: list[dict[str, str]]) -> str:
    """Render extracted emails as a Markdown document."""

    if not emails:
        return "No emails matched the requested date range and filters.\n"

    chunks: list[str] = []
    for email_record in emails:
        from_line = email_record["from_name"]
        if email_record["from_email"]:
            if from_line:
                from_line = f'{from_line} <{email_record["from_email"]}>'
            else:
                from_line = email_record["from_email"]

        chunks.append("---")
        chunks.append(f"**From:** {from_line or 'Unknown Sender'}  ")
        chunks.append(f'**Date:** {email_record["date"]}  ')
        chunks.append(f'**Subject:** {email_record["subject"]}')
        chunks.append("")
        chunks.append(email_record["body"])
        chunks.append("")

    return "\n".join(chunks).rstrip() + "\n"


def run_extraction(config: Config) -> ExtractionResult:
    """Run the full extraction workflow and return emails plus mailbox metadata."""

    global RUNTIME_PROFILE_OVERRIDE, RUNTIME_ACCOUNT_OVERRIDE

    validated = validate_config(config)
    RUNTIME_PROFILE_OVERRIDE = validated.thunderbird_profile
    RUNTIME_ACCOUNT_OVERRIDE = validated.thunderbird_account

    profile_path = find_thunderbird_profile()
    account = resolve_mail_account(profile_path, validated.thunderbird_account)
    mbox_path = find_mbox_path(profile_path, validated.mail_folder)

    try:
        if mbox_path.stat().st_size == 0:
            raise RuntimeError(
                f"The mail file {mbox_path} is empty. Sync Thunderbird and try again."
            )
    except PermissionError as exc:
        raise PermissionError(
            f"Permission denied when reading {mbox_path}. Check file permissions."
        ) from exc

    emails = extract_emails(
        mbox_path,
        validated.days_back,
        validated.max_body_length,
        sender_filter=validated.sender_filter,
        subject_filter=validated.subject_filter,
        body_mode=validated.body_mode,
    )
    if LAST_EXTRACTION_STATS.total_in_mbox == 0:
        raise RuntimeError(
            f"The mail file {mbox_path} does not contain any synced messages yet. "
            "Open Thunderbird, click the folder, and let it sync."
        )

    return ExtractionResult(
        profile_path=profile_path,
        account_path=account.account_path,
        mbox_path=mbox_path,
        emails=emails,
        stats=ExtractionStats(
            total_in_mbox=LAST_EXTRACTION_STATS.total_in_mbox,
            matched_in_range=LAST_EXTRACTION_STATS.matched_in_range,
            matched_after_filters=LAST_EXTRACTION_STATS.matched_after_filters,
            skipped_malformed=LAST_EXTRACTION_STATS.skipped_malformed,
        ),
    )


def main() -> None:
    """Run the Thunderbird mbox extraction workflow."""

    try:
        config = validate_config(parse_args())
        if config.list_accounts:
            global RUNTIME_PROFILE_OVERRIDE, RUNTIME_ACCOUNT_OVERRIDE
            RUNTIME_PROFILE_OVERRIDE = config.thunderbird_profile
            RUNTIME_ACCOUNT_OVERRIDE = ""
            profile_path = find_thunderbird_profile()
            print(render_account_discovery(profile_path, discover_mail_accounts(profile_path)), end="")
            return

        result = run_extraction(config)
        output_path = Path(config.output_file).expanduser()
        has_content_filters = bool(config.sender_filter.strip() or config.subject_filter.strip())

        eprint(f"Thunderbird account path: {result.account_path}")
        eprint(f"Emails found in mbox: {result.stats.total_in_mbox}")
        eprint(f"Emails matching date filter: {result.stats.matched_in_range}")
        if has_content_filters:
            eprint(f"Emails matching sender/subject filters: {result.stats.matched_after_filters}")
        if config.as_json:
            eprint(f"Output file path: {output_path} (not written because --json was used)")
        else:
            eprint(f"Output file path: {output_path}")
        if result.stats.skipped_malformed:
            eprint(f"Malformed emails skipped: {result.stats.skipped_malformed}")

        if not result.emails:
            if has_content_filters:
                eprint(
                    f"No emails found in the last {config.days_back} days matching the current filters. "
                    "Try increasing --days or relaxing --sender/--subject."
                )
            else:
                eprint(f"No emails found in the last {config.days_back} days. Try increasing --days.")
            if config.as_json:
                write_json(result.emails)
            else:
                write_markdown(result.emails, output_path)
            return

        if config.as_json:
            write_json(result.emails)
        else:
            write_markdown(result.emails, output_path)
    except FileNotFoundError as exc:
        eprint(f"Error: {exc}")
        raise SystemExit(1) from exc
    except PermissionError as exc:
        eprint(f"Error: {exc}")
        raise SystemExit(1) from exc
    except RuntimeError as exc:
        eprint(f"Error: {exc}")
        raise SystemExit(1) from exc
    except ValueError as exc:
        eprint(f"Error: {exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
