from __future__ import annotations

import asyncio
import json
import mailbox
import shutil
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import extract_emails as extractor
import thunderbird_email_mcp as mcp_app


class FixedDateTime(datetime):
    @classmethod
    def now(cls, tz=None):  # type: ignore[override]
        return cls(2026, 3, 25, 12, 0, 0, tzinfo=tz)


@pytest.fixture
def fixture_source_profile() -> Path:
    return Path(__file__).parent / "fixtures" / "thunderbird_profile" / "ImapMail"


@pytest.fixture
def fixture_profile(tmp_path: Path, fixture_source_profile: Path) -> Path:
    profile_path = tmp_path / "demo.default-release"
    shutil.copytree(fixture_source_profile, profile_path / "ImapMail")
    return profile_path


@pytest.fixture
def frozen_now(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(extractor, "datetime", FixedDateTime)


def add_mailbox_message(
    mailbox_path: Path,
    *,
    subject: str,
    date: str = "Tue, 24 Mar 2026 10:00:00 -0400",
    sender: str = "Folder Test <folder@example.com>",
    body: str = "Folder body",
) -> None:
    mailbox_path.parent.mkdir(parents=True, exist_ok=True)
    mbox = mailbox.mbox(str(mailbox_path))
    message = EmailMessage()
    message["From"] = sender
    message["Date"] = date
    message["Subject"] = subject
    message.set_content(body)
    mbox.add(message)
    mbox.flush()
    mbox.close()


def parse_text_payload(result: list[object]) -> dict[str, object]:
    assert len(result) == 1
    text = getattr(result[0], "text")
    return json.loads(text)


def test_build_tools_exposes_account_discovery_and_fetch() -> None:
    tool_names = [tool.name for tool in mcp_app.build_tools()]

    assert "list_thunderbird_mail_accounts" in tool_names
    assert "fetch_thunderbird_local_emails" in tool_names


def test_build_config_rejects_non_array_folder_globs() -> None:
    with pytest.raises(ValueError, match="folder_globs must be an array of strings"):
        mcp_app.build_config({"folder_globs": "1*"})


def test_list_accounts_tool_returns_structured_accounts(fixture_profile: Path) -> None:
    result = asyncio.run(
        mcp_app.call_tool_impl(
            "list_thunderbird_mail_accounts",
            {"profile": str(fixture_profile), "format": "json"},
        )
    )

    payload = parse_text_payload(result)
    assert payload["format"] == "json"
    assert payload["profile_path"] == str(fixture_profile)
    assert payload["accounts"] == [
        {
            "account_name": "mail.example.invalid",
            "storage_root": "ImapMail",
            "account_path": str(fixture_profile / "ImapMail" / "mail.example.invalid"),
        }
    ]


def test_fetch_tool_supports_recursive_folder_globs(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    account_path = fixture_profile / "ImapMail" / "mail.example.invalid"
    top_level = account_path / "1Inbox"
    nested = account_path / "1Inbox.sbd" / "ProjectA"
    exact_match = account_path / "A1"
    ignored = account_path / "Archive"
    add_mailbox_message(top_level, subject="Top-level active folder")
    add_mailbox_message(nested, subject="Nested active folder")
    add_mailbox_message(exact_match, subject="Exact folder")
    add_mailbox_message(ignored, subject="Ignored folder")

    result = asyncio.run(
        mcp_app.call_tool_impl(
            "fetch_thunderbird_local_emails",
            {
                "profile": str(fixture_profile),
                "format": "json",
                "folder_globs": ["1*", "A1"],
                "recursive_folders": True,
            },
        )
    )

    payload = parse_text_payload(result)
    subjects = [email["subject"] for email in payload["emails"]]
    assert subjects == ["Top-level active folder", "Nested active folder", "Exact folder"]
    assert payload["mbox_paths"] == [str(top_level), str(nested), str(exact_match)]
