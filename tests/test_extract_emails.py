from __future__ import annotations

import json
import shutil
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import extract_emails as app


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
    monkeypatch.setattr(app, "datetime", FixedDateTime)


def add_second_account(profile_path: Path) -> Path:
    second_account_path = profile_path / "ImapMail" / "mail.backup.invalid"
    shutil.copytree(profile_path / "ImapMail" / "mail.example.invalid", second_account_path)
    return second_account_path


def test_extract_emails_filters_date_and_decodes_content(
    fixture_profile: Path,
    frozen_now: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(app, "RUNTIME_PROFILE_OVERRIDE", "")
    monkeypatch.setattr(app, "RUNTIME_ACCOUNT_OVERRIDE", "")
    mbox_path = app.find_mbox_path(fixture_profile, "INBOX")

    emails = app.extract_emails(mbox_path, days_back=7, max_body=1000)

    assert len(emails) == 2
    assert [email["subject"] for email in emails] == [
        "Weekly sync notes",
        "Möbius meeting assets",
    ]
    assert emails[0]["from_name"] == "Alice Example"
    assert emails[1]["from_email"] == "no-reply@zoom.us"
    assert emails[1]["body"] == "Hello\\\nworld\n\nRésumé"
    assert app.LAST_EXTRACTION_STATS.total_in_mbox == 4
    assert app.LAST_EXTRACTION_STATS.matched_in_range == 2
    assert app.LAST_EXTRACTION_STATS.matched_after_filters == 2
    assert app.LAST_EXTRACTION_STATS.skipped_malformed == 1


def test_run_extraction_returns_profile_mbox_and_stats(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file="~/emails_last_week.md",
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
        )
    )

    assert result.profile_path == fixture_profile
    assert result.account_path == fixture_profile / "ImapMail" / "mail.example.invalid"
    assert result.mbox_path == fixture_profile / "ImapMail" / "mail.example.invalid" / "INBOX"
    assert result.stats.total_in_mbox == 4
    assert result.stats.matched_in_range == 2
    assert result.stats.matched_after_filters == 2
    assert result.stats.skipped_malformed == 1


def test_extract_emails_filters_by_partial_sender(
    fixture_profile: Path,
    frozen_now: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(app, "RUNTIME_PROFILE_OVERRIDE", "")
    monkeypatch.setattr(app, "RUNTIME_ACCOUNT_OVERRIDE", "")
    mbox_path = app.find_mbox_path(fixture_profile, "INBOX")

    emails = app.extract_emails(mbox_path, days_back=7, max_body=1000, sender_filter="@zoom.us")

    assert len(emails) == 1
    assert emails[0]["from_email"] == "no-reply@zoom.us"
    assert emails[0]["subject"] == "Möbius meeting assets"
    assert app.LAST_EXTRACTION_STATS.matched_in_range == 2
    assert app.LAST_EXTRACTION_STATS.matched_after_filters == 1


def test_extract_emails_filters_by_fuzzy_subject(
    fixture_profile: Path,
    frozen_now: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(app, "RUNTIME_PROFILE_OVERRIDE", "")
    monkeypatch.setattr(app, "RUNTIME_ACCOUNT_OVERRIDE", "")
    mbox_path = app.find_mbox_path(fixture_profile, "INBOX")

    emails = app.extract_emails(mbox_path, days_back=7, max_body=1000, subject_filter="meetng assets")

    assert len(emails) == 1
    assert emails[0]["subject"] == "Möbius meeting assets"
    assert app.LAST_EXTRACTION_STATS.matched_in_range == 2
    assert app.LAST_EXTRACTION_STATS.matched_after_filters == 1


def test_run_extraction_combines_sender_and_subject_filters(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file="~/emails_last_week.md",
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
            sender_filter="@zoom.us",
            subject_filter="meetng assets",
        )
    )

    assert len(result.emails) == 1
    assert result.emails[0]["subject"] == "Möbius meeting assets"
    assert result.stats.matched_in_range == 2
    assert result.stats.matched_after_filters == 1


def test_run_extraction_with_body_mode_both_includes_markdown_and_text(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file="~/emails_last_week.md",
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
            body_mode="both",
        )
    )

    assert len(result.emails) == 2
    assert result.emails[0]["body_markdown"] == "Plain text body for the weekly sync.\nContinuation line."
    assert result.emails[0]["body_text"] == "Plain text body for the weekly sync.\nContinuation line."
    assert result.emails[1]["body"] == "Hello\\\nworld\n\nRésumé"
    assert result.emails[1]["body_markdown"] == "Hello\\\nworld\n\nRésumé"
    assert result.emails[1]["body_text"] == "Hello\nworld\n\nRésumé"


def test_get_plain_text_body_converts_html_to_markdown_with_structure() -> None:
    message = EmailMessage()
    message.set_type("multipart/alternative")
    message.add_alternative(
        "<h1>Status</h1><p>Overview</p><ul><li>First</li><li>Second</li></ul>",
        subtype="html",
    )

    body = app.get_plain_text_body(message, 1000)

    assert body == "# Status\n\nOverview\n\n- First\n- Second"


def test_get_plain_text_body_preserves_links_breaks_and_tables_from_html() -> None:
    message = EmailMessage()
    message.set_type("multipart/alternative")
    message.add_alternative(
        """
        <h2>Runbook</h2>
        <p>Open the <a href="https://example.com/runbook">runbook</a><br>before deploy.</p>
        <table>
          <tr><td>Name</td><td>Status</td></tr>
          <tr><td>Node A</td><td>Healthy</td></tr>
        </table>
        <p><a href="https://example.com/zoom" title="View in Zoom"></a></p>
        """,
        subtype="html",
    )

    body = app.get_plain_text_body(message, 1000)

    assert body == (
        "## Runbook\n\n"
        "Open the [runbook](https://example.com/runbook)\\\n"
        "before deploy.\n\n"
        "| Name | Status |\n"
        "| --- | --- |\n"
        "| Node A | Healthy |\n\n"
        "[View in Zoom](https://example.com/zoom)"
    )


def test_get_body_variants_returns_markdown_and_plain_text_for_html() -> None:
    message = EmailMessage()
    message.set_type("multipart/alternative")
    message.add_alternative(
        """
        <h2>Runbook</h2>
        <p>Open the <a href="https://example.com/runbook">runbook</a><br>before deploy.</p>
        <table>
          <tr><td>Name</td><td>Status</td></tr>
          <tr><td>Node A</td><td>Healthy</td></tr>
        </table>
        <p><a href="https://example.com/zoom" title="View in Zoom"></a></p>
        """,
        subtype="html",
    )

    body_markdown, body_text = app.get_body_variants(message, 1000)

    assert body_markdown == (
        "## Runbook\n\n"
        "Open the [runbook](https://example.com/runbook)\\\n"
        "before deploy.\n\n"
        "| Name | Status |\n"
        "| --- | --- |\n"
        "| Node A | Healthy |\n\n"
        "[View in Zoom](https://example.com/zoom)"
    )
    assert body_text == (
        "Runbook\n\n"
        "Open the runbook (https://example.com/runbook)\n"
        "before deploy.\n\n"
        "Name | Status\n"
        "Node A | Healthy\n\n"
        "View in Zoom (https://example.com/zoom)"
    )


def test_find_thunderbird_profile_auto_detects_first_profile(
    tmp_path: Path,
    fixture_source_profile: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    profiles_root = tmp_path / "Library" / "Thunderbird" / "Profiles"
    first_profile = profiles_root / "a.default-release"
    second_profile = profiles_root / "z.default-release"
    shutil.copytree(fixture_source_profile, first_profile / "ImapMail")
    shutil.copytree(fixture_source_profile, second_profile / "ImapMail")

    monkeypatch.setattr(app, "RUNTIME_PROFILE_OVERRIDE", "")
    monkeypatch.setattr(app.Path, "home", classmethod(lambda cls: tmp_path))

    assert app.find_thunderbird_profile() == first_profile


def test_discover_mail_accounts_lists_profile_accounts(fixture_profile: Path) -> None:
    accounts = app.discover_mail_accounts(fixture_profile)

    assert accounts == [
        app.MailAccount(
            profile_path=fixture_profile,
            account_path=fixture_profile / "ImapMail" / "mail.example.invalid",
            storage_root="ImapMail",
            account_name="mail.example.invalid",
        )
    ]


def test_find_mbox_path_requires_account_when_multiple_accounts(
    fixture_profile: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    add_second_account(fixture_profile)
    monkeypatch.setattr(app, "RUNTIME_ACCOUNT_OVERRIDE", "")

    with pytest.raises(ValueError, match="Multiple Thunderbird mail accounts were found"):
        app.find_mbox_path(fixture_profile, "INBOX")


def test_find_mbox_path_uses_selected_account(
    fixture_profile: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    second_account_path = add_second_account(fixture_profile)
    monkeypatch.setattr(app, "RUNTIME_ACCOUNT_OVERRIDE", "mail.backup.invalid")

    assert app.find_mbox_path(fixture_profile, "INBOX") == second_account_path / "INBOX"


def test_render_account_discovery_includes_accounts(fixture_profile: Path) -> None:
    add_second_account(fixture_profile)

    rendered = app.render_account_discovery(
        fixture_profile,
        app.discover_mail_accounts(fixture_profile),
    )

    assert f"Thunderbird profile: {fixture_profile}" in rendered
    assert "- mail.backup.invalid (ImapMail)" in rendered
    assert "- mail.example.invalid (ImapMail)" in rendered
    assert "--list-accounts" not in rendered
    assert "--account <account-name>" in rendered


def test_render_markdown_and_write_markdown(
    fixture_profile: Path,
    frozen_now: None,
    tmp_path: Path,
) -> None:
    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file=str(tmp_path / "emails.md"),
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
        )
    )

    markdown = app.render_markdown(result.emails)
    app.write_markdown(result.emails, tmp_path / "emails.md")

    assert markdown.startswith("---\n**From:** Alice Example <alice@example.com>")
    assert "**Subject:** Möbius meeting assets" in markdown
    assert (tmp_path / "emails.md").read_text(encoding="utf-8") == markdown


def test_write_json_outputs_expected_array(
    fixture_profile: Path,
    frozen_now: None,
    capsys: pytest.CaptureFixture[str],
) -> None:
    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file="~/emails_last_week.md",
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
        )
    )

    app.write_json(result.emails)
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert len(payload) == 2
    assert payload[0]["subject"] == "Weekly sync notes"
    assert payload[1]["body"] == "Hello\\\nworld\n\nRésumé"


def test_run_extraction_raises_for_empty_mailbox(
    tmp_path: Path,
) -> None:
    profile_path = tmp_path / "empty.default-release"
    mbox_root = profile_path / "ImapMail" / "mail.example.invalid"
    mbox_root.mkdir(parents=True)
    (mbox_root / "INBOX").touch()

    with pytest.raises(RuntimeError, match="is empty"):
        app.run_extraction(
            app.Config(
                days_back=7,
                mail_folder="INBOX",
                output_file="~/emails_last_week.md",
                max_body_length=1000,
                thunderbird_profile=str(profile_path),
            )
        )


def test_run_extraction_requires_account_when_multiple_accounts(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    add_second_account(fixture_profile)

    with pytest.raises(ValueError, match="Multiple Thunderbird mail accounts were found"):
        app.run_extraction(
            app.Config(
                days_back=7,
                mail_folder="INBOX",
                output_file="~/emails_last_week.md",
                max_body_length=1000,
                thunderbird_profile=str(fixture_profile),
            )
        )


def test_run_extraction_uses_requested_account_when_multiple_accounts(
    fixture_profile: Path,
    frozen_now: None,
) -> None:
    second_account_path = add_second_account(fixture_profile)

    result = app.run_extraction(
        app.Config(
            days_back=7,
            mail_folder="INBOX",
            output_file="~/emails_last_week.md",
            max_body_length=1000,
            thunderbird_profile=str(fixture_profile),
            thunderbird_account="mail.backup.invalid",
        )
    )

    assert result.account_path == second_account_path
    assert result.mbox_path == second_account_path / "INBOX"
