from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
from pathlib import Path
from typing import Any

import mcp.server.stdio
import mcp.types as types
from mcp.server import NotificationOptions, Server
from mcp.server.models import InitializationOptions

from extract_emails import (
    Config,
    render_markdown,
    run_extraction,
    write_markdown,
)

logger = logging.getLogger("thunderbird_email_mcp")
SERVER_NAME = "thunderbird-local-emails"
SERVER_VERSION = "0.1.0"


def parse_arguments() -> argparse.Namespace:
    """Parse server CLI arguments."""

    parser = argparse.ArgumentParser()
    parser.add_argument("--log-level", default=os.environ.get("LOG_LEVEL", "INFO"))
    return parser.parse_args()


def configure_logging() -> None:
    """Configure server logging."""

    args = parse_arguments()
    log_level = getattr(logging, args.log_level.upper(), logging.INFO)
    logging.basicConfig(level=log_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    logger.setLevel(log_level)


def build_tools() -> list[types.Tool]:
    """Return the MCP tools exposed by this server."""

    return [
        types.Tool(
            name="fetch_thunderbird_local_emails",
            description=(
                "Read recent emails directly from Thunderbird's local macOS mbox storage and return "
                "them as Markdown or structured JSON, with optional sender and subject filtering "
                "and no cloud access."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {
                        "type": "integer",
                        "description": "Number of days back to include. Default 7.",
                    },
                    "folder": {
                        "type": "string",
                        "description": 'Mailbox filename inside Thunderbird. Default "INBOX".',
                    },
                    "output_file": {
                        "type": "string",
                        "description": "Optional Markdown output file path used when write_to_file is true.",
                    },
                    "max_body": {
                        "type": "integer",
                        "description": "Maximum body length per email. Default 1000.",
                    },
                    "profile": {
                        "type": "string",
                        "description": "Optional Thunderbird profile path override.",
                    },
                    "sender": {
                        "type": "string",
                        "description": 'Optional case-insensitive sender filter. Partial matches are supported, for example "@zoom.us".',
                    },
                    "subject": {
                        "type": "string",
                        "description": 'Optional subject filter. Case-insensitive contains matches are supported, with fuzzy fallback for close matches like "Meeting Assets".',
                    },
                    "body_mode": {
                        "type": "string",
                        "enum": ["rendered", "both"],
                        "description": 'Body output mode. "rendered" keeps the existing body field, while "both" also includes body_markdown and body_text.',
                    },
                    "format": {
                        "type": "string",
                        "enum": ["markdown", "json"],
                        "description": "Return emails as Markdown or JSON. Default markdown.",
                    },
                    "write_to_file": {
                        "type": "boolean",
                        "description": "When true and format is markdown, also write the Markdown file to disk.",
                    },
                },
            },
        )
    ]


def text_result(text: str) -> list[types.TextContent]:
    """Wrap plain text as an MCP result."""

    return [types.TextContent(type="text", text=text)]


def _validate_optional_int(arguments: dict[str, Any], name: str, default: int) -> int:
    value = arguments.get(name, default)
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError(f"{name} must be an integer.")
    return value


def _validate_optional_str(arguments: dict[str, Any], name: str, default: str) -> str:
    value = arguments.get(name, default)
    if not isinstance(value, str):
        raise ValueError(f"{name} must be a string.")
    return value


def _validate_optional_bool(arguments: dict[str, Any], name: str, default: bool) -> bool:
    value = arguments.get(name, default)
    if not isinstance(value, bool):
        raise ValueError(f"{name} must be a boolean.")
    return value


def build_config(arguments: dict[str, Any] | None) -> tuple[Config, str, bool]:
    """Build extractor config plus MCP-only options."""

    raw_arguments = arguments or {}
    output_format = _validate_optional_str(raw_arguments, "format", "markdown").strip().lower()
    if output_format not in {"markdown", "json"}:
        raise ValueError("format must be either 'markdown' or 'json'.")

    write_to_file = _validate_optional_bool(raw_arguments, "write_to_file", False)
    if write_to_file and output_format != "markdown":
        raise ValueError("write_to_file can only be used when format is 'markdown'.")

    config = Config(
        days_back=_validate_optional_int(raw_arguments, "days", 7),
        mail_folder=_validate_optional_str(raw_arguments, "folder", "INBOX"),
        output_file=_validate_optional_str(raw_arguments, "output_file", "~/emails_last_week.md"),
        max_body_length=_validate_optional_int(raw_arguments, "max_body", 1000),
        thunderbird_profile=_validate_optional_str(raw_arguments, "profile", ""),
        sender_filter=_validate_optional_str(raw_arguments, "sender", ""),
        subject_filter=_validate_optional_str(raw_arguments, "subject", ""),
        body_mode=_validate_optional_str(raw_arguments, "body_mode", "rendered"),
        as_json=output_format == "json",
    )
    return config, output_format, write_to_file


def build_markdown_payload(config: Config, write_to_file: bool) -> dict[str, object]:
    """Run extraction and build a Markdown payload."""

    result = run_extraction(config)
    markdown = render_markdown(result.emails)
    output_file: str | None = None

    if write_to_file:
        write_markdown(result.emails, config_output_path(config))
        output_file = str(config_output_path(config))

    return {
        "format": "markdown",
        "profile_path": str(result.profile_path),
        "mbox_path": str(result.mbox_path),
        "stats": {
            "total_in_mbox": result.stats.total_in_mbox,
            "matched_in_range": result.stats.matched_in_range,
            "matched_after_filters": result.stats.matched_after_filters,
            "skipped_malformed": result.stats.skipped_malformed,
        },
        "output_file": output_file,
        "markdown": markdown,
    }


def build_json_payload(config: Config) -> dict[str, object]:
    """Run extraction and build a JSON payload."""

    result = run_extraction(config)
    return {
        "format": "json",
        "profile_path": str(result.profile_path),
        "mbox_path": str(result.mbox_path),
        "stats": {
            "total_in_mbox": result.stats.total_in_mbox,
            "matched_in_range": result.stats.matched_in_range,
            "matched_after_filters": result.stats.matched_after_filters,
            "skipped_malformed": result.stats.skipped_malformed,
        },
        "emails": result.emails,
    }


def config_output_path(config: Config) -> Path:
    """Expand the configured output path."""

    return Path(os.path.expanduser(config.output_file))


async def call_tool_impl(name: str, arguments: dict[str, Any] | None) -> list[types.TextContent]:
    """Handle MCP tool calls."""

    if name != "fetch_thunderbird_local_emails":
        return text_result(json.dumps({"error": f"Unknown tool: {name}"}, indent=2))

    try:
        config, output_format, write_to_file = build_config(arguments)
        if output_format == "markdown":
            payload = await asyncio.to_thread(build_markdown_payload, config, write_to_file)
        else:
            payload = await asyncio.to_thread(build_json_payload, config)
        return text_result(json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception as exc:  # noqa: BLE001
        logger.exception("Thunderbird extraction failed")
        return text_result(json.dumps({"error": str(exc)}, ensure_ascii=False, indent=2))


async def main() -> None:
    """Run the Thunderbird email MCP server."""

    configure_logging()
    server = Server(SERVER_NAME)

    @server.list_tools()
    async def handle_list_tools() -> list[types.Tool]:
        return build_tools()

    @server.call_tool()
    async def handle_call_tool(name: str, arguments: dict[str, Any] | None) -> list[types.TextContent]:
        return await call_tool_impl(name, arguments)

    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name=SERVER_NAME,
                server_version=SERVER_VERSION,
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


if __name__ == "__main__":
    asyncio.run(main())
