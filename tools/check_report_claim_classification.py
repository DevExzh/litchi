#!/usr/bin/env python3
"""Validate the bounded classification registry for the two REPORT tables.

The performance report predates the strict claim registry and contains a large
amount of useful historical material.  This checker gives that material an
explicit disposition without treating a Markdown paragraph as claim evidence.
It deliberately parses only the two audited tables.  Every table header,
ordinal, label, and row digest is bound by the sidecar registry, so a changed
or reordered row fails closed instead of silently changing meaning.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import stat
import sys
from pathlib import Path
from typing import Any

if __package__:
    from tools import check_perf_claims
else:
    import check_perf_claims


SCHEMA_VERSION = 1
REGISTRY_KIND = "litchi-performance-report-claim-classification"
REPORT_RELATIVE_PATH = Path("docs/performance/REPORT.md")
CLAIM_REGISTRY_RELATIVE_PATH = Path("docs/performance/claim-registry-v1.json")
STATES = ("strict_claim", "historical", "descriptive", "withheld")
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
ATX_HEADING_RE = re.compile(r"^[ ]{0,3}#{1,6}(?:[ \t]+|$)")
SETEXT_UNDERLINE_RE = re.compile(r"^[ ]{0,3}(?:=+|-+)[ \t]*$")
FENCE_RE = re.compile(r"^[ ]{0,3}(`{3,}|~{3,})(.*)$")
RAW_HTML_TAG_RE = re.compile(
    r"</?\s*(pre|script|style|div|details|table|section|article|blockquote|aside|figure|main|nav|p|h[1-6])\b[^>]*>",
    re.IGNORECASE,
)
HTML_COMMENT_OPEN = "<!--"
HTML_COMMENT_CLOSE = "-->"
EXPECTED_TABLES = (
    {
        "id": "historical-stable-tranche",
        "section_heading": "### Historical stable-tranche table (descriptive; not current claims)",
        "preamble": (
            "",
            "This aggregate table is the only stable-tranche table covered by the",
            "classification sidecar; surrounding report prose and change sections remain",
            "outside its scope.",
            "",
        ),
        "header": [
            "Change",
            "Historical/descriptive evidence (not a current claim)",
            "Historical scope / limitation",
        ],
        "row_count": 88,
    },
    {
        "id": "historical-accepted-results",
        "section_heading": "## Historical accepted results (descriptive; not current claims)",
        "preamble": (
            "",
            "The measurements below are retained historical results, presented for",
            "descriptive context rather than as current performance claims. They are warm-",
            "memory release-build p50 results from matched before/after binaries. Each",
            "linked change record contains raw-sample counts, ABBA ordering, mean or",
            "interval context, hashes, and memory profiles; this table does not promote any",
            "row into the strict claim registry.",
            "",
        ),
        "header": [
            "Workload group",
            "Historical before",
            "Historical after",
            "Historical result (not a current claim)",
            "Historical memory result",
        ],
        "row_count": 79,
    },
)


class ClassificationError(Exception):
    """Base class for malformed input or a classification policy failure."""


class ClassificationInputError(ClassificationError):
    """Input is unavailable, malformed, or does not match the audited shape."""


class ClassificationPolicyError(ClassificationError):
    """The registry contains an invalid classification or linkage."""


def _reject_constant(value: str) -> Any:
    raise ValueError(f"non-finite JSON constant {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _validate_json_tree(value: Any, location: str) -> None:
    if value is None or isinstance(value, (bool, str, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ClassificationInputError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, child in enumerate(value):
            _validate_json_tree(child, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, child in value.items():
            if not isinstance(key, str):
                raise ClassificationInputError(f"{location} has a non-string key")
            _validate_json_tree(child, f"{location}.{key}")
        return
    raise ClassificationInputError(f"{location} contains an unsupported JSON value")


def load_json(path: Path, *, location: str | None = None) -> Any:
    label = location or str(path)
    try:
        raw = path.read_bytes()
    except OSError as error:
        raise ClassificationInputError(f"cannot read {label}: {error}") from error
    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        raise ClassificationInputError(f"invalid JSON in {label}: {error}") from error
    _validate_json_tree(value, label)
    return value


def canonical_bytes(value: Any) -> bytes:
    _validate_json_tree(value, "JSON")
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise ClassificationInputError(f"cannot canonicalize JSON: {error}") from error


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def cells_sha256(cells: list[str]) -> str:
    """Hash trimmed Markdown cells using the sidecar's v1 canonicalization."""

    return sha256_bytes(canonical_bytes(cells))


def _split_pipe_row(line: str, *, location: str) -> list[str]:
    """Split one pipe-table row while preserving escaped pipe characters."""

    value = line.rstrip("\r\n")
    if not value.startswith("|") or not value.endswith("|"):
        raise ClassificationInputError(f"{location} is not a closed Markdown table row")
    cells: list[str] = []
    start = 1
    escaped = False
    for index in range(1, len(value)):
        char = value[index]
        if char == "|" and not escaped:
            cells.append(value[start:index].strip())
            start = index + 1
        if char == "\\" and not escaped:
            escaped = True
        else:
            escaped = False
    if len(cells) < 2 or any("\n" in cell or "\r" in cell for cell in cells):
        raise ClassificationInputError(f"{location} has malformed Markdown table cells")
    return cells


def _is_separator(cells: list[str]) -> bool:
    return all(re.fullmatch(r":?-{3,}:?", cell) is not None for cell in cells)


def _hidden_lines(lines: list[str]) -> list[bool]:
    """Mask fenced and paired raw HTML blocks that are not rendered Markdown tables.

    A report is Markdown, but a pipe-shaped table in a fenced block, HTML
    comment, or raw HTML container is not a rendered table.  Keeping one
    fail-closed mask for all of these containers also prevents headings inside
    them from changing the section boundary.
    """

    masked = [False] * len(lines)
    fence: tuple[str, int] | None = None
    html_comment = False
    raw_stack: list[str] = []
    ambiguous_html = False
    for index, line in enumerate(lines):
        if fence is not None:
            masked[index] = True
            match = FENCE_RE.fullmatch(line)
            marker = match.group(1) if match else None
            if marker is not None and marker[0] == fence[0] and len(marker) >= fence[1]:
                if not match.group(2).strip():
                    fence = None
            continue

        if ambiguous_html:
            masked[index] = True
            continue

        if not raw_stack and not html_comment:
            fence_match = FENCE_RE.fullmatch(line)
            if fence_match is not None:
                masked[index] = True
                marker = fence_match.group(1)
                fence = (marker[0], len(marker))
                continue

        masked[index] = bool(raw_stack or html_comment)
        cursor = 0
        while cursor < len(line):
            if html_comment:
                close_index = line.find(HTML_COMMENT_CLOSE, cursor)
                if close_index < 0:
                    break
                html_comment = False
                cursor = close_index + len(HTML_COMMENT_CLOSE)
                continue

            comment_index = line.find(HTML_COMMENT_OPEN, cursor)
            stray_close_index = line.find(HTML_COMMENT_CLOSE, cursor)
            tag_match = RAW_HTML_TAG_RE.search(line, cursor)
            events = [
                (position, "comment", None)
                for position in (comment_index,)
                if position >= 0
            ]
            events.extend(
                (position, "stray_close", None)
                for position in (stray_close_index,)
                if position >= 0
            )
            if tag_match is not None:
                events.append((tag_match.start(), "tag", tag_match))
            if not events:
                break
            position, event_kind, event_value = min(events, key=lambda event: event[0])
            if event_kind == "comment":
                masked[index] = True
                html_comment = True
                cursor = position + len(HTML_COMMENT_OPEN)
                continue
            if event_kind == "stray_close":
                masked[index] = True
                ambiguous_html = True
                break

            assert event_value is not None
            masked[index] = True
            token = event_value.group(0)
            tag = event_value.group(1).lower()
            if token.lstrip().startswith("</"):
                if not raw_stack or raw_stack[-1] != tag:
                    ambiguous_html = True
                    break
                raw_stack.pop()
            elif not token.rstrip().endswith("/>"):
                raw_stack.append(tag)
            cursor = event_value.end()

        if ambiguous_html:
            masked[index] = True
        continue
    return masked


def _parse_table_candidates(lines: list[str], hidden: list[bool]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for index in range(len(lines) - 1):
        if hidden[index] or hidden[index + 1]:
            continue
        if not lines[index].startswith("|") or not lines[index + 1].startswith("|"):
            continue
        header = _split_pipe_row(lines[index], location=f"REPORT.md:{index + 1}")
        separator = _split_pipe_row(lines[index + 1], location=f"REPORT.md:{index + 2}")
        if len(header) != len(separator) or not _is_separator(separator):
            continue
        rows: list[list[str]] = []
        cursor = index + 2
        while cursor < len(lines) and not hidden[cursor] and lines[cursor].startswith("|"):
            row = _split_pipe_row(lines[cursor], location=f"REPORT.md:{cursor + 1}")
            if len(row) != len(header):
                raise ClassificationInputError(
                    f"REPORT.md:{cursor + 1} has {len(row)} cells; expected {len(header)}"
                )
            rows.append(row)
            cursor += 1
        candidates.append(
            {
                "start_line": index + 1,
                "header": header,
                "separator": separator,
                "rows": rows,
                "end_line": cursor,
            }
        )
    return candidates


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ClassificationInputError(f"{location} must be an object")
    return value


def _require_list(value: Any, location: str) -> list[Any]:
    if not isinstance(value, list):
        raise ClassificationInputError(f"{location} must be an array")
    return value


def _require_string(value: Any, location: str) -> str:
    if not isinstance(value, str) or not value:
        raise ClassificationInputError(f"{location} must be a non-empty string")
    return value


def _require_int(value: Any, location: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ClassificationInputError(f"{location} must be an integer")
    return value


def _require_exact_keys(value: dict[str, Any], expected: set[str], location: str) -> None:
    actual = set(value)
    missing = expected - actual
    extra = actual - expected
    if missing or extra:
        bits = []
        if missing:
            bits.append(f"missing {sorted(missing)}")
        if extra:
            bits.append(f"unexpected {sorted(extra)}")
        raise ClassificationInputError(f"{location} has invalid keys ({'; '.join(bits)})")


def _safe_relative_path(value: Any, *, location: str, repo_root: Path) -> Path:
    raw = _require_string(value, location)
    if raw.startswith("/") or "\\" in raw or "\x00" in raw:
        raise ClassificationInputError(f"{location} must be a safe relative path")
    if any(part in {"", ".", ".."} for part in raw.split("/")):
        raise ClassificationInputError(f"{location} must be a safe relative path")
    relative = Path(raw)
    if relative.is_absolute():
        raise ClassificationInputError(f"{location} must be a safe relative path")
    resolved = (repo_root / relative).resolve()
    root = repo_root.resolve()
    try:
        resolved.relative_to(root)
    except ValueError as error:
        raise ClassificationInputError(f"{location} escapes the repository") from error
    return relative


def _resolved_repo_root(repo_root: Path, *, location: str) -> Path:
    try:
        root = repo_root.resolve(strict=True)
    except (OSError, RuntimeError) as error:
        raise ClassificationInputError(f"{location} is not a readable repository root") from error
    try:
        root_mode = root.stat().st_mode
    except OSError as error:
        raise ClassificationInputError(f"{location} cannot be inspected") from error
    if not stat.S_ISDIR(root_mode):
        raise ClassificationInputError(f"{location} must be a directory")
    return root


def _checked_repo_relative_file(
    relative: Path, *, repo_root: Path, location: str
) -> Path:
    """Return a regular file without traversing a symlinked component."""

    root = _resolved_repo_root(repo_root, location="repo_root")
    if relative.is_absolute() or not relative.parts or any(
        part in {"", ".", ".."} for part in relative.parts
    ):
        raise ClassificationInputError(f"{location} must be a safe repository-relative file")
    current = root
    last_index = len(relative.parts) - 1
    for index, part in enumerate(relative.parts):
        current /= part
        try:
            mode = current.lstat().st_mode
        except OSError as error:
            raise ClassificationInputError(f"{location} cannot be read") from error
        if stat.S_ISLNK(mode):
            raise ClassificationInputError(
                f"{location} must not traverse a symbolic link ({part!r})"
            )
        if index == last_index:
            if not stat.S_ISREG(mode):
                raise ClassificationInputError(f"{location} must be a regular file")
        elif not stat.S_ISDIR(mode):
            raise ClassificationInputError(f"{location} has a non-directory component")
    return current


def _checked_sidecar_file(registry_path: Path, *, repo_root: Path) -> Path:
    """Check the CLI sidecar path lexically before opening it."""

    raw = str(registry_path)
    if not raw or "\\" in raw or "\x00" in raw:
        raise ClassificationInputError("classification sidecar path is not a safe POSIX path")
    root = _resolved_repo_root(repo_root, location="repo_root")
    candidate = registry_path if registry_path.is_absolute() else root / registry_path
    try:
        relative = candidate.relative_to(root)
    except ValueError as error:
        raise ClassificationInputError(
            "classification sidecar must be inside the repository"
        ) from error
    return _checked_repo_relative_file(
        relative, repo_root=root, location="classification sidecar"
    )


def _find_heading(lines: list[str], heading: str, hidden: list[bool]) -> int:
    matches = [index for index, line in enumerate(lines) if not hidden[index] and line == heading]
    if len(matches) != 1:
        raise ClassificationInputError(
            f"REPORT.md must contain exactly one heading {heading!r}; found {len(matches)}"
        )
    return matches[0]


def _next_heading_index(lines: list[str], heading_index: int, hidden: list[bool]) -> int:
    for index in range(heading_index + 1, len(lines)):
        if hidden[index]:
            continue
        if ATX_HEADING_RE.match(lines[index]):
            return index
        if (
            index > heading_index + 1
            and SETEXT_UNDERLINE_RE.fullmatch(lines[index])
            and not hidden[index - 1]
            and lines[index - 1].strip()
        ):
            return index - 1
    return len(lines)


def _validate_claim_linkage(registry: dict[str, Any], *, repo_root: Path) -> None:
    linkage = _require_object(registry["strict_claim_registry"], "strict_claim_registry")
    _require_exact_keys(linkage, {"path", "links"}, "strict_claim_registry")
    relative = _safe_relative_path(
        linkage["path"], location="strict_claim_registry.path", repo_root=repo_root
    )
    if relative != CLAIM_REGISTRY_RELATIVE_PATH:
        raise ClassificationInputError(
            f"strict_claim_registry.path must be {CLAIM_REGISTRY_RELATIVE_PATH}"
        )
    links = _require_list(linkage["links"], "strict_claim_registry.links")
    claim_ids: list[str] = []
    link_keys: list[tuple[str, int]] = []
    link_rows: dict[tuple[str, int], dict[str, Any]] = {}
    for index, link_value in enumerate(links):
        link = _require_object(link_value, f"strict_claim_registry.links[{index}]")
        _require_exact_keys(
            link,
            {"claim_id", "table_id", "ordinal", "row_sha256"},
            f"strict_claim_registry.links[{index}]",
        )
        claim_id = _require_string(link["claim_id"], f"strict_claim_registry.links[{index}].claim_id")
        table_id = _require_string(link["table_id"], f"strict_claim_registry.links[{index}].table_id")
        ordinal = _require_int(link["ordinal"], f"strict_claim_registry.links[{index}].ordinal")
        row_sha256 = _require_string(link["row_sha256"], f"strict_claim_registry.links[{index}].row_sha256")
        if not SHA256_RE.fullmatch(row_sha256):
            raise ClassificationPolicyError(f"strict_claim_registry.links[{index}] has an invalid row SHA-256")
        key = (table_id, ordinal)
        if key in link_rows or claim_id in claim_ids:
            raise ClassificationPolicyError("strict claim links must be duplicate-free")
        claim_ids.append(claim_id)
        link_keys.append(key)
        link_rows[key] = link
    if link_keys != sorted(link_keys) or claim_ids != sorted(claim_ids):
        raise ClassificationPolicyError("strict claim links must be sorted by table/ordinal and claim ID")
    strict_rows: list[tuple[str, int, str, str]] = []
    for table in registry["report"]["tables"]:
        for row in table["rows"]:
            if row["state"] == "strict_claim":
                claim_id = row.get("claim_id")
                if not isinstance(claim_id, str) or not claim_id:
                    raise ClassificationPolicyError(
                        f"strict row {table['id']} ordinal {row['ordinal']} lacks claim_id"
                    )
                strict_rows.append((table["id"], row["ordinal"], claim_id, row["row_sha256"]))
            elif "claim_id" in row:
                raise ClassificationPolicyError(
                    f"non-strict row {table['id']} ordinal {row['ordinal']} has claim_id"
                )
    expected_links = [
        (table_id, ordinal, claim_id, row_sha256)
        for table_id, ordinal, claim_id, row_sha256 in strict_rows
    ]
    actual_links = [
        (link["table_id"], link["ordinal"], link["claim_id"], link["row_sha256"])
        for link in links
    ]
    if expected_links != actual_links:
        raise ClassificationPolicyError(
            "strict claim links must exactly biject strict rows, including row digests"
        )

    try:
        claim_registry = check_perf_claims.load_json(
            _checked_repo_relative_file(
                CLAIM_REGISTRY_RELATIVE_PATH,
                repo_root=repo_root,
                location=str(CLAIM_REGISTRY_RELATIVE_PATH),
            ),
            location=str(CLAIM_REGISTRY_RELATIVE_PATH),
        )
        _, _, canonical_claims = check_perf_claims.validate_registry(
            claim_registry, repo_root=repo_root
        )
    except check_perf_claims.ClaimRegistryError as error:
        raise ClassificationInputError(
            f"canonical strict claim registry failed validation: {error}"
        ) from error

    for claim_id in claim_ids:
        parsed = canonical_claims.get(claim_id)
        if parsed is None:
            raise ClassificationPolicyError(f"strict claim {claim_id!r} is not in the claim registry")
        claim = parsed["value"]
        if claim.get("status") != "landed" or claim.get("code_state") != "landed":
            raise ClassificationPolicyError(
                f"strict claim {claim_id!r} is not both status=landed and code_state=landed"
            )


def validate_registry(registry: Any, *, repo_root: Path) -> dict[str, Any]:
    root = _require_object(registry, "registry")
    _require_exact_keys(
        root,
        {
            "schema_version",
            "registry_kind",
            "canonicalization",
            "scope",
            "report",
            "strict_claim_registry",
        },
        "registry",
    )
    if root["schema_version"] != SCHEMA_VERSION:
        raise ClassificationInputError("unsupported report-claim-classification schema_version")
    if root["registry_kind"] != REGISTRY_KIND:
        raise ClassificationInputError("unsupported report-claim-classification registry_kind")
    scope = root.get("scope")
    if scope != {
        "kind": "two_historical_report_tables",
        "description": (
            "Only the two audited historical tables in docs/performance/REPORT.md "
            "are classified; other report prose and tables are outside this registry."
        ),
    }:
        raise ClassificationInputError("registry scope must cover only the two audited historical tables")
    canonicalization = _require_object(root["canonicalization"], "canonicalization")
    _require_exact_keys(canonicalization, {"algorithm", "hash"}, "canonicalization")
    if canonicalization != {
        "algorithm": "sorted-json-utf8-compact-v1",
        "hash": "sha256",
    }:
        raise ClassificationInputError("unsupported report row canonicalization")
    report = _require_object(root["report"], "report")
    _require_exact_keys(report, {"path", "tables"}, "report")
    report_path = _safe_relative_path(report["path"], location="report.path", repo_root=repo_root)
    if report_path != REPORT_RELATIVE_PATH:
        raise ClassificationInputError(f"report.path must be {REPORT_RELATIVE_PATH}")
    tables = _require_list(report["tables"], "report.tables")
    if len(tables) != len(EXPECTED_TABLES):
        raise ClassificationInputError(f"report.tables must contain exactly {len(EXPECTED_TABLES)} tables")
    expected_by_id = {item["id"]: item for item in EXPECTED_TABLES}
    table_ids: list[str] = []
    for index, table_value in enumerate(tables):
        table = _require_object(table_value, f"report.tables[{index}]")
        _require_exact_keys(
            table,
            {"id", "section_heading", "header", "header_sha256", "row_count", "rows"},
            f"report.tables[{index}]",
        )
        table_id = _require_string(table["id"], f"report.tables[{index}].id")
        if table_id in table_ids:
            raise ClassificationPolicyError(f"duplicate report table id {table_id!r}")
        table_ids.append(table_id)
        expected = expected_by_id.get(table_id)
        if expected is None:
            raise ClassificationInputError(f"unexpected report table id {table_id!r}")
        if table["section_heading"] != expected["section_heading"]:
            raise ClassificationInputError(f"{table_id}.section_heading does not match the audited heading")
        header = _require_list(table["header"], f"{table_id}.header")
        if header != expected["header"] or any(not isinstance(cell, str) for cell in header):
            raise ClassificationInputError(f"{table_id}.header does not match the audited header")
        header_digest = _require_string(table["header_sha256"], f"{table_id}.header_sha256")
        if not SHA256_RE.fullmatch(header_digest) or header_digest != cells_sha256(header):
            raise ClassificationPolicyError(f"{table_id}.header_sha256 does not match header")
        row_count = _require_int(table["row_count"], f"{table_id}.row_count")
        if row_count != expected["row_count"]:
            raise ClassificationInputError(f"{table_id}.row_count does not match the audited count")
        rows = _require_list(table["rows"], f"{table_id}.rows")
        if len(rows) != row_count:
            raise ClassificationPolicyError(f"{table_id}.rows length does not equal row_count")
        seen_labels: set[str] = set()
        seen_digests: set[str] = set()
        for ordinal, row_value in enumerate(rows, start=1):
            row = _require_object(row_value, f"{table_id}.rows[{ordinal - 1}]")
            allowed_row_keys = {"ordinal", "label", "row_sha256", "state", "claim_id"}
            if not set(row).issubset(allowed_row_keys) or not {"ordinal", "label", "row_sha256", "state"}.issubset(row):
                raise ClassificationInputError(f"{table_id}.rows[{ordinal - 1}] has invalid keys")
            if _require_int(row["ordinal"], f"{table_id}.rows[{ordinal - 1}].ordinal") != ordinal:
                raise ClassificationPolicyError(f"{table_id} row ordinals must be contiguous from one")
            label = _require_string(row["label"], f"{table_id}.rows[{ordinal - 1}].label")
            digest = _require_string(row["row_sha256"], f"{table_id}.rows[{ordinal - 1}].row_sha256")
            if not SHA256_RE.fullmatch(digest):
                raise ClassificationPolicyError(f"{table_id} row {ordinal} has an invalid SHA-256")
            state = _require_string(row["state"], f"{table_id}.rows[{ordinal - 1}].state")
            if state not in STATES:
                raise ClassificationPolicyError(f"{table_id} row {ordinal} has unknown state {state!r}")
            if label in seen_labels:
                raise ClassificationPolicyError(f"{table_id} contains duplicate row label {label!r}")
            if digest in seen_digests:
                raise ClassificationPolicyError(f"{table_id} contains duplicate row digest {digest!r}")
            seen_labels.add(label)
            seen_digests.add(digest)
    if table_ids != [item["id"] for item in EXPECTED_TABLES]:
        raise ClassificationPolicyError("report tables must be in audited order")

    report_path_absolute = _checked_repo_relative_file(
        report_path, repo_root=repo_root, location=str(report_path)
    )
    try:
        lines = report_path_absolute.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        raise ClassificationInputError(f"cannot read {report_path}: {error}") from error
    hidden = _hidden_lines(lines)
    candidates = _parse_table_candidates(lines, hidden)
    for expected, table in zip(EXPECTED_TABLES, tables):
        heading_index = _find_heading(lines, expected["section_heading"], hidden)
        next_heading_index = _next_heading_index(lines, heading_index, hidden)
        preamble = list(expected["preamble"])
        preamble_start = heading_index + 1
        preamble_end = preamble_start + len(preamble)
        if lines[preamble_start:preamble_end] != preamble:
            raise ClassificationInputError(
                f"{expected['id']} section preamble does not match the audited report"
            )
        header_line = preamble_end
        if header_line >= len(lines) or header_line >= next_heading_index:
            raise ClassificationInputError(
                f"{expected['id']} table header is not at the audited line"
            )
        matches = [
            candidate
            for candidate in candidates
            if candidate["header"] == expected["header"]
            and candidate["start_line"] == header_line + 1
        ]
        if len(matches) != 1:
            raise ClassificationInputError(
                f"{expected['id']} must have exactly one matching Markdown table; found {len(matches)}"
            )
        candidate = matches[0]
        if len(candidate["rows"]) != expected["row_count"]:
            raise ClassificationInputError(
                f"{expected['id']} Markdown row count is {len(candidate['rows'])}, expected {expected['row_count']}"
            )
        registry_rows = table["rows"]
        for ordinal, (cells, row) in enumerate(zip(candidate["rows"], registry_rows), start=1):
            if row["label"] != cells[0]:
                raise ClassificationPolicyError(f"{expected['id']} row {ordinal} label does not match REPORT.md")
            digest = cells_sha256(cells)
            if row["row_sha256"] != digest:
                raise ClassificationPolicyError(f"{expected['id']} row {ordinal} digest does not match REPORT.md")
    _validate_claim_linkage(root, repo_root=repo_root)
    counts = {state: 0 for state in STATES}
    for table in tables:
        for row in table["rows"]:
            counts[row["state"]] += 1
    return {"table_count": len(tables), "row_count": sum(counts.values()), "state_counts": counts}


def lint_registry(registry_path: Path, *, repo_root: Path) -> tuple[int, str]:
    try:
        registry = load_json(
            _checked_sidecar_file(registry_path, repo_root=repo_root),
            location="classification sidecar",
        )
        summary = validate_registry(registry, repo_root=repo_root)
    except ClassificationError as error:
        return 1, f"ERROR: {error}"
    return (
        0,
        "OK: {row_count} REPORT rows classified across {table_count} tables "
        "({counts})".format(
            row_count=summary["row_count"],
            table_count=summary["table_count"],
            counts=", ".join(
                f"{state}={summary['state_counts'][state]}" for state in STATES
            ),
        ),
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--registry",
        type=Path,
        default=Path("docs/performance/report-claim-classification-v1.json"),
    )
    parser.add_argument("--repo-root", type=Path, default=Path("."))
    args = parser.parse_args(argv)
    repo_root = args.repo_root.resolve()
    registry_path = args.registry
    if not registry_path.is_absolute():
        registry_path = repo_root / registry_path
    status, message = lint_registry(registry_path, repo_root=repo_root)
    stream = sys.stdout if status == 0 else sys.stderr
    print(message, file=stream)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
