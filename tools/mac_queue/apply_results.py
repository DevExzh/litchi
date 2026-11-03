#!/usr/bin/env python3
"""Fold a macOS queue run back into `docs/unified-ops/0363-mac-queue.md`.

Record 0430. Reads the generator's `manifest.json` and the runner's
`results.tsv` and fills the Result and Artifact SHA-256 columns of every
queue row whose packages all have a result:

- every package passed: Result `pass (0430: <App> <version>, <date>)`,
  Artifact SHA-256 the re-saved package hashes (`P0001 <sha>`, ...);
- any package failed: Result `FAIL (0430: P0001 <stage>, ...)`, Artifact
  SHA-256 the hashes the run produced (or `-`).

Rows with a package that has no result yet keep their cells. Rows that are
not `pending` are left alone unless `--force`. Dry-run lines and later
results for the same package (the last line wins) are handled; a result
whose package hash differs from the manifest is an error.

    python3 tools/mac_queue/apply_results.py --manifest OUT/manifest.json \\
        --results OUT/results.tsv [--dry-run] [--force]
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import rows as mac_queue  # noqa: E402  (tools/mac_queue/rows.py)

RESULT_COLUMNS = (
    "package_id", "package_sha256", "app", "app_version", "macos", "result", "stage",
    "resaved_sha256", "resaved_kind", "pdf_sha256", "started_utc", "message",
)
PENDING = "pending"


def read_results(path: Path) -> dict[str, dict[str, str]]:
    """Package id -> its last non-dry-run result line."""
    results: dict[str, dict[str, str]] = {}
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.reader(handle, delimiter="\t", quoting=csv.QUOTE_NONE)
        header = next(reader, None)
        if header is None or tuple(header[: len(RESULT_COLUMNS)]) != RESULT_COLUMNS:
            raise ValueError(f"{path}: unexpected header {header}")
        for number, cells in enumerate(reader, start=2):
            if not cells or not any(cells):
                continue
            if len(cells) < len(RESULT_COLUMNS):
                raise ValueError(f"{path}:{number}: {len(cells)} cells, expected {len(RESULT_COLUMNS)}")
            line = dict(zip(RESULT_COLUMNS, cells))
            if line["result"] == "dry-run":
                continue
            if line["result"] not in ("pass", "fail"):
                raise ValueError(f"{path}:{number}: unknown result {line['result']!r}")
            results[line["package_id"]] = line
    return results


def cells_for(entry: dict, packages: dict[str, dict], results: dict[str, dict[str, str]]) -> tuple[str, str] | None:
    """The Result and Artifact SHA-256 cells of one manifest row, or None
    while a package of the row has no result."""
    ids = entry["packages"]
    if not ids or any(identifier not in results for identifier in ids):
        return None
    lines = [results[identifier] for identifier in ids]
    for identifier, line in zip(ids, lines):
        expected = packages[identifier]["sha256"]
        if line["package_sha256"] != expected:
            raise ValueError(
                f"{identifier}: result for package {line['package_sha256']}, manifest has {expected}"
            )
    artifacts = ", ".join(
        f"{identifier} {line['resaved_sha256']}" if len(ids) > 1 else line["resaved_sha256"]
        for identifier, line in zip(ids, lines)
        if line["resaved_sha256"] not in ("", "-")
    ) or "-"
    failed = [(identifier, line) for identifier, line in zip(ids, lines) if line["result"] != "pass"]
    if failed:
        detail = ", ".join(f"{identifier} {line['stage']}" for identifier, line in failed)
        return f"FAIL (0430: {detail})", artifacts
    apps = sorted({f"{line['app']} {line['app_version']}" for line in lines})
    date = max(line["started_utc"] for line in lines)[:10]
    return f"pass (0430: {', '.join(apps)}, {date})", artifacts


def apply(queue_text: str, manifest: dict, results: dict[str, dict[str, str]], *, force: bool = False) -> tuple[str, list[str]]:
    """Return the updated queue text and one change line per updated row."""
    packages = {package["id"]: package for package in manifest["packages"]}
    for identifier in results:
        if identifier not in packages:
            raise ValueError(f"result for unknown package {identifier}")
    by_id = {entry["id"]: entry for entry in manifest["rows"]}
    rows = {row.line: row for row in mac_queue.parse_queue(queue_text)}
    known = {row.id for row in rows.values()}
    stale = sorted(set(by_id) - known)
    if stale:
        raise ValueError(
            f"{len(stale)} manifest rows are no longer in the queue (first {stale[0]}); regenerate the manifest"
        )
    lines = queue_text.split("\n")
    changes: list[str] = []
    for number, row in rows.items():
        entry = by_id.get(row.id)
        if entry is None:
            continue
        cells = cells_for(entry, packages, results)
        if cells is None:
            continue
        if row.cells["result"] != PENDING and not force:
            continue
        result, artifact = cells
        if (row.cells["result"], row.cells["artifact sha-256"]) == (result, artifact):
            continue
        values = [row.cells[column] for column in mac_queue.COLUMNS[:5]] + [result, artifact]
        lines[number - 1] = "| " + " | ".join(values) + " |"
        changes.append(f"line {number} {row.id}: {result}")
    return "\n".join(lines), changes


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--queue", type=Path, default=mac_queue.ROOT / mac_queue.MAC_QUEUE)
    parser.add_argument("--dry-run", action="store_true", help="print the changes, do not write")
    parser.add_argument("--force", action="store_true", help="also rewrite rows that are not pending")
    args = parser.parse_args(argv)

    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    results = read_results(args.results)
    text = args.queue.read_text(encoding="utf-8")
    try:
        updated, changes = apply(text, manifest, results, force=args.force)
    except ValueError as error:
        print(f"apply_results: {error}", file=sys.stderr)
        return 1
    for change in changes:
        print(change)
    passed = sum(1 for line in results.values() if line["result"] == "pass")
    print(f"{len(results)} package results ({passed} pass, {len(results) - passed} fail); "
          f"{len(changes)} queue rows updated{' (dry run)' if args.dry_run else ''}")
    if not args.dry_run and updated != text:
        args.queue.write_text(updated, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
