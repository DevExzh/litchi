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

`--evidence LABEL=RESULTS` (repeatable, oldest run first) folds several runs by
package SHA-256 instead (record 0433). A regenerated queue's packages keep their
pass only where the bytes ran: every packaged row is rewritten to `pass` when
each of its packages' bytes last passed, `FAIL` when any last failed, and
`pending` naming the packages no run has seen. Rows with no package keep their
cells. `--unproven FILE` writes those packages' ids, one per line, for the next
run's kit.

    python3 tools/mac_queue/apply_results.py --manifest OUT/manifest.json \\
        --evidence "run 1=results-1.tsv" --evidence "run 2=results-2.tsv" \\
        --record 0433 --next-run "run 3" --unproven unproven.txt
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


def evidence_by_digest(runs: list[tuple[str, dict[str, dict[str, str]]]]) -> tuple[dict[str, dict[str, str]], list[str]]:
    """Package SHA-256 -> its last result across `runs`, oldest first, with the
    run's label in `run`. A package a later run disagrees with is listed: the
    later result stands, and the disagreement needs a reader."""
    latest: dict[str, dict[str, str]] = {}
    conflicts: list[str] = []
    for label, results in runs:
        seen: dict[str, dict[str, str]] = {}
        for line in results.values():
            digest = line["package_sha256"]
            earlier = seen.get(digest)
            if earlier is not None and earlier["result"] != line["result"]:
                conflicts.append(f"{digest}: {label} disagrees with itself")
            seen[digest] = dict(line, run=label)
        for digest, line in seen.items():
            earlier = latest.get(digest)
            if earlier is not None and earlier["result"] != line["result"]:
                conflicts.append(f"{digest}: {earlier['run']} {earlier['result']}, {label} {line['result']}")
            latest[digest] = line
    return latest, conflicts


def evidence_cells(entry: dict, packages: dict[str, dict], evidence: dict[str, dict[str, str]],
                   record: str, next_run: str) -> tuple[str, str, list[str]]:
    """Result and Artifact SHA-256 cells of one packaged row, and its packages
    that no run has seen."""
    ids = entry["packages"]
    lines = [evidence.get(packages[identifier]["sha256"]) for identifier in ids]
    unproven = [identifier for identifier, line in zip(ids, lines) if line is None]
    failed = [(identifier, line) for identifier, line in zip(ids, lines)
              if line is not None and line["result"] != "pass"]
    artifacts = ", ".join(
        f"{identifier} {line['resaved_sha256']}" if len(ids) > 1 else line["resaved_sha256"]
        for identifier, line in zip(ids, lines)
        if line is not None and line["resaved_sha256"] not in ("", "-")
    ) or "-"
    if failed:
        detail = ", ".join(f"{identifier} {line['stage']} in {line['run']}" for identifier, line in failed)
        return f"FAIL ({record}: {detail})", artifacts, unproven
    if unproven:
        return f"pending ({record}: {next_run}: {', '.join(unproven)})", artifacts, unproven
    runs = sorted({line["run"] for line in lines})
    apps = sorted({f"{line['app']} {line['app_version']}" for line in lines})
    date = max(line["started_utc"] for line in lines)[:10]
    return f"pass ({record}: {' + '.join(runs)}: {', '.join(apps)}, {date})", artifacts, []


def apply_evidence(queue_text: str, manifest: dict, evidence: dict[str, dict[str, str]], *,
                   record: str, next_run: str) -> tuple[str, list[str], list[str]]:
    """Rewrite every packaged row from `evidence`: the updated queue, one change
    line per row that changed, and the packages the next run needs, in order."""
    packages = {package["id"]: package for package in manifest["packages"]}
    by_id = {entry["id"]: entry for entry in manifest["rows"]}
    rows = {row.line: row for row in mac_queue.parse_queue(queue_text)}
    known = {row.id for row in rows.values()}
    stale = sorted(set(by_id) - known)
    if stale:
        raise ValueError(
            f"{len(stale)} manifest rows are no longer in the queue (first {stale[0]}); regenerate the manifest"
        )
    missing = sorted(row.id for row in rows.values() if row.id not in by_id)
    if missing:
        raise ValueError(f"{len(missing)} queue rows are not in the manifest (first {missing[0]}); regenerate it")
    lines = queue_text.split("\n")
    changes: list[str] = []
    needed: list[str] = []
    for number, row in rows.items():
        entry = by_id[row.id]
        if not entry["packages"]:
            continue
        result, artifact, unproven = evidence_cells(entry, packages, evidence, record, next_run)
        for identifier in unproven:
            if identifier not in needed:
                needed.append(identifier)
        if (row.cells["result"], row.cells["artifact sha-256"]) == (result, artifact):
            continue
        values = [row.cells[column] for column in mac_queue.COLUMNS[:5]] + [result, artifact]
        lines[number - 1] = "| " + " | ".join(values) + " |"
        changes.append(f"line {number} {row.id}: {result}")
    return "\n".join(lines), changes, needed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--results", type=Path)
    parser.add_argument("--evidence", action="append", default=[], metavar="LABEL=RESULTS",
                        help="one run's results, oldest first; folds by package SHA-256")
    parser.add_argument("--record", default="0433", help="the record cited in --evidence cells")
    parser.add_argument("--next-run", default="run 3", help="the run named in pending cells")
    parser.add_argument("--unproven", type=Path, help="with --evidence: write the packages no run has seen")
    parser.add_argument("--queue", type=Path, default=mac_queue.ROOT / mac_queue.MAC_QUEUE)
    parser.add_argument("--dry-run", action="store_true", help="print the changes, do not write")
    parser.add_argument("--force", action="store_true", help="also rewrite rows that are not pending")
    args = parser.parse_args(argv)

    if bool(args.results) == bool(args.evidence):
        parser.error("give either --results or --evidence")
    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    text = args.queue.read_text(encoding="utf-8")
    if args.evidence:
        runs = []
        for item in args.evidence:
            label, separator, path = item.partition("=")
            if not separator:
                parser.error(f"--evidence {item}: expected LABEL=RESULTS")
            runs.append((label, read_results(Path(path))))
        evidence, conflicts = evidence_by_digest(runs)
        try:
            updated, changes, needed = apply_evidence(
                text, manifest, evidence, record=args.record, next_run=args.next_run)
        except ValueError as error:
            print(f"apply_results: {error}", file=sys.stderr)
            return 1
        for conflict in conflicts:
            print(f"conflict: {conflict}")
        for change in changes:
            print(change)
        print(f"{len(evidence)} packages with a result; {len(changes)} queue rows updated; "
              f"{len(needed)} packages need {args.next_run}{' (dry run)' if args.dry_run else ''}")
        if not args.dry_run:
            if updated != text:
                args.queue.write_text(updated, encoding="utf-8")
            if args.unproven:
                args.unproven.write_text("".join(f"{identifier}\n" for identifier in needed), encoding="utf-8")
        return 0
    results = read_results(args.results)
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
