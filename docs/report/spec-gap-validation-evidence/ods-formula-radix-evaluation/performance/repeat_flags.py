#!/usr/bin/env python3
"""Repeat selected radix profile flags in interleaved, status-checked rounds."""

from __future__ import annotations

import argparse
import csv
import json
import runpy
import statistics
from pathlib import Path


PERF = Path(__file__).resolve().parent
RUNNER = PERF / "radix-harness" / "run.py"
BINARY_NAME = "ods-formula-radix-evaluation-profile"
METRICS = (
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "max_rss_kib",
    "alloc_calls_p50",
    "requested_bytes_p50",
    "peak_live_delta_p50",
)


def load_flags(path: Path) -> list[dict[str, object]]:
    if not path.is_file():
        raise SystemExit(
            f"missing {path}; create the bounded post-capture flag list before repeating it"
        )
    flags = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(flags, list) or not flags:
        raise SystemExit(f"{path} must contain a non-empty JSON array")
    normalized: list[dict[str, object]] = []
    seen: set[tuple[str, str, str, int]] = set()
    for index, flag in enumerate(flags):
        if not isinstance(flag, dict):
            raise SystemExit(f"flag {index} is not an object")
        group = str(flag.get("group", "comparable"))
        phase = str(flag.get("phase", ""))
        case = str(flag.get("case", ""))
        try:
            repeat = int(flag["repeat"])
        except (KeyError, TypeError, ValueError) as error:
            raise SystemExit(f"flag {index} has no integer repeat") from error
        if group not in {"comparable", "radix"}:
            raise SystemExit(f"flag {index} has unsupported group {group!r}")
        if phase not in {"parse", "evaluate", "parse-evaluate"}:
            raise SystemExit(f"flag {index} has unsupported phase {phase!r}")
        if not case or repeat <= 0:
            raise SystemExit(f"flag {index} has an empty case or non-positive repeat")
        key = (group, phase, case, repeat)
        if key in seen:
            raise SystemExit(f"duplicate flag {key!r}")
        seen.add(key)
        normalized.append({"group": group, "phase": phase, "case": case, "repeat": repeat})
    return normalized


def revisions_for(flag: dict[str, object]) -> tuple[str, ...]:
    # Radix functions are candidate-only; comparable controls have both sides.
    return ("candidate",) if flag["group"] == "radix" else ("baseline", "candidate")


def check_row_status(
    row: dict[str, str], sequence: dict[str, str], folder: Path
) -> None:
    if row.get("status") != "0" or sequence.get("status") != "0":
        raise RuntimeError(
            f"subprocess failed for {sequence.get('revision')}/{sequence.get('group')}/"
            f"{sequence.get('phase')}/{sequence.get('case')}: row={row.get('status')!r} "
            f"sequence={sequence.get('status')!r}"
        )
    status_name = sequence.get("status")
    if not status_name:
        raise RuntimeError(f"missing status sidecar for {sequence}")
    status_path = folder / f"{sequence['revision']}-{sequence['phase']}-{sequence['case']}.status"
    if not status_path.is_file() or status_path.read_text(encoding="utf-8").strip() != "0":
        raise RuntimeError(f"nonzero or missing status sidecar: {status_path}")


def check_written_rows(rows: list[dict[str, str]], sequence: list[dict[str, str]]) -> None:
    if len(rows) != len(sequence) or not rows:
        raise RuntimeError(f"row/sequence count mismatch: rows={len(rows)} sequence={len(sequence)}")
    if any(row.get("status") != "0" for row in rows):
        raise RuntimeError("raw.csv contains a nonzero subprocess status")
    if any(item.get("status") != "0" for item in sequence):
        raise RuntimeError("sequence.json contains a nonzero subprocess status")


def metric_summary(
    rows: list[dict[str, str]],
    flag: dict[str, object],
) -> dict[str, object]:
    selected = [
        row
        for row in rows
        if row["group"] == flag["group"]
        and row["phase"] == flag["phase"]
        and row["case"] == flag["case"]
        and int(row["repeat"]) == flag["repeat"]
    ]
    summary: dict[str, object] = {
        "group": flag["group"],
        "phase": flag["phase"],
        "case": flag["case"],
        "repeat": flag["repeat"],
    }
    for metric in METRICS:
        values: dict[str, float] = {}
        for revision in revisions_for(flag):
            revision_rows = [row for row in selected if row["revision"] == revision]
            if not revision_rows:
                raise RuntimeError(f"missing {revision} row for {flag}")
            values[revision] = statistics.median(float(row[metric]) for row in revision_rows)
        item: dict[str, float | None] = {**values}
        if "baseline" in values and "candidate" in values:
            base = values["baseline"]
            item["delta_pct"] = (values["candidate"] / base - 1) * 100 if base else None
        else:
            item["delta_pct"] = None
        summary[metric] = item
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--flags",
        type=Path,
        default=PERF / "initial-flags.json",
        help="JSON array of bounded rows selected for repeat (default: initial-flags.json)",
    )
    parser.add_argument("--rounds", type=int, default=4)
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    if args.rounds <= 0 or args.warmups < 0 or args.iterations <= 0:
        raise SystemExit("rounds and iterations must be positive; warmups must be nonnegative")

    flags = load_flags(args.flags.resolve())
    runner = runpy.run_path(str(RUNNER))
    output = PERF / "abab"
    output.mkdir(exist_ok=True)
    combined: list[dict[str, str]] = []
    for round_number in range(1, args.rounds + 1):
        folder = output / f"r{round_number}"
        folder.mkdir(exist_ok=True)
        (folder / "commands.txt").write_text("", encoding="utf-8")
        rows: list[dict[str, str]] = []
        sequence: list[dict[str, str]] = []
        for flag in flags:
            for revision in revisions_for(flag):
                binary = PERF / revision / BINARY_NAME
                if not binary.is_file():
                    raise SystemExit(f"missing profile binary: {binary}")
                row, item = runner["run_one"](
                    binary,
                    folder,
                    revision,
                    str(flag["group"]),
                    str(flag["phase"]),
                    str(flag["case"]),
                    int(flag["repeat"]),
                    args.warmups,
                    args.iterations,
                )
                check_row_status(row, item, folder)
                rows.append(row)
                sequence.append(item)
                combined.append({"round": str(round_number), **row})
        with (folder / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=runner["FIELDS"], lineterminator="\n")
            writer.writeheader()
            writer.writerows(rows)
        (folder / "sequence.json").write_text(
            json.dumps(sequence, indent=2) + "\n", encoding="utf-8"
        )
        check_written_rows(
            list(csv.DictReader((folder / "raw.csv").open(encoding="utf-8"))),
            json.loads((folder / "sequence.json").read_text(encoding="utf-8")),
        )
        print(f"round {round_number} {len(rows)} rows", flush=True)

    with (output / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(
            stream, fieldnames=["round", *runner["FIELDS"]], lineterminator="\n"
        )
        writer.writeheader()
        writer.writerows(combined)
    check_written_rows(
        list(csv.DictReader((output / "raw.csv").open(encoding="utf-8"))),
        [{"status": row["status"]} for row in combined],
    )
    (output / "summary.json").write_text(
        json.dumps([metric_summary(combined, flag) for flag in flags], indent=2) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
