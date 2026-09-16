#!/usr/bin/env python3
"""Summarize change 0638's first descriptive baseline.

Reads the schema-1 reports the two repeats produced and prints:

* one row per facade case with its p50/mean/p95/p99, the corpus identity and
  the frozen observation, plus the R1/R2 paired p50 spread, which is the A/A
  floor for this window because the two repeats are the same binary on the same
  bytes;
* one row per ordinary-save case with the same statistics and, for the counting
  phase, the byte split;
* the cross-phase attribution the record reads: the lifecycle median against
  the sum of its edit and atomic-publish medians.

It asserts, rather than assumes, the gates every case carries: every retained
sample reproduced its frozen observation, every publication reproduced the
corpus reference, and every corpus proved the 0625/0631 repeated-save
invariant.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def load(directory: Path, label: str) -> dict[str, dict]:
    rows: dict[str, dict] = {}
    for path in sorted(directory.glob(f"{label}-*.json")):
        report = json.loads(path.read_text(encoding="utf-8"))
        for result in report["results"]:
            key = f"{result['case']}@{result['corpus']['archive_sha256'][:12]}"
            if key in rows:
                raise SystemExit(f"duplicate case identity {key} in {path}")
            rows[key] = result
    return rows


def statistics(result: dict) -> tuple[int, float, int, int]:
    elapsed = result["elapsed_ns"]
    return elapsed["p50"], elapsed["mean"], elapsed["p95"], elapsed["p99"]


def percent(after: float, before: float) -> float:
    return 0.0 if before == 0 else (after - before) / before * 100.0


def main() -> int:
    if len(sys.argv) != 2:
        raise SystemExit("usage: summarize.py <raw-directory>")
    raw = Path(sys.argv[1])
    labels = sorted({path.name.split("-", 1)[0] for path in raw.glob("R*-*.json")})
    if len(labels) < 2:
        raise SystemExit("an A/A floor needs at least two repeats")
    repeats = {label: load(raw, label) for label in labels}
    first = repeats[labels[0]]
    for label in labels[1:]:
        if set(repeats[label]) != set(first):
            raise SystemExit(f"repeat {label} measured different case identities")

    def spread(key: str) -> float:
        """Widest p50 disagreement between repeats, as a percentage of the
        smallest. The repeats are the same binary on the same bytes, so this is
        the A/A floor for that row in this window."""
        values = [statistics(repeats[label][key])[0] for label in labels]
        return percent(max(values), min(values))

    over_floor: list[tuple[str, float]] = []

    facade_floor: list[float] = []
    save_floor: list[float] = []

    print("== facade selectors")
    print(
        f"{'case':38} {'fixture':26} {'bytes':>9} {'p50 ns':>10} "
        f"{'p95 ns':>10} {'p50 spread %':>13}  observation"
    )
    for key in sorted(first):
        result = first[key]
        source = result.get("source") or {}
        facade = source.get("facade_ole2")
        if facade is None:
            continue
        p50, _mean, p95, _p99 = statistics(result)
        delta = spread(key)
        facade_floor.append(delta)
        if delta > 5.0:
            over_floor.append((result["case"], delta))
        if not facade["observations_identical"]:
            raise SystemExit(f"{key} did not reproduce its frozen observation")
        fixture = Path(facade["corpus"]["real_file"]["path"]).name
        print(
            f"{result['case']:38} {fixture:26} "
            f"{facade['corpus']['real_file']['bytes']:9d} {p50:10d} {p95:10d} "
            f"{delta:13.2f}  {facade['observation'][:56]}"
        )

    print()
    print("== ordinary documented save")
    print(
        f"{'case':46} {'origin':10} {'p50 ns':>12} {'p95 ns':>12} "
        f"{'p50 spread %':>13}  edit"
    )
    medians: dict[tuple[str, str, str], int] = {}
    for key in sorted(first):
        result = first[key]
        source = result.get("source") or {}
        save = source.get("ordinary_save")
        if save is None:
            continue
        p50, _mean, p95, _p99 = statistics(result)
        delta = spread(key)
        save_floor.append(delta)
        if delta > 5.0:
            over_floor.append((result["case"], delta))
        if not save["publications_identical"]:
            raise SystemExit(f"{key} published an artifact that differs")
        if not save["edit_outcomes_identical"]:
            raise SystemExit(f"{key} produced an unstable edit outcome")
        corpus = save["corpus"]
        if not corpus["repeated_cycles_identical"]:
            raise SystemExit(f"{key} failed the repeated-cycle invariant")
        if not corpus["repeated_saves_identical"]:
            raise SystemExit(f"{key} failed the repeated-save invariant")
        medians[(save["format"], save["origin"], save["phase"])] = p50
        print(
            f"{result['case']:46} {save['origin'][:10]:10} {p50:12d} {p95:12d} "
            f"{delta:13.2f}  {corpus['edit_outcome'][:44]}"
        )

    print()
    print("== byte split (counting phase, one sample per corpus)")
    header = (
        f"{'format/origin':34} {'source':>9} {'output':>9} {'deflate':>9} "
        f"{'stored':>9} {'identical':>10} {'regen':>8} {'regen in':>9} "
        f"{'framing':>8} {'members':>8}"
    )
    print(header)
    for key in sorted(first):
        result = first[key]
        source = result.get("source") or {}
        save = source.get("ordinary_save")
        if save is None or save["phase"] != "serialize-to-counting-sink":
            continue
        corpus = save["corpus"]
        split = save["sample_byte_split"] or corpus["byte_split"]
        name = f"{save['format']}/{save['origin']}"
        print(
            f"{name:34} {corpus['source_archive_bytes']:9d} "
            f"{split['output_total_bytes']:9d} {split['payload_bytes_deflated']:9d} "
            f"{split['payload_bytes_stored']:9d} "
            f"{split['payload_bytes_identical_to_source']:10d} "
            f"{split['payload_bytes_regenerated']:8d} "
            f"{split['uncompressed_payload_bytes_regenerated']:9d} "
            f"{split['framing_bytes']:8d} "
            f"{split['members_identical_to_source']}/{split['output_member_count']:<6}"
        )

    print()
    print("== phase attribution (p50 ns)")
    print(
        f"{'format/origin':34} {'lifecycle':>12} {'edit':>12} {'atomic':>12} "
        f"{'counting':>12} {'edit+atomic/lifecycle':>22}"
    )
    for fmt in ("DOCX", "XLSX", "PPTX"):
        for origin in ("generated-harness-corpus", "caller-named-real-file"):
            life = medians.get((fmt, origin, "open+edit+save"))
            edit = medians.get((fmt, origin, "edit"))
            atomic = medians.get((fmt, origin, "save-to-path"))
            counting = medians.get((fmt, origin, "serialize-to-counting-sink"))
            if life is None:
                continue
            share = (edit + atomic) / life * 100.0
            print(
                f"{fmt + '/' + origin:34} {life:12d} {edit:12d} {atomic:12d} "
                f"{counting:12d} {share:21.1f}%"
            )

    def floor(values: list[float]) -> str:
        if not values:
            return "n/a"
        ordered = sorted(values)
        mid = ordered[len(ordered) // 2]
        return f"p50 {mid:.2f}%, max {ordered[-1]:.2f}% over {len(ordered)} rows"

    print()
    print(f"repeats compared: {', '.join(labels)}")
    print(f"A/A floor, facade rows:        {floor(facade_floor)}")
    print(f"A/A floor, ordinary-save rows: {floor(save_floor)}")
    print(f"A/A floor, all rows:           {floor(facade_floor + save_floor)}")
    if over_floor:
        print()
        print("rows whose A/A spread exceeds 5% (a review trigger, reported not hidden):")
        for case, delta in sorted(over_floor, key=lambda row: -row[1]):
            print(f"  {case:46} {delta:8.2f}%")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
