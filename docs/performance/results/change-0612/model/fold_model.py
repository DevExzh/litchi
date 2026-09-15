#!/usr/bin/env python3
"""Reprint every table the change 0612 record cites from `globals-skip.json`.

Pure standard library; reads only this directory.
"""

from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DOC = json.load(open(os.path.join(HERE, "globals-skip.json")))
ROWS = [r for r in DOC["files"] if "error" not in r and r.get("globals_end")]
BAD = [r for r in DOC["files"] if "error" in r]

# Change 0564 derived ~116 ns per positional read on this host; change 0574's
# corpus regression measured 53.4 ns per KiB of globals payload.
NS_PER_READ = 116.0
NS_PER_KIB = 53.4


def net_ns(row, label):
    return (
        row[f"{label}_delta_reads"] * NS_PER_READ
        + row[f"{label}_delta_bytes"] / 1024.0 * NS_PER_KIB
    )


def main():
    s = DOC["summary"]
    print("=" * 78)
    print("1. Corpus composition: what a source-backed open reads and what it uses")
    print("=" * 78)
    print(f"  fixtures scanned            {s['fixtures_scanned']:>12,}")
    print(f"  modelled                    {s['fixtures_modelled']:>12,}")
    print(f"  refused (encrypted/no BOF)  {s['fixtures_refused']:>12,}")
    print()
    print(f"  globals bytes framed        {s['total_globals_bytes']:>12,}   100.00%")
    print(
        f"  four-byte record headers    {s['total_header_bytes']:>12,}"
        f"   {100 * s['total_header_bytes'] / s['total_globals_bytes']:>6.2f}%"
    )
    print(
        f"  consumed record payloads    {s['total_consumed_payload_bytes']:>12,}"
        f"   {100 * s['total_consumed_payload_bytes'] / s['total_globals_bytes']:>6.2f}%"
    )
    print(
        f"  touched closure             {s['total_closure_bytes']:>12,}"
        f"   {100 * s['closure_share']:>6.2f}%"
    )
    print(
        f"  payloads never interpreted  {s['total_skipped_payload_bytes']:>12,}"
        f"   {100 * s['total_skipped_payload_bytes'] / s['total_globals_bytes']:>6.2f}%"
    )
    print()
    print(f"  fixtures consuming over half their globals   {s['fixtures_over_half_consumed']:>4} / {s['fixtures_modelled']}")
    print(f"  fixtures with a skipped payload >= 1 KiB     {s['fixtures_with_any_skipped_payload_over_1k']:>4} / {s['fixtures_modelled']}")
    print(f"  fixtures with a skipped payload >= 4 KiB     {s['fixtures_with_any_skipped_payload_over_4k']:>4} / {s['fixtures_modelled']}")
    print(f"  fixtures with a skipped payload >= 8 KiB     {s['fixtures_with_any_skipped_payload_over_8k']:>4} / {s['fixtures_modelled']}")
    print()

    print("=" * 78)
    print("2. Concentration: where the skippable bytes live")
    print("=" * 78)
    top = sorted(ROWS, key=lambda r: -r["skipped_payload_bytes"])
    total_skipped = s["total_skipped_payload_bytes"]
    running = 0
    n90 = None
    for i, r in enumerate(top, 1):
        running += r["skipped_payload_bytes"]
        if n90 is None and running >= 0.90 * total_skipped:
            n90 = i
    print(f"  90% of all skippable bytes live in {n90} of {len(ROWS)} fixtures")
    print()
    print("  %-52s %10s %9s %7s" % ("fixture", "skippable", "globals", "share"))
    for r in top[:12]:
        print(
            "  %-52s %10d %9d %6.1f%%"
            % (
                r["path"][-52:],
                r["skipped_payload_bytes"],
                r["globals_end"],
                100 * r["skipped_payload_bytes"] / r["globals_end"],
            )
        )
    print()

    print("=" * 78)
    print("3. Today's schedule, replayed from the source")
    print("=" * 78)
    print("  %-52s %6s %10s %7s" % ("fixture", "fills", "bytes", "records"))
    for r in sorted(ROWS, key=lambda r: -r["globals_end"])[:10]:
        print(
            "  %-52s %6d %10d %7d"
            % (r["path"][-52:], r["today_reads"], r["today_bytes"], r["globals_records"])
        )
    print(
        f"\n  corpus total: {s['total_today_reads']:,} fills, "
        f"{s['total_today_bytes']:,} bytes"
    )
    print()

    print("=" * 78)
    print("4. Candidate schedules: the read-for-bytes trade, corpus totals")
    print("=" * 78)
    print(
        "  %-9s %8s %7s %12s %9s %8s %8s %11s"
        % ("gate", "reads", "delta", "bytes", "saved%", "seeks", "fx.seek", "net ns")
    )
    for g in DOC["gates"]:
        lbl = g["label"]
        reads = s[f"{lbl}_total_reads"]
        by = s[f"{lbl}_total_bytes"]
        print(
            "  %-9s %8d %+7d %12d %8.1f%% %8d %8d %+11.0f"
            % (
                lbl,
                reads,
                reads - s["total_today_reads"],
                by,
                100 * (s["total_today_bytes"] - by) / s["total_today_bytes"],
                s[f"{lbl}_total_seeks"],
                s[f"{lbl}_fixtures_any_seek"],
                s[f"{lbl}_modelled_net_ns_file_source"],
            )
        )
    print()
    print("  gate definitions:")
    for g in DOC["gates"]:
        print(
            "    %-9s skip_min=%-6d dense_mean=%-20s reset_target_on_skip=%s"
            % (
                g["label"],
                g["skip_min_bytes"],
                "off" if g["dense_mean_bytes"] > 2**40 else g["dense_mean_bytes"],
                g["reset_target_on_skip"],
            )
        )
    print()
    print(
        "  net ns = delta_reads x %.0f ns + delta_bytes / 1024 x %.1f ns, a FILE source."
        % (NS_PER_READ, NS_PER_KIB)
    )
    print("  Negative is faster. An owned in-memory source pays no per-read syscall,")
    print("  so its trade is strictly better; a range source pays far more.")
    print()

    print("=" * 78)
    print("5. Every fixture the recommended gate (g1k) touches")
    print("=" * 78)
    print(
        "  %-46s %6s %6s %10s %10s %10s"
        % ("fixture", "r.now", "r.new", "b.now", "b.new", "net ns")
    )
    touched = [r for r in ROWS if r["g1k_seeks"]]
    for r in sorted(touched, key=lambda r: net_ns(r, "g1k")):
        print(
            "  %-46s %6d %6d %10d %10d %+10.0f"
            % (
                r["path"][-46:],
                r["today_reads"],
                r["g1k_reads"],
                r["today_bytes"],
                r["g1k_bytes"],
                net_ns(r, "g1k"),
            )
        )
    untouched = len(ROWS) - len(touched)
    print(
        f"\n  {len(touched)} fixtures take at least one seek; {untouched} are "
        "byte-for-byte and read-for-read unchanged."
    )
    worse = [r for r in ROWS if net_ns(r, "g1k") > 0]
    print(f"  {len(worse)} fixtures are modelled net worse on a file source:")
    for r in sorted(worse, key=lambda r: -net_ns(r, "g1k")):
        print(
            "    %-50s %+d reads, %+d bytes, %+.0f ns"
            % (
                r["path"][-50:],
                r["g1k_delta_reads"],
                r["g1k_delta_bytes"],
                net_ns(r, "g1k"),
            )
        )
    print()

    print("=" * 78)
    print("6. The three profiled fixtures")
    print("=" * 78)
    stems = {
        "test-data/ole/xls/ConditionalFormattingSamples.xls": "flagship",
        "test-data/ole/xls/WithCustomViews.xls": "WithCustomViews",
        "test-data/poi/test-data/spreadsheet/54016.xls": "54016",
    }
    for path, stem in stems.items():
        r = next((x for x in ROWS if x["path"] == path), None)
        if r is None:
            print(f"  {stem}: not modelled")
            continue
        print(f"  {stem} ({path})")
        print(
            f"    globals {r['globals_end']:,} B over {r['globals_records']:,} records; "
            f"closure {r['closure_bytes']:,} B ({100 * r['closure_share_of_globals']:.2f}%)"
        )
        print(
            f"    skippable {r['skipped_payload_bytes']:,} B; largest skipped payload "
            f"{r['max_skipped_payload']:,} B; {r['skipped_payloads_over_1k']} payloads >= 1 KiB "
            f"holding {r['skipped_bytes_in_payloads_over_1k']:,} B"
        )
        print(f"    top skipped kinds: {r['top_skipped']}")
        print(f"    continuation bytes by owning record: {r['continuation_bytes_by_owner']}")
        print(f"    today:  {r['today_reads']} fills, {r['today_bytes']:,} bytes")
        for g in DOC["gates"]:
            lbl = g["label"]
            print(
                "    %-8s %4d fills (%+d), %9d bytes (%+d), %d seeks, net %+0.0f ns"
                % (
                    lbl,
                    r[f"{lbl}_reads"],
                    r[f"{lbl}_delta_reads"],
                    r[f"{lbl}_bytes"],
                    r[f"{lbl}_delta_bytes"],
                    r[f"{lbl}_seeks"],
                    net_ns(r, lbl),
                )
            )
        print()

    if BAD:
        print("=" * 78)
        print("7. Fixtures not modelled")
        print("=" * 78)
        for r in BAD:
            print("  %-56s %s" % (r["path"][-56:], r["error"]))
        print()


if __name__ == "__main__":
    sys.exit(main())
