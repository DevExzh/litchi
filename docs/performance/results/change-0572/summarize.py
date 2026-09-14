#!/usr/bin/env python3
"""Summarise the classified capture into the tables the change record cites.

Reads the output of ``classify_zip_requests.py`` and reports, without
interpretation:

  * the request count per (scenario, fixture, policy), with the cross-transport
    and cross-route checks the plan's gates require;
  * per-ZIP-region request and byte totals for the ``exact`` arms;
  * the contiguous local-header/descriptor sweep and where it sits relative to
    the first payload read;
  * delayed-transport wall-clock medians.

Only the Python standard library is used.
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from pathlib import Path

ZERO = "zero_delay"
CAP = "zero_delay_64KiB_cap"
DELAYED = "delayed_1ms_100MiBps_64KiB"
POLICIES = ("exact", "forward_start(4096)", "forward_start(65536)")


def key(r):
    return (r["scenario"], r["fixture"], r["policy"], r["route"], r["transport"])


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--attribution", required=True)
    ap.add_argument("--capture", required=True)
    ap.add_argument("--json-out")
    args = ap.parse_args(argv)
    doc = json.loads(Path(args.attribution).read_text())
    capture = json.loads(Path(args.capture).read_text())
    opens = {}
    for run in capture["runs"]:
        opens[(run["scenario"], run["fixture"], run["policy"], run["route"], run["transport"])] = [
            rep["open_requests"] for rep in run["repeats"]
        ]
    by = {key(r): r for r in doc["results"]}
    layouts = doc["layouts"]

    order = []
    for r in doc["results"]:
        k = (r["scenario"], r["fixture"])
        if k not in order:
            order.append(k)

    print("== request counts (route = package_then_adopt except exact/native) ==")
    print(
        f"{'scenario':<28}{'fixture':<30}{'members':>8}{'desc':>6}"
        f"{'exact/native':>14}{'exact/adopt':>13}{'fs(4096)':>10}{'fs(64K)':>9}"
    )
    rows = []
    for scenario, fixture in order:
        lay = layouts[fixture]
        def n(policy, route):
            r = by.get((scenario, fixture, policy, route, ZERO))
            return r["summary"]["request_count"] if r else None
        row = {
            "scenario": scenario,
            "fixture": fixture,
            "members": lay["member_count"],
            "descriptor_members": lay["descriptor_members"],
            "size": lay["size"],
            "exact": n("exact", "native_leaf"),
            "exact_native": n("exact", "native_leaf"),
            "exact_adopt": n("exact", "package_then_adopt"),
            "fs4096": n("forward_start(4096)", "package_then_adopt"),
            "fs65536": n("forward_start(65536)", "package_then_adopt"),
        }
        for policy, name in (("exact", "exact"), ("forward_start(4096)", "fs4096"),
                             ("forward_start(65536)", "fs65536")):
            route = "native_leaf" if policy == "exact" else "package_then_adopt"
            r = by.get((scenario, fixture, policy, route, ZERO))
            row[name + "_bytes"] = r["summary"]["requested_bytes"] if r else None
            o = opens.get((scenario, fixture, policy, route, ZERO))
            row[name + "_open_requests"] = o[0] if o else None
            row[name + "_open_consistent"] = (len(set(o)) == 1) if o else None
        rows.append(row)
        print(
            f"{scenario:<28}{fixture:<30}{row['members']:>8}{row['descriptor_members']:>6}"
            f"{row['exact_native']:>14}{row['exact_adopt']:>13}{row['fs4096']:>10}{row['fs65536']:>9}"
        )

    print()
    print("== requested bytes and open/read phase split (zero delay) ==")
    print(f"{'fixture':<30}{'size':>8}"
          f"{'exact':>7}{'open':>6}{'read':>6}{'bytes':>9}"
          f"{'fs4k':>6}{'open':>6}{'bytes':>9}"
          f"{'fs64k':>7}{'open':>6}{'bytes':>9}{'amp':>7}")
    for row in rows:
        amp = row["fs65536_bytes"] / row["size"] if row["size"] else 0
        print(f"{row['fixture']:<30}{row['size']:>8}"
              f"{row['exact']:>7}{row['exact_open_requests']:>6}"
              f"{row['exact'] - row['exact_open_requests']:>6}{row['exact_bytes']:>9}"
              f"{row['fs4096']:>6}{row['fs4096_open_requests']:>6}{row['fs4096_bytes']:>9}"
              f"{row['fs65536']:>7}{row['fs65536_open_requests']:>6}{row['fs65536_bytes']:>9}"
              f"{amp:>6.2f}x")

    print()
    print("== gate: control (native leaf exact == two-step adopt exact) ==")
    control = []
    for scenario, fixture in order:
        a = by[(scenario, fixture, "exact", "native_leaf", ZERO)]
        b = by[(scenario, fixture, "exact", "package_then_adopt", ZERO)]
        same = a["sequence_digest"] == b["sequence_digest"]
        control.append({"scenario": scenario, "fixture": fixture, "identical": same,
                        "native": a["sequence_digest"], "adopt": b["sequence_digest"]})
        print(f"  {'PASS' if same else 'FAIL':<5} {scenario:<28}{fixture:<30}"
              f"{a['sequence_digest']} vs {b['sequence_digest']}")

    print()
    print("== gate: determinism across transports (same policy/route) ==")
    det = []
    for scenario, fixture in order:
        for policy in POLICIES:
            for route in ("native_leaf", "package_then_adopt"):
                arms = [by.get((scenario, fixture, policy, route, t))
                        for t in (ZERO, CAP, DELAYED)]
                if any(a is None for a in arms):
                    continue
                digests = [a["sequence_digest"] for a in arms]
                repeats_ok = all(a["identical_across_repeats"] for a in arms)
                same = len(set(digests)) == 1
                det.append({"scenario": scenario, "fixture": fixture, "policy": policy,
                            "route": route, "identical_across_transports": same,
                            "identical_across_repeats": repeats_ok, "digests": digests})
                if not (same and repeats_ok):
                    print(f"  DIFFER {scenario} {fixture} {policy} {route} {digests} "
                          f"repeats_ok={repeats_ok}")
    bad = [d for d in det if not (d["identical_across_transports"] and d["identical_across_repeats"])]
    print(f"  arms checked: {len(det)}   divergent: {len(bad)}")

    print()
    print("== gate: attribution (exact policy, native leaf, zero delay) ==")
    print(f"{'fixture':<30}{'reqs':>6}{'eocd':>6}{'cd':>5}{'probe':>7}{'lhdr':>6}{'lvar':>6}"
          f"{'payl':>6}{'desc':>6}{'GAP':>5}{'pastEOF':>8}{'bytes':>10}")
    attribution = []
    for scenario, fixture in order:
        r = by[(scenario, fixture, "exact", "native_leaf", ZERO)]
        s = r["summary"]
        p = s["requests_by_role"]
        attribution.append({"scenario": scenario, "fixture": fixture, **s})
        print(f"{fixture:<30}{s['request_count']:>6}{p['eocd']:>6}{p['central_directory']:>5}"
              f"{p['local_record_probe']:>7}{p['local_header']:>6}{p['local_header_var']:>6}"
              f"{p['payload']:>6}{p['data_descriptor']:>6}{s['requests_in_gap']:>5}"
              f"{s['bytes_beyond_eof']:>8}{s['requested_bytes']:>10}")

    print()
    print("== strict-layout proof: longest run of consecutive local-record probes ==")
    print(f"{'fixture':<30}{'mem':>5}{'sweep':>7}{'probes':>8}{'desc':>6}{'frame':>7}"
          f"{'distinct':>9}{'asc':>6}{'idx':>11}{'bytes':>8}{'ofTotal':>9}  next request")
    sweeps = []
    for scenario, fixture in order:
        r = by[(scenario, fixture, "exact", "native_leaf", ZERO)]
        sw = r["longest_header_sweep"]
        total = r["summary"]["request_count"]
        longest = sw["longest"]
        entry = {"scenario": scenario, "fixture": fixture, "total_requests": total, **sw}
        sweeps.append(entry)
        if longest is None:
            print(f"{fixture:<30}{sw['member_count']:>5}{'none':>7}")
            continue
        nxt = sw["first_request_after_longest_run"]
        nxt_s = (f"#{nxt['index']} off={nxt['offset']} len={nxt['length']} "
                 f"{nxt['primary_region']} {','.join(str(m) for m in nxt['members'])}") if nxt else "-"
        print(f"{fixture:<30}{sw['member_count']:>5}{longest['request_count']:>7}"
              f"{longest['probe_count']:>8}{longest['descriptor_count']:>6}"
              f"{longest['framing_count'] - longest['descriptor_count']:>7}"
              f"{longest['distinct_members_probed']:>9}"
              f"{str(longest['layout_order_ascending']):>6}"
              f"{(str(longest['start_index']) + '..' + str(longest['end_index'])):>11}"
              f"{longest['requested_bytes']:>8}"
              f"{100.0 * longest['request_count'] / total:>8.1f}%  {nxt_s}")

    print()
    print("== delayed-transport wall clock, medians over 5 repeats (ms) ==")
    print(f"{'scenario':<28}{'fixture':<30}{'exact':>10}{'fs(4096)':>10}{'fs(64K)':>10}"
          f"{'exact/req':>11}")
    timing = []
    for scenario, fixture in order:
        def med(policy, route="package_then_adopt"):
            r = by.get((scenario, fixture, policy, route, DELAYED))
            return statistics.median(r["elapsed_ns"]) if r else None
        e = med("exact", "native_leaf")
        f4 = med("forward_start(4096)")
        f6 = med("forward_start(65536)")
        n = by[(scenario, fixture, "exact", "native_leaf", DELAYED)]["summary"]["request_count"]
        timing.append({"scenario": scenario, "fixture": fixture,
                       "exact_ns": e, "fs4096_ns": f4, "fs65536_ns": f6,
                       "exact_requests": n})
        print(f"{scenario:<28}{fixture:<30}{e/1e6:>10.2f}{f4/1e6:>10.2f}{f6/1e6:>10.2f}"
              f"{e/n/1e6:>10.3f}ms")

    if args.json_out:
        Path(args.json_out).write_text(json.dumps({
            "schema_version": 1,
            "record_kind": "litchi-perf-0572-summary",
            "request_counts": rows,
            "control_gate": control,
            "determinism_gate": det,
            "attribution_gate": attribution,
            "header_sweeps": sweeps,
            "timing_medians_delayed": timing,
        }, indent=1) + "\n")
        print(f"\nwrote {args.json_out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
