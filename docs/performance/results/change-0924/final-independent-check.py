#!/usr/bin/env python3
"""Final read-only cross-check of frozen 0924 attribution and phase outputs."""
from __future__ import annotations

import collections
import hashlib
import json
import pathlib
import subprocess

BASE = pathlib.Path("/tmp/litchi-performance-0924")


def digest(path: pathlib.Path) -> dict[str, object]:
    digest_value = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest_value.update(block)
            size += len(block)
    return {"path": str(path), "bytes": size, "sha256": digest_value.hexdigest()}


def main() -> int:
    attribution_path = BASE / "stack-attribution-v2.json"
    rows_path = BASE / "stack-rows-v2.json"
    phase_path = BASE / "phase-share-v2.json"
    attribution = json.loads(attribution_path.read_text(encoding="utf-8"))
    rows = json.loads(rows_path.read_text(encoding="utf-8"))
    assert attribution["schema"] == "litchi.performance.0924.heaptrack-stack-attribution.v2"
    assert attribution["profile_count"] == 25 == len(rows)
    assert len({profile["id"] for profile in attribution["profiles"]}) == 25
    assert len({item["id"] for item in rows}) == 25
    assert attribution["stack_rows_external"]["descriptor"] == digest(rows_path)

    counts = collections.Counter(
        (profile["route"], profile["variant"]) for profile in attribution["profiles"]
    )
    assert counts == collections.Counter(
        {
            ("generic_read_at", "baseline"): 5,
            ("generic_read_at", "candidate"): 4,
            ("owned_vec", "baseline"): 4,
            ("owned_vec", "candidate"): 4,
            ("owned_reader", "baseline"): 4,
            ("owned_reader", "candidate"): 4,
        }
    )

    # Recheck the final external stack rows independently of their group totals.
    groups = collections.defaultdict(
        lambda: {"bytes": 0, "calls": 0, "read_bytes": 0, "read_calls": 0, "profiles": 0}
    )
    for item in rows:
        key = item["route"], item["variant"]
        group = groups[key]
        group["profiles"] += 1
        for row in item["rows"]:
            group["bytes"] += row["bytes"]
            group["calls"] += row["calls"]
            if row.get("contains_read_limited"):
                group["read_bytes"] += row["bytes"]
                group["read_calls"] += row["calls"]
    for (route, variant), group in groups.items():
        expected = attribution["groups"][f"{route}:{variant}"]
        assert (
            group["profiles"],
            group["bytes"],
            group["calls"],
            group["read_bytes"],
            group["read_calls"],
        ) == (
            expected["profiles"],
            expected["requested_bytes"],
            expected["allocation_event_count"],
            expected["read_limited_bytes"],
            expected["read_limited_calls"],
        )

    # Run the independent raw parser, which is bound to final v2, and compare
    # its pair counters with the final attribution summary.
    replay = json.loads(
        subprocess.check_output(
            ["python3", "-B", str(BASE / "attribution-independent-replay.py")],
            text=True,
        )
    )
    assert replay["summary_schema"] == attribution["schema"]
    assert replay["summary"] == digest(attribution_path)
    assert replay["raw_descriptors_verified_before_parse"] is True
    for variant in ("baseline", "candidate"):
        pair = attribution["owned_reader_vs_owned_vec"][variant]
        raw_pair = replay["pair_checks"][variant]
        for field in (
            "total_delta_bytes",
            "primary_read_limited_bytes",
            "primary_read_limited_calls",
            "common_stack_delta_bytes",
            "common_stack_delta_calls",
            "reader_requested_bytes",
            "vec_requested_bytes",
        ):
            assert pair[field] == raw_pair[field], (variant, field)
        assert raw_pair["accounting_check"]["pass"] is True
        assert pair["accounting_check"]["pass"] is True

    # Binary and parser-custody witnesses remain bound by the final summary.
    for variant in ("baseline", "candidate"):
        proof = attribution["binary_proof"][variant]
        assert proof["present"] is True
        assert proof["actual"] == digest(pathlib.Path(proof["actual"]["path"]))
    assert attribution["parser_controls"]["status"] == "pass"
    assert attribution["parser_controls"]["control_count"] == 10

    phase = json.loads(phase_path.read_text(encoding="utf-8"))
    assert phase["schema"] == "litchi.performance.0924.raw-phase-share-audit.v2"
    for key in (
        "all_candidate_report_hashes_match_root_raw_sample_summary",
        "all_candidate_report_sizes_match_root_raw_sample_summary",
        "all_candidate_reports_have_exact_binary_identity",
        "all_candidate_reports_have_exact_case_and_route",
        "all_phase_vectors_have_200_samples",
        "all_sample_orders_are_permutations",
        "all_phase_sums_equal_elapsed_samples",
    ):
        assert phase["validation"][key] is True, key
    assert phase["all_routes_sum_of_raw_samples"]["sample_count"] == 4800
    expected_open = {
        "generic_read_at": 0.010086770954901587,
        "owned_vec": 0.012168307399429819,
        "owned_reader": 0.09965195084465603,
    }
    for route, expected in expected_open.items():
        assert phase["routes"][route]["phase_shares"]["open_ns"] == expected
        route_data = phase["routes"][route]
        assert route_data["source_identities"] == [
            {
                "bytes": 2101427,
                "sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
            }
        ]
        assert route_data["output_identities"] == [
            {
                "source_bytes": 2101427,
                "sha256": "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
            }
        ]

    output = {
        "status": "verified",
        "attribution": digest(attribution_path),
        "stack_rows": digest(rows_path),
        "phase": digest(phase_path),
        "raw_replay": replay,
        "checks": {
            "profiles": 25,
            "raw_descriptors_bound_before_parse": True,
            "pair_accounting": True,
            "binary_proof": True,
            "parser_controls": 10,
            "phase_reports": 24,
            "phase_samples": 4800,
            "phase_corruption_binding": True,
        },
    }
    print(json.dumps(output, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
