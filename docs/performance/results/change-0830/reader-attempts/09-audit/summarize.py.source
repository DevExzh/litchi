"""Compact, reproducible projections of the complete compressed analysis."""
import gzip
import json

import driver as d


def main():
    source = d.P / "analysis.json.gz"
    a = json.loads(gzip.decompress(source.read_bytes()))
    repeats = []
    for repeat in a["heaptrack"]["repeats"]:
        parser = repeat["parser"]
        rows = parser["full_leaf"]["rows"]
        owner_rows = [r for r in rows if r["owner_calls"]]
        selected = [r for r in owner_rows if any(
            (f["function"] or "").startswith("litchi_xlsx::column::Assignments$LT$T$GT$::new::")
            for f in r["frames_leaf_to_root"])]
        assert len(selected) == 3
        selected_calls = sum(r["owner_calls"] for r in selected)
        selected_bytes = sum(r["owner_requested_bytes"] for r in selected)
        total = repeat["metrics"]["owner"]
        repeats.append({
            "repeat": repeat["repeat"], "whole": repeat["metrics"]["whole"], "owner": total,
            "owner_average_per_edit": {k: v / 5 for k, v in total.items()},
            "column_assignments": {"calls": selected_calls, "requested_bytes": selected_bytes,
                "fraction_of_owner_requested_bytes": selected_bytes / total["requested_bytes"],
                "traces": [{"trace_id": r["trace_id"], "calls": r["owner_calls"],
                            "requested_bytes": r["owner_requested_bytes"],
                            "project_frames": [f for f in r["frames_leaf_to_root"]
                                               if (f["function"] or "").startswith("litchi_xlsx::")]}
                           for r in selected]},
            "owner_calls_with_unknown_function": sum(r["owner_calls"] for r in owner_rows
                if any(f["function"] is None for f in r["frames_leaf_to_root"])),
            "owner_max_expanded_frame_depth": max(len(r["frames_leaf_to_root"]) for r in owner_rows),
            "diagnostics": parser["diagnostics"],
        })
    result = {"schema": "litchi.performance.0830.summary.v1", "analysis_sha256": d.sha(source),
              "reports": 23, "samples": 559, "heaptrack": repeats,
              "native_medians": {arm: {key: row["midpoint_median"] for key, row in value["metrics"].items()}
                                 for arm, value in a["native"]["arms"].items()},
              "native_paired_ratios": {key: value["bootstrap"] for key, value in a["native"]["paired_ratios"].items()},
              "native_spread_flags": a["native"]["spread_flags"], "native_tail_flags": a["native"]["tail_flags"]}
    d.write(d.P / "summary.json", result)
    print(json.dumps({"status": "pass", "column_assignment_bytes_per_capture": [r["column_assignments"]["requested_bytes"] for r in repeats]}))


if __name__ == "__main__":
    main()
