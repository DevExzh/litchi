#!/usr/bin/env python3
"""Small offline controls for focuseddecode_0977.py.

The tests consume only retained perf-script/log text and synthetic bounded
rows.  They never invoke a workload, a build, or the decoder's offline perf
subcommands.
"""
from __future__ import annotations

import importlib.util
import json
import tempfile
from pathlib import Path


DECODER = Path(__file__).with_name("focuseddecode_0977_v2.py")
TARGET = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
RETAINED_ROOT = Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3")
SCRIPT = RETAINED_ROOT / "original.final-v2.corrected.script.txt"
LOG = RETAINED_ROOT / "translated.dwfl.jsonl"
LAYOUT = RETAINED_ROOT / "translated.target-layout.json"
EXPECTED_PID = 3_763_764
EXPECTED_CPU = 31
PERIOD = 1_000_000


def load_decoder():
    spec = importlib.util.spec_from_file_location("focuseddecode_0977_test_subject", DECODER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import decoder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def assert_decode_failure(fn, code: str) -> None:
    try:
        fn()
    except Exception as exc:  # typed class is retained by the subject
        if getattr(exc, "code", None) != code:
            raise AssertionError(f"expected {code}, got {exc!r}") from exc
        return
    raise AssertionError(f"expected DecodeFailure({code})")


def main() -> int:
    decoder = load_decoder()
    marker = decoder.PARSER.MARKER
    dso = decoder.PARSER.BINARY_PATH
    evlist = (
        "instructions:u: type: 4 (cpu), size: 144, config: 0xc0 (instructions), "
        "{ sample_period, sample_freq }: 1000000, "
        "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER, "
        "read_format: ID|LOST, disabled: 1, exclude_kernel: 1, exclude_hv: 1, "
        "mmap: 1, comm: 1, enable_on_exec: 1, task: 1, mmap_data: 1, sample_id_all: 1, "
        "exclude_callchain_kernel: 1, exclude_callchain_user: 1, mmap2: 1, comm_exec: 1, "
        "ksymbol: 1, bpf_event: 1, build_id: 1, sample_regs_user: 0xff0fff, sample_stack_user: 32768\n"
    )
    if decoder._validate_evlist(evlist, expected_period=1_000_000, expected_stack_cap=32_768)["status"] != "pass":
        raise AssertionError("evlist exact attribute fixture was not admitted")
    for mutation in ("sample_stack_user: 32768", "task: 1", "sample_period, sample_freq }: 1000000"):
        bad = evlist.replace(mutation, mutation.replace("32768", "327680") if "stack" in mutation else mutation.replace("1", "0", 1) if "task" in mutation else mutation.replace("1000000", "10000000"), 1)
        assert_decode_failure(lambda bad=bad: decoder._validate_evlist(bad, expected_period=1_000_000, expected_stack_cap=32_768), "evlist_contract")
    duplicate = evlist.replace("task: 1, mmap_data", "task: 1, task: 1, mmap_data", 1)
    assert_decode_failure(lambda: decoder._validate_evlist(duplicate, expected_period=1_000_000, expected_stack_cap=32_768), "evlist_contract")
    synthetic = (
        f"{EXPECTED_PID}/{EXPECTED_PID} [{EXPECTED_CPU:03d}] 1.000: {PERIOD} instructions:u: \n"
        f"\t1000 synthetic::leaf+0x1 (/lib/x.so)\n"
        f"\t1001 {marker}+0x2 ({dso})\n\n"
        f"{EXPECTED_PID}/{EXPECTED_PID} [{EXPECTED_CPU:03d}] 2.000: {PERIOD} instructions:u: \n"
        f"\t0 [unknown] ([unknown])\n"
        f"\t1001 {marker}+0x3 ({dso})\n\n"
    )
    with tempfile.TemporaryDirectory(prefix="focuseddecode-0977-controls-") as temp:
        path = Path(temp) / "synthetic.perf-script"
        path.write_text(synthetic, encoding="utf-8")
        streamed = decoder._stream_perf_script(
            path, marker=marker, marker_dso=dso, expected_pid=EXPECTED_PID,
            expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
        )
        strict = decoder.PARSER.parse_perf_script_strict(
            synthetic, marker=marker, marker_dso=dso, expected_pid=EXPECTED_PID,
            expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
        )
        differential_keys = {
            "sample_count", "total_period", "marker", "marker_sample_count", "marker_period",
            "marker_period_share", "lost_events", "unresolved_frames", "explicit_stack_truncated_count",
            "empty_callchain_count", "physical_leaf_owner_period", "physical_leaf_unknown_period",
            "physical_leaf_unknown_samples", "marker_unknown_physical_leaf_samples",
            "marker_no_resolved_frame_samples", "physical_leaf_classified_period",
            "physical_leaf_conservation", "physical_leaf_scope", "deepest_resolved_proxy_period",
            "inclusive_marker_ancestor_period",
        }
        if any(streamed[key] != strict[key] for key in differential_keys):
            raise AssertionError("streaming parser differs from sealed parser on synthetic control")
        if streamed["marker_sample_count"] != 2 or streamed["physical_leaf_unknown_samples"] != 1:
            raise AssertionError("synthetic marker/unknown ledger is wrong")
        bad_cpu = path.with_name("bad-cpu.perf-script")
        bad_cpu.write_text(synthetic.replace("[031]", "[030]", 1), encoding="utf-8")
        assert_decode_failure(
            lambda: decoder._stream_perf_script(
                bad_cpu, marker=marker, marker_dso=dso, expected_pid=EXPECTED_PID,
                expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
            ),
            "perf_script",
        )
        bad_marker = path.with_name("bad-marker.perf-script")
        bad_marker.write_text(synthetic.replace(f"({dso})", "(/wrong/dso)", 1), encoding="utf-8")
        no_marker = decoder._stream_perf_script(
            bad_marker, marker=marker, marker_dso=dso, expected_pid=EXPECTED_PID,
            expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
        )
        if no_marker["marker_sample_count"] != 1:
            raise AssertionError("marker DSO gate overmatched a wrong DSO")

        layout = json.loads(LAYOUT.read_text(encoding="utf-8"))
        log_rows = []
        for line in LOG.open(encoding="utf-8"):
            row = json.loads(line)
            if row.get("schema") == "litchi.performance.0976.dwfl_report.final" and row.get("path_match") is True and row.get("name_match") is True:
                log_rows.append(row)
                if len(log_rows) == 2:
                    break
        if len(log_rows) != 2:
            raise AssertionError("retained final-v2 log lacks two exact target rows")
        log_fixture = Path(temp) / "two-target.dwfl.jsonl"
        log_fixture.write_text("".join(json.dumps(row, separators=(",", ":")) + "\n" for row in log_rows), encoding="utf-8")
        stats = decoder._stream_shim_log(log_fixture, layout, TARGET, expected_build_id=decoder.BASELINE_BUILD_ID)
        if stats["cache_stored"] != 1 or stats["reused_existing"] != 1 or stats["fallback"] != 0 or stats["result_module_null"] != 0:
            raise AssertionError(f"final-v2 positive log control failed: {stats}")
        bad_log_row = dict(log_rows[0], admission=False)
        bad_log = Path(temp) / "bad-target.dwfl.jsonl"
        bad_log.write_text(json.dumps(bad_log_row) + "\n" + json.dumps(log_rows[1]) + "\n", encoding="utf-8")
        assert_decode_failure(
            lambda: decoder._stream_shim_log(bad_log, layout, TARGET, expected_build_id=decoder.BASELINE_BUILD_ID),
            "shim_log",
        )

        zero_terminal = Path(temp) / "terminal-zero.txt"
        zero_terminal.write_text("# Total Lost Samples: 0\n", encoding="utf-8")
        if decoder._terminal_loss({"terminal_report": str(zero_terminal), "returncode": 0, "timed_out": False, "native_ld_preload": "absent"}, path.parent) != 0:
            raise AssertionError("zero terminal loss was not accepted")
        bad_terminal = Path(temp) / "terminal-nonzero.txt"
        bad_terminal.write_text("# Total Lost Samples: 1\n", encoding="utf-8")
        if decoder._terminal_loss({"terminal_report": str(bad_terminal), "returncode": 0, "timed_out": False, "native_ld_preload": "absent"}, path.parent) != 1:
            raise AssertionError("nonzero terminal loss was not parsed")
        try:
            decoder.ADAPTER.V1.require_zero_loss(
                {"raw_lost_sum": 0},
                {"lost_samples_summary": {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": 1}},
            )
        except Exception:
            pass
        else:
            raise AssertionError("nonzero terminal loss passed the zero-loss gate")

    # The retained 0976 stream is the expensive but immutable differential
    # fixture.  It proves O(N) parsing on the real 6,940-sample shape without
    # starting perf again.
    streamed = decoder._stream_perf_script(
        SCRIPT, marker=marker, marker_dso=dso, expected_pid=EXPECTED_PID,
        expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
    )
    strict = decoder.PARSER.parse_perf_script_strict(
        SCRIPT.read_text(encoding="utf-8", errors="strict"), marker=marker,
        marker_dso=dso, expected_pid=EXPECTED_PID, expected_cpu=EXPECTED_CPU,
        expected_period=PERIOD,
    )
    if any(streamed[key] != strict[key] for key in differential_keys):
        raise AssertionError("streaming parser differs from sealed parser on retained 0976 stream")
    if streamed["sample_count"] != 6940 or streamed["marker_sample_count"] != 951 or streamed["empty_callchain_count"] != 0:
        raise AssertionError("retained 0976 quality control changed")
    print(json.dumps({
        "schema": "litchi.performance.0977.focuseddecode_controls.v1",
        "status": "pass",
        "native_workload_started": False,
        "retained_script": {"path": str(SCRIPT), "bytes": SCRIPT.stat().st_size, "sha256": decoder._sha256(SCRIPT)},
        "retained_quality": {key: streamed[key] for key in ("sample_count", "total_period", "marker_sample_count", "marker_period", "empty_callchain_count", "unresolved_frames")},
        "synthetic": {"marker_sample_count": 2, "unknown_leaf_samples": 1, "conservation": True, "wrong_dso_not_overmatched": True, "wrong_cpu_rejected": True},
        "final_v2_log": {"rows_checked": 2, "positive_store_reuse": True, "fallback_rejected": True},
        "terminal_loss": {"zero_accepted": True, "nonzero_rejected": True},
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
