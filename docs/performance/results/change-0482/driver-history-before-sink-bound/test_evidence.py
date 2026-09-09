#!/usr/bin/env python3
"""Focused, data-only tests for the 0482 evidence scripts."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest

import analyze
import verify


ROOT = Path(__file__).resolve().parent


def _runner_module():
    spec = importlib.util.spec_from_file_location("change0482_runner", ROOT / "run-measurements.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load run-measurements.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _meta(path: Path) -> dict[str, object]:
    return {"bytes": path.stat().st_size, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def _inventory(directory: Path) -> dict[str, dict[str, object]]:
    return {
        path.relative_to(directory).as_posix(): _meta(path)
        for path in sorted(directory.rglob("*"))
        if path.is_file()
    }


def _inventory_digest(values: dict[str, dict[str, object]]) -> str:
    encoded = json.dumps(values, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def _fixture(root: Path) -> None:
    """Create a tiny complete bundle using the real verifier and analyzer."""

    for name in ("common.py", "gate.py", "build.py", "run-gates.py", "run-measurements.py", "analyze.py", "verify.py", "test_evidence.py", "fuzz.py"):
        shutil.copy2(ROOT / name, root / name)
    final_gate_commands = {
        label: ["python3", "-B", str(root / "run-gates.py"), label]
        for label in analyze.REQUIRED_VALIDATION_LABELS
        if label not in ("build-normal", "build-allocator")
    }
    (root / "run-gates.py").write_text(
        "GATES = " + repr(final_gate_commands) + "\n",
        encoding="utf-8",
    )
    binary = root / "xml_stream_audit"
    shutil.copy2("/bin/true", binary)
    binary.chmod(0o755)
    binary_record = {"path": str(binary), **_meta(binary)}
    xml_assets = {
        f"crates/synthetic/src/assets/asset-{index:02d}.xml": "0" * 64
        for index in range(1, 78)
    }
    inventory = {
        "schema": "xml-stream-audit-xml-assets-v1",
        "files": len(xml_assets),
        "assets": xml_assets,
    }
    (root / "xml-assets.json").write_text(json.dumps(inventory, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    target_source_hash = "1" * 64
    fuzz_root = root / "fuzz"
    fuzz_final = fuzz_root / "final"
    fuzz_inputs = fuzz_final / "build-inputs"
    fuzz_inputs.mkdir(parents=True, exist_ok=True)
    manifest_bytes = b"[package]\nname = \"synthetic-fuzz\"\nversion = \"0.0.0\"\n"
    lock_bytes = b"# synthetic lock\nversion = 4\n"
    (fuzz_inputs / "Cargo.toml").write_bytes(manifest_bytes)
    (fuzz_inputs / "Cargo.lock.txt").write_bytes(lock_bytes)
    seed_names = (
        "binary/invalid-utf8.bin",
        "binary/truncated-utf8.bin",
        "xml/attributes-over.xml",
        "xml/authored-space.xml",
        "xml/bom-comment.xml",
        "xml/depth-over.xml",
        "xml/end-tag-space.xml",
        "xml/malformed-reference.xml",
        "xml/token-over.xml",
        "xml/truncated.xml",
        "xml/valid-mixed.xml",
    )
    seed_directory = fuzz_root / "seeds"
    for index, name in enumerate(seed_names):
        path = seed_directory / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(f"synthetic seed {index}\n".encode())
    seed_inventory = _inventory(seed_directory)
    corpus_before = {name.replace("/", "-"): value for name, value in seed_inventory.items()}
    (fuzz_final / "post-run" / "corpus").mkdir(parents=True, exist_ok=True)
    (fuzz_final / "post-run" / "artifacts").mkdir(parents=True, exist_ok=True)
    for index, name in enumerate(seed_names):
        (fuzz_final / "post-run" / "corpus" / name.replace("/", "-")).write_bytes(f"synthetic seed {index}\n".encode())
    (fuzz_final / "post-run" / "artifacts" / "no-crash").write_bytes(b"")
    post_run_inventory = _inventory(fuzz_final / "post-run")
    manifest_hash = _meta(fuzz_inputs / "Cargo.toml")["sha256"]
    manifest = {
        "synthetic.rs": "0" * 64,
        "crates/xml-minifier/fuzz/fuzz_targets/minify_xml.rs": target_source_hash,
        "docs/performance/results/change-0482/fuzz/final/build-inputs/Cargo.toml": manifest_hash,
        **xml_assets,
    }
    manifest_bytes = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode()
    manifest_digest = hashlib.sha256(manifest_bytes).hexdigest()
    source_manifest = root / "validation-sources" / f"{manifest_digest}.json"
    source_manifest.parent.mkdir(parents=True, exist_ok=True)
    source_manifest.write_bytes(manifest_bytes)
    source_snapshot = {"path": source_manifest.relative_to(root).as_posix(), "sha256": hashlib.sha256(source_manifest.read_bytes()).hexdigest(), "files": len(manifest)}

    runner = _runner_module()
    recorded_repo = Path("/recorded/litchi")
    fuzz_work = Path("/recorded/work/fuzz-xml-final")
    fuzz_binary = {"path": str(fuzz_work / "minify_xml"), **_meta(binary)}
    prepared = {
        "prepared_utc": "2026-09-08T00:00:01+00:00",
        "target_source": {"bytes": 1, "sha256": target_source_hash},
        "manifest": _meta(fuzz_inputs / "Cargo.toml"),
        "lock": _meta(fuzz_inputs / "Cargo.lock.txt"),
        "seeds": seed_inventory,
        "corpus_before": corpus_before,
    }
    (fuzz_final / "prepared.json").write_text(json.dumps(prepared, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    fuzz_build_argv = [
        "env",
        f"CARGO_TARGET_DIR={recorded_repo / 'target/fuzz-asan'}",
        "RUSTC_BOOTSTRAP=1",
        f"RUSTFLAGS={verify.FUZZ_FLAGS}",
        "cargo", "build", "--release", "--locked", "--manifest-path", str(fuzz_work / "Cargo.toml"),
        "--target", analyze.FUZZ_TARGET, "--bin", "minify_xml",
    ]
    fuzz_build = {
        "argv": fuzz_build_argv,
        "cwd": str(recorded_repo),
        "started_utc": "2026-09-08T00:00:02.100000+00:00",
        "finished_utc": "2026-09-08T00:00:02.200000+00:00",
        "source_snapshot": source_snapshot,
        "inputs": prepared,
        "binary": fuzz_binary,
    }
    (fuzz_final / "build.json").write_text(json.dumps(fuzz_build, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    fuzz_smoke_argv = [
        str(fuzz_work / "minify_xml"), str(fuzz_work / "corpus"), "-runs=10000", "-seed=482",
        "-max_len=65536", "-timeout=10", f"-artifact_prefix={fuzz_work / 'artifacts'}/",
    ]
    fuzz_smoke = {
        "argv": fuzz_smoke_argv,
        "cwd": str(recorded_repo),
        "started_utc": "2026-09-08T00:00:02.300000+00:00",
        "finished_utc": "2026-09-08T00:00:02.400000+00:00",
        "exit_code": 0,
        "binary_before": fuzz_binary,
        "binary_after": fuzz_binary,
        "corpus_before": corpus_before,
        "retained": post_run_inventory,
    }
    (fuzz_final / "smoke.json").write_text(json.dumps(fuzz_smoke, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    def fuzz_ref(relative: str) -> dict[str, object]:
        return {"path": relative, **_meta(root / relative)}

    fuzz_reference = {
        "schema": analyze.FUZZ_SCHEMA,
        "cwd": str(recorded_repo),
        "helper": fuzz_ref("fuzz.py"),
        "prepared": fuzz_ref("fuzz/final/prepared.json"),
        "build": fuzz_ref("fuzz/final/build.json"),
        "smoke": fuzz_ref("fuzz/final/smoke.json"),
        "manifest": fuzz_ref("fuzz/final/build-inputs/Cargo.toml"),
        "lock": fuzz_ref("fuzz/final/build-inputs/Cargo.lock.txt"),
        "seed_inventory": {"path": "fuzz/seeds", "files": len(seed_inventory), "sha256": _inventory_digest(seed_inventory)},
        "post_run_inventory": {"path": "fuzz/final/post-run", "files": len(post_run_inventory), "sha256": _inventory_digest(post_run_inventory)},
    }
    binaries = {"normal": binary_record, "allocator": binary_record}
    build_records = {}
    (root / "builds").mkdir(parents=True, exist_ok=True)
    for instrumentation in runner.INSTRUMENTATIONS:
        build_argv = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", "xml_stream_audit"]
        if instrumentation == "allocator":
            build_argv.extend(["--features", "allocator-metrics"])
        build_records[instrumentation] = {
            "schema": "xml-stream-audit-build-v1",
            "instrumentation": instrumentation,
            "source_snapshot": source_snapshot,
            "argv": build_argv,
            "git_revision": "0" * 40,
            "environment": {key: runner.ENV[key] for key in runner.ENV_KEYS},
            "started_utc": "2026-09-08T00:00:00+00:00",
            "finished_utc": "2026-09-08T00:00:01+00:00",
            "binary": binary_record,
        }
        (root / "builds" / f"{instrumentation}.json").write_text(json.dumps(build_records[instrumentation], indent=2, sort_keys=True) + "\n", encoding="utf-8")
    captures = runner._capture_rows("normal", binaries) + runner._capture_rows("allocator", binaries)
    protocol = {
        "schema": runner.SCHEMA,
        "report_schema": runner.REPORT_SCHEMA,
        "version": 1,
        "samples": runner.SAMPLES,
        "warmups": runner.WARMUPS,
        "cpu": runner.CPU,
        "sizes": list(runner.SIZES),
        "modes": list(runner.MODES),
        "arms": list(runner.ARMS),
        "repeats": list(runner.REPEATS),
        "sequence": "A1/B1/B2/A2",
        "comparison": "synthetic same-build route comparison",
        "performance_claim": "primitive-enabler-only",
        "no_cross_revision_claim": True,
        "process_rss": "GNU time -v",
        "normal_and_allocator_timings_separate": True,
        "pair_review_percent": 5,
        "required_validation_labels": list(analyze.REQUIRED_VALIDATION_LABELS),
        "environment": {key: runner.ENV[key] for key in runner.ENV_KEYS},
        "source_snapshot": source_snapshot,
        "binaries": binaries,
        "builds": {name: {"path": f"builds/{name}.json", "sha256": hashlib.sha256((root / "builds" / f"{name}.json").read_bytes()).hexdigest()} for name in runner.INSTRUMENTATIONS},
        "scripts": {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in ("common.py", "gate.py", "build.py", "run-gates.py", "run-measurements.py", "analyze.py", "verify.py", "test_evidence.py", "fuzz.py")},
        "xml_asset_inventory": {"path": "xml-assets.json", "sha256": hashlib.sha256((root / "xml-assets.json").read_bytes()).hexdigest(), "files": 77},
        "fuzz": fuzz_reference,
        "captures": captures,
    }
    (root / "protocol.json").write_text(json.dumps(protocol, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    protocol_hash = hashlib.sha256((root / "protocol.json").read_bytes()).hexdigest()

    (root / "validation").mkdir(parents=True, exist_ok=True)
    for instrumentation in runner.INSTRUMENTATIONS:
        label = f"build-{instrumentation}"
        recorded_repo = Path("/recorded/litchi")
        gate_argv = ["python3", "-B", str(recorded_repo / "docs/performance/results/change-0482/build.py"), instrumentation]
        gate_started = {"argv": gate_argv, "cwd": str(recorded_repo), "environment": {key: runner.ENV[key] for key in runner.ENV_KEYS}, "driver_sha256": hashlib.sha256((root / "gate.py").read_bytes()).hexdigest(), "common_sha256": hashlib.sha256((root / "common.py").read_bytes()).hexdigest(), "started_utc": "2026-09-08T00:00:00+00:00", "source_before": source_snapshot}
        gate_prefix = root / "validation" / label
        gate_prefix.with_suffix(".stdout").write_bytes(b"")
        gate_prefix.with_suffix(".stderr").write_bytes(b"")
        gate_artifacts = {path.name: {"bytes": path.stat().st_size, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()} for path in (gate_prefix.with_suffix(".stdout"), gate_prefix.with_suffix(".stderr"))}
        gate_terminal = dict(gate_started, exit_code=0, finished_utc="2026-09-08T00:00:02+00:00", source_after=source_snapshot, source_unchanged=True, artifacts=gate_artifacts)
        gate_prefix.with_suffix(".started.json").write_text(json.dumps(gate_started, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        gate_prefix.with_suffix(".json").write_text(json.dumps(gate_terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    for label in analyze.REQUIRED_VALIDATION_LABELS:
        if label in ("build-normal", "build-allocator"):
            continue
        gate_argv = ["python3", "-B", str(root / "run-gates.py"), label]
        gate_started = {
            "argv": gate_argv,
            "cwd": str(root),
            "environment": {key: runner.ENV[key] for key in runner.ENV_KEYS},
            "driver_sha256": hashlib.sha256((root / "gate.py").read_bytes()).hexdigest(),
            "common_sha256": hashlib.sha256((root / "common.py").read_bytes()).hexdigest(),
            "started_utc": "2026-09-08T00:00:00+00:00",
            "source_before": source_snapshot,
        }
        gate_prefix = root / "validation" / label
        gate_prefix.with_suffix(".stdout").write_bytes(b"")
        gate_prefix.with_suffix(".stderr").write_bytes(b"")
        gate_artifacts = {
            path.name: {"bytes": path.stat().st_size, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
            for path in (gate_prefix.with_suffix(".stdout"), gate_prefix.with_suffix(".stderr"))
        }
        gate_terminal = dict(
            gate_started,
            exit_code=0,
            finished_utc="2026-09-08T00:00:02+00:00",
            source_after=source_snapshot,
            source_unchanged=True,
            artifacts=gate_artifacts,
        )
        gate_prefix.with_suffix(".started.json").write_text(json.dumps(gate_started, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        gate_prefix.with_suffix(".json").write_text(json.dumps(gate_terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    for spec in captures:
        prefix = root / "captures" / spec["label"]
        prefix.parent.mkdir(parents=True, exist_ok=True)
        for suffix, content in ((".stdout", b""), (".stderr", b""), (".resource", b"Maximum resident set size (kbytes): 1\n")):
            prefix.with_suffix(suffix).write_bytes(content)
        size = spec["size_bytes"]
        mode = spec["mode"]
        samples = []
        for index in range(30):
            allocation = None
            if spec["instrumentation"] == "allocator":
                allocation = {
                    "status": "measured",
                    "scope": "operation_global_system_allocator",
                    "allocation_calls": 1,
                    "deallocation_calls": 1,
                    "reallocation_calls": 0,
                    "failed_allocation_calls": 0,
                    "allocated_bytes": 100,
                    "deallocated_bytes": 100,
                    "live_bytes_before": 100,
                    "live_bytes_after": 100,
                    "peak_live_bytes_before": 100,
                    "peak_live_bytes_after": 100,
                    "region_peak_live_bytes": 100,
                }
            samples.append({
                "sample": index,
                "elapsed_ns": 1000 + index,
                "actual_bytes": size,
                "source_bytes_match": True,
                "outcome": {"status": "success", "report": {"attributes": 0, "bytes": size, "events": 3, "max_depth": 2, "text_bytes": size - 12}},
                "allocation": allocation,
                "operation": {
                    "generator_buffer_bytes": 16384,
                    "generator_position": size,
                    "source_materialization_bytes": size if mode == "materialized" else None,
                    "entry_live_bytes": 100 if allocation else None,
                    "exit_live_bytes": 100 if allocation else None,
                    "net_live_bytes": 0 if allocation else None,
                    "zero_net_live": True if allocation else None,
                },
            })
        report = {
            "schema": runner.REPORT_SCHEMA,
            "benchmark": "litchi-xml-repetitive-root-items-v1",
            "binary": {
                "binary": "litchi-perf-baseline" if spec["instrumentation"] == "normal" else "litchi-perf-baseline-alloc",
                "allocator": "Rust system allocator" if spec["instrumentation"] == "normal" else "CountingSystemAllocator(std::alloc::System)",
                "instrumentation": "none" if spec["instrumentation"] == "normal" else "system_allocator_operation_scoped",
                "counter_revision": None if spec["instrumentation"] == "normal" else "serialized_region_peak_v3",
            },
            "modes": [mode], "sizes": [size], "samples": 30, "warmups": 3,
            "timed_scope": "generator_create_read_materialize_or_stream_audit_and_drop",
            "corpus_scope": "deterministic_generator_hash_and_layout_are_precomputed_outside_timed_regions",
            "cases": [{
                "mode": mode,
                "size_bytes": size,
                "source": {
                    "requested_bytes": size,
                    "generated_bytes": size,
                    "sha256": analyze._expected_source_sha256(size),
                    "actual_generator_bytes": size,
                    "actual_generator_sha256": analyze._expected_source_sha256(size),
                    "actual_generator_oracle_verified": True,
                    "generator": "litchi-xml-repetitive-root-items-v1",
                    "chunk_bytes": 16384,
                    "source_materialization": "materialized_leg_only",
                    "layout": analyze._expected_layout(size),
                },
                "limits": {
                    "max_bytes": size,
                    "max_depth": 8,
                    "max_events": 4_000_000,
                    "max_attributes": 16,
                    "max_token_bytes": 16384,
                    "max_text_bytes": size,
                    "streaming_memory_upper_bound": 1,
                },
                "warmups": 3,
                "samples": samples,
            }],
        }
        report_path = prefix.with_suffix(".report.json")
        report_path.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        artifacts = {path.name: {"path": str(path), **_meta(path)} for path in (prefix.with_suffix(".stdout"), prefix.with_suffix(".stderr"), prefix.with_suffix(".resource"), report_path)}
        started = {
            "schema": runner.SCHEMA, "capture": spec, "cwd": str(root), "started_utc": "2026-09-08T00:00:03+00:00", "protocol_sha256": protocol_hash,
            "source_before": source_snapshot, "binary_before": binary_record, "environment": {"LC_ALL": "C"},
        }
        terminal = dict(started, exit_code=0, finished_utc="2026-09-08T00:00:04+00:00", source_after=source_snapshot, binary_after=binary_record, source_unchanged=True, binary_unchanged=True, artifacts=artifacts)
        (prefix.with_suffix(".started.json")).write_text(json.dumps(started, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        (prefix.with_suffix(".json")).write_text(json.dumps(terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    (root / "pilots").mkdir(parents=True, exist_ok=True)
    for instrumentation in runner.INSTRUMENTATIONS:
        for pilot in runner._pilot_rows(instrumentation, binaries[instrumentation], 1):
            prefix = root / "pilots" / pilot["label"]
            formal_label = f"{instrumentation}-{pilot['mode']}-r1-{pilot['size_bytes']}"
            formal_report = json.loads((root / "captures" / f"{formal_label}.report.json").read_text(encoding="utf-8"))
            formal_report["samples"] = 1
            formal_report["warmups"] = 1
            formal_report["cases"][0]["warmups"] = 1
            formal_report["cases"][0]["samples"] = [formal_report["cases"][0]["samples"][0]]
            pilot_report = json.dumps(formal_report, indent=2, sort_keys=True).encode() + b"\n"
            for suffix, content in ((".stdout", b""), (".stderr", b""), (".resource", b"Maximum resident set size (kbytes): 1\n"), (".report.json", pilot_report)):
                prefix.with_suffix(suffix).write_bytes(content)
            artifacts = {path.name: {"bytes": path.stat().st_size, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()} for path in (prefix.with_suffix(".stdout"), prefix.with_suffix(".stderr"), prefix.with_suffix(".resource"), prefix.with_suffix(".report.json"))}
            started = {"schema": runner.SCHEMA, "capture": pilot, "cwd": str(root), "started_utc": "2026-09-08T00:00:00+00:00", "source_before": source_snapshot, "binary_before": binaries[instrumentation], "environment": {key: runner.ENV[key] for key in runner.ENV_KEYS}}
            terminal = dict(started, exit_code=0, finished_utc="2026-09-08T00:00:00.500000+00:00", source_after=source_snapshot, binary_after=binaries[instrumentation], source_unchanged=True, binary_unchanged=True, artifacts=artifacts)
            prefix.with_suffix(".started.json").write_text(json.dumps(started, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            prefix.with_suffix(".json").write_text(json.dumps(terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    old_root = analyze.ROOT
    analyze.ROOT = root
    try:
        rows = analyze.load_rows(protocol)
        summary = analyze.summarize(protocol, rows)
    finally:
        analyze.ROOT = old_root
    (root / "summary.json").write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")


class EvidenceTests(unittest.TestCase):
    def test_abba_plan_is_twelve_rows_per_instrumentation(self) -> None:
        runner = _runner_module()
        self.assertEqual(len(analyze.REQUIRED_VALIDATION_LABELS), 24)
        self.assertIn("xml-fuzz-build", analyze.REQUIRED_VALIDATION_LABELS)
        self.assertIn("xml-fuzz-smoke", analyze.REQUIRED_VALIDATION_LABELS)
        binaries = {name: {"path": f"/tmp/{name}", "bytes": 1, "sha256": "0" * 64} for name in runner.INSTRUMENTATIONS}
        rows = runner._capture_rows("normal", binaries)
        self.assertEqual(len(rows), 12)
        self.assertEqual([row["arm"] for row in rows], ["materialized"] * 3 + ["streaming"] * 3 + ["streaming"] * 3 + ["materialized"] * 3)
        self.assertEqual([row["size_bytes"] for row in rows], [*runner.SIZES, *runner.SIZES, *reversed(runner.SIZES), *reversed(runner.SIZES)])
        self.assertIn("--sizes", rows[0]["argv"])
        self.assertIn("--warmup", rows[0]["argv"])

    def test_stats_are_reproducible_and_include_ci_percentiles(self) -> None:
        result = analyze._stats([10, 20, 30, 40, 50])
        self.assertEqual(result["n"], 5)
        self.assertEqual(result["p50_ns"], 30)
        self.assertEqual(result["p95_ns"], 50)
        self.assertEqual(result["p99_ns"], 50)
        self.assertEqual(len(result["ci95_mean_ns"]), 2)

    def test_generator_digest_known_u8_cycle_boundaries(self) -> None:
        # These literal digests were computed independently.  Each payload
        # has at least 257 full records, so records 255 and 256 exercise the
        # Rust `(record as u8) % 26` wrap rather than Python's direct modulo.
        self.assertEqual(
            analyze._expected_source_sha256(13 + 257 * 1024),
            "7157e703b2c93a9aba24ba3b43f9ab6b3f7d87c81703d3cd17be83fe8884877b",
        )
        self.assertEqual(
            analyze._expected_source_sha256(13 + 258 * 1024),
            "8c31397d8ffab4cb91aad65ab3b1544f7cfbfef5678058a6fb2b96230f96bef7",
        )

    def test_verifier_rejects_missing_capture_and_tampered_summary(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0482-evidence-") as temporary:
            root = Path(temporary)
            _fixture(root)
            protocol = json.loads((root / "protocol.json").read_text(encoding="utf-8"))
            missing = root / "captures" / (protocol["captures"][0]["label"] + ".json")
            missing.unlink()
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)

            _fixture(root)
            binary_path = Path(json.loads((root / "protocol.json").read_text(encoding="utf-8"))["binaries"]["normal"]["path"])
            binary_path.unlink()
            self.assertTrue(verify.verify_bundle(root)["data_only_recomputed"])
            with tempfile.TemporaryDirectory(prefix="litchi-0482-copy-") as copied_directory:
                copied_root = Path(copied_directory) / "bundle"
                shutil.copytree(root, copied_root)
                self.assertTrue(verify.verify_bundle(copied_root)["data_only_recomputed"])

            _fixture(root)
            for instrumentation in ("normal", "allocator"):
                for suffix in (".started.json", ".json"):
                    receipt_path = root / "validation" / f"build-{instrumentation}{suffix}"
                    receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
                    receipt["argv"][2] = "docs/performance/results/change-0482/build.py"
                    receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            self.assertTrue(verify.verify_bundle(root)["data_only_recomputed"])

            _fixture(root)
            fuzz_smoke_path = root / "fuzz/final/smoke.json"
            fuzz_smoke = json.loads(fuzz_smoke_path.read_text(encoding="utf-8"))
            fuzz_smoke["argv"][2] = "-runs=1"
            fuzz_smoke_path.write_text(json.dumps(fuzz_smoke, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)

            _fixture(root)
            (root / "xml-assets.json").unlink()
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)

            _fixture(root)
            summary_path = root / "summary.json"
            summary = json.loads(summary_path.read_text(encoding="utf-8"))
            summary["rows"][0]["latency"]["mean_ns"] += 1
            summary_path.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)

            _fixture(root)
            protocol_path = root / "protocol.json"
            protocol = json.loads(protocol_path.read_text(encoding="utf-8"))
            protocol["required_validation_labels"].remove("registry-strict")
            protocol_path.write_text(json.dumps(protocol, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)

            _fixture(root)
            gate_started_path = root / "validation" / "format-xml-final.started.json"
            gate_terminal_path = root / "validation" / "format-xml-final.json"
            gate_started = json.loads(gate_started_path.read_text(encoding="utf-8"))
            gate_terminal = json.loads(gate_terminal_path.read_text(encoding="utf-8"))
            gate_started["argv"] = ["true"]
            gate_terminal["argv"] = ["true"]
            gate_started_path.write_text(json.dumps(gate_started, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            gate_terminal_path.write_text(json.dumps(gate_terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_bundle(root)


if __name__ == "__main__":
    unittest.main()
