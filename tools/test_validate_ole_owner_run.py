#!/usr/bin/env python3
"""Build-free round-trip and mutation tests for the OLE run validator."""

from __future__ import annotations

import copy
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
from unittest import mock

from tools.native_qualification import ole_owner_paired_runner as runner
from tools.native_qualification import prepare_ole_owner_inputs as preparation
from tools.native_qualification import validate_ole_owner_run as validator


def _git(root: Path, *args: str) -> str:
    return subprocess.run(
        ["git", "-C", str(root), *args], check=True, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, text=True,
    ).stdout.strip()


def _source(root: Path, kind: str) -> str:
    root.mkdir()
    _git(root, "init", "--quiet")
    _git(root, "config", "user.email", "validator@test.invalid")
    _git(root, "config", "user.name", "OLE validator")
    (root / "source.txt").write_text(kind, encoding="utf-8")
    cfb = root / preparation.CFB_SOURCE_PATH
    cfb.parent.mkdir(parents=True)
    if kind == "baseline":
        cfb.write_text(
            "pub fn open(reader: R) -> Result<Self, OleError> {}\n"
            "pub fn open_with_limits(mut reader: R, limits: OleFileLimits) {}\n",
            encoding="utf-8",
        )
    else:
        cfb.write_text("pub fn open(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError> {}\n", encoding="utf-8")
    _git(root, "add", ".")
    _git(root, "commit", "--quiet", "-m", kind)
    return _git(root, "rev-parse", "HEAD")


def _sha(path: Path) -> tuple[str, int]:
    data = path.read_bytes()
    return hashlib.sha256(data).hexdigest(), len(data)


def _record(path: Path) -> dict[str, object]:
    digest, size = _sha(path)
    return {"path": str(path), "sha256": digest, "bytes": size}


def _write(path: Path, data: bytes, executable: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    if executable:
        path.chmod(0o755)


class ValidateOleOwnerRunTests(unittest.TestCase):
    def _fixture(self, root: Path) -> tuple[dict[str, Path], Path]:
        baseline = root / "baseline"
        candidate = root / "candidate"
        baseline_ref = _source(baseline, "baseline")
        candidate_ref = _source(candidate, "candidate")
        markers = "\n".join(preparation.REQUIRED_PROBE_MARKERS)
        candidate_probe = candidate / preparation.PROBE_PATH
        candidate_probe.parent.mkdir(parents=True)
        candidate_probe.write_text(
            markers + "\nOleFile::open(\n        Cursor::new(package.into_inner()),\n        OleFileLimits::default(),\nOleFileLimits::default()\n",
            encoding="utf-8",
        )
        _git(candidate, "add", ".")
        _git(candidate, "commit", "--quiet", "-m", "candidate probe")
        candidate_ref = _git(candidate, "rev-parse", "HEAD")
        reviewed = root / "reviewed.rs"
        reviewed.write_text(markers + "\nOleFile::open(Cursor::new(package.into_inner()))\n", encoding="utf-8")
        clone = root / "clone.rs"
        clone.write_text("\n".join(preparation.REQUIRED_CLONE_MARKERS), encoding="utf-8")
        oracle = root / "oracle.rs"
        oracle.write_text("fn check_ole_outputs\nallocator_evidence\nowner_target\n", encoding="utf-8")
        reviewed_lock = root / "reviewed.Cargo.lock"
        reviewed_lock.write_text(
            "version = 4\n\n"
            "[[package]]\nname = \"bitflags\"\nversion = \"2.0.0\"\n"
            "source = \"registry+https://example.invalid/index\"\n"
            "checksum = \"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\"\n\n"
            "[[package]]\nname = \"sha2\"\nversion = \"0.11.0\"\n"
            "source = \"registry+https://example.invalid/index\"\n"
            "checksum = \"bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb\"\n\n"
            "[[package]]\nname = \"litchi-core\"\nversion = \"0.0.1\"\n\n"
            "[[package]]\nname = \"litchi-odraw\"\nversion = \"0.0.1\"\ndependencies = [\n \"bitflags\",\n]\n\n"
            "[[package]]\nname = \"litchi-perf-baseline\"\nversion = \"0.1.0\"\ndependencies = [\n \"litchi-odraw\",\n]\n",
            encoding="utf-8",
        )
        baseline_lock = root / "baseline.Cargo.lock"
        baseline_lock.write_text(
            reviewed_lock.read_text(encoding="utf-8").replace(
                ' "litchi-odraw",\n]\n', ' "litchi-odraw",\n "sha2 0.11.0",\n]\n'
            ),
            encoding="utf-8",
        )
        candidate_lock = root / "candidate.Cargo.lock"
        candidate_lock.write_text(
            baseline_lock.read_text(encoding="utf-8").replace(
                'name = "litchi-odraw"\nversion = "0.0.1"\ndependencies = [\n "bitflags",\n]',
                'name = "litchi-odraw"\nversion = "0.0.1"\ndependencies = [\n "bitflags",\n "litchi-core",\n]'
            ),
            encoding="utf-8",
        )
        pins = {
            "BASELINE_COMMIT": baseline_ref,
            "REVIEWED_PROBE_SHA256": _sha(reviewed)[0],
            "CANDIDATE_ADAPTER_SHA256": _sha(candidate_probe)[0],
            "CLONE_PROBE_SHA256": _sha(clone)[0],
            "ORACLE_SOURCE_SHA256": _sha(oracle)[0],
            "REVIEWED_LOCK_SHA256": _sha(reviewed_lock)[0],
            "BASELINE_LOCK_SHA256": _sha(baseline_lock)[0],
            "CANDIDATE_LOCK_SHA256": _sha(candidate_lock)[0],
        }
        prep_path = root / "preparation.json"
        with mock.patch.multiple(preparation, **pins):
            preparation.prepare(
                baseline_root=baseline,
                candidate_root=candidate,
                candidate_ref=candidate_ref,
                reviewed_probe=reviewed,
                clone_probe=clone,
                oracle_probe=oracle,
                reviewed_lock=reviewed_lock,
                baseline_lock=baseline_lock,
                candidate_lock=candidate_lock,
                drivers_root=root / "drivers",
                output=prep_path,
            )
        manifest = json.loads((Path(__file__).parents[0] / "native_qualification/manifests/ole-owner.json").read_text())
        manifest["source_policy"]["baseline_commit"] = baseline_ref
        manifest["input_pins"].update({
            "reviewed_probe_sha256": pins["REVIEWED_PROBE_SHA256"],
            "candidate_adapter_sha256": pins["CANDIDATE_ADAPTER_SHA256"],
            "clone_probe_sha256": pins["CLONE_PROBE_SHA256"],
            "oracle_source_sha256": pins["ORACLE_SOURCE_SHA256"],
            "reviewed_lock_sha256": pins["REVIEWED_LOCK_SHA256"],
            "baseline_lock_sha256": pins["BASELINE_LOCK_SHA256"],
            "candidate_lock_sha256": pins["CANDIDATE_LOCK_SHA256"],
        })
        manifest_path = root / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, sort_keys=True), encoding="utf-8")
        build_dir = root / "run"
        target_root = build_dir / "targets"
        build_dir.mkdir()
        preparation_document = json.loads(prep_path.read_text())

        identities: dict[str, dict[str, object]] = {}
        for role, endpoint in (("before", "baseline"), ("after", "candidate")):
            source = json.loads(prep_path.read_text())[endpoint]
            source_keys = ("root", "head", "tree", "status", "status_paths", "clean", "status_sha256", "diff_sha256", "production_tree_sha256")
            identity: dict[str, object] = {
                "source": {key: source[key] for key in source_keys},
                "driver_root": source["driver"]["root"],
                "driver_files": source["driver"]["files"],
            }
            identity["identity_sha256"] = hashlib.sha256(json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
            identities[role] = identity

        cargo_tool = Path(shutil.which("cargo") or os.sys.executable).absolute()
        rustc_tool = Path(shutil.which("rustc") or os.sys.executable).absolute()
        path_env = os.environ.get("PATH", "")
        toolchain = {
            "path_env": path_env,
            "cargo": {**_record(cargo_tool), "resolved_path": str(cargo_tool.resolve()), "version": "fixture"},
            "rustc": {**_record(rustc_tool), "resolved_path": str(rustc_tool.resolve()), "version": "fixture"},
        }
        build_events = []
        builds: dict[str, object] = {}
        for key in runner.BUILD_KEYS:
            role, variant = key.split("/")
            target = target_root / f"{role}-{variant}"
            binaries = {
                "probe": target / "release/examples/ole_w5_probe",
                "clone": target / "release/examples/ole_package_clone_probe",
                "oracle": target / "release/examples/ole_owner_oracle",
            }
            artifacts = {}
            outputs = []
            for name, path in binaries.items():
                _write(path, f"{key}-{name}".encode(), executable=True)
                artifacts[name] = _record(path)
                outputs.append({**_record(path), "exists": True})
            stdout = build_dir / "commands" / f"{key.replace('/', '-')}.stdout"
            stderr = build_dir / "commands" / f"{key.replace('/', '-')}.stderr"
            _write(stdout, b"")
            _write(stderr, b"")
            driver_root = Path(preparation_document["baseline" if role == "before" else "candidate"]["driver"]["root"])
            command = [
                str(cargo_tool), "build", "--release", "--offline", "--locked", "--manifest-path",
                str(driver_root / "Cargo.toml"), "--no-default-features", "--features",
                validator.EXPECTED_FEATURES[key], "--example", "ole_w5_probe", "--example",
                "ole_package_clone_probe", "--example", "ole_owner_oracle",
            ]
            event = {
                "event_id": f"build-{key.replace('/', '-')}", "phase": "build", "role": role, "variant": variant,
                "argv": command, "cwd": str(driver_root),
                "source_identity_sha256": identities[role]["identity_sha256"],
                "source_identity_before": identities[role], "source_identity_after": identities[role],
                "returncode": 0, "runner_error": None, "process_guard_violations": [],
                "post_exit_descendants": [], "post_exit_offenders": [],
                "environment": {
                    "CARGO_TARGET_DIR": str(target), "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0",
                    "LITCHI_MEM_MAX": "20G", "PATH": path_env, "CARGO": str(cargo_tool),
                    "RUSTC": str(rustc_tool), "RUSTC_WRAPPER": None, "RUSTC_WORKSPACE_WRAPPER": None,
                    "CARGO_BUILD_RUSTC": None, "RUSTUP_TOOLCHAIN": None,
                },
                "toolchain": toolchain,
                "stdout": _record(stdout), "stderr": _record(stderr), "outputs": outputs,
                "target": {"path": str(target), "bytes_before": 0, "bytes_after": 0},
            }
            build_events.append(event)
            builds[key] = {"role": role, "variant": variant, "features": validator.EXPECTED_FEATURES[key], "target": str(target), "command": command, "event_id": event["event_id"], "artifacts": artifacts}
        build_events_path = build_dir / "events-build.jsonl"
        build_events_path.write_text("\n".join(json.dumps(event, sort_keys=True) for event in build_events) + "\n", encoding="utf-8")
        event_raw = build_events_path.read_bytes()
        build_receipt = {
            "schema": runner.SCHEMA, "phase": "build", "status": "passed",
            "manifest": {"path": str(manifest_path), "sha256": _sha(manifest_path)[0]},
            "preparation": {"path": str(prep_path), "sha256": _sha(prep_path)[0]},
            "runner": _record(Path(runner.__file__)), "limits": {"target_preflight_bytes": validator.TARGET_PREFLIGHT_BYTES, "target_hard_bytes": validator.TARGET_HARD_BYTES, "memory_bytes": validator.MEMORY_BYTES},
            "claim_scope": validator.CLAIM_FREE["build"], "target_root": str(target_root),
            "event_journal": {"path": str(build_events_path), "count": len(build_events), "event_ids": [e["event_id"] for e in build_events], "sha256": hashlib.sha256(json.dumps(build_events, sort_keys=True, separators=(",", ":")).encode()).hexdigest(), "file_sha256": hashlib.sha256(event_raw).hexdigest(), "bytes": len(event_raw)},
            "toolchain": toolchain,
            "identities": identities, "builds": builds,
        }
        build_path = build_dir / "build-receipt.json"
        build_path.write_text(json.dumps(build_receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")

        reports = {}
        replacement_payload = b"ole-w5-replacement"
        metadata = "\n".join(["schema\t1", "owner_streams\t65", "owner_target\tole_w5_edit_target.bin", "owner_original_bytes\t15", f"owner_replacement_bytes\t{len(replacement_payload)}", "properties\tvalidated", "doc\trevision\tInsertion\t0\t24\tW5Oracle", "xls\tCells\t0\t0\t4242"]) + "\n"
        allocator_header = "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches"
        allocator_rows = [allocator_header]
        for route in validator.ALLOCATOR_ROUTES:
            retained = "129\t128" if route in {"doc_history", "xls_history"} else "0\t0"
            allocator_rows.append(f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{retained}")
        doc_proof = ["kind\tindex\ttext"]
        for kind in ("source", "edit"):
            for index in range(256):
                text = f"ole-w5 tracked paragraph {index:03}"
                if kind == "edit" and index == 0:
                    text = "ole-w5-tracked-revision " + text
                doc_proof.append(f"{kind}\t{index}\t{text}")
        xls_proof = ["kind\trow\tvalue_bits"]
        import struct
        for kind in ("source", "edit"):
            for row in range(256):
                value = 4242.0 if kind == "edit" and row == 0 else float(row)
                xls_proof.append(f"{kind}\t{row}\t{struct.unpack('<Q', struct.pack('<d', value))[0]}")
        xls_proof.append("opaque\tbytes\t98304")
        inventory_header = "kind\tname\tentry_type\tsid\tsid_left\tsid_right\tsid_child\tclsid\tstate_bits\tcreation_time\tmodified_time\tsize\tpayload_hex"
        inventory_rows = [inventory_header]
        entries = [("Root Entry", 5, 0, 0, "")] + [
            (f"ole_w5_sibling_{index:03}", 2, index + 1, len(f"opaque sibling payload {index:03}"), f"opaque sibling payload {index:03}".encode().hex())
            for index in range(64)
        ] + [("ole_w5_edit_target.bin", 2, 65, 15, b"ole-w5-original".hex())]
        for kind in ("source", "replacement"):
            for name, entry_type, sid, size, payload in entries:
                if kind == "replacement" and name == "ole_w5_edit_target.bin":
                    size, payload = len(replacement_payload), replacement_payload.hex()
                fields = (kind, name, str(entry_type), str(sid), "0", "0", "0", "", "0", "0", "0", str(size), payload)
                inventory_rows.append("\t".join(fields))
        history_proof = "kind\tpackages\tpatches\tforward\tinverse\ndoc\t129\t128\tvalidated\tvalidated\nxls\t129\t128\tvalidated\tvalidated\n"
        for key in runner.BUILD_KEYS:
            role, variant = key.split("/")
            directory = build_dir / "oracle" / key.replace("/", "-")
            files = {}
            values = {
                "ole-source.bin": b"source", "ole-noop.bin": b"source", "ole-replacement.bin": b"replacement",
                "properties.tsv": b"link_base\thttps://example.invalid/base/\nhyperlinks\t1\nhyperlink\t7\t11\t0\thttps://example.invalid/target\tbookmark\nopaque_property\t42\t128\n", "doc-source.bin": b"doc-source", "doc-edit.bin": b"doc-edit",
                "doc-proof.tsv": ("\n".join(doc_proof) + "\n").encode(), "xls-source.bin": b"xls-source", "xls-edit.bin": b"xls-edit",
                "xls-proof.tsv": ("\n".join(xls_proof) + "\n").encode(), "xls-opaque-source.bin": b"x" * 98304, "xls-opaque-edit.bin": b"x" * 98304,
                "doc-history-final.bin": b"doc-history", "xls-history-final.bin": b"xls-history", "history-proof.tsv": history_proof.encode(),
                "ole-inventory.tsv": ("\n".join(inventory_rows) + "\n").encode(), "oracle.tsv": metadata.encode(),
                "allocator.tsv": (b"status\tnative-uninstrumented\n" if variant == "normal" else ("\n".join(allocator_rows) + "\n").encode()),
            }
            for name, data in values.items():
                path = directory / name
                _write(path, data)
                files[name] = _record(path)
            reports[key] = {"role": role, "variant": variant, "directory": str(directory), "files": files, "event_id": f"oracle-{key.replace('/', '-') }"}

        timing_batches = {"clone": {"before": 1, "after": 1, "run": 1}}
        for route in validator.ROUTES:
            timing_batches[route] = {"before": 1, "after": 1, "run": 1}
        timing = {"clone": [], "routes": {}, "batches": timing_batches}
        for block in range(validator.BLOCKS):
            legs = []
            for index, role in enumerate(validator.ABBA_ORDER):
                rows = [{"size": size, "sample": sample, "nanoseconds": validator.MIN_SPAN_NS, "output_bytes": 1}
                        for size in ("small", "large") for sample in range(validator.SAMPLES)]
                legs.append({"role": role, "rows": rows, "event_id": f"run-clone-b{block:02d}-{index}-{role}"})
            timing["clone"].append({"block": block, "legs": legs})
        for route in validator.ROUTES:
            timing["routes"][route] = []
            for block in range(validator.BLOCKS):
                legs = []
                for index, role in enumerate(validator.ABBA_ORDER):
                    rows = [{"sample": sample, "nanoseconds": validator.MIN_SPAN_NS, "output_bytes": 1}
                            for sample in range(validator.SAMPLES)]
                    legs.append({"role": role, "rows": rows, "event_id": f"run-route-{route}-b{block:02d}-{index}-{role}"})
                timing["routes"][route].append({"block": block, "legs": legs})
        allocation = {"abba_order": list(validator.ABBA_ORDER), "elapsed_claimable": False, "routes": {}}
        for route in validator.ALLOCATOR_ROUTES:
            allocation["routes"][route] = {"legs": [{"role": role, "report": {"output_bytes": 1}, "event_id": f"run-alloc-{route}-{index}-{role}"} for index, role in enumerate(validator.ABBA_ORDER)]}
        measurements = {"schema": validator.MEASUREMENTS_SCHEMA, "status": "complete; coordinator review pending", "routes": list(validator.ROUTES), "abba_order": list(validator.ABBA_ORDER), "blocks": validator.BLOCKS, "warmup": validator.WARMUP, "samples": validator.SAMPLES, "minimum_span_ns": validator.MIN_SPAN_NS, "timing": timing, "allocation": allocation, "oracle": {"status": "preflight-validated", "files": reports, "claim_scope": validator.CLAIM_FREE["oracle"]}, "claim_scope": validator.CLAIM_FREE["measurements"]}
        measurements_path = build_dir / "measurements.json"
        measurements_path.write_text(json.dumps(measurements, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        prep_roles = {"before": "baseline", "after": "candidate"}
        def make_event(event_id: str, role: str, variant: str, argv: list[str], stdout_data: bytes) -> dict[str, object]:
            stdout = build_dir / "runs" / f"{event_id}.stdout"
            stderr = build_dir / "runs" / f"{event_id}.stderr"
            _write(stdout, stdout_data)
            _write(stderr, b"")
            target = target_root / f"{role}-{variant}"
            driver_root = Path(preparation_document[prep_roles[role]]["driver"]["root"])
            event = {
                "event_id": event_id, "phase": "abba", "role": role, "variant": variant,
                "argv": argv, "cwd": str(driver_root),
                "source_identity_sha256": identities[role]["identity_sha256"],
                "source_identity_before": identities[role], "source_identity_after": identities[role],
                "returncode": 0, "runner_error": None, "process_guard_violations": [],
                "post_exit_descendants": [], "post_exit_offenders": [],
                "environment": {
                    "CARGO_TARGET_DIR": str(target), "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0",
                    "LITCHI_MEM_MAX": "20G", "PATH": path_env, "CARGO": str(cargo_tool),
                    "RUSTC": str(rustc_tool), "RUSTC_WRAPPER": None, "RUSTC_WORKSPACE_WRAPPER": None,
                    "CARGO_BUILD_RUSTC": None, "RUSTUP_TOOLCHAIN": None,
                },
                "toolchain": toolchain,
                "stdout": _record(stdout), "stderr": _record(stderr), "outputs": [],
                "target": {"path": str(target), "bytes_before": 0, "bytes_after": 0},
            }
            return event

        abba_events_list = []
        for key in runner.BUILD_KEYS:
            role, variant = key.split("/")
            report = reports[key]
            oracle_binary = Path(builds[key]["artifacts"]["oracle"]["path"])
            directory = Path(report["directory"])
            manifest_lines = [f"ole_owner_oracle,{directory}"]
            for name in runner.ORACLE_FILES:
                record = report["files"][name]
                manifest_lines.append(f"oracle_output\t{name}\t{record['bytes']}\t{record['sha256']}")
            event = make_event(f"oracle-{key.replace('/', '-')}", role, variant, [str(oracle_binary), "--output", str(directory)], ("\n".join(manifest_lines) + "\n").encode())
            event["outputs"] = [{**report["files"][name], "exists": True} for name in runner.ORACLE_FILES]
            abba_events_list.append(event)

        for role in validator.ROLES:
            clone_binary = Path(builds[f"{role}/normal"]["artifacts"]["clone"]["path"])
            for attempt in (1,):
                event_id = f"run-cal-{role}-clone-{attempt:03d}"
                abba_events_list.append(make_event(event_id, role, "normal", [str(clone_binary), "--warmup", "1", "--samples", "1", "--batch", "1"], b"route,size,sample,nanoseconds,batch,archive_bytes\nclone_drop,small,0,250000000,1,1\nclone_drop,large,0,250000000,1,1\n"))

        for block in range(validator.BLOCKS):
            for index, role in enumerate(validator.ABBA_ORDER):
                event_id = f"run-clone-b{block:02d}-{index}-{role}"
                clone_binary = Path(builds[f"{role}/normal"]["artifacts"]["clone"]["path"])
                rows = [f"clone_drop,{size},{sample},{validator.MIN_SPAN_NS},1,1" for size in ("small", "large") for sample in range(validator.SAMPLES)]
                abba_events_list.append(make_event(event_id, role, "normal", [str(clone_binary), "--warmup", "2", "--samples", "3", "--batch", "1"], ("route,size,sample,nanoseconds,batch,archive_bytes\n" + "\n".join(rows) + "\n").encode()))
        for route in validator.ROUTES:
            for role in validator.ROLES:
                probe_binary = Path(builds[f"{role}/normal"]["artifacts"]["probe"]["path"])
                event_id = f"run-cal-{role}-{route}-001"
                abba_events_list.append(make_event(event_id, role, "normal", [str(probe_binary), "--route", route, "--warmup", "1", "--samples", "1", "--batch", "1"], f"route,sample,nanoseconds,output_bytes\n{route},0,250000000,1\n".encode()))
        for route in validator.ROUTES:
            for block in range(validator.BLOCKS):
                for index, role in enumerate(validator.ABBA_ORDER):
                    event_id = f"run-route-{route}-b{block:02d}-{index}-{role}"
                    probe_binary = Path(builds[f"{role}/normal"]["artifacts"]["probe"]["path"])
                    rows = [f"{route},{sample},{validator.MIN_SPAN_NS},1" for sample in range(validator.SAMPLES)]
                    abba_events_list.append(make_event(event_id, role, "normal", [str(probe_binary), "--route", route, "--warmup", "2", "--samples", "3", "--batch", "1"], ("route,sample,nanoseconds,output_bytes\n" + "\n".join(rows) + "\n").encode()))
        for route in validator.ALLOCATOR_ROUTES:
            for index, role in enumerate(validator.ABBA_ORDER):
                event_id = f"run-alloc-{route}-{index}-{role}"
                probe_binary = Path(builds[f"{role}/allocator"]["artifacts"]["probe"]["path"])
                allocator_header = "w5_alloc,phase,route,iteration,delta_calls,delta_bytes,live_calls,live_bytes,peak_calls,peak_bytes,peak_delta_calls,peak_delta_bytes,base_live_calls,base_live_bytes,base_peak_calls,base_peak_bytes"
                stdout = f"route,sample,nanoseconds,output_bytes\n{route},0,1,1\n{allocator_header}\nw5_alloc,full,{route},0,1,1,1,1,1,1,1,1,0,0,0,0\n"
                abba_events_list.append(make_event(event_id, role, "allocator", [str(probe_binary), "--route", route, "--warmup", "0", "--samples", "1", "--batch", "1"], stdout.encode()))
        abba_events = build_dir / "events-abba.jsonl"
        abba_events.write_text("".join(json.dumps(event, sort_keys=True) + "\n" for event in abba_events_list), encoding="utf-8")
        abba_raw = abba_events.read_bytes()
        abba_receipt = {"schema": runner.SCHEMA, "phase": "abba", "status": "passed", "manifest": {"path": str(manifest_path), "sha256": _sha(manifest_path)[0]}, "preparation": {"path": str(prep_path), "sha256": _sha(prep_path)[0]}, "runner": _record(Path(runner.__file__)), "claim_scope": "orchestration and source-bound evidence only; no performance acceptance claim", "build_receipt_sha256": _sha(build_path)[0], "event_journal": {"path": str(abba_events), "count": len(abba_events_list), "event_ids": [event["event_id"] for event in abba_events_list], "sha256": hashlib.sha256(json.dumps(abba_events_list, sort_keys=True, separators=(",", ":")).encode()).hexdigest(), "file_sha256": hashlib.sha256(abba_raw).hexdigest(), "bytes": len(abba_raw)}, "measurements": {"path": str(measurements_path), "sha256": _sha(measurements_path)[0], "bytes": _sha(measurements_path)[1]}}
        abba_path = build_dir / "abba-receipt.json"
        abba_path.write_text(json.dumps(abba_receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return {"manifest": manifest_path, "preparation": prep_path, "build": build_path, "abba": abba_path, "measurements": measurements_path}, runner.Path(runner.__file__) if False else Path(runner.__file__)

    def _validate_fixture(self, paths: dict[str, Path], runner_path: Path) -> dict[str, object]:
        pins = json.loads(paths["manifest"].read_text())["input_pins"]
        with mock.patch.multiple(
            preparation,
            REVIEWED_PROBE_SHA256=pins["reviewed_probe_sha256"],
            CANDIDATE_ADAPTER_SHA256=pins["candidate_adapter_sha256"],
            CLONE_PROBE_SHA256=pins["clone_probe_sha256"],
            ORACLE_SOURCE_SHA256=pins["oracle_source_sha256"],
            REVIEWED_LOCK_SHA256=pins["reviewed_lock_sha256"],
            BASELINE_LOCK_SHA256=pins["baseline_lock_sha256"],
            CANDIDATE_LOCK_SHA256=pins["candidate_lock_sha256"],
        ):
            with mock.patch.multiple(
                runner.preparation,
                REVIEWED_PROBE_SHA256=pins["reviewed_probe_sha256"],
                CANDIDATE_ADAPTER_SHA256=pins["candidate_adapter_sha256"],
                CLONE_PROBE_SHA256=pins["clone_probe_sha256"],
                ORACLE_SOURCE_SHA256=pins["oracle_source_sha256"],
                REVIEWED_LOCK_SHA256=pins["reviewed_lock_sha256"],
                BASELINE_LOCK_SHA256=pins["baseline_lock_sha256"],
                CANDIDATE_LOCK_SHA256=pins["candidate_lock_sha256"],
            ):
                with mock.patch.multiple(
                    validator,
                    REVIEWED_LOCK_SHA256=pins["reviewed_lock_sha256"],
                    BASELINE_LOCK_SHA256=pins["baseline_lock_sha256"],
                    CANDIDATE_LOCK_SHA256=pins["candidate_lock_sha256"],
                ):
                    return validator.validate(
                        manifest_path=paths["manifest"], preparation_path=paths["preparation"], build_path=paths["build"],
                        abba_path=paths["abba"], measurements_path=paths["measurements"], runner_path=runner_path,
                    )

    def test_build_free_preparation_to_validator_round_trip(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            result = self._validate_fixture(paths, runner_path)
            self.assertEqual(result["status"], "validated")

    def test_rejects_mutated_immutable_build_receipt_hash(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            receipt = json.loads(paths["abba"].read_text())
            receipt["build_receipt_sha256"] = "0" * 64
            paths["abba"].write_text(json.dumps(receipt))
            with self.assertRaisesRegex(validator.ValidationError, "build receipt hash"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_build_feature_or_log_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            receipt = json.loads(paths["build"].read_text())
            receipt["builds"]["after/normal"]["features"] = "legacy-owner-api"
            paths["build"].write_text(json.dumps(receipt))
            with self.assertRaisesRegex(validator.ValidationError, "features"):
                self._validate_fixture(paths, runner_path)

        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            log = Path(directory) / "run/commands/before-normal.stdout"
            log.write_bytes(b"tampered\n")
            with self.assertRaisesRegex(validator.ValidationError, "stdout SHA-256"):
                self._validate_fixture(paths, runner_path)

    def test_allocator_conservation_rejects_counter_mutation(self) -> None:
        header = "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches"
        rows = [
            "read\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1080\t110\t1100\t0\t0",
        ]
        for route in validator.ALLOCATOR_ROUTES[1:]:
            retained = "129\t128" if route in {"doc_history", "xls_history"} else "0\t0"
            rows.append(f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{retained}")
        with self.assertRaisesRegex(validator.ValidationError, "live-byte conservation"):
            validator._validate_allocator_rows([header, *rows], "fixture")

    def test_rejects_oracle_artifact_not_attested_by_invocation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            artifact = Path(directory) / "run/oracle/before-normal/ole-source.bin"
            artifact.write_bytes(b"arbitrary replacement bytes")
            with self.assertRaisesRegex(validator.ValidationError, "oracle.before/normal.ole-source.bin SHA-256"):
                self._validate_fixture(paths, runner_path)

    def test_independent_lock_check_rejects_registry_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            root = Path(directory)
            self._fixture(root)
            reviewed = root / "reviewed.Cargo.lock"
            candidate = root / "candidate.Cargo.lock"
            mutated = root / "candidate-mutated.Cargo.lock"
            mutated.write_text(
                candidate.read_text(encoding="utf-8").replace("b" * 64, "c" * 64),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(validator.ValidationError, "registry package identity"):
                validator._validate_lock_patch(reviewed, mutated, "candidate")

    def test_rejects_oracle_report_without_journal_event_link(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            measurements = json.loads(paths["measurements"].read_text())
            measurements["oracle"]["files"]["before/normal"]["event_id"] = "abba-fixture"
            paths["measurements"].write_text(json.dumps(measurements, sort_keys=True))
            receipt = json.loads(paths["abba"].read_text())
            receipt["measurements"]["sha256"] = _sha(paths["measurements"])[0]
            receipt["measurements"]["bytes"] = _sha(paths["measurements"])[1]
            paths["abba"].write_text(json.dumps(receipt, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "oracle.before/normal event ID differs"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_manifest_route_policy_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            manifest = json.loads(paths["manifest"].read_text())
            manifest["routes"][0] = "bogus"
            paths["manifest"].write_text(json.dumps(manifest, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "manifest policy validation failed"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_manifest_candidate_commit_policy_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            manifest = json.loads(paths["manifest"].read_text())
            manifest["candidate_commit"] = "deadbeef"
            paths["manifest"].write_text(json.dumps(manifest, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "manifest policy validation failed"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_build_artifact_outside_expected_target(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            build = json.loads(paths["build"].read_text())
            artifact = Path(build["builds"]["before/normal"]["artifacts"]["probe"]["path"])
            misplaced = Path(directory) / "misplaced-probe"
            _write(misplaced, artifact.read_bytes(), executable=True)
            build["builds"]["before/normal"]["artifacts"]["probe"] = _record(misplaced)
            paths["build"].write_text(json.dumps(build, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "artifact before/normal/probe path"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_cargo_argv_not_bound_to_toolchain_path(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            build = json.loads(paths["build"].read_text())
            alternate = Path(directory) / "alternate-cargo"
            _write(alternate, b"alternate cargo\n", executable=True)
            build["builds"]["before/normal"]["command"][0] = str(alternate)
            paths["build"].write_text(json.dumps(build, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "Cargo dispatch path"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_rustc_executable_or_version_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            build = json.loads(paths["build"].read_text())
            original = build["toolchain"]["rustc"]
            alternate = Path(directory) / "alternate-rustc"
            _write(alternate, Path(original["resolved_path"]).read_bytes(), executable=True)
            build["toolchain"]["rustc"] = {
                **_record(alternate),
                "resolved_path": str(alternate.resolve()),
                "version": original["version"],
            }
            paths["build"].write_text(json.dumps(build, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "build event before/normal environment differs"):
                self._validate_fixture(paths, runner_path)

        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            build = json.loads(paths["build"].read_text())
            build["toolchain"]["rustc"]["version"] = "a different rustc identity"
            paths["build"].write_text(json.dumps(build, sort_keys=True))
            with self.assertRaisesRegex(validator.ValidationError, "build event before/normal toolchain identity"):
                self._validate_fixture(paths, runner_path)

    def test_rejects_duplicate_allocator_route(self) -> None:
        header = "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches"
        row = "read\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t0\t0"
        lines = [header, row, row]
        for route in validator.ALLOCATOR_ROUTES[2:]:
            retained = "129\t128" if route in {"doc_history", "xls_history"} else "0\t0"
            lines.append(f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{retained}")
        with self.assertRaisesRegex(validator.ValidationError, "route order and uniqueness"):
            validator._validate_allocator_rows(lines, "duplicate-route")

    def test_oracle_metadata_matches_payload_and_rejects_size_drift(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, _ = self._fixture(Path(directory))
            measurements = json.loads(paths["measurements"].read_text())
            files = measurements["oracle"]["files"]["before/normal"]["files"]
            metadata_path = Path(files["oracle.tsv"]["path"])
            metadata = metadata_path.read_text()
            oracle_directory = metadata_path.parent
            validator._validate_oracle_proofs(files, "native-payload")
            runner.PairedRunner._oracle_metadata(oracle_directory)
            for changed in (
                metadata.replace("owner_replacement_bytes\t18", "owner_replacement_bytes\t19"),
                metadata.replace("ole_w5_edit_target.bin", "ole-w5_edit_target.bin"),
            ):
                with self.subTest(metadata=changed):
                    metadata_path.write_text(changed)
                    with self.assertRaisesRegex(validator.ValidationError, "metadata"):
                        validator._validate_oracle_proofs(files, "wrong-metadata")
                    with self.assertRaisesRegex(runner.RunnerError, "metadata"):
                        runner.PairedRunner._oracle_metadata(oracle_directory)
            metadata_path.write_text(metadata)
            inventory_path = Path(files["ole-inventory.tsv"]["path"])
            lines = inventory_path.read_text().splitlines()
            for index, line in enumerate(lines):
                if line.startswith("replacement\tole_w5_edit_target.bin\t"):
                    fields = line.split("\t")
                    fields[-2] = "19"
                    lines[index] = "\t".join(fields)
                    break
            else:
                self.fail("replacement target missing from oracle fixture")
            inventory_path.write_text("\n".join(lines) + "\n")
            with self.assertRaisesRegex(validator.ValidationError, "stream size"):
                validator._validate_oracle_proofs(files, "wrong-size")
            with self.assertRaisesRegex(runner.RunnerError, "stream size"):
                runner.PairedRunner._oracle_metadata(oracle_directory)

    def test_rejects_duplicate_doc_xls_indices_and_root_identity(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, _ = self._fixture(Path(directory))
            measurements = json.loads(paths["measurements"].read_text())
            files = measurements["oracle"]["files"]["before/normal"]["files"]
            doc_path = Path(files["doc-proof.tsv"]["path"])
            doc_lines = doc_path.read_text().splitlines()
            original_doc_lines = list(doc_lines)
            doc_lines[2] = f"source\t0\t{doc_lines[1].split(chr(9), 2)[2]}"
            doc_path.write_text("\n".join(doc_lines) + "\n")
            files["doc-proof.tsv"] = _record(doc_path)
            with self.assertRaisesRegex(validator.ValidationError, "DOC source indices"):
                validator._validate_oracle_proofs(files, "duplicate-doc")
            doc_path.write_text("\n".join(original_doc_lines) + "\n")
            files["doc-proof.tsv"] = _record(doc_path)

            xls_path = Path(files["xls-proof.tsv"]["path"])
            xls_lines = xls_path.read_text().splitlines()
            xls_lines[2] = f"source\t0\t{xls_lines[1].split(chr(9), 2)[2]}"
            xls_path.write_text("\n".join(xls_lines) + "\n")
            files["xls-proof.tsv"] = _record(xls_path)
            with self.assertRaisesRegex(validator.ValidationError, "XLS source indices"):
                validator._validate_oracle_proofs(files, "duplicate-xls")

            inventory_path = Path(files["ole-inventory.tsv"]["path"])
            inventory_lines = inventory_path.read_text().splitlines()
            inventory_lines[1] = inventory_lines[1].replace("\t5\t", "\t2\t", 1)
            inventory_path.write_text("\n".join(inventory_lines) + "\n")
            files["ole-inventory.tsv"] = _record(inventory_path)
            with self.assertRaisesRegex(validator.ValidationError, "Root Entry storage identity"):
                validator._validate_oracle_proofs(files, "wrong-root")


if __name__ == "__main__":
    unittest.main()
