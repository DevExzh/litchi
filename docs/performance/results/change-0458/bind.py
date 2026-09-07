#!/usr/bin/env python3
"""Freeze a successful build's copied executables and capture inputs."""
import json
from pathlib import Path
import shutil

import capture

ROOT = Path(__file__).resolve().parent
TASK = Path("/tmp/litchi-goal-0458")


def main():
    build_path = ROOT / "checks/build.json"
    build = json.loads(build_path.read_text())
    assert build["status"] == "pass" and build["source_unchanged"] is True
    assert build["source_before"] == build["source_after"]
    assert not (ROOT / "binary-binding.json").exists()
    assert not (ROOT / "protocol.json").exists()
    TASK.mkdir()
    binaries = {}
    for mode, name in [("normal", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")]:
        source = capture.REPO / "tools/perf-baseline/target/release" / name
        target = TASK / name
        shutil.copy2(source, target)
        assert capture.sha(source) == capture.sha(target)
        binaries[mode] = {"path": str(target), "source": str(source.relative_to(capture.REPO)), "sha256": capture.sha(target), "bytes": target.stat().st_size}
    binding = {"schema": "litchi-0458-binary-binding-v1", "change": 458, "revision": build["revision"],
               "build_receipt": capture.artifact(build_path), "source_manifest": build["source_after"], "binaries": binaries}
    capture.write(ROOT / "binary-binding.json", binding)
    order = []
    for repeat in ["R1", "R2"]:
        modes = ["normal", "allocator"] if repeat == "R1" else ["allocator", "normal"]
        shapes = ["tiny", "medium", "large"] if repeat == "R1" else ["large", "medium", "tiny"]
        scopes = ["lifecycle", "phases"] if repeat == "R1" else ["phases", "lifecycle"]
        for instrumentation in modes:
            for shape in shapes:
                for scope in scopes:
                    order.append({"id": f"{repeat}-{instrumentation}-{shape}-{scope}", "repeat": repeat,
                                  "instrumentation": instrumentation, "shape": shape, "scope": scope})
    argv = ["{binary}", "odp-append-attribution", "--mode", "{scope}", "--shape", "{shape}",
            "--warmup", "3", "--samples", "30", "--repeat", "{repeat}", "--output", "{report}"]
    diagnostic_argv = ["100" if part == "30" else part for part in argv]
    protocol = {"schema": "litchi-0458-protocol-v1", "change": 458, "revision": build["revision"],
                "cpu": 2, "workers": 1, "samples": 30, "warmup": 3, "reports": 24, "retained_samples": 720,
                "order": order, "argv": argv, "diagnostic_argv": diagnostic_argv,
                "environment": {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"},
                "binary_binding_sha256": capture.sha(ROOT / "binary-binding.json"),
                "capture_sha256": capture.sha(ROOT / "capture.py"), "profile_sha256": capture.sha(ROOT / "profile.py"),
                "oracle_sha256": capture.sha(ROOT / "oracle.py"), "prior_control_binding_sha256": capture.sha(ROOT / "prior-control-bindings.json"),
                "host_sha256": capture.sha(ROOT / "host.json"),
                "claims": "public API phase diagnostic and same-binary instrumentation overhead; no production optimization, bounded memory, physical I/O, or worker scaling claim"}
    capture.write(ROOT / "protocol.json", protocol)
    print(json.dumps({"status": "bound", "binaries": binaries, "protocol_sha256": capture.sha(ROOT / "protocol.json")}, indent=2))


if __name__ == "__main__":
    main()
