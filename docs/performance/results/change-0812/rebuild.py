#!/usr/bin/env python3
"""Root-only reconstruction of the sealed 0811 frame-pointer probe.

This driver has one execution lane: it rebuilds the original 0811 probe into
the original temporary target, copies the resulting executable to ``fp``, and
compares the complete file with the sealed 0811 witness.  Only after an exact
byte/length match does it retain the scanner symbol and its bounded
disassembly.  It never runs the probe workload and never modifies the sealed
0811 packet.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / "change-0811"
TARGET = Path("/home/zhuhe/code/litchi-target-0811")
MANIFEST = OLD / "probe-src/Cargo.toml"
PROBE = MANIFEST.parent
OUT = P / "rebuild"
SCRIPT = P / "rebuild.py"
PLAN = P / "plan.md"
SYMBOL = "litchi_pptx::notes::codec::scan_processed_xml"
MANGLED_SCAN = re.compile(
    r"^_ZN11litchi_pptx5notes5codec18scan_processed_xml17h[0-9a-f]{16}E$"
)
SOURCE_PREFIXES = (
    "crates",
    "Cargo.toml",
    "clippy.toml",
    ".cargo/config.toml",
    "rust-toolchain.toml",
)
FORBIDDEN_COMPILER_ENV = (
    "RUSTUP_TOOLCHAIN",
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "CARGO_BUILD_RUSTFLAGS",
    "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS",
    "RUSTC_WRAPPER",
    "RUSTC_WORKSPACE_WRAPPER",
)


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str):
    return json.loads(Path(path).read_text())


def artifact(path: Path | str) -> dict[str, object]:
    path = Path(path)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def write_once(path: Path, value: object) -> None:
    assert not path.exists(), f"refusing to overwrite {path}"
    path.write_text(value if isinstance(value, str) else json.dumps(value, indent=2, sort_keys=True) + "\n")


def source_manifest() -> dict[str, object]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", *SOURCE_PREFIXES], cwd=ROOT
    )
    names = [name for name in raw.decode().split("\0") if name]
    assert len(names) == 9196, f"expected 9196 tracked production files, found {len(names)}"
    assert len(set(names)) == len(names)
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def probe_manifest() -> dict[str, str]:
    return {
        str(path.relative_to(PROBE)): sha(path)
        for path in sorted(PROBE.rglob("*"))
        if path.is_file()
    }


def command_receipt(
    command: list[str],
    *,
    env: dict[str, str] | None = None,
    capture: bool = True,
) -> dict[str, object]:
    started = time.time()
    result = subprocess.run(
        command,
        cwd=ROOT,
        env=env,
        capture_output=capture,
        text=True,
        check=False,
    )
    return {
        "command": command,
        "cwd": str(ROOT),
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "stdout": result.stdout if capture else "",
        "stderr": result.stderr if capture else "",
    }


def selected_environment(env: dict[str, str]) -> dict[str, object]:
    return {
        "CARGO_TARGET_DIR": env.get("CARGO_TARGET_DIR"),
        "CARGO_BUILD_JOBS": env.get("CARGO_BUILD_JOBS"),
        "CARGO_INCREMENTAL": env.get("CARGO_INCREMENTAL"),
        "RUSTFLAGS": env.get("RUSTFLAGS"),
        "inherited_compiler_environment": {
            name: os.environ.get(name) for name in FORBIDDEN_COMPILER_ENV
        },
    }


def verify_sealed_inputs() -> dict[str, object]:
    """Check the old seal before creating the new target or output packet."""
    seal_path = OLD / "seal.json"
    seal = read(seal_path)
    assert seal["schema"] == "litchi.performance.0811.seal.v1"
    checked = []
    for name, digest in sorted(seal["files"].items()):
        path = ROOT / name
        assert path.is_file(), name
        assert sha(path) == digest, name
        checked.append({"path": name, "sha256": digest})

    build_path = OLD / "build/build.json"
    source_path = OLD / "build/source.json"
    frozen_path = OLD / "build/frozen-inputs.json"
    probe_path = OLD / "build/probe.json"
    cleanup_path = OLD / "cleanup.json"
    build = read(build_path)
    source = read(source_path)
    frozen = read(frozen_path)
    probe = read(probe_path)
    cleanup = read(cleanup_path)

    assert build["schema"] == "litchi.performance.0811.build.v1"
    assert artifact(source_path) == build["source"]
    assert artifact(frozen_path) == build["frozen_inputs"]
    assert build["probe"] == probe
    assert build["architecture"] == frozen["architecture"]
    assert build["unrelated"] == frozen["unrelated"]
    assert build["root_inputs"] == frozen["root_inputs"]
    assert cleanup["schema"] == "litchi.performance.0811.cleanup.v1"
    assert cleanup["target"] == str(TARGET)
    assert cleanup["target_removed"] is True
    assert cleanup["binaries_verified_before_removal"] is True
    assert not TARGET.exists(), "0811 target must remain absent before reconstruction"

    cleanup_binaries = {
        row["path"]: row for row in cleanup["removed_binaries"]
    }
    assert set(build["binaries"]) == {"control", "profile", "fp"}
    for name, witness in build["binaries"].items():
        assert witness == cleanup_binaries[witness["path"]], name
        assert witness["bytes"] > 0 and len(witness["sha256"]) == 64
    assert build["binaries"]["fp"] == cleanup_binaries[str(TARGET / "fp")]

    # Verify the current root inputs, architecture identities, and unrelated
    # workspace files against the already sealed 0811 records.
    for name, digest in frozen["root_inputs"].items():
        assert sha(ROOT / name) == digest, name
    architecture_path = P / "architecture-inputs.json"
    assert read(architecture_path) == frozen["architecture"]
    for name, digest in frozen["architecture"].items():
        assert sha(ROOT / name) == digest, name
    for name, digest in frozen["unrelated"].items():
        assert sha(ROOT / name) == digest, name

    actual_probe = probe_manifest()
    assert len(actual_probe) == 6, actual_probe
    assert actual_probe == probe
    assert actual_probe == frozen["probe"]
    assert MANIFEST == OLD / "probe-src/Cargo.toml"

    return {
        "seal": artifact(seal_path),
        "sealed_files": len(checked),
        "build_manifest": artifact(build_path),
        "source_manifest": artifact(source_path),
        "frozen_inputs": artifact(frozen_path),
        "probe_manifest": artifact(probe_path),
        "cleanup": artifact(cleanup_path),
        "fp_binary_witness": build["binaries"]["fp"],
        "architecture_files": len(frozen["architecture"]),
        "probe_files": len(actual_probe),
        "source_files": len(source["files"]),
        "sealed_file_hashes": checked,
    }


def main() -> int:
    started = time.time()
    assert not OUT.exists(), f"refusing an existing output directory: {OUT}"
    assert not TARGET.exists(), f"refusing an existing target: {TARGET}"
    assert MANIFEST.is_file()
    assert SCRIPT.is_file()
    assert (P / "architecture-inputs.json").is_file()

    sealed = verify_sealed_inputs()
    expected_source = read(OLD / "build/source.json")
    source_before = source_manifest()
    assert source_before["files"] == expected_source["files"]
    assert len(source_before["files"]) == 9196
    actual_probe = probe_manifest()

    for name in FORBIDDEN_COMPILER_ENV:
        assert not os.environ.get(name), f"unexpected inherited build override: {name}"

    tool_commands = {
        "rustc": ["rustc", "--version", "--verbose"],
        "cargo": ["cargo", "--version"],
        "nm": ["nm", "--version"],
        "objdump": ["objdump", "--version"],
    }
    tools = {
        name: command_receipt(command)
        for name, command in tool_commands.items()
    }
    assert all(row["exit_code"] == 0 for row in tools.values())
    old_toolchain = read(OLD / "toolchain.json")
    old_by_command = {
        tuple(row["command"]): row for row in old_toolchain["commands"]
    }
    for name in ("rustc", "cargo"):
        old = old_by_command[tuple(tool_commands[name])]
        assert tools[name]["stdout"] == old["stdout"]
        assert tools[name]["stderr"] == old["stderr"]

    command = [
        "cargo",
        "build",
        "--offline",
        "--locked",
        "--release",
        "--manifest-path",
        str(MANIFEST),
        "--features",
        "capture-profile",
    ]
    build_env = os.environ.copy()
    build_env.update(
        {
            "CARGO_TARGET_DIR": str(TARGET),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0",
            "RUSTFLAGS": "-C force-frame-pointers=yes",
        }
    )
    prefreeze = {
        "schema": "litchi.performance.0812.rebuild-prefreeze.v1",
        "started": started,
        "source": source_before,
        "source_reference": artifact(OLD / "build/source.json"),
        "probe": actual_probe,
        "probe_reference": sealed["probe_manifest"],
        "architecture": read(OLD / "build/frozen-inputs.json")["architecture"],
        "unrelated_reference": read(OLD / "build/frozen-inputs.json")["unrelated"],
        "root_inputs_reference": read(OLD / "build/frozen-inputs.json")["root_inputs"],
        "sealed_0811": sealed,
        "driver": artifact(SCRIPT),
        "plan": artifact(PLAN) if PLAN.is_file() else None,
        "manifest": artifact(MANIFEST),
        "target": str(TARGET),
        "commands": {
            "tools": tool_commands,
            "build": command,
        },
        "tool_versions": tools,
        "environment": selected_environment(build_env),
        "contract": {
            "offline": True,
            "locked": True,
            "release": True,
            "features": ["capture-profile"],
            "rustflags": "-C force-frame-pointers=yes",
            "cargo_build_jobs": 2,
            "cargo_incremental": 0,
            "workload": False,
        },
    }
    OUT.mkdir()
    write_once(OUT / "source-before.json", source_before)
    prefreeze["source_before_artifact"] = artifact(OUT / "source-before.json")
    write_once(OUT / "prefreeze.json", prefreeze)

    log_path = OUT / "build.log"
    build_started = time.time()
    with log_path.open("w") as stream:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=build_env,
            stdout=stream,
            stderr=subprocess.STDOUT,
            check=False,
        )
    build_ended = time.time()
    source_after = source_manifest()
    write_once(OUT / "source-after.json", source_after)

    build_receipt: dict[str, object] = {
        "command": command,
        "cwd": str(ROOT),
        "started": build_started,
        "ended": build_ended,
        "exit_code": result.returncode,
        "environment": selected_environment(build_env),
        "log": artifact(log_path),
        "source_before": artifact(OUT / "source-before.json"),
        "source_after": artifact(OUT / "source-after.json"),
        "source_files_equal": source_after["files"] == source_before["files"],
    }
    command_rows: dict[str, object] = {
        "schema": "litchi.performance.0812.rebuild-commands.v1",
        "tool_versions": tools,
        "build": build_receipt,
    }

    status = "build-failed"
    binary = None
    comparison = None
    assembly = None
    nm_receipt = None
    objdump_receipt = None

    executable = TARGET / "release/namespace-uri-probe"
    destination = TARGET / "fp"
    if (
        result.returncode == 0
        and source_after["files"] == source_before["files"]
        and executable.is_file()
    ):
        assert not destination.exists()
        shutil.copy2(executable, destination)
        binary = artifact(destination)
        expected_binary = sealed["fp_binary_witness"]
        comparison = {
            "expected": expected_binary,
            "actual": binary,
            "bytes_equal": binary["bytes"] == expected_binary["bytes"],
            "sha256_equal": binary["sha256"] == expected_binary["sha256"],
            "exact": binary == expected_binary,
        }
        if comparison["exact"]:
            status = "binary-exact"
            nm_command = ["nm", "-S", "--defined-only", str(destination)]
            nm_started = time.time()
            nm_result = subprocess.run(
                nm_command,
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            nm_lines = [
                line
                for line in nm_result.stdout.splitlines()
                if line.split(maxsplit=3)[-1:] and MANGLED_SCAN.fullmatch(
                    line.split(maxsplit=3)[-1]
                )
            ]
            nm_receipt = {
                "command": nm_command,
                "cwd": str(ROOT),
                "started": nm_started,
                "ended": time.time(),
                "exit_code": nm_result.returncode,
                "stderr": nm_result.stderr,
                "matched_count": len(nm_lines),
            }
            if nm_result.returncode == 0 and len(nm_lines) == 1:
                matched_line = nm_lines[0]
                fields = matched_line.split(maxsplit=3)
                assert len(fields) == 4
                address, size, kind, mangled = fields
                assert kind.lower() == "t"
                address_int = int(address, 16)
                size_int = int(size, 16)
                assert size_int > 0
                write_once(OUT / "scan-symbol.txt", matched_line + "\n")
                objdump_command = [
                    "objdump",
                    "-d",
                    "--demangle",
                    "--no-show-raw-insn",
                    f"--start-address=0x{address_int:x}",
                    f"--stop-address=0x{address_int + size_int:x}",
                    str(destination),
                ]
                objdump_started = time.time()
                objdump_result = subprocess.run(
                    objdump_command,
                    cwd=ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )
                objdump_receipt = {
                    "command": objdump_command,
                    "cwd": str(ROOT),
                    "started": objdump_started,
                    "ended": time.time(),
                    "exit_code": objdump_result.returncode,
                    "stderr": objdump_result.stderr,
                }
                if objdump_result.returncode == 0 and not objdump_result.stderr:
                    write_once(
                        OUT / "scan-processed-xml-assembly.txt",
                        objdump_result.stdout,
                    )
                    assembly = {
                        "symbol": SYMBOL,
                        "mangled_symbol": mangled,
                        "address_hex": f"0x{address_int:x}",
                        "size_hex": f"0x{size_int:x}",
                        "nm_line": artifact(OUT / "scan-symbol.txt"),
                        "objdump": artifact(
                            OUT / "scan-processed-xml-assembly.txt"
                        ),
                    }
                else:
                    status = "objdump-failed"
            else:
                status = "symbol-not-found"
        else:
            status = "binary-mismatch"
    elif result.returncode == 0:
        status = "build-output-invalid"
    elif source_after["files"] != source_before["files"]:
        status = "source-changed-during-build"

    command_rows["nm"] = nm_receipt
    command_rows["objdump"] = objdump_receipt
    write_once(OUT / "commands.json", command_rows)
    receipt = {
        "schema": "litchi.performance.0812.rebuild.v1",
        "status": status,
        "scope": "Exact 0811 frame-pointer probe rebuild and post-build scanner disassembly only; no workload.",
        "started": started,
        "ended": time.time(),
        "target": str(TARGET),
        "manifest": str(MANIFEST),
        "prefreeze": artifact(OUT / "prefreeze.json"),
        "commands": artifact(OUT / "commands.json"),
        "build": build_receipt,
        "source": {
            "before": artifact(OUT / "source-before.json"),
            "after": artifact(OUT / "source-after.json"),
            "file_count": len(source_before["files"]),
            "hashes_equal": source_after["files"] == source_before["files"],
        },
        "binary": binary,
        "binary_comparison": comparison,
        "assembly": assembly,
        "historical_mapping_authorized": bool(
            status == "binary-exact" and assembly is not None
        ),
    }
    write_once(OUT / "receipt.json", receipt)

    if status != "binary-exact" or assembly is None:
        print(f"0812 rebuild retained failure receipt: {status}", flush=True)
        return 1
    print("0812 exact 0811 fp rebuild and scanner assembly retained", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
