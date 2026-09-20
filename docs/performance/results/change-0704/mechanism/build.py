#!/usr/bin/env python3
"""Build the temporary 0704 MCE trace binary and restore the codec exactly."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import time
from pathlib import Path

P = Path(__file__).resolve().parent
PACKET = P.parent
ROOT = P.parents[4]
TARGET = ROOT.parent / "litchi-target-0704"
BIN = ROOT.parent / "litchi-0704-bin"
CODEC = ROOT / "crates" / "litchi-ooxml-common" / "src" / "mce" / "codec.rs"
SOURCE_CENSUS = PACKET / "source-census-candidate.json"
PATCH = P / "0704-mce-trace.patch"
MANIFEST = P / "probe" / "Cargo.toml"
BINARY = BIN / "probe0704mce"


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def map_digest(mapping: dict[str, str]) -> str:
    digest = hashlib.sha256()
    for name, value in sorted(mapping.items()):
        digest.update(name.encode())
        digest.update(b"\0")
        digest.update(bytes.fromhex(value))
    return digest.hexdigest()


def candidate_source_map() -> dict[str, str]:
    census = json.loads(SOURCE_CENSUS.read_text())
    expected = census["source_sha256"]
    if census["source_file_count"] != 603 or len(expected) != 603:
        raise AssertionError("candidate source census is not the expected 603-file map")
    current: dict[str, str] = {}
    for owner in ("litchi-pptx", "litchi-ooxml-common", "litchi-opc"):
        for path in (ROOT / "crates" / owner).rglob("*.rs"):
            current[str(path.relative_to(ROOT))] = sha(path)
    if current != expected:
        missing = sorted(set(expected) - set(current))
        added = sorted(set(current) - set(expected))
        changed = sorted(name for name in set(expected) & set(current) if expected[name] != current[name])
        raise AssertionError(
            f"candidate source census mismatch: missing={missing[:3]} "
            f"added={added[:3]} changed={changed[:3]}"
        )
    return dict(sorted(current.items()))


def production_source_map() -> dict[str, str]:
    result: dict[str, str] = {}
    for owner in ("litchi-pptx", "litchi-ooxml-common", "litchi-opc"):
        for path in (ROOT / "crates" / owner).rglob("*.rs"):
            result[str(path.relative_to(ROOT))] = sha(path)
    return dict(sorted(result.items()))


def main() -> None:
    source_map = candidate_source_map()
    original = CODEC.read_bytes()
    original_sha = sha_bytes(original)
    expected_codec = source_map["crates/litchi-ooxml-common/src/mce/codec.rs"]
    if original_sha != expected_codec:
        raise AssertionError(f"codec source is not the candidate census: {original_sha}")
    if original_sha != "a5b5b0aca3ec5a392bc7ae1ea6ca482653bd0a72cb9bdd4c30b8ea9faff87bee":
        raise AssertionError("unexpected production codec hash")

    workspace_lock = (ROOT / "Cargo.lock").read_bytes()
    if not PATCH.is_file() or not MANIFEST.is_file() or not (P / "probe" / "Cargo.lock").is_file():
        raise AssertionError("standalone mechanism probe or patch is incomplete")

    env = os.environ | {"RUSTFLAGS": "-D warnings"}
    steps: list[dict[str, object]] = []

    def run(name: str, command: list[str]) -> None:
        log = P / f"{name}.log"
        started = time.monotonic()
        with log.open("w") as output:
            result = subprocess.run(command, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT)
        steps.append(
            {
                "name": name,
                "command": command,
                "exit_code": result.returncode,
                "seconds": time.monotonic() - started,
                "log_sha256": sha(log),
            }
        )
        if result.returncode:
            raise RuntimeError(log.read_text())

    receipt: dict[str, object] = {
        "schema": "litchi-0704-mce-mechanism-build-v1",
        "diagnostic_only": True,
        "performance_claim": "none",
        "baseline_head": json.loads((PACKET / "baseline.json").read_text())["baseline_head"],
        "candidate_source_file_count": len(source_map),
        "candidate_source_sha256": source_map,
        "candidate_source_map_sha256": map_digest(source_map),
        "candidate_census_sha256": sha(SOURCE_CENSUS),
        "original_codec_sha256": original_sha,
        "patch_sha256": sha(PATCH),
        "probe_sha256": {
            str(path.relative_to(P)): sha(path)
            for path in sorted((P / "probe").rglob("*"))
            if path.is_file()
        },
        "environment": {"RUSTFLAGS": env["RUSTFLAGS"]},
        "steps": steps,
    }
    try:
        run("patch-check", ["git", "apply", "--check", str(PATCH)])
        run("patch-apply", ["git", "apply", str(PATCH)])
        receipt["instrumented_codec_sha256"] = sha(CODEC)
        instrumented_source = production_source_map()
        receipt["instrumented_source_sha256"] = instrumented_source
        receipt["instrumented_source_map_sha256"] = map_digest(instrumented_source)
        run(
            "build-trace",
            [
                "cargo",
                "build",
                "--release",
                "--locked",
                "--offline",
                "--manifest-path",
                str(MANIFEST),
                "--target-dir",
                str(TARGET),
                "-j",
                "2",
            ],
        )
        built = TARGET / "release" / "probe0704mce"
        if not built.is_file():
            raise AssertionError(f"Cargo completed without {built}")
        BIN.mkdir(parents=True, exist_ok=True)
        shutil.copy2(built, BINARY)
        receipt["binary"] = str(BINARY)
        receipt["binary_sha256"] = sha(BINARY)
        receipt["status"] = "completed"
    finally:
        CODEC.write_bytes(original)
        receipt["restored_codec_sha256"] = sha(CODEC)
        restored_source = candidate_source_map()
        receipt["postrestore_candidate_source_sha256"] = restored_source
        receipt["postrestore_candidate_source_map_sha256"] = map_digest(restored_source)
        receipt["workspace_lock_sha256"] = sha(ROOT / "Cargo.lock")
        receipt["candidate_source_restored"] = restored_source == source_map
        receipt["steps"] = steps
        (P / "build.json").write_text(json.dumps(receipt, indent=2) + "\n")
        if (ROOT / "Cargo.lock").read_bytes() != workspace_lock:
            raise AssertionError("workspace Cargo.lock changed during mechanism build")
        if receipt["restored_codec_sha256"] != original_sha:
            raise AssertionError("codec was not restored byte-for-byte")

    print("PASS: mechanism binary built; candidate census and production codec restored", flush=True)


if __name__ == "__main__":
    main()
