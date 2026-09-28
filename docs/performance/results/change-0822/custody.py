"""Fail-closed input custody for the 0822 PPTX edit profile.

This packet deliberately separates packet-local probe inputs from the tracked
production and perf-baseline source.  The root agent owns Cargo, workload,
perf, symbol, and decode execution; these helpers only freeze and recheck
identities around those child processes.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TOOL = ROOT / "tools/perf-baseline"
TARGET = Path("/home/zhuhe/code/litchi-target-0822")
SCRATCH: Path | None = None
PLAN_PATH = P / "plan.json"
ORIGIN_PATH = P / "origin.json"
REPAIR_SOURCE_PATH = ROOT / "docs/performance/results/change-0820/repair/source.json"
PREVIOUS_PACKET = ROOT / "docs/performance/results/change-0821"
PREVIOUS_COMMIT = "353aa00a7da2795e6b4c28708138a103a143917e"

ROOT_INPUTS = {
    "Cargo.lock": P / "inputs/root-Cargo.lock",
    "rustfmt.toml": P / "inputs/rustfmt.toml",
}
TOOL_LOCK = P / "inputs/tool-Cargo.lock"

DRIVERS = (
    "custody.py",
    "quality.py",
    "build.py",
    "capture.py",
    "decode.py",
)
STATIC_PACKET_FILES = (
    "plan.json",
    "origin.json",
    "corpus-inputs.json",
    "input-inventory.json",
    "provenance.json",
    "architecture-inputs.json",
    "host.json",
    "toolchain.json",
    "lock-parity.json",
    "cgroup-limits.json",
    "root-inputs.json",
    "inputs/root-Cargo.lock",
    "inputs/rustfmt.toml",
    "inputs/tool-Cargo.lock",
    "profile-design-review.md",
    "source-review.md",
)
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md":
        "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md":
        "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json":
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}
PRODUCTION_PREFIXES = (
    "crates",
    "Cargo.toml",
    "clippy.toml",
    ".cargo/config.toml",
    "rust-toolchain.toml",
)
REFERENCE = ROOT / "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx"
REFERENCE_SOURCE = ROOT / "test-data/ooxml/pptx/shapes.pptx"
REFERENCE_SOURCE_COPY = ROOT / "docs/performance/results/change-0821/artifacts/real-002-pptx/source.pptx"
REFERENCE_REPORT = ROOT / (
    "docs/performance/results/change-0821/native/"
    "00-pptx_real_file_ordinary_save_lifecycle__default.json"
)


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path: Path | str, value: Any) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def artifact(path: Path | str) -> dict[str, object]:
    path = Path(path)
    assert path.is_file() and not path.is_symlink(), f"missing artifact {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def metadata_artifact(path: Path) -> dict[str, object]:
    """Return the repository-relative identity used by static manifests."""
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def _tracked(prefixes: tuple[str, ...]) -> list[str]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", *prefixes], cwd=ROOT
    )
    return [name for name in raw.decode().split("\0") if name]


def source() -> dict[str, object]:
    names = _tracked(PRODUCTION_PREFIXES)
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def tool_source() -> dict[str, str]:
    names = _tracked(("tools/perf-baseline",))
    assert names, "perf-baseline source census is empty"
    return {name: sha(ROOT / name) for name in names}


def probe_files() -> dict[str, str]:
    root = P / "probe-src"
    assert root.is_dir() and not root.is_symlink(), f"missing probe source {root}"
    result: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            assert not path.is_symlink(), f"probe source symlink: {path}"
            result[str(path.relative_to(root))] = sha(path)
    assert result, "probe source is empty"
    return result


def root_input_hashes() -> dict[str, str]:
    return {name: sha(path) for name, path in ROOT_INPUTS.items()}


def assert_root_inputs() -> dict[str, str]:
    declared = read(P / "root-inputs.json")
    actual = root_input_hashes()
    assert declared == actual, "root input manifest differs from packet copies"
    for name, copy in ROOT_INPUTS.items():
        source_path = ROOT / name
        assert source_path.is_file() and not source_path.is_symlink()
        assert copy.is_file() and not copy.is_symlink()
        assert sha(source_path) == actual[name], f"root input changed: {name}"
    return actual


def lock_identity() -> dict[str, dict[str, object]]:
    root_lock = ROOT / "Cargo.lock"
    tool_lock = TOOL / "Cargo.lock"
    assert root_lock.is_file() and tool_lock.is_file() and TOOL_LOCK.is_file()
    assert sha(TOOL_LOCK) == sha(tool_lock), "tool lock differs from packet copy"
    probe_lock = P / "probe-src/Cargo.lock"
    assert probe_lock.is_file() and not probe_lock.is_symlink()
    return {
        "root": artifact(root_lock),
        "tool": artifact(tool_lock),
        "packet_tool": artifact(TOOL_LOCK),
        "probe": artifact(probe_lock),
    }


def architecture_hashes() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert isinstance(declared, dict) and len(declared) == 35
    actual = {name: sha(ROOT / name) for name in declared}
    assert actual == declared, "normative architecture input changed"
    return actual


def assert_host() -> str:
    value = read(P / "host.json")
    assert value.get("affinity_selected") == [12]
    assert 12 in value.get("affinity_available", [])
    return sha(P / "host.json")


def assert_corpus_inputs() -> dict[str, object]:
    value = read(P / "corpus-inputs.json")
    assert value.get("schema") == "litchi.performance.0822.corpus-inputs.v1"
    expected = {
        "path": str(REFERENCE.relative_to(ROOT)),
        "bytes": 68284,
        "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
    }
    source_expected = {
        "path": str(REFERENCE_SOURCE.relative_to(ROOT)),
        "bytes": 68822,
        "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
    }
    reference = value.get("reference")
    source_value = value.get("source")
    assert isinstance(reference, dict) and isinstance(source_value, dict)
    assert reference["path"] == expected["path"]
    assert source_value["path"] == source_expected["path"]
    assert metadata_artifact(REFERENCE) == expected
    assert metadata_artifact(REFERENCE_SOURCE) == source_expected
    assert metadata_artifact(REFERENCE_SOURCE_COPY) == {
        "path": str(REFERENCE_SOURCE_COPY.relative_to(ROOT)),
        "bytes": 68822,
        "sha256": source_expected["sha256"],
    }
    return {"reference": expected, "source": source_expected}


def assert_provenance(corpus: dict[str, object] | None = None) -> dict[str, object]:
    value = read(P / "provenance.json")
    assert value.get("schema") == "litchi.performance.0822.provenance.v1"
    assert value.get("quality_reuse") == "docs/performance/results/change-0821/quality.json"
    reference = value.get("reference_archive")
    source_value = value.get("source_archive")
    assert isinstance(reference, dict) and isinstance(source_value, dict)
    assert metadata_artifact(REFERENCE) == reference
    assert metadata_artifact(REFERENCE_SOURCE) == source_value
    source_copy = value.get("source_copy")
    assert isinstance(source_copy, dict) and metadata_artifact(REFERENCE_SOURCE_COPY) == source_copy
    report = value.get("reference_report")
    assert isinstance(report, dict) and metadata_artifact(REFERENCE_REPORT) == report
    if corpus is not None:
        assert corpus["reference"] == reference and corpus["source"] == source_value
    return value


def assert_unrelated() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    assert actual == UNRELATED, "unrelated workspace file changed"
    return actual


def repair_source_witness() -> dict[str, object]:
    prior = read(REPAIR_SOURCE_PATH)
    expected_production = prior["production"]
    expected_tool = prior["tool"]
    current_production = source()
    current_tool = tool_source()
    assert len(current_production["files"]) == 9197
    assert len(current_tool) == 87
    assert current_production["files"] == expected_production["files"]
    assert current_tool == expected_tool
    return {
        "prior_revision": expected_production["revision"],
        "current_revision": current_production["revision"],
        "production_files": len(current_production["files"]),
        "tool_files": len(current_tool),
        "source": artifact(REPAIR_SOURCE_PATH),
    }


def packet_hashes() -> dict[str, str]:
    names = (*STATIC_PACKET_FILES, *DRIVERS)
    return {name: sha(P / name) for name in names}


def driver_hashes() -> dict[str, str]:
    return {name: sha(P / name) for name in DRIVERS}


def previous_seal_witness() -> dict[str, object]:
    seal_path = PREVIOUS_PACKET / "seal.json"
    committed = subprocess.check_output(
        ["git", "show", f"{PREVIOUS_COMMIT}:docs/performance/results/change-0821/seal.json"],
        cwd=ROOT,
    )
    seal = json.loads(committed.decode())
    assert seal.get("schema") == "litchi.performance.0821.seal.v1"
    files = seal.get("files")
    assert isinstance(files, dict) and len(files) >= 770
    for name, expected in files.items():
        assert not Path(name).is_absolute() and ".." not in Path(name).parts
        blob = subprocess.check_output(["git", "show", f"{PREVIOUS_COMMIT}:{name}"], cwd=ROOT)
        assert hashlib.sha256(blob).hexdigest() == expected, f"prior seal mismatch: {name}"
    assert seal_path.is_file() and sha(seal_path) == hashlib.sha256(committed).hexdigest()
    return {
        "schema": seal["schema"],
        "commit": PREVIOUS_COMMIT,
        "path": str(seal_path),
        "bytes": len(committed),
        "sha256": hashlib.sha256(committed).hexdigest(),
        "files": len(files),
    }


def check_no_overrides() -> None:
    for name in (
        "RUSTUP_TOOLCHAIN",
        "RUSTFLAGS",
        "RUSTDOCFLAGS",
        "CARGO_ENCODED_RUSTFLAGS",
        "CARGO_BUILD_RUSTFLAGS",
        "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS",
        "RUSTC_WRAPPER",
        "RUSTC_WORKSPACE_WRAPPER",
    ):
        assert not os.environ.get(name), f"unexpected inherited build override: {name}"


def stable(frozen: dict[str, object] | None = None) -> None:
    """Recheck all immutable inputs around a root-owned child process."""
    current = source()
    current_tool = tool_source()
    expected_source = None if frozen is None else frozen.get("source")
    expected_tool = None if frozen is None else frozen.get("tool")
    if expected_source is not None:
        assert current["files"] == expected_source["files"]
    else:
        repair_source_witness()
    if expected_tool is not None:
        assert current_tool == expected_tool
    else:
        assert len(current_tool) == 87
    assert_root_inputs()
    locks = lock_identity()
    if frozen is not None:
        assert locks == frozen["locks"]
    corpus = assert_corpus_inputs()
    assert_provenance(corpus)
    architecture = architecture_hashes()
    if frozen is not None:
        assert architecture == frozen["architecture"]
    host = assert_host()
    if frozen is not None:
        assert host == frozen["host"]
    unrelated = assert_unrelated()
    if frozen is not None:
        assert unrelated == frozen["unrelated"]
    if frozen is not None and "probe" in frozen:
        assert probe_files() == frozen["probe"]
    if frozen is not None:
        assert sha(PLAN_PATH) == frozen["plan"]
        assert sha(ORIGIN_PATH) == frozen["origin"]
        assert driver_hashes() == frozen["drivers"]
        assert packet_hashes() == frozen["static_packet"]


def freeze() -> dict[str, object]:
    """Return the complete immutable witness used by build/capture/decode."""
    assert read(PLAN_PATH)["base"] == PREVIOUS_COMMIT
    origin = read(ORIGIN_PATH)
    assert origin["base"] == PREVIOUS_COMMIT
    repair_source_witness()
    corpus = assert_corpus_inputs()
    provenance = assert_provenance(corpus)
    witness = {
        "schema": "litchi.performance.0822.frozen-inputs.v1",
        "source": source(),
        "tool": tool_source(),
        "probe": probe_files(),
        "root_inputs": assert_root_inputs(),
        "locks": lock_identity(),
        "architecture": architecture_hashes(),
        "corpus": corpus,
        "provenance": provenance,
        "host": assert_host(),
        "unrelated": assert_unrelated(),
        "previous_seal": previous_seal_witness(),
        "plan": sha(PLAN_PATH),
        "origin": sha(ORIGIN_PATH),
        "drivers": driver_hashes(),
        "static_packet": packet_hashes(),
    }
    stable(witness)
    return witness


def check_probe_report(path: Path, *, mode: str, samples: int, warmup: int) -> dict[str, object]:
    value = read(path)
    assert isinstance(value, dict)
    assert value.get("schema") == "litchi.performance.0822.pptx-edit-profile.v1"
    assert value.get("tool") == "pptx-edit-profile-0822"
    assert value.get("base_revision") == "353aa00a7d"
    assert value.get("mode") == mode
    def file_identity(raw: object, expected_path: Path, expected_bytes: int, expected_sha: str) -> None:
        assert isinstance(raw, dict)
        raw_path = raw.get("path")
        assert isinstance(raw_path, str)
        path = Path(raw_path)
        if not path.is_absolute():
            path = ROOT / path
        assert path.resolve() == expected_path.resolve()
        assert raw.get("bytes") == expected_bytes
        assert raw.get("sha256") == expected_sha

    file_identity(
        value.get("input"), ROOT / "test-data/ooxml/pptx/shapes.pptx", 68822,
        "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
    )
    file_identity(value.get("reference"), REFERENCE, 68284,
                  "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf")
    output = value.get("output")
    assert output == {
        "bytes": 68284,
        "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
    }
    assert value.get("marker") == "litchi-perf-0638-ordinary-save"
    assert value.get("target") == {"slide": 0, "shape": 0}
    assert value.get("target_text") == value["marker"]
    assert isinstance(value.get("full_text_sha256"), str)
    assert value.get("full_text_digest") == value["full_text_sha256"]
    assert value.get("slide_count") == 6
    assert value.get("samples_requested") == samples
    assert value.get("warmup") == warmup
    assert value.get("warmup_verified") is True
    assert value.get("all_verified") is True
    observed = value.get("samples")
    assert isinstance(observed, list) and len(observed) == samples
    elapsed = value.get("elapsed_ns")
    assert isinstance(elapsed, dict)
    assert elapsed.get("unit") == "ns"
    values = elapsed.get("samples")
    assert isinstance(values, list) and len(values) == samples
    assert elapsed.get("sample_order") == list(range(samples))
    verification_fields = {
        "all_verified", "input_hash_verified", "reference_hash_verified",
        "output_hash_verified", "output_size_verified", "output_bytes_verified",
        "reopened", "marker_verified", "target_verified",
        "full_text_digest_verified", "slide_count_verified",
    }
    for index, sample in enumerate(observed):
        assert isinstance(sample, dict)
        assert sample.get("index") == index
        assert isinstance(sample.get("elapsed_ns"), int) and sample["elapsed_ns"] > 0
        assert sample["elapsed_ns"] == values[index]
        assert sample.get("output") == output
        checks = sample.get("verification")
        assert isinstance(checks, dict) and set(checks) == verification_fields
        assert all(checks.values())
    return value
