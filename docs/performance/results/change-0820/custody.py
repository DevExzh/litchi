"""Input and source custody for the 0820 durability-attribution packet.

The root agent owns all Cargo, exporter, and benchmark execution.  These
helpers make every driver re-check the same production, harness, corpus,
packet, and host identities before and after each child process.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TOOL = ROOT / "tools/perf-baseline"

ORIGIN_PATH = P / "origin.json"
ORIGIN = json.loads(ORIGIN_PATH.read_text()) if ORIGIN_PATH.is_file() else {}
TARGET = Path(ORIGIN.get("target", "/home/zhuhe/code/litchi-target-0820"))
SCRATCH = Path(ORIGIN.get("scratch", "/home/zhuhe/code/litchi-fs-0820"))

ROOT_INPUTS = {
    "Cargo.lock": P / "inputs/root-Cargo.lock",
    "rustfmt.toml": P / "inputs/rustfmt.toml",
}
TOOL_LOCK = P / "inputs/tool-Cargo.lock"

# The packet is deliberately explicit.  A frozen run must fail if a review or
# input witness is missing instead of silently falling back to a live file.
DRIVERS = ("custody.py", "build.py", "capture.py", "admission.py", "quality.py")
FROZEN_PACKET_FILES = (
    "plan.json",
    *DRIVERS,
    "architecture-inputs.json",
    "corpus-inputs.json",
    "input-inventory.json",
    "provenance.json",
    "host.json",
    "origin.json",
    "toolchain.json",
    "lock-parity.json",
    "cgroup-limits.json",
    "root-inputs.json",
    "inputs/root-Cargo.lock",
    "inputs/rustfmt.toml",
    "inputs/tool-Cargo.lock",
    "protocol-review.md",
    "source-review.md",
    "artifact_audit.py",
    "preservation.py",
)

UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md":
        "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md":
        "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json":
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str):
    return json.loads(Path(path).read_text())


def write(path: Path | str, value) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def artifact(path: Path | str) -> dict[str, object]:
    path = Path(path)
    assert path.is_file() and not path.is_symlink(), f"missing artifact {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def owned_directory(path: Path, marker_name: str, marker_text: str) -> dict[str, object]:
    """Create an owned directory, refusing a pre-existing foreign root."""
    marker = path / marker_name
    if path.exists():
        assert path.is_dir() and not path.is_symlink(), f"foreign owned directory {path}"
        assert marker.is_file() and marker.read_text() == marker_text, (
            f"refusing pre-existing unowned directory {path}"
        )
    else:
        path.mkdir(parents=True)
        marker.write_text(marker_text)
    return artifact(marker)


def assert_origin() -> dict[str, object]:
    origin = read(ORIGIN_PATH)
    assert origin["schema"] == "litchi.performance.0820.origin.v1"
    assert origin["base"] == read(P / "plan.json")["base"]
    assert origin["production_changed"] is False
    assert origin["tool_changed"] is False
    assert origin["runtime_harness_changed"] is False
    assert origin["target"] == str(TARGET)
    assert origin["scratch"] == str(SCRATCH)
    assert origin["tool_allowlist"] == []
    return origin


def root_input_hashes() -> dict[str, str]:
    return {name: sha(path) for name, path in ROOT_INPUTS.items()}


def assert_root_inputs() -> dict[str, str]:
    frozen = root_input_hashes()
    declared = read(P / "root-inputs.json")
    assert declared == frozen, "root input manifest differs from packet copies"
    for name, copy in ROOT_INPUTS.items():
        source = ROOT / name
        assert source.is_file(), f"missing root input {source}"
        assert sha(source) == frozen[name], f"root input changed: {name}"
        assert copy.is_file(), f"missing packet input {copy}"
    return frozen


def lock_identity() -> dict[str, dict[str, object]]:
    live_root = ROOT / "Cargo.lock"
    live_tool = TOOL / "Cargo.lock"
    assert live_root.is_file(), f"missing root lock {live_root}"
    assert live_tool.is_file(), f"missing tool lock {live_tool}"
    assert TOOL_LOCK.is_file(), f"missing packet tool lock {TOOL_LOCK}"
    assert sha(TOOL_LOCK) == sha(live_tool), "tool lock differs from packet copy"
    assert TOOL_LOCK.stat().st_size == live_tool.stat().st_size
    return {
        "root": {"path": str(live_root), "bytes": live_root.stat().st_size, "sha256": sha(live_root)},
        "tool": {"path": str(live_tool), "bytes": live_tool.stat().st_size, "sha256": sha(live_tool)},
        "packet_tool": {"path": str(TOOL_LOCK), "bytes": TOOL_LOCK.stat().st_size, "sha256": sha(TOOL_LOCK)},
    }


def assert_corpus_inputs() -> dict[str, dict[str, object]]:
    declared = read(P / "corpus-inputs.json")
    assert set(declared) == {
        "test-data/ooxml/docx/documentProperties.docx",
        "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
        "test-data/ooxml/pptx/shapes.pptx",
    }, "corpus input set changed"
    actual = {}
    for name, expected in declared.items():
        path = ROOT / name
        assert path.is_file() and not path.is_symlink(), f"missing corpus input {path}"
        observed = {"bytes": path.stat().st_size, "sha256": sha(path)}
        assert observed == expected, f"corpus input changed: {name}"
        actual[name] = observed
    return actual


def assert_provenance(corpus: dict[str, dict[str, object]] | None = None) -> dict[str, object]:
    """Verify both the packet witness and its live checked-in reference."""
    provenance = read(P / "provenance.json")
    assert provenance["schema"] == "litchi.performance.0820.provenance.v1"
    reference = provenance["reference"]
    reference_path = ROOT / reference["path"]
    assert reference_path.is_file() and not reference_path.is_symlink()
    assert sha(reference_path) == reference["sha256"], "provenance reference changed"
    assert isinstance(provenance["corpus"], dict)
    if corpus is not None:
        assert provenance["corpus"] == corpus, "provenance corpus differs from corpus manifest"
    return provenance


def tracked_source_names() -> list[str]:
    raw = subprocess.check_output(
        [
            "git",
            "ls-files",
            "-z",
            "--",
            "crates",
            "Cargo.toml",
            "clippy.toml",
            ".cargo/config.toml",
            "rust-toolchain.toml",
        ],
        cwd=ROOT,
    )
    return [name for name in raw.decode().split("\0") if name]


def source() -> dict[str, object]:
    names = tracked_source_names()
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def tracked_tool_names() -> list[str]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", "tools/perf-baseline"], cwd=ROOT
    )
    return [name for name in raw.decode().split("\0") if name]


def tool_source() -> dict[str, str]:
    names = tracked_tool_names()
    assert names, "perf-baseline has no tracked source files"
    return {name: sha(ROOT / name) for name in names}


def unchanged(frozen: dict[str, object]) -> None:
    production = frozen.get("production", frozen)
    tool = frozen.get("tool")
    assert isinstance(production, dict), "production source witness is malformed"
    assert source() == production, "production source changed"
    if tool is not None:
        assert isinstance(tool, dict), "tool source witness is malformed"
        assert tool_source() == tool, "perf-baseline tool source changed"


def architecture_hashes() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert isinstance(declared, dict) and len(declared) == 35, (
        f"expected 35 architecture inputs, found {len(declared)}"
    )
    actual = {name: sha(ROOT / name) for name in declared}
    assert actual == declared, "architecture input changed after packet admission"
    return actual


def host_hash() -> str:
    return sha(P / "host.json")


def assert_host(expected: str | None = None) -> str:
    actual = host_hash()
    if expected is not None:
        assert actual == expected, "host descriptor changed after admission"
    return actual


def packet_hashes() -> dict[str, str]:
    result = {}
    for name in FROZEN_PACKET_FILES:
        path = P / name
        assert path.is_file() and not path.is_symlink(), f"missing packet input {path}"
        result[name] = sha(path)
    return result


def driver_hashes() -> dict[str, str]:
    return {name: sha(P / name) for name in DRIVERS}


def assert_unrelated() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    assert actual == UNRELATED, "an unrelated workspace file changed"
    return actual


def plan_cases(plan: dict[str, object]) -> list[dict[str, object]]:
    cases = plan["cases"]
    assert isinstance(cases, list) and len(cases) == 24
    required = {"id", "case", "format", "phase", "input", "policy"}
    for case in cases:
        assert isinstance(case, dict) and set(case) == required, f"invalid plan case {case!r}"
        assert case["format"] in {"docx", "xlsx", "pptx"}
        assert case["phase"] in {"lifecycle", "atomic_publish"}
        assert case["policy"] in {"default", "full", "file-only", "no-sync"}
        assert case["case"] == f"{case['format']}_real_file_ordinary_save_{case['phase']}"
        assert case["id"] == f"{case['case']}__{case['policy']}"
    assert len({case["id"] for case in cases}) == 24
    assert {
        (case["format"], case["phase"], case["policy"])
        for case in cases
    } == {
        (format_name, phase, policy)
        for format_name in ("docx", "xlsx", "pptx")
        for phase in ("lifecycle", "atomic_publish")
        for policy in ("default", "full", "file-only", "no-sync")
    }
    return cases


def ordered_cases(
    plan: dict[str, object], lane: str, block: int
) -> list[dict[str, object]]:
    """Expand one block's six-group order and rotated policy order.

    The plan lists each (format, phase, policy) row once.  Blocks choose the
    six format/phase groups with ``forward``/``reverse`` and rotate the four
    durability policies inside each group.  Keeping this expansion here makes
    capture and offline readers agree on the exact acquisition order.
    """
    cases = plan_cases(plan)
    lane_plan = plan["lanes"][lane]
    orders = lane_plan["orders"]
    policy_orders = lane_plan["policy_orders"]
    assert 0 <= block < len(orders)
    assert len(orders) == len(policy_orders)
    assert orders[block] in {"forward", "reverse"}
    policies = tuple(plan["policies"])
    policy_order = tuple(policy_orders[block])
    assert len(policy_order) == len(policies)
    assert set(policy_order) == set(policies)

    groups: dict[tuple[str, str], dict[str, dict[str, object]]] = {}
    group_order: list[tuple[str, str]] = []
    for case in cases:
        key = (case["format"], case["phase"])
        if key not in groups:
            groups[key] = {}
            group_order.append(key)
        policy = case["policy"]
        assert policy not in groups[key]
        groups[key][policy] = case
    assert len(groups) == 6
    assert all(set(group) == set(policies) for group in groups.values())
    if orders[block] == "reverse":
        group_order.reverse()
    result = [groups[key][policy] for key in group_order for policy in policy_order]
    assert len(result) == len(cases) == 24
    assert {case["id"] for case in result} == {case["id"] for case in cases}
    return result


def check_stable(
    frozen: dict[str, object],
    root_inputs: dict[str, str],
    locks: dict[str, dict[str, object]],
    architecture: dict[str, str],
    corpus: dict[str, dict[str, object]],
    host: str,
    packet: dict[str, str],
    drivers: dict[str, str],
    unrelated: dict[str, str],
) -> None:
    unchanged(frozen)
    assert assert_root_inputs() == root_inputs
    assert lock_identity() == locks
    assert architecture_hashes() == architecture
    assert assert_corpus_inputs() == corpus
    assert_provenance(corpus)
    assert assert_host(host) == host
    assert packet_hashes() == packet
    assert driver_hashes() == drivers
    assert assert_unrelated() == unrelated
