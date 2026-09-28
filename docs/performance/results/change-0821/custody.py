"""Input and source custody for the 0821 durability-attribution packet.

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
PREVIOUS_PACKET = ROOT / "docs/performance/results/change-0820"
REPAIR_PACKET = PREVIOUS_PACKET / "repair"
REPAIR_SOURCE_PATH = REPAIR_PACKET / "source.json"
REPAIR_QUALITY_PATH = REPAIR_PACKET / "quality.json"
REPAIR_TEST_SUMMARY_PATH = REPAIR_PACKET / "test-summary.json"
REPAIR_INPUTS_PATH = REPAIR_PACKET / "inputs.json"
PREVIOUS_SEAL_PATH = PREVIOUS_PACKET / "seal.json"

ORIGIN_PATH = P / "origin.json"
ORIGIN = json.loads(ORIGIN_PATH.read_text()) if ORIGIN_PATH.is_file() else {}
TARGET = Path(ORIGIN.get("target", "/home/zhuhe/code/litchi-target-0821"))
SCRATCH = Path(ORIGIN.get("scratch", "/home/zhuhe/code/litchi-fs-0821"))

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
    assert origin["schema"] == "litchi.performance.0821.origin.v1"
    assert origin["base"] == read(P / "plan.json")["base"]
    assert origin["production_changed"] is False
    assert origin["tool_changed"] is False
    assert origin["runtime_harness_changed"] is False
    assert origin["target"] == str(TARGET)
    assert origin["scratch"] == str(SCRATCH)
    assert origin["tool_allowlist"] == []
    reuse = origin.get("quality_reuse")
    assert isinstance(reuse, dict)
    assert reuse.get("schema") == "litchi.performance.0821.quality-reuse.v1"
    assert reuse.get("cargo_commands_executed") is False
    assert reuse == read(P / "plan.json")["quality_reuse"]
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
    assert provenance["schema"] == "litchi.performance.0821.provenance.v1"
    assert provenance.get("quality_reuse") == "docs/performance/results/change-0820/repair/quality.json"
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


def _sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def repair_source_witness() -> dict[str, object]:
    """Require current source bytes to equal the committed 0820 repair source.

    The repair packet recorded the source hashes after the allocator test fix.
    The 0821 base advances the revision identifier, while every tracked
    production and perf-baseline file must remain byte-identical.
    """
    prior = read(REPAIR_SOURCE_PATH)
    assert isinstance(prior, dict)
    expected_production = prior.get("production")
    expected_tool = prior.get("tool")
    assert isinstance(expected_production, dict)
    assert isinstance(expected_tool, dict)
    current_production = source()
    current_tool = tool_source()
    assert current_production["files"] == expected_production["files"], (
        "current production source bytes differ from the 0820 repair source"
    )
    assert current_tool == expected_tool, (
        "current perf-baseline source bytes differ from the 0820 repair source"
    )
    assert current_production["revision"] == read(ORIGIN_PATH)["base"]
    return {
        "prior_revision": expected_production.get("revision"),
        "current_revision": current_production["revision"],
        "production_files": len(current_production["files"]),
        "tool_files": len(current_tool),
        "source": artifact(REPAIR_SOURCE_PATH),
    }


def _descriptor(value: object, label: str, *, expected: Path | None = None) -> dict[str, object]:
    assert isinstance(value, dict), f"{label} descriptor is malformed"
    raw_path = value.get("path")
    assert isinstance(raw_path, str) and raw_path
    path = Path(raw_path)
    if expected is not None:
        assert path.resolve() == expected.resolve(), f"{label} path changed"
    actual = artifact(path)
    assert actual == {
        "path": str(path),
        "bytes": value.get("bytes"),
        "sha256": value.get("sha256"),
    }, f"{label} descriptor changed"
    return actual


def previous_seal(*, require_worktree: bool = False) -> dict[str, object]:
    """Replay every 0820 sealed blob from its committed base.

    ``require_worktree`` is used before the 0821 documentation indexes are
    edited.  Later stability checks use the base commit so legitimate 0821
    index updates do not invalidate the prior packet's custody evidence.
    """
    origin = read(ORIGIN_PATH)
    base = origin["base"]
    seal = read(PREVIOUS_SEAL_PATH)
    assert seal.get("schema") == "litchi.performance.0820.seal.v1"
    files = seal.get("files")
    assert isinstance(files, dict) and len(files) >= 80
    for name, expected in files.items():
        relative = Path(name)
        assert not relative.is_absolute() and ".." not in relative.parts
        blob = subprocess.check_output(["git", "show", f"{base}:{name}"], cwd=ROOT)
        assert _sha_bytes(blob) == expected, f"0820 sealed base blob changed: {name}"
        if require_worktree:
            path = ROOT / relative
            assert path.is_file() and not path.is_symlink()
            assert sha(path) == expected, f"0820 sealed worktree payload changed: {name}"
    seal_blob = subprocess.check_output(
        ["git", "show", f"{base}:docs/performance/results/change-0820/seal.json"],
        cwd=ROOT,
    )
    assert _sha_bytes(seal_blob) == sha(PREVIOUS_SEAL_PATH)
    return {
        "path": str(PREVIOUS_SEAL_PATH),
        "bytes": PREVIOUS_SEAL_PATH.stat().st_size,
        "sha256": sha(PREVIOUS_SEAL_PATH),
        "base": base,
        "schema": seal["schema"],
        "files": len(files),
        "worktree_checked": require_worktree,
    }


def original_repair_inputs() -> dict[str, object]:
    """Replay the repair runner's complete original frozen-input witness."""
    inputs = read(REPAIR_INPUTS_PATH)
    assert inputs.get("schema") == "litchi.performance.0820.repair-inputs.v1"
    descriptors = {}
    for name in ("before", "origin", "original_frozen_inputs", "runner", "source"):
        descriptors[name] = _descriptor(inputs.get(name), f"0820 repair input {name}")
    frozen_path = Path(descriptors["original_frozen_inputs"]["path"])
    frozen = read(frozen_path)
    assert frozen.get("schema") == "litchi.performance.0820.frozen-inputs.v1"
    assert frozen.get("root_inputs") == assert_root_inputs()
    frozen_locks = frozen.get("locks")
    live_locks = lock_identity()
    assert isinstance(frozen_locks, dict)
    for name in ("root", "tool"):
        assert frozen_locks.get(name) == live_locks.get(name), (
            f"0820 frozen {name} lock identity changed"
        )
    assert frozen_locks["packet_tool"]["bytes"] == live_locks["packet_tool"]["bytes"]
    assert frozen_locks["packet_tool"]["sha256"] == live_locks["packet_tool"]["sha256"]
    assert Path(frozen_locks["packet_tool"]["path"]).is_file()
    assert Path(live_locks["packet_tool"]["path"]).is_file()
    assert frozen.get("architecture") == architecture_hashes()
    assert frozen.get("corpus") == assert_corpus_inputs()
    assert frozen.get("unrelated") == assert_unrelated()
    old_host = PREVIOUS_PACKET / "host.json"
    assert frozen.get("host") == sha(old_host)
    old_packet = frozen.get("packet")
    old_drivers = frozen.get("drivers")
    assert isinstance(old_packet, dict) and isinstance(old_drivers, dict)
    for name, expected in old_packet.items():
        assert sha(PREVIOUS_PACKET / name) == expected, f"0820 frozen packet changed: {name}"
    for name, expected in old_drivers.items():
        assert sha(PREVIOUS_PACKET / name) == expected, f"0820 frozen driver changed: {name}"
    return {
        "inputs": artifact(REPAIR_INPUTS_PATH),
        "descriptors": descriptors,
        "original_frozen_inputs": descriptors["original_frozen_inputs"],
        "schema": frozen["schema"],
    }


def assert_quality_reuse(*, require_worktree_seal: bool = False) -> dict[str, object]:
    """Check the exact committed 0820 repair result used by the 0821 adapter."""
    plan = read(P / "plan.json")
    reuse = plan.get("quality_reuse")
    assert isinstance(reuse, dict)
    assert reuse.get("schema") == "litchi.performance.0821.quality-reuse.v1"
    assert reuse.get("mode") == "committed-receipt-replay"
    assert reuse.get("required_gate_count") == 6
    assert reuse.get("cargo_commands_executed") is False

    quality = read(REPAIR_QUALITY_PATH)
    assert quality.get("schema") == "litchi.performance.0820.repair-quality.v1"
    assert quality.get("status") == "pass" and quality.get("gate_count") == 6
    source_descriptor = _descriptor(quality.get("source"), "0820 repair quality source",
                                    expected=REPAIR_SOURCE_PATH)
    inputs_descriptor = _descriptor(quality.get("inputs"), "0820 repair quality inputs",
                                    expected=REPAIR_INPUTS_PATH)
    gates = quality.get("gates")
    assert isinstance(gates, list) and len(gates) == 6
    gate_names = []
    gate_descriptors = []
    for gate in gates:
        assert isinstance(gate, dict)
        result_descriptor = _descriptor(gate, "0820 repair gate result")
        result = read(Path(result_descriptor["path"]))
        assert result.get("schema") == "litchi.performance.0820.repair-command.v1"
        assert result.get("exit_code") == 0
        assert isinstance(result.get("name"), str)
        assert result["name"] not in gate_names
        gate_names.append(result["name"])
        log_descriptor = _descriptor(result.get("log"), f"0820 repair {result['name']} log")
        assert isinstance(result.get("command"), list)
        assert isinstance(result.get("started"), (int, float))
        assert isinstance(result.get("ended"), (int, float))
        assert result["started"] <= result["ended"]
        gate_descriptors.append({"result": result_descriptor, "log": log_descriptor})
    assert gate_names == ["fmt", "check", "tests", "clippy", "rustdoc", "boundaries"]

    focused_descriptor = _descriptor(quality.get("focused"), "0820 focused result")
    focused = read(Path(focused_descriptor["path"]))
    assert focused.get("schema") == "litchi.performance.0820.repair-command.v1"
    assert focused.get("name") == "focused" and focused.get("exit_code") == 0
    focused_log = _descriptor(focused.get("log"), "0820 focused log")

    summary = read(REPAIR_TEST_SUMMARY_PATH)
    assert summary == {
        "failed": 0,
        "ignored": 1,
        "log": summary["log"],
        "passed": 641,
        "schema": "litchi.performance.0820.repair-tests.v1",
        "scope": "full perf-baseline all-features tests including doctest invocation",
        "suites": 28,
    }
    summary_log = _descriptor(summary.get("log"), "0820 repair test summary log")

    original_inputs = original_repair_inputs()
    source_witness = repair_source_witness()
    seal_witness = previous_seal(require_worktree=require_worktree_seal)
    return {
        "schema": reuse["schema"],
        "mode": reuse["mode"],
        "source": source_descriptor,
        "inputs": inputs_descriptor,
        "quality": artifact(REPAIR_QUALITY_PATH),
        "test_summary": artifact(REPAIR_TEST_SUMMARY_PATH),
        "gates": gate_descriptors,
        "focused": {"result": focused_descriptor, "log": focused_log},
        "test_summary_log": summary_log,
        "original_inputs": original_inputs,
        "source_witness": source_witness,
        "seal": seal_witness,
        "cargo_commands_executed": False,
    }


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
