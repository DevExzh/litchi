"""Freeze 0827 custody before the root installs the exact before source."""

from __future__ import annotations

import custody as c
import driver_preflight as preflight


assert not (c.P / "freeze.json").exists(), "refusing to overwrite freeze.json"
assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
assert {path.name for path in c.TARGET.iterdir()} == {"quality"}, (
    "before freeze permits only the completed quality target"
)
c.assert_static()
quality = c.assert_quality()
preflight_path = c.P / "preflight.json"
assert preflight_path.is_file() and not preflight_path.is_symlink(), (
    "0827 driver preflight must pass before freeze"
)
preflight_value = preflight.derive()
assert preflight_value["status"] == "pass"
assert preflight_value["reports"] == 325
assert preflight_value["sample_envelopes"] == 6733
assert c.read(preflight_path) == preflight_value, (
    "preflight.json is not the canonical result of the required freeze gate"
)
suite_path = c.P / "preflight-suite.json"
suite = c.read(suite_path)
assert suite["schema"] == "litchi.performance.0827.preflight-suite.v1"
assert set(suite["checks"]) == {"driver_preflight.py", "admission_preflight.py", "reader_preflight.py"}
for name, descriptor in suite["checks"].items():
    receipt_path = c.resolve_descriptor(descriptor)
    assert c.artifact(receipt_path) == descriptor
    receipt = c.read(receipt_path)
    assert receipt["exit_code"] == 0 and receipt["command"][2] == str(c.P / name)
    log_path = c.resolve_descriptor(receipt["log"])
    assert c.artifact(log_path) == receipt["log"]
    dependencies = {
        "driver_preflight.py": ("driver_preflight.py", "custody.py"),
        "admission_preflight.py": ("admission_preflight.py", "admission.py", "preservation.py", "artifact_audit.py"),
        "reader_preflight.py": ("reader_preflight.py", "analysis.py", "raw_audit.py"),
    }[name]
    required = {str(c.P / item) for item in dependencies} | {str(c.ROOT / "tools/perf_allocation_schema.py")}
    recorded = {row["original"]["path"]: row for row in receipt["sources"]}
    assert required <= recorded.keys()
    for source in required:
        original = recorded[source]["original"]
        assert c.artifact(c.resolve_descriptor(original)) == original, source
        snapshot = recorded[source]["snapshot"]
        assert c.artifact(c.resolve_descriptor(snapshot)) == snapshot, source
        assert snapshot["sha256"] == original["sha256"]
current = c.source()
assert current["revision"] == c.BASE
assert {name: current["files"][name] for name in c.ALLOWLIST} == c.candidate_files("after")

frozen = {
    "schema": "litchi.performance.0827.freeze.v1",
    "base": c.BASE,
    "source": current,
    "after_source": current,
    "before_source_contract": c.candidate_files("before"),
    "after_source_contract": c.candidate_files("after"),
    "tool": c.tool_source(),
    "candidate_archives": c.candidate_descriptors(),
    "root_inputs": c.root_inputs(),
    "architecture": c.architecture(),
    "corpus": c.corpus(),
    "unrelated": c.unrelated(),
    "provenance": c.artifact(c.P / "provenance.json"),
    "plan": c.artifact(c.P / "plan.json"),
    "quality": c.artifact(c.P / "quality.json"),
    "preflight": c.artifact(preflight_path),
    "preflight_suite": c.artifact(suite_path),
    "allocation_schema": c.allocation_schema_source(),
    "host": c.artifact(c.P / "host.json"),
    "toolchain": c.artifact(c.P / "toolchain.json"),
    "root_input_manifest": c.artifact(c.P / "root-inputs.json"),
    "architecture_manifest": c.artifact(c.P / "architecture-inputs.json"),
    "corpus_manifest": c.artifact(c.P / "corpus-inputs.json"),
    "lock_parity": c.artifact(c.P / "lock-parity.json"),
    "origin": c.artifact(c.P / "origin.json"),
    "drivers": c.packet_files(c.DRIVERS),
    "support": c.support_files(),
}
c.write(c.P / "freeze.json", frozen)
c.stable_inputs(frozen)
print("0827 freeze PASS: shipped after source, exact before archives, quality, allocation schema, corpus, locks, and preflight bound")
