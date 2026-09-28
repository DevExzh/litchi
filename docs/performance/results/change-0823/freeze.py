"""Freeze 0823 custody before the baseline build and native lane."""

from __future__ import annotations

import custody as c


assert not (c.P / "freeze.json").exists(), "refusing to overwrite freeze.json"
assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
assert {path.name for path in c.TARGET.iterdir()} == {"quality"}, (
    "before freeze permits only the completed baseline quality target"
)
c.assert_quality_before()
c.assert_static()
source = c.source()
assert source["revision"] == c.BASE, "freeze must start at the declared base commit"
frozen = {
    "schema": "litchi.performance.0823.freeze.v1",
    "base": c.BASE,
    "source": source,
    "tool": c.tool_source(),
    "probes": {name: c.probe_files(name) for name in c.PROBES},
    "candidate": c.packet_files(c.CANDIDATE_FILES),
    "drivers": c.packet_files(c.DRIVER_FILES),
    "root_inputs": c.root_inputs(),
    "architecture": c.architecture(),
    "unrelated": c.unrelated(),
    "plan": c.artifact(c.P / "plan.json"),
    "host": c.artifact(c.P / "host.json"),
    "toolchain": c.artifact(c.P / "toolchain.json"),
    "root_input_manifest": c.artifact(c.P / "root-inputs.json"),
    "architecture_manifest": c.artifact(c.P / "architecture-inputs.json"),
    "corpus": c.artifact(c.P / "corpus-inputs.json"),
    "lock_parity": c.artifact(c.P / "lock-parity.json"),
}
c.write(c.P / "freeze.json", frozen)
c.stable(frozen)
print("0823 freeze PASS: source, probes, locks, architecture, host, toolchain, and corpus bound")
