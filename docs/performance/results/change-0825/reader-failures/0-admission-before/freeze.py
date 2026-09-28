"""Freeze 0825 custody before the root installs the exact before source."""

from __future__ import annotations

import custody as c


assert not (c.P / "freeze.json").exists(), "refusing to overwrite freeze.json"
assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
assert {path.name for path in c.TARGET.iterdir()} == {"quality"}, (
    "before freeze permits only the completed quality target"
)
c.assert_static()
quality = c.assert_quality()
current = c.source()
assert current["revision"] == c.BASE
assert {name: current["files"][name] for name in c.ALLOWLIST} == c.candidate_files("after")

frozen = {
    "schema": "litchi.performance.0825.freeze.v1",
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
print("0825 freeze PASS: shipped after source, exact before archives, quality, corpus, locks, and host bound")
