#!/usr/bin/env python3
"""Exercise the actual 0730 analyzer against temporary evidence mutations."""

from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import io
import json
import shutil
import tempfile
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
spec = importlib.util.spec_from_file_location("change0730_analyzer", P / "analyze.py")
assert spec and spec.loader
analyzer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(analyzer)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def copy_packet(target: Path) -> None:
    target.mkdir(exist_ok=True)
    for child in P.iterdir():
        if child.name in {"captures", "before", "__pycache__"}:
            continue
        if child.is_file():
            shutil.copy2(child, target / child.name)
        elif child.is_dir():
            shutil.copytree(child, target / child.name)
    shutil.copytree(P / "captures", target / "captures")
    shutil.copytree(P / "before", target / "before")


def main() -> None:
    checks: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="litchi-0730-negative-") as temporary:
        packet = Path(temporary)
        copy_packet(packet)
        freeze_path = packet / "freeze.json"
        frozen = json.loads(freeze_path.read_text())
        source_bindings = frozen.get("bindings", frozen)
        prefix = str(P) + "/"
        rebased = {key.replace(prefix, str(packet) + "/", 1) if key.startswith(prefix) else key: value
                   for key, value in source_bindings.items()}
        if "bindings" in frozen:
            frozen["bindings"] = rebased
        else:
            frozen = rebased
        freeze_path.write_text(json.dumps(frozen))
        analyzer.P = packet
        analyzer.ROOT = ROOT
        analyzer.CAPTURES = packet / "captures"
        analyzer.BEFORE = packet / "before"
        manifest_path = packet / "captures" / "manifest.json"
        original_manifest = manifest_path.read_bytes()
        manifest = json.loads(original_manifest)
        target = packet / "captures" / manifest["runs"][0]["output"]
        original_report = target.read_bytes()

        def run(label: str, accepted: bool) -> None:
            okay = True
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    analyzer.main()
            except (analyzer.Failure, AssertionError, KeyError, OSError, TypeError, ValueError):
                okay = False
            assert okay is accepted, label
            checks.append({"name": label, "accepted": okay, "expected": accepted})

        def restore() -> None:
            target.write_bytes(original_report)
            manifest_path.write_bytes(original_manifest)

        run("complete packet accepted", True)

        # Raw evidence and its manifest digest are coupled.
        target.write_bytes(original_report + b"\n")
        run("raw report hash mismatch rejected", False)
        restore()

        def mutate_report(label: str, mutate) -> None:
            report = json.loads(original_report)
            mutate(report)
            target.write_text(json.dumps(report), encoding="utf-8")
            changed_manifest = json.loads(original_manifest)
            changed_manifest["runs"][0]["sha256"] = digest(target)
            manifest_path.write_text(json.dumps(changed_manifest), encoding="utf-8")
            run(label, False)
            restore()

        mutate_report("false expected semantic boolean rejected",
                      lambda x: x["expected_oracle"].__setitem__("semantic_reopen_ok", False))
        mutate_report("missing semantic witness rejected",
                      lambda x: x["samples"][0]["oracle"].__setitem__("semantic_witness", {}))
        mutate_report("physical-only length proof rejected",
                      lambda x: x["changed_length_proof"].__setitem__(
                          "logical_stream_length_change_proven", False))
        mutate_report("missing format-specific length proof rejected",
                      lambda x: x["changed_length_proof"].__setitem__(
                          "format_specific_semantic_length_proven", False))
        mutate_report("missing named negative control rejected",
                      lambda x: x["oracle_controls"].pop())
        mutate_report("accepted negative control rejected",
                      lambda x: x["oracle_controls"][0].update(status="unexpectedly_accepted",
                                                                  rejected=False,
                                                                  failure_reasons=[]))
        mutate_report("wrong source identity rejected",
                      lambda x: x.__setitem__("source_sha256", "0" * 64))
        mutate_report("wrong output identity rejected",
                      lambda x: x.__setitem__("expected_output_sha256", "0" * 64))
        mutate_report("stream inventory mutation rejected",
                      lambda x: x["samples"][0]["output_inventory"]["streams"][0].__setitem__(
                          "sha256", "0" * 64))
        mutate_report("allocation field removal rejected",
                      lambda x: x["samples"][0].__setitem__("allocations", {}))
        mutate_report("timing field mutation rejected",
                      lambda x: x["samples"][0].__setitem__("phase_ns", {"whole_ns": 0}))

        def mutate_manifest(label: str, mutate) -> None:
            changed = json.loads(original_manifest)
            mutate(changed)
            manifest_path.write_text(json.dumps(changed), encoding="utf-8")
            run(label, False)
            restore()

        mutate_manifest("duplicate process rejected",
                        lambda x: x["runs"].__setitem__(1, x["runs"][0]))
        mutate_manifest("command order mutation rejected",
                        lambda x: x["runs"][0]["command"].__setitem__(-1, "999"))
        mutate_manifest("missing process rejected", lambda x: x["runs"].pop())

        def mutate_packet(name: str, label: str, mutate) -> None:
            path = packet / name
            original = path.read_bytes()
            value = json.loads(original)
            mutate(value)
            path.write_text(json.dumps(value), encoding="utf-8")
            run(label, False)
            path.write_bytes(original)

        mutate_packet("cleanup.json", "partial binary cleanup witness rejected",
                      lambda x: x["identities"].pop())
        def mutate_freeze(value) -> None:
            bindings = value.get("bindings", value)
            key = next(iter(bindings))
            bindings[key] = "0" * 64

        mutate_packet("freeze.json", "frozen custody mutation rejected", mutate_freeze)

        def mutate_quality_failure() -> None:
            qualification_path = packet / "qualification.json"
            qualification_original = qualification_path.read_bytes()
            qualification = json.loads(qualification_original)
            quality_relative = next(relative for relative in qualification["files"]
                                    if relative.startswith("quality-") and relative.endswith("/manifest.json"))
            quality_path = packet / quality_relative
            quality_original = quality_path.read_bytes()
            quality = json.loads(quality_original)
            quality["runs"][0]["exit_code"] = 1
            quality_path.write_text(json.dumps(quality), encoding="utf-8")
            qualification["files"][quality_relative] = digest(quality_path)
            qualification_path.write_text(json.dumps(qualification), encoding="utf-8")
            run("quality manifest failure rejected", False)
            quality_path.write_bytes(quality_original)
            qualification_path.write_bytes(qualification_original)

        mutate_quality_failure()
        mutate_packet("oracle-contract.json", "oracle contract mutation rejected",
                      lambda x: x["docfloat"]["control_names"].append("unexpected"))

        receipt = {"analyzer_sha256": digest(P / "analyze.py"),
                   "script_sha256": digest(Path(__file__)), "checks": checks}
        (P / "negative-checks.json").write_text(json.dumps(receipt, indent=2) + "\n",
                                                encoding="utf-8")
        print(f"PASS {len(checks)} actual-analyzer mutation controls")


if __name__ == "__main__":
    main()
