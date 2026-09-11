#!/usr/bin/env python3
"""Recheck source-bound gates and rerun the retained XML schema evidence."""

import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import zipfile

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    gates = HERE / "gates"
    receipt = json.loads((gates / "receipt.json").read_text())
    assert receipt["passed"] and receipt["source_unchanged"]
    assert len(receipt["commands"]) == 7
    assert all(row["exit_code"] == 0 for row in receipt["commands"])
    sources = json.loads((gates / "source-hashes.json").read_text())
    for name, expected in sources.items():
        assert digest(ROOT / name) == expected, name
    reviewed = json.loads((HERE / "review-source-hashes.json").read_text())
    for name, expected in reviewed.items():
        assert sources[name] == expected, f"reviewed source differs from gate: {name}"
    for name, expected in receipt["logs"].items():
        assert digest(gates / name) == expected, name
    totals = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", (gates / "tests.log").read_text())
    actual = dict(zip(["passed", "failed", "ignored"], [sum(int(row[i]) for row in totals) for i in range(3)]))
    assert actual == receipt["test_totals_including_doctests"]
    caller = json.loads((HERE / "caller-receipt.json").read_text())
    assert caller["passed"] and caller["source_unchanged"] and caller["owned_temporary_target_removed"]
    assert caller["production_source_manifest_sha256"] == hashlib.sha256(json.dumps(sources, sort_keys=True).encode()).hexdigest()
    assert digest(HERE / "caller.log") == caller["log_sha256"]
    for name, expected in {**caller["caller_inputs"], **caller["outputs"]}.items():
        assert digest(HERE / name) == expected, name
    schema_reports = 0
    for name in ["native-schema-validation.json", "caller-schema-validation.json"]:
        previous = json.loads((HERE / name).read_text())
        paths = list(dict.fromkeys(row["path"] for row in previous["reports"]))
        current = json.loads(subprocess.check_output([sys.executable, str(HERE / "validate_schema.py"), *paths], cwd=ROOT))
        assert current["reports"] == previous["reports"], name
        schema_reports += len(current["reports"])
    result = {"passed": True, "gate_input_files": len(sources), "reviewed_files": len(reviewed),
              "tests_including_doctests": actual, "schema_reports": schema_reports}
    if "--performance" in sys.argv:
        profile = json.loads(subprocess.check_output(
            [sys.executable, str(HERE / "performance" / "verify_root.py")], cwd=ROOT))
        assert profile["passed"]
        raw_samples = 0
        raw_processes = 0
        native = ROOT / "test-data/ooxml/xlsb/date.xlsb"
        with zipfile.ZipFile(native) as archive:
            native_theme_hash = hashlib.sha256(archive.read("xl/theme/theme1.xml")).hexdigest()
        for path in sorted((HERE / "performance" / "results").glob("*-p[123].json")):
            raw = json.loads(path.read_text())
            assert len(raw["samples"]) == 30
            if raw["fixture"] == "native":
                assert raw["package_sha256"] == digest(native)
                assert raw["theme_sha256"] == native_theme_hash
            for sample in raw["samples"]:
                assert all(sample[key] is True for key in ["semantic_ok",
                    "forward_preservation_ok", "preservation_ok", "inverse_ok",
                    "changed_ok", "alloc_balance_ok"]), path.name
                if raw["operation"] == "family_clone":
                    assert sample["alloc_calls"] == sample["realloc_calls"] == 0
                raw_samples += 1
            raw_processes += 1
        assert raw_processes == 54 and raw_samples == profile["samples"] == 1620
        result["performance"] = {key: profile[key] for key in
            ["passed", "source_manifest_sha256", "manifest_paths_checked", "lanes", "samples"]}
        result["performance"]["independently_checked_raw_processes"] = raw_processes
        result["performance"]["verification_receipt_sha256"] = digest(
            HERE / "performance" / "root-verification.json")
    (HERE / "root-verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
