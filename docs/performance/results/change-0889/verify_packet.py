#!/usr/bin/env python3
"""Replay 0889 command custody, frozen source overlays, and negative controls."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess

PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def descriptor(path):
    data = path.read_bytes()
    return {"path": str(path.relative_to(PACKET)), "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest()}


def read_attempt(number):
    path = PACKET / "attempts" / f"validation-{number}" / "validation-receipt.json"
    return json.loads(path.read_text()), descriptor(path)


def check_commands(receipt):
    names = set()
    for command in receipt["commands"]:
        require(command["name"] not in names, "duplicate command name")
        names.add(command["name"])
        identity = descriptor(PACKET / command["log"])
        require(identity["sha256"] == command["log_sha256"] and
                identity["bytes"] == command["log_bytes"], "command log changed")
        require(command["cwd"] == str(ROOT), "unexpected command directory")
        expected_pass = (command["returncode"] == 0 if command["expected"] == "pass"
                         else command["returncode"] == 1)
        require(command["pass"] is expected_pass, "command exit interpretation changed")
    return names


def check_source(receipt, current=False):
    overlay = receipt.get("source_overlay")
    require(overlay is not None, "source overlay absent")
    overlay_path = PACKET / overlay["path"]
    manifest = json.loads(overlay_path.read_text())
    require(manifest["base_commit"] == receipt["before"]["git_head"], "overlay base mismatch")
    mapped = {entry["path"]: entry for entry in manifest["files"]}
    require(len(mapped) == 4, "unexpected changed source population")
    for entry in receipt["before"]["source"]:
        name = entry["path"]
        if name in mapped:
            require(entry == mapped[name], "overlay descriptor differs")
            data = (overlay_path.parent / "source" / name).read_bytes()
        else:
            data = subprocess.check_output(["git", "show", f"{manifest['base_commit']}:{name}"], cwd=ROOT)
        require(len(data) == entry["bytes"] and hashlib.sha256(data).hexdigest() == entry["sha256"],
                f"source replay differs: {name}")
        if current:
            require((ROOT / name).read_bytes() == data, f"current source differs: {name}")
    return len(receipt["before"]["source"])


def negative_record(receipt, expected_stages):
    matches = []
    for retained in receipt["negative_reports"]:
        identity = retained["json"]
        path = PACKET / identity["path"]
        require(descriptor(path) == identity, "negative report content changed")
        report = json.loads(path.read_text())
        stages = [failure["stage"] for failure in report["input_inventory"]["failures"]]
        if stages != expected_stages:
            continue
        require(report["qualification_complete"] is False and report["control_failure_count"] == 0,
                "negative control failed for an unrelated reason")
        require(report["input_inventory"]["complete"] is False and
                report["input_inventory"]["read_failure_count"] == len(stages), "negative accounting")
        require(report["generated"] == report["aliasing"] == 0, "negative generated population")
        require(retained["failures"] == report["input_inventory"]["failures"], "negative receipt mismatch")
        matches.append(identity)
    require(len(matches) == 1, f"negative population differs: {expected_stages}")
    return matches[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--check-current-source", action="store_true")
    parser.add_argument("--require-cleanup", action="store_true")
    args = parser.parse_args()
    attempts = {}
    identities = []
    for number in (1, 2, 3, 5, 6, 7):
        receipt, identity = read_attempt(number)
        attempts[number] = receipt
        identities.append(identity)
        check_commands(receipt)
    broad, unreadable, final = (attempts[number] for number in (3, 5, 7))
    require([attempts[number]["status"] for number in (1, 2, 3, 5, 6, 7)] ==
            ["fail", "fail", "fail", "fail", "pass", "pass"], "historical attempt statuses changed")
    required_prefix = {"fmt_check", "cargo_test_bin", "clippy_bin_no_deps", "clippy_bin_tests_no_deps",
                       "rustdoc_bin", "crate_boundary_check", "cargo_build_bin", "differential_a",
                       "differential_b", "negative_missing_root"}
    require({c["name"] for c in broad["commands"]} == required_prefix and
            all(c["pass"] for c in broad["commands"]), "broad passed prefix incomplete")
    require(final["status"] == "pass" and len(final["commands"]) == 1 and
            final["commands"][0]["name"] == "negative_corrupt_member" and
            final["commands"][0]["returncode"] == 1, "final corrupt control absent")
    binary = broad["runtime_binary"]
    for receipt in (unreadable, attempts[6], final):
        require(receipt["before"]["source"] == broad["before"]["source"], "source drift between runs")
        require(receipt["before"]["binary"] == binary, "executable drift between runs")
        require(receipt["environment"] == broad["environment"], "environment drift between runs")
        require(receipt["before"]["rustc_vv"] == broad["before"]["rustc_vv"], "toolchain drift")
    require(final["after"]["source"] == broad["before"]["source"] and
            final["after"]["binary"] == binary, "source/executable changed by completion")
    sources = check_source(broad, args.check_current_source)
    check_source(final)
    tests = next(c for c in broad["commands"] if c["name"] == "cargo_test_bin")
    require("24 passed; 0 failed; 0 ignored" in (PACKET / tests["log"]).read_text(), "focused test count")
    negatives = [negative_record(broad, ["root_stat", "no_ooxml_package"]),
                 negative_record(unreadable, ["package_scan_read", "no_ooxml_package"]),
                 negative_record(final, ["member_read"])]
    oracle_path = PACKET / "oracle-verification.json"
    oracle = json.loads(oracle_path.read_text())
    require(oracle["status"] == "pass" and oracle["oracle_records"] == 200201 and
            oracle["controls"] == 10 and oracle["members"] == 67 and oracle["repeat_byte_equal"] is True,
            "independent oracle verification missing")
    for key in ("compressed_first", "compressed_second"):
        report = broad["differential"][key]
        path = PACKET / report["path"]
        require(descriptor(path)["sha256"] == report["sha256"], "compressed report changed")
    require(broad["differential"]["first_sha256"] == broad["differential"]["second_sha256"] ==
            oracle["reports"][0]["json_sha256"] == oracle["reports"][1]["json_sha256"],
            "repeat content identity mismatch")
    if args.require_cleanup:
        require(not (ROOT / "target/litchi-0889-validation").exists(), "owned target remains")
        require(not list((PACKET / "attempts").glob("*/runtime/**/*.xlsx")), "temporary archive remains")
    result = {"schema": "litchi.mce-packet-verification-0889.v1", "status": "pass",
              "source_files_reconstructed": sources, "focused_tests_passed": 24,
              "broad_runs": 2, "oracle_records_per_run": 200201, "controls_per_run": 10,
              "required_negative_controls": negatives, "attempt_receipts": identities,
              "oracle_verification": descriptor(oracle_path), "binary": binary,
              "current_source_checked": args.check_current_source, "cleanup_checked": args.require_cleanup,
              "historical_failures_retained": True,
              "additional_negative_runs_retained": ["validation-6 unreadable", "validation-6 corrupt"]}
    with args.output.open("x") as output:
        json.dump(result, output, indent=2)
        output.write("\n")
    print(json.dumps({key: value for key, value in result.items() if key not in
                      {"attempt_receipts", "required_negative_controls", "binary"}}, sort_keys=True))


if __name__ == "__main__":
    main()
