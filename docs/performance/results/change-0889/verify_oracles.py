#!/usr/bin/env python3
"""Verify complete v2 MCE population against an independent ZIP census."""
import argparse
from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import zipfile

from input_census import ROOT, census

MCE_LIMITS = {
    "max_input_bytes": 268435456, "max_output_bytes": 536870912,
    "max_depth": 256, "max_namespace_bindings": 4096,
    "max_directive_tokens": 4096, "max_choices_per_alternate": 1024,
    "max_attributes_per_element": 1024,
}
STREAM_LIMITS = {
    "processing": MCE_LIMITS, "max_events": 1000000,
    "max_event_bytes": 8388608, "max_attributes_per_event": 4096,
    "max_attribute_bytes_per_event": 4194304, "max_context_bytes": 16777216,
    "max_name_bytes": 1048576,
}
REPORT_FIELDS = {
    "alternate_content_count", "selected_choices", "selected_fallbacks",
    "ignored_elements", "ignored_attributes", "preserved_elements",
    "preserved_attributes", "unwrapped_elements",
}
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
ATTRIBUTE_CASES = (
    ("no_mce_zero_attributes", 0, 1024),
    ("attributes_exact_one", 1, 1024),
    ("attributes_exact_four", 4, 1024),
    ("attributes_exact_five", 5, 1024),
    ("attributes_exact_1024", 1024, 1024),
    ("attributes_over_1024", 1025, 1024),
    ("attributes_exact_4096", 4096, 4096),
    ("attributes_over_4096", 4097, 4096),
)


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest_valid(value):
    return (isinstance(value, str) and len(value) == 64
            and all(char in "0123456789abcdef" for char in value))


def check_error(error):
    if error is None:
        return
    require(isinstance(error, dict), "unstructured error")
    require(error.get("kind") in {"non_conformant", "must_understand", "limit_exceeded",
                                 "xml", "allocation", "input", "mce", "callback"},
            "unknown error kind")
    for field in ("error", "prior_mce_error"):
        if field in error:
            check_error(error[field])


def read_report(path):
    raw = path.read_bytes()
    data = gzip.decompress(raw) if path.suffix == ".gz" else raw
    return json.loads(data), {
        "path": str(path), "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "json_bytes": len(data), "json_sha256": hashlib.sha256(data).hexdigest(),
    }


def check_controls(report):
    require(report["qualification_complete"] is True and report["control_failure_count"] == 0,
            "control qualification failed")
    controls = report["controls"]
    expected_names = {row[0] for row in ATTRIBUTE_CASES} | {
        "semantic_bare", "semantic_understood_extension",
    }
    require(len(controls) == len(expected_names) and
            {control["name"] for control in controls} == expected_names,
            "control population mismatch")
    by_name = {control["name"]: control for control in controls}
    for control in controls:
        require(control["expectation_met"] is True, "control expectation failed")
        for operation in ("codec", "stream"):
            record = control[operation]
            require(record["qualification_complete"] is True, "incomplete control oracle")
            require(record["kind"] == "mce_" + operation, "control oracle kind")
            check_error(record["error"])
        require(control["codec"]["input"] == control["stream"]["input"], "control inputs differ")
        require(control["stream"]["semantic"]["unknown_events"] == 0, "unknown control event")
        require(control["stream"]["raw_unknown_kinds"] == 0, "unknown raw control kind")
    for name, count, limit in ATTRIBUTE_CASES:
        control = by_name[name]
        require(control["family"] == "attribute_cardinality" and
                control["attribute_count_including_xmlns"] == count and
                control["configured_limit"] == limit, "attribute control configuration")
        document = ("<r/>" if count == 0 else '<r xmlns:mc="' + MC + '"' +
                    ''.join(f' a{index}="v"' for index in range(count - 1)) + '/>').encode()
        require(control["codec"]["input"] == {
            "bytes": len(document), "sha256": hashlib.sha256(document).hexdigest(),
        }, "attribute control input")
        processing = dict(MCE_LIMITS, max_attributes_per_element=limit)
        require(control["codec"]["limits"] == processing and
                control["stream"]["limits"] == dict(STREAM_LIMITS, processing=processing),
                "attribute control limits")
        success = count <= limit
        require(control["expectation"] == ("success" if success else "limit"), "control expectation")
        for operation in ("codec", "stream"):
            record = control[operation]
            require(record["capabilities"] == "bare" and record["success"] is success,
                    "attribute control outcome")
            if success:
                expected_report = dict.fromkeys(REPORT_FIELDS, 0)
                if operation == "stream":
                    expected_report["events"] = 1
                    require(record["raw"]["events"] == record["semantic"]["events"] == 1,
                            "attribute callback counts")
                else:
                    output = document if count == 0 else document[:-2] + b'></r>'
                    require(record["output"] == {
                        "bytes": len(output), "sha256": hashlib.sha256(output).hexdigest(),
                    }, "attribute output preservation")
                require(record["error"] is None and record["report"] == expected_report,
                        "attribute report differs")
            else:
                message = ("stream attributes per event" if operation == "stream" and limit == 4096
                           else "attributes per element")
                error = {"kind": "limit_exceeded", "message": message}
                if operation == "stream":
                    error = {"kind": "mce", "error": error, "prior_mce_error": None,
                             "raw_callback_error": False, "active_callback_error": False}
                require(record["error"] == error and record["report"] is None,
                        "attribute refusal precedence")
    semantic_input = by_name["semantic_bare"]["codec"]["input"]
    for name, understood in (("semantic_bare", False), ("semantic_understood_extension", True)):
        control = by_name[name]
        require(control["family"] == "semantic_selection", "semantic control family")
        for operation in ("codec", "stream"):
            record = control[operation]
            require(record["input"] == semantic_input and record["success"] is True and
                    record["error"] is None, "semantic control success/input")
            require(record["capabilities"] == ("understood:urn:i" if understood else "bare"),
                    "semantic control capability")
            require(record["limits"] == (MCE_LIMITS if operation == "codec" else STREAM_LIMITS),
                    "semantic control limits")
            expected = {"alternate_content_count": 1, "selected_choices": int(understood),
                        "selected_fallbacks": int(not understood), "ignored_elements": int(not understood),
                        "preserved_elements": 0 if understood else 2,
                        "unwrapped_elements": int(not understood)}
            require(all(record["report"][key] == value for key, value in expected.items()),
                    "semantic control selection report")
    return len(controls)


def check(report, generated, aliasing):
    expected = census()
    require(report["schema"] == "litchi.xml-attribute-bounds.differential.v2", "schema")
    require(report["incomplete_oracle_count"] == 0, "incomplete oracle population")
    control_count = check_controls(report)
    require(report["generated"] == generated and report["aliasing"] == aliasing,
            "generated population differs")
    inventory = report["input_inventory"]
    require([str((ROOT / path).resolve()) for path in inventory["roots"]] ==
            [str((ROOT / package["path"]).resolve()) for package in expected["packages"]],
            "requested root population/order differs")
    require(inventory["complete"] is True, "incomplete inventory")
    require(inventory["failures"] == [] and inventory["read_failure_count"] == 0,
            "input failures")
    require(inventory["failure_counts"] == {}, "nonempty failure counts")
    require(inventory["scanned_file_count"] == 3 and
            inventory["excluded_file_count"] == 0 and inventory["excluded_files"] == [],
            "unexpected scanned or excluded files")
    require(report["packages"] == inventory["package_count"] == 3, "package count")
    require(report["members"] == inventory["read_member_count"] ==
            inventory["selected_member_count"] == 67, "member count")
    expected_packages = {str((ROOT / p["path"]).resolve()): p for p in expected["packages"]}
    require(len(inventory["packages"]) == 3, "duplicate package inventory")
    seen = set()
    oracle_keys = {}
    extra_result_keys = set()
    for package in inventory["packages"]:
        absolute = str((ROOT / package["path"]).resolve())
        require(absolute in expected_packages and absolute not in seen, "package identity")
        seen.add(absolute)
        wanted = expected_packages[absolute]
        with zipfile.ZipFile(absolute) as archive:
            require(package["archive_member_count"] == len(archive.infolist()),
                    "archive population count differs")
        require(package["bytes"] == wanted["bytes"] and
                package["sha256"] == wanted["sha256"], "package content")
        require(package["selected_member_count"] == package["read_member_count"] ==
                len(wanted["members"]), "package member count")
        require(sorted(package["members"], key=lambda item: item["name"]) == wanted["members"],
                "member identities differ from independent ZIP reader")
        for member in wanted["members"]:
            prefix = package["path"] + "::" + member["name"]
            identity = {key: member[key] for key in ("bytes", "sha256")}
            for suffix, capability in (("mce_baseline", "ooxml_baseline"),
                                       ("mce_bare", "bare"),
                                       ("stream_baseline", "ooxml_baseline")):
                oracle_keys[prefix + "::" + suffix] = (capability, identity)
            for suffix in ("audit_source", "audit_authored", "audit_compact"):
                extra_result_keys.add(prefix + "::" + suffix)
            if member["name"].lower().endswith("styles.xml"):
                extra_result_keys.add(prefix + "::docx_styles")
    for group, count, profiles in (
        ("generated", generated, (("bare", "bare"), ("understood", "understood:urn:w"))),
        ("aliasing", aliasing, (("bare", "bare"), ("baseline", "ooxml_baseline"),
                               ("understood", "understood:urn:a"))),
    ):
        for index in range(count):
            for suffix, capability in profiles:
                for operation in ("mce", "stream"):
                    oracle_keys[f"{group}::{index:06}::{operation}_{suffix}"] = (capability, None)
    require(set(report["oracles"]) == set(oracle_keys), "oracle key population mismatch")
    require(set(report["results"]) == set(oracle_keys) | extra_result_keys,
            "legacy result key population mismatch")
    counts = Counter()
    generated_inputs = {}
    for key, (capability, expected_input) in oracle_keys.items():
        record = report["oracles"][key]
        require(record["qualification_complete"] is True, "incomplete oracle")
        stream = key.rsplit("::", 1)[-1].startswith("stream_")
        require(record["kind"] == ("mce_stream" if stream else "mce_codec"), "oracle kind")
        require(record["capabilities"] == capability, "capability mismatch")
        require(record["limits"] == (STREAM_LIMITS if stream else MCE_LIMITS), "limit mismatch")
        identity = record["input"]
        require(isinstance(identity["bytes"], int) and identity["bytes"] > 0 and
                digest_valid(identity["sha256"]), "invalid input identity")
        if expected_input is not None:
            require(identity == expected_input, "oracle input differs from ZIP member")
        else:
            prefix = key.rsplit("::", 1)[0]
            require(generated_inputs.setdefault(prefix, identity) == identity,
                    "capability variants use different inputs")
        require(type(record["success"]) is bool, "missing success boolean")
        counts[("stream" if stream else "codec") + ("_success" if record["success"] else "_refusal")] += 1
        if stream:
            require(record["raw_unknown_kinds"] == 0, "unsupported raw element kind")
            require(record["semantic"]["unknown_events"] == 0, "unsupported semantic event")
            for channel in ("raw", "semantic"):
                require(digest_valid(record[channel]["sha256"]) and
                        type(record[channel]["events"]) is int and record[channel]["events"] >= 0,
                        "invalid callback transcript")
        if record["success"]:
            require(record["error"] is None, "success carries error")
            require(set(record["report"]) == REPORT_FIELDS | ({"events"} if stream else set()),
                    "incomplete report fields")
            require(all(type(value) is int and value >= 0 for value in record["report"].values()),
                    "invalid report count")
            if not stream:
                require(digest_valid(record["output"]["sha256"]) and
                        type(record["output"]["bytes"]) is int, "invalid output identity")
        else:
            require(record["report"] is None and isinstance(record["error"], dict), "refusal fields")
            check_error(record["error"])
            if stream:
                error = record["error"]
                require(error["raw_callback_error"] is False and error["active_callback_error"] is False,
                        "infallible callback reported failure")
                require("prior_mce_error" in error or error["kind"] == "callback", "missing prior error")
    return {"packages": 3, "members": 67, "generated": generated, "aliasing": aliasing,
            "controls": control_count,
            "oracle_records": len(oracle_keys), "legacy_results": len(report["results"]),
            "distinct_generated_inputs": {
                group: len({identity["sha256"] for key, identity in generated_inputs.items()
                            if key.startswith(group + "::")})
                for group in ("generated", "aliasing")
            },
            "outcomes": dict(sorted(counts.items()))}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--first", type=Path, required=True)
    parser.add_argument("--second", type=Path, required=True)
    parser.add_argument("--generated", type=int, default=20000)
    parser.add_argument("--aliasing", type=int, default=20000)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    first, first_descriptor = read_report(args.first)
    result = check(first, args.generated, args.aliasing)
    second, second_descriptor = read_report(args.second)
    require(first == second, "repeat outputs differ")
    require(first_descriptor["json_sha256"] == second_descriptor["json_sha256"], "repeat bytes differ")
    # Prove the population check detects an omitted outcome, independently of repeat equality.
    dropped_key, dropped_value = first["oracles"].popitem()
    try:
        check(first, args.generated, args.aliasing)
    except ValueError as error:
        require(str(error) == "oracle key population mismatch", "unexpected negative check")
    else:
        raise ValueError("missing oracle was accepted")
    first["oracles"][dropped_key] = dropped_value
    saved_controls = first["controls"]
    first["controls"] = []
    try:
        check(first, args.generated, args.aliasing)
    except ValueError as error:
        require(str(error) == "control population mismatch", "unexpected control negative check")
    else:
        raise ValueError("missing controls were accepted")
    first["controls"] = saved_controls
    result.update({"schema": "litchi.mce-oracle-verification-0889.v1", "status": "pass",
                   "repeat_byte_equal": True, "missing_oracle_refused": True,
                   "missing_controls_refused": True,
                   "reports": [first_descriptor, second_descriptor]})
    with args.output.open("x") as output:
        json.dump(result, output, indent=2)
        output.write("\n")
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
