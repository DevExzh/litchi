"""Pure mutation tests for the independent 0844 reader.

The test process never starts Cargo, the probe, Valgrind, or Heaptrack. It
creates small ZIP and Callgrind fixtures in a temporary directory, then checks
that the reader accepts the exact contract and rejects mutations of each
independent oracle.
"""

from __future__ import annotations

import copy
import io
import json
import struct
import tempfile
from pathlib import Path
import zipfile

import reader as r


HERE = Path(__file__).resolve().parent


def _reject(name: str, callback, rejected: list[str]) -> None:
    try:
        callback()
    except r.ReaderError:
        rejected.append(name)
    else:
        raise AssertionError(f"accepted malformed mutation: {name}")


def _archive(parts: int, changed: bool, shape: str = "shared") -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED,
                         strict_timestamps=False) as output:
        for name, payload in r._metadata_payloads(parts, shape).items():
            output.writestr(name, payload)
        for index in range(parts):
            output.writestr(r.part_archive_name(index, shape),
                            r.part_payload(index, changed and index == parts - 1, parts))
    return stream.getvalue()


def _logical(data: bytes) -> dict[str, object]:
    manifest = r.archive_manifest(data)
    return r._archive_logical_identity(manifest, data, "fixture")


def _report(parts: int, mode: str, source: bytes, output: bytes,
            samples: int = 1, warmup: int = 0,
            shape: str = "shared") -> dict[str, object]:
    source_id = r.archive_manifest(source)["archive"]
    output_id = r.archive_manifest(output)["archive"]
    source_logical = _logical(source)
    output_logical = _logical(output)
    return {
        "schema": r.REPORT_SCHEMA,
        "tool": r.REPORT_TOOL,
        "shape": shape,
        "mode": mode,
        "timing_description": r.TIMING_SCOPE[mode],
        "generator": {
            "shape": shape,
            "part_count": parts,
            "member_count": parts + 2,
            "payload_bytes": r.PART_BYTES,
            "part_name_pattern": {
                "shared": "/custom/p{index:05}.bin",
                "root": "/p{index:05}.bin",
                "deep-shared": "/a/b/c/d/e/f/g/h/p{index:05}.bin",
                "deep-unique": "/u{index:05}/b/c/d/e/f/g/h/p{index:05}.bin",
            }[shape],
            "depth": r.shape_depth(shape),
            "payload_formula": "byte[i] = (part_index * 17 + i * 13) % 251",
            "target_part_index": parts - 1,
            "target_part_name": r.part_name(parts - 1, shape),
            "replacement": "target payload with byte[0] XOR 1",
            "metadata_members": list(r.METADATA_ARCHIVE_NAMES),
        },
        "warmup": warmup,
        "samples_requested": samples,
        "source": source_id,
        "source_logical": source_logical,
        "expected_output": output_id,
        "expected_output_logical": output_logical,
        "samples": [
            {
                "index": index,
                "elapsed_ns": 1000 + index,
                "output": output_id,
                "logical": output_logical,
                "member_count": parts + 2,
                "hash_witness": {
                    "output_sha256": output_id["sha256"],
                    "logical_sha256": output_logical["sha256"],
                    "output_exact_expected": True,
                    "logical_exact_expected": True,
                },
            }
            for index in range(samples)
        ],
        "artifact": output_id,
        "artifact_role": "changed-output-archive",
        "source_artifact": source_id,
    }


def _zip_checks(rejected: list[str], passed: list[str]) -> None:
    parts = 4
    with tempfile.TemporaryDirectory(prefix="litchi-0844-reader-") as temporary:
        root = Path(temporary)
        for shape in r.SHAPES:
            source = _archive(parts, False, shape)
            output = _archive(parts, True, shape)
            source_manifest = r.archive_manifest(source)
            output_manifest = r.archive_manifest(output)
            assert source_manifest["member_count"] == output_manifest["member_count"] == parts + 2
            assert [row["name"] for row in source_manifest["members"]] == [
                *r.METADATA_ARCHIVE_NAMES,
                *(r.part_archive_name(index, shape) for index in range(parts)),
            ]
            r._generator_payload_check(source_manifest, source, parts, False, "source", shape)
            r._generator_payload_check(output_manifest, output, parts, True, "output", shape)
            spans = r.compare_untouched_spans(source_manifest, output_manifest, parts,
                                              shape=shape)
            target_metadata = r.compare_changed_target_metadata(
                source_manifest, output_manifest, parts, shape=shape)
            assert spans["count"] == parts + 1
            assert target_metadata["target"] == r.part_archive_name(parts - 1, shape)
            passed.append(f"{shape}-zip-crc-payload-and-raw-metadata")
            passed.append(f"{shape}-untouched-compressed-span-proof")

            source_path = root / f"{shape}-source.zip"
            output_path = root / f"{shape}-artifact.zip"
            report_path = root / f"{shape}-report.json"
            source_path.write_bytes(source)
            output_path.write_bytes(output)
            original = _report(parts, "publish", source, output, shape=shape)

            def check(value: dict[str, object], artifact: Path = output_path,
                      source_file: Path = source_path, report_file: Path = report_path) -> None:
                report_file.write_text(json.dumps(value), encoding="utf-8")
                r.validate_report(report_file, parts, "publish", 1, 0,
                                  artifact, source_path=source_file, lane="qualification",
                                  shape=shape)

            check(original)
            passed.append(f"{shape}-qualification-report-valid")
            native = copy.deepcopy(original)
            for key in ("artifact", "artifact_role", "source_artifact"):
                native.pop(key)
            r.validate_report(native, parts, "publish", 1, 0,
                              expected=original, lane="native", shape=shape)
            passed.append(f"{shape}-native-identities-bind-to-qualification")
            if shape == "shared":
                bad_expected_source_logical = copy.deepcopy(original)
                bad_expected_source_logical["source_logical"]["sha256"] = "0" * 64
                _reject("expected-source-logical", lambda: r.validate_report(
                    native, parts, "publish", 1, 0,
                    expected=bad_expected_source_logical, lane="native", shape=shape), rejected)
                bad_expected_output_logical = copy.deepcopy(original)
                bad_expected_output_logical["expected_output_logical"]["sha256"] = "0" * 64
                _reject("expected-output-logical", lambda: r.validate_report(
                    native, parts, "publish", 1, 0,
                    expected=bad_expected_output_logical, lane="native", shape=shape), rejected)
                _reject("qualification-without-archives", lambda: r.validate_report(
                    original, parts, "publish", 1, 0, lane="qualification", shape=shape), rejected)
                mutations = [
                    ("schema", lambda value: value.update(schema="wrong")),
                    ("unknown-top-level", lambda value: value.update(unexpected=True)),
                    ("shape", lambda value: value.update(shape="root")),
                    ("mode", lambda value: value.update(mode="lifecycle")),
                    ("timing-description", lambda value: value.update(timing_description="wrong")),
                    ("generator-count", lambda value: value["generator"].update(part_count=64)),
                    ("generator-pattern", lambda value: value["generator"].update(payload_formula="wrong")),
                    ("generator-depth", lambda value: value["generator"].update(depth=9)),
                    ("warmup", lambda value: value.update(warmup=1)),
                    ("sample-count", lambda value: value.update(samples=[])),
                    ("sample-elapsed", lambda value: value["samples"][0].update(elapsed_ns=0)),
                    ("output-identity", lambda value: value["expected_output"].update(sha256="0" * 64)),
                    ("source-identity", lambda value: value["source"].update(sha256="0" * 64)),
                    ("output-witness", lambda value: value["samples"][0]["hash_witness"].update(
                        output_exact_expected=False)),
                    ("logical-witness", lambda value: value["samples"][0]["hash_witness"].update(
                        logical_sha256="0" * 64)),
                    ("artifact-role", lambda value: value.update(artifact_role="source-archive")),
                    ("source-artifact-identity", lambda value: value["source_artifact"].update(
                        sha256="0" * 64)),
                    ("legacy-sample-fields", lambda value: (
                        value["samples"][0].update(archive=value["samples"][0].pop("output")),
                        value["samples"][0].update(archive_exact_match=True))),
                ]
                for name, mutate in mutations:
                    value = copy.deepcopy(original)
                    mutate(value)
                    _reject("report-" + name, lambda value=value: check(value), rejected)

                damaged_output = bytearray(output)
                damaged_output[-1] ^= 1
                damaged_output_path = root / "damaged-output.zip"
                damaged_output_path.write_bytes(damaged_output)
                _reject("output-byte-corruption",
                        lambda: check(original, damaged_output_path), rejected)
                _reject("target-not-changed",
                        lambda: check(original, source_path), rejected)

                altered = io.BytesIO()
                with zipfile.ZipFile(altered, "w", compression=zipfile.ZIP_DEFLATED,
                                     strict_timestamps=False) as archive:
                    for name, payload in r._metadata_payloads(parts, shape).items():
                        archive.writestr(name, payload)
                    for index in range(parts):
                        payload = r.part_payload(index, index == parts - 1, parts)
                        if index == 0:
                            payload = bytes([payload[0] ^ 1]) + payload[1:]
                        archive.writestr(r.part_archive_name(index, shape), payload)
                altered_path = root / "altered-untouched.zip"
                altered_path.write_bytes(altered.getvalue())
                altered_report = copy.deepcopy(original)
                altered_id = r.archive_manifest(altered.getvalue())["archive"]
                altered_report["artifact"] = altered_id
                altered_report["expected_output"] = altered_id
                altered_report["samples"][0]["output"] = altered_id
                altered_report["samples"][0]["hash_witness"]["output_sha256"] = altered_id["sha256"]
                _reject("untouched-span-mutation",
                        lambda: check(altered_report, altered_path), rejected)

                target = output_manifest["members"][-1]
                local_mutated = bytearray(output)
                local_offset = target["local"]["offset"]
                current_mtime = struct.unpack_from("<H", local_mutated, local_offset + 10)[0]
                struct.pack_into("<H", local_mutated, local_offset + 10, current_mtime ^ 1)
                local_mutated_path = root / "altered-target-local-header.zip"
                local_mutated_path.write_bytes(local_mutated)
                local_mutated_report = copy.deepcopy(original)
                local_mutated_id = r.archive_manifest(bytes(local_mutated))["archive"]
                local_mutated_report["artifact"] = local_mutated_id
                local_mutated_report["expected_output"] = local_mutated_id
                local_mutated_report["samples"][0]["output"] = local_mutated_id
                local_mutated_report["samples"][0]["hash_witness"]["output_sha256"] = local_mutated_id["sha256"]
                _reject("changed-target-local-metadata-mutation",
                        lambda: check(local_mutated_report, local_mutated_path), rejected)


def _profile_fixture(root: Path) -> tuple[Path, Path, list[str]]:
    root.mkdir(parents=True, exist_ok=True)
    numbered = root / "profile.callgrind.1"
    terminal = root / "profile.callgrind"
    command = [
        "taskset", "-c", "12", "valgrind", *r.PROFILE_OPTIONS,
        "/tmp/opc-preservation-probe", "--shape", "shared", "--parts", "4", "--mode", "publish",
        "--samples", "1", "--warmup", "0", "--output", "/tmp/report.json",
    ]
    target = " ".join(command[command.index("/tmp/opc-preservation-probe"):])
    common = [
        "version: 1",
        "creator: callgrind-3.21.0",
        "cmd: " + target,
        "events: Ir",
    ]
    numbered.write_text("\n".join(common + [
        "part: 1",
        "desc: Trigger: --dump-after=" + r.PROFILE_OWNER,
        "summary: 10",
        "totals: 10",
        "fl=probe/src/main.rs",
        "fn=(1) root",
        "1 2",
        "cfn=(2) " + r.PROFILE_OWNER,
        "calls=1 1",
        "1 10",
        "fn=(2) " + r.PROFILE_OWNER,
        "1 3",
        "cfn=(3) child",
        "calls=1 1",
        "1 7",
        "fl=probe/src/archive.rs",
        "fn=(3) child",
        "1 5",
    ]) + "\n", encoding="utf-8")
    terminal.write_text("\n".join(common + [
        "part: 2",
        "desc: Trigger: Program termination",
        "summary: 0",
        "totals: 0",
    ]) + "\n", encoding="utf-8")
    return numbered, terminal, command


def _profile_checks(rejected: list[str], passed: list[str]) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0844-profile-") as temporary:
        root = Path(temporary)
        numbered, terminal, command = _profile_fixture(root)
        result = r.validate_profile(numbered, terminal, command=command)
        assert result["validation"]["file_source_attribution_conserved"]
        assert result["owner"]["name"] == r.PROFILE_OWNER
        assert result["call_rows"][0]["calls"] == 1
        passed.append("callgrind-owner-and-cost-conservation")
        passed.append("callgrind-file-source-attribution")
        passed.append("callgrind-command-normalization")

        mutations = [
            ("duplicate-events", lambda path: path.write_text(
                path.read_text() + "events: Ir\n", encoding="utf-8")),
            ("zero-summary", lambda path: path.write_text(
                path.read_text().replace("summary: 10", "summary: 0"), encoding="utf-8")),
            ("owner-call-count", lambda path: path.write_text(
                path.read_text().replace("calls=1 1", "calls=2 1", 1), encoding="utf-8")),
            ("negative-cost", lambda path: path.write_text(
                path.read_text().replace("1 10", "1 -10"), encoding="utf-8")),
            ("terminal-nonzero", lambda path: path.write_text(
                path.read_text().replace("summary: 0", "summary: 1"), encoding="utf-8")),
            ("command-mismatch", lambda path: path.write_text(
                path.read_text().replace("--mode publish", "--mode lifecycle"), encoding="utf-8")),
            ("missing-source-attribution", lambda path: path.write_text(
                path.read_text().replace("fl=probe/src/main.rs\n", "").replace(
                    "fl=probe/src/archive.rs\n", ""), encoding="utf-8")),
            ("self-cost-conservation", lambda path: path.write_text(
                path.read_text().replace("1 5", "1 6"), encoding="utf-8")),
        ]
        for name, mutate in mutations:
            fresh_numbered, fresh_terminal, fresh_command = _profile_fixture(root / name)
            target = fresh_terminal if name in {"terminal-nonzero", "command-mismatch"} else fresh_numbered
            mutate(target)
            _reject("callgrind-" + name,
                    lambda n=fresh_numbered, t=fresh_terminal, c=fresh_command:
                    r.validate_profile(n, t, command=c), rejected)

        fresh_numbered, fresh_terminal, fresh_command = _profile_fixture(root / "extra-dump")
        extra = fresh_numbered.with_name("profile.callgrind.2")
        extra.write_text(fresh_numbered.read_text(), encoding="utf-8")
        _reject("callgrind-unexpected-second-dump",
                lambda: r.validate_profile(fresh_numbered, fresh_terminal, command=fresh_command),
                rejected)


def _retained_qualification_check(passed: list[str]) -> int:
    count = 0
    for report_path in sorted((HERE / "runs").glob("qualification-*/report.json")):
        report = json.loads(report_path.read_text(encoding="utf-8"))
        folder = report_path.parent
        result = r.validate_report(
            report_path,
            report["generator"]["part_count"],
            report["mode"],
            report["samples_requested"],
            report["warmup"],
            folder / "artifact.zip",
            source_path=folder / "source.zip",
            lane="qualification",
            shape=report["shape"],
        )
        assert result["verification_ok"]
        count += 1
    if count:
        passed.append("retained-qualification-reports")
    return count


def main() -> int:
    passed: list[str] = []
    rejected: list[str] = []
    _zip_checks(rejected, passed)
    _profile_checks(rejected, passed)
    qualification_count = _retained_qualification_check(passed)
    assert r.PROFILE_REPORTS == 0
    assert r.QUALIFICATION_REPORTS == r.HEAP_REPORTS == 64
    assert r.NATIVE_REPORTS == 384
    assert r.EXPECTED_TOTAL_REPORTS == 512
    assert r.EXPECTED_TOTAL_SAMPLES == 11648
    assert len(rejected) >= 22, rejected
    print(json.dumps({
        "status": "pass",
        "reader": str(HERE / "reader.py"),
        "qualification_reports_checked": qualification_count,
        "checks": passed,
        "mutations": [{"name": name, "rejected": True, "error": "ReaderError"}
                      for name in rejected],
        "mutation_count": len(rejected),
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
