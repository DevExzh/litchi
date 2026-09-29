"""Mutation tests for the independent 0842 evidence reader.

The tests copy retained qualification JSON and ZIP artifacts into temporary
files before mutating them.  Frozen packet inputs are never edited.  The small
ZIP fixtures exercise the raw compressed-range and CRC checks without starting
Cargo, a probe, or a workload.
"""

from __future__ import annotations

import copy
import io
import json
import math
import tempfile
import warnings
import zipfile
from pathlib import Path

import reader as r


HERE = Path(__file__).resolve().parent


def _expect_reject(name: str, callback, failures: list[str]) -> None:
    try:
        callback()
    except r.ReaderError:
        failures.append(name)
    else:
        raise AssertionError(f"accepted malformed mutation: {name}")


def _zip_fixture() -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED, strict_timestamps=False) as archive:
        archive.writestr("one.txt", b"one")
        archive.writestr("two.txt", b"two" * 20)
    return stream.getvalue()


def _archive_tests(passed: list[str]) -> None:
    data = _zip_fixture()
    manifest = r.archive_manifest(data)
    assert manifest["member_count"] == 2
    assert [row["name"] for row in manifest["members"]] == ["one.txt", "two.txt"]
    assert all(row["raw"]["bytes"] > 0 for row in manifest["members"])
    assert all(row["central"]["compress_type"] == zipfile.ZIP_DEFLATED
               for row in manifest["members"])
    passed.append("zip-logical-crc-raw-metadata")

    duplicate = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        with zipfile.ZipFile(duplicate, "w", compression=zipfile.ZIP_STORED) as archive:
            archive.writestr("same", b"a")
            archive.writestr("same", b"b")
    _expect_reject("zip-duplicate-member", lambda: r.archive_manifest(duplicate.getvalue()), passed)

    traversal = io.BytesIO()
    with zipfile.ZipFile(traversal, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr("../escape", b"x")
    _expect_reject("zip-unsafe-member", lambda: r.archive_manifest(traversal.getvalue()), passed)

    damaged = bytearray(data)
    member = manifest["members"][1]
    offset = member["local"]["offset"] + 30 + member["local"]["name_bytes"] + member["local"]["extra_bytes"]
    damaged[offset] ^= 0x01
    changed_manifest = r.archive_manifest(bytes(damaged))
    assert changed_manifest["members"][1]["raw"]["sha256"] != manifest["members"][1]["raw"]["sha256"]
    _expect_reject(
        "zip-compressed-byte-metadata-mismatch",
        lambda: r.compare_archive_manifest(changed_manifest, manifest),
        passed,
    )


def _find_qualification() -> tuple[Path, Path] | None:
    candidates = []
    for report in sorted((HERE / "runs").rglob("report.json")) if (HERE / "runs").is_dir() else []:
        try:
            value = json.loads(report.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        if value.get("samples_requested") != r.QUALIFICATION_SAMPLES or value.get("warmup") != r.QUALIFICATION_WARMUP:
            continue
        if value.get("shape") not in r.SHAPES or value.get("mode") not in r.MODES:
            continue
        artifact = report.parent / "artifact.zip"
        if artifact.is_file() and not artifact.is_symlink():
            candidates.append((report, artifact))
            continue
        zips = sorted(report.parent.glob("*.zip"))
        if len(zips) == 1 and zips[0].is_file() and not zips[0].is_symlink():
            candidates.append((report, zips[0]))
    return candidates[0] if candidates else None


def _report_mutation_tests(passed: list[str]) -> None:
    found = _find_qualification()
    if found is None:
        raise AssertionError("qualification artifacts required for the full mutation suite")
    report_path, artifact_path = found
    original = json.loads(report_path.read_text(encoding="utf-8"))
    shape = original["shape"]
    mode = original["mode"]
    with tempfile.TemporaryDirectory(prefix="litchi-0842-reader-") as temporary:
        root = Path(temporary)
        report_copy = root / "report.json"
        artifact_copy = root / "artifact.zip"
        artifact_copy.write_bytes(artifact_path.read_bytes())

        def check(value, artifact: Path = artifact_copy, samples: int = 1, warmup: int = 0) -> None:
            report_copy.write_text(json.dumps(value, allow_nan=True), encoding="utf-8")
            r.validate_report(report_copy, shape, mode, samples, warmup, artifact_path=artifact)

        check(original)
        passed.append("qualification-valid")
        mutations = [
            ("missing-schema", lambda value: value.pop("schema")),
            ("wrong-schema", lambda value: value.update(schema="wrong")),
            ("wrong-shape", lambda value: value.update(shape="xml" if shape != "xml" else "tiny")),
            ("wrong-part-count", lambda value: value.update(part_count=3)),
            ("wrong-member-dimension", lambda value: value.update(member_payload_bytes=1)),
            ("wrong-warmup", lambda value: value.update(warmup=1)),
            ("wrong-sample-count", lambda value: value.update(samples=[])),
            ("zero-elapsed", lambda value: value["samples"][0].update(elapsed_ns=0)),
            ("nonfinite-elapsed", lambda value: value["samples"][0].update(elapsed_ns=float("nan"))),
            ("wrong-source-hash", lambda value: value["samples"][0].update(input_archive_sha256="0" * 64)),
            ("false-archive-witness", lambda value: value["samples"][0].update(archive_exact_match=False)),
            ("false-logical-witness", lambda value: value["samples"][0].update(logical_exact_match=False)),
            ("wrong-report-artifact", lambda value: value["artifact"].update(sha256="0" * 64)),
        ]
        for name, mutate in mutations:
            value = copy.deepcopy(original)
            mutate(value)
            _expect_reject(name, lambda value=value: check(value), passed)

        damaged = bytearray(artifact_copy.read_bytes())
        damaged[-1] ^= 1
        damaged_copy = root / "damaged.zip"
        damaged_copy.write_bytes(damaged)
        _expect_reject("artifact-byte-corruption", lambda: check(original, damaged_copy), passed)


def main() -> int:
    passed: list[str] = []
    _archive_tests(passed)
    _report_mutation_tests(passed)
    assert len(passed) >= 5, passed
    print(json.dumps({"status": "pass", "checks": passed}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
