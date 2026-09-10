#!/usr/bin/env python3
"""Independent receipt, raw-statistic, and native cached-value verification.

The small wire oracle below reads only the two retained fixture worksheets and
their cached scalar fields. It neither imports the Rust harness verifier nor
implements general XLSB parsing, formatting, or formula evaluation.
"""
import hashlib
import json
import math
from pathlib import Path
import re
import statistics
import struct
import sys
import zipfile

BASE = Path(__file__).resolve().parent
ROOT = BASE.parents[3]
SEED = 0xCBF29CE484222325
MASK = (1 << 64) - 1
CASES = ("indexed_cold", "materialize_cold", "indexed_warm", "materialize_warm")
FIXTURES = {
    "sparse": ROOT / "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
    "large": ROOT / "test-data/ooxml/xlsb/62815.xlsb",
    "synthetic-131072": None,
}


def fnv(data, initial=SEED):
    for byte in data:
        initial = ((initial ^ byte) * 0x100000001B3) & MASK
    return initial


def records(data):
    offset = 0
    while offset < len(data):
        fields = []
        for maximum in (2, 4):
            value = 0
            for index in range(maximum):
                byte = data[offset]
                offset += 1
                value |= (byte & 127) << (7 * index)
                if byte < 128:
                    break
            else:
                raise AssertionError("unterminated BIFF12 header")
            fields.append(value)
        kind, length = fields
        end = offset + length
        assert end <= len(data)
        yield kind, data[offset:end]
        offset = end


def wide(payload, offset):
    units = struct.unpack_from("<I", payload, offset)[0]
    end = offset + 4 + 2 * units
    assert end <= len(payload)
    return payload[offset + 4:end].decode("utf-16le")


def native_cells(path):
    errors = {0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!",
              29: "#NAME?", 36: "#NUM!", 42: "#N/A", 43: "#GETTING_DATA"}
    with zipfile.ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.bin" in archive.namelist():
            strings = [wide(p, 1) for k, p in records(archive.read("xl/sharedStrings.bin")) if k == 19]
        rows = records(archive.read("xl/worksheets/sheet1.bin"))
        cells, in_data, row = {}, False, None
        for kind, payload in rows:
            if kind == 0x91:
                in_data = True
            elif kind == 0x92:
                in_data = False
            elif in_data and kind == 0:
                row = struct.unpack_from("<I", payload)[0]
            elif in_data and (1 <= kind <= 11 or kind == 62):
                assert row is not None
                column = struct.unpack_from("<I", payload)[0]
                if kind == 1:
                    value = (1, b"")
                elif kind == 2:
                    word = struct.unpack_from("<I", payload, 8)[0]
                    if word & 2:
                        number = float(struct.unpack("<i", struct.pack("<I", word))[0] >> 2)
                    else:
                        number = struct.unpack("<d", struct.pack("<Q", (word & ~3) << 32))[0]
                    if word & 1:
                        number /= 100
                    value = (4, struct.pack("<d", number))
                elif kind in (5, 9):
                    value = (4, payload[8:16])
                elif kind in (4, 10):
                    assert payload[8] in (0, 1)
                    value = (2, payload[8:9])
                elif kind in (3, 11):
                    value = (6, errors[payload[8]].encode())
                elif kind == 7:
                    value = (5, strings[struct.unpack_from("<I", payload, 8)[0]].encode())
                else:
                    value = (5, wide(payload, 9 if kind == 62 else 8).encode())
                assert (row, column) not in cells
                cells[row, column] = value
    return cells


def value_digest(row, column, value, initial=SEED):
    initial = fnv(struct.pack("<II", row, column), initial)
    return fnv(bytes([value[0]]) + value[1] if value is not None else b"\0", initial)


def allocation_check(value):
    assert value["invalid"] is False and value["failed"] == 0
    assert value["live_before"] + value["allocated_bytes"] - value["deallocated_bytes"] == value["live_after"]
    assert value["peak_before"] == value["live_before"]
    assert value["peak_after"] >= max(value["peak_before"], value["live_after"])


def main():
    if not __debug__:
        raise RuntimeError("Verification requires Python assertions enabled")
    raw = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else BASE / "performance/raw/final-release"
    gates = BASE / "gates"
    source_manifest = (gates / "source-hashes-before.txt").read_bytes()
    assert source_manifest == (gates / "source-hashes-after.txt").read_bytes()
    for line in source_manifest.decode().splitlines():
        expected, name = line.split(maxsplit=1)
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected, name
    assert (gates / "workspace-Cargo.lock").read_bytes() == (ROOT / "Cargo.lock").read_bytes()
    before = (gates / "git-rev-before.txt").read_bytes()
    assert before == (gates / "git-rev-after.txt").read_bytes()
    build_state = (raw / "release-source-state-before.txt").read_bytes()
    for name in ("release-source-state-after.txt", "source-state-before.txt", "source-state-after.txt"):
        assert build_state == (raw / name).read_bytes()
    for line in build_state.decode().splitlines():
        expected, name = line.split(maxsplit=1)
        if expected == "git_head":
            assert name == before.decode().strip()
        else:
            assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected, name
    manifest_path = raw / "release-build-manifest.txt"
    manifest = dict(line.split(maxsplit=1) for line in manifest_path.read_text().splitlines())
    binary_hash = manifest["profile_binary_sha256"]
    for name in ("binary-sha256-before.txt", "binary-sha256-after.txt", "release-binary-sha256.txt"):
        assert (raw / name).read_text().split()[0] == binary_hash
    manifest_hash = hashlib.sha256(manifest_path.read_bytes()).hexdigest()
    for name in ("external-manifest-sha256-before.txt", "external-manifest-sha256-after.txt"):
        assert (raw / name).read_text().split()[0] == manifest_hash
    assert manifest["source_state_sha256"] == hashlib.sha256(build_state).hexdigest()
    assert manifest["build_log_sha256"] == hashlib.sha256((raw / "release-build.log").read_bytes()).hexdigest()
    binary = Path(manifest["profile_binary"])
    if binary.exists():
        assert hashlib.sha256(binary.read_bytes()).hexdigest() == binary_hash
    counts = {}
    for name in ("xlsb-all-features-all-targets-incremental0.log", "xlsb-all-features-doctests-incremental0.log"):
        results = re.findall(r"test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored;", (gates / name).read_text())
        assert results and all(status == "ok" and failed == "0" for status, _, failed, _ in results)
        counts[name] = {"passed": sum(int(r[1]) for r in results), "ignored": sum(int(r[3]) for r in results)}
    expected_names = {f"{fixture}-{case}-p{process}.json" for fixture in FIXTURES for case in CASES for process in range(1, 4)}
    assert {p.name for p in raw.glob("*-p*.json")} == expected_names
    groups, oracle_counts = {}, {}
    for label, path in FIXTURES.items():
        cells = native_cells(path) if path else None
        oracle_counts[label] = len(cells) if cells is not None else 131072
        identity = None
        for case in CASES:
            reports = []
            for process in range(1, 4):
                report = json.loads((raw / f"{label}-{case}-p{process}.json").read_text())
                assert report["schema"] == "xlsb-binary-index-profile-v1" and report["case"] == case
                assert report["warmup"] == 3 and report["sample_count"] == 30 and report["sheet"] == 0
                for gate in ("semantic_ok", "digest_stable", "source_observation_stable"):
                    assert report[gate] is True
                sample_identity = [report[k] for k in ("input_bytes", "input_digest", "cell_count", "dimensions", "targets")]
                if identity is None:
                    identity = sample_identity
                assert identity == sample_identity
                assert report["cell_count"] == oracle_counts[label]
                if path:
                    source = path.read_bytes()
                    assert report["input_bytes"] == len(source) and report["input_digest"] == fnv(source)
                digest = SEED
                for target in report["targets"]:
                    row, column = target["row"], target["column"]
                    if cells is not None:
                        value = cells.get((row, column))
                    else:
                        index = row * 32 + column
                        value = (4, struct.pack("<d", float(index))) if column < 32 and index < 131072 else None
                    assert target["expected_present"] is (value is not None)
                    assert target["expected_digest"] == value_digest(row, column, value)
                    digest = value_digest(row, column, value, digest)
                samples = report["samples"]
                assert len(samples) == 30
                times = sorted(s["elapsed_ns"] for s in samples)
                assert all(type(t) is int and t > 0 for t in times)
                for quantile in (50, 95, 99):
                    assert report[f"p{quantile}_ns"] == times[math.ceil(30 * quantile / 100) - 1]
                assert abs(report["mean_ns"] - statistics.fmean(times)) <= 0.001
                observation = None
                for sample in samples:
                    assert sample["semantic_ok"] is True and sample["digest"] == digest
                    allocation_check(sample["allocation"])
                    for prefix in ("reads", "cache"):
                        total, start, delta = (sample[f"{prefix}_{part}"] for part in ("total", "before", "operation"))
                        assert total.keys() == start.keys() == delta.keys()
                        assert all(total[k] == start[k] + delta[k] for k in total)
                    current = [sample[k] for k in ("reads_operation", "cache_operation", "retained_entries", "retained_bytes")]
                    if observation is None:
                        observation = current
                    assert current == observation
                if case.endswith("_warm"):
                    allocation_check(report["setup_allocation"])
                reports.append(report)
            groups[f"{label}/{case}"] = {
                "median_process_p50_ns": statistics.median(r["p50_ns"] for r in reports),
                "median_allocated_bytes": statistics.median(s["allocation"]["allocated_bytes"] for r in reports for s in r["samples"]),
                "median_peak_extra_requested_live_bytes": statistics.median(s["allocation"]["peak_after"] - s["allocation"]["live_before"] for r in reports for s in r["samples"]),
            }
    receipt = {"source_files_verified": len(source_manifest.decode().splitlines()), "baseline_commit": before.decode().strip(),
               "release_binary_sha256": binary_hash, "release_build_manifest_sha256": manifest_hash,
               "tests": counts, "reports_verified": 36, "timed_samples_verified": 1080,
               "independent_wire_oracle_cell_counts": oracle_counts, "groups": groups}
    (BASE / "root-verification.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps(receipt, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
