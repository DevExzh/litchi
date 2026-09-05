#!/usr/bin/env python3
"""Verify capture bindings and cross-implementation ZIP64 observations."""
import hashlib
import json
from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def read(path):
    return json.loads((ROOT / path).read_text())


def sha(data):
    return hashlib.sha256(data).hexdigest()


def timing(path):
    text = (ROOT / path).read_text()
    assert re.search(r"Exit status:\s*0\s*$", text)
    return int(re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)[1])


def main():
    identities = read("identities.json")
    for name, expected in identities["probe_sources"].items():
        assert sha((ROOT / name).read_bytes()) == expected, name
    for name, expected in identities["candidate_sources"].items():
        result = subprocess.run(["git", "show", f'{identities["candidate_revision"]}:{name}'],
                                cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        actual = result.stdout if result.returncode == 0 else (REPO / name).read_bytes()
        assert sha(actual) == expected, name
    expected_order = [(leg, mode, payload, size) for leg in ("A1", "B1", "B2", "A2")
                      for mode in ("borrowed", "owned") for payload in ("zeros", "mixed")
                      for size in (16384, 1048576)]
    for directory, samples, warmups in ((Path(), 300, 30), (Path("followup"), 1000, 100)):
        capture = read(directory / "capture.json")
        for role in ("control", "candidate"):
            item = capture["identities"][role]
            assert item["sha256"] == identities["binaries_sha256"][role]
            assert item["revision"] == identities[f"{role}_revision"]
        assert len(capture["runs"]) == len(expected_order)
        previous = ""
        for run, (leg, mode, payload, size) in zip(capture["runs"], expected_order):
            assert run["stem"] == f"{leg}-{mode}-{payload}-{size}"
            assert run["exit_code"] == 0 and run["stderr"] == ""
            assert previous <= run["started"] <= run["finished"]
            previous = run["finished"]
            assert run["argv"][-6:] == ["guard", mode, payload, str(size), str(samples), str(warmups)]
            timing(directory / "guards" / f'{run["stem"]}.time.txt')
    observed = read("observations/capture.json")
    assert observed["binary_sha256"] == identities["binaries_sha256"]["candidate"]
    assert len(observed["runs"]) == 18
    previous = ""
    for run in observed["runs"]:
        assert run["exit_code"] == 0 and run["stderr"] == ""
        assert previous <= run["started"] <= run["finished"]
        previous = run["finished"]
    for mode in ("borrowed", "owned"):
        for size in (67108864, 268435456, 4294967297):
            stem = f"observations/{mode}-{size}"
            writer, rust, python = (read(f"{stem}-{kind}.json") for kind in ("write", "readback", "python"))
            assert writer["input_bytes"] == size
            assert writer["output_bytes"] == rust["archive_bytes"] == python["archive_bytes"]
            assert len(rust["entries"]) == len(python["entries"]) == 1
            assert rust["entries"][0]["uncompressed_bytes"] == size
            for key in ("name", "compressed_bytes", "uncompressed_bytes", "crc32"):
                assert rust["entries"][0][key] == python["entries"][0][key]
            assert python["entries"][0]["flags"] & 8
            assert rust["archive_zip64"] == (size >= 2**32 - 1)
            for kind in ("write", "readback", "python"):
                timing(f"{stem}-{kind}.time.txt")
    large, python = read("large-output.json"), read("large-output-python.json")
    executable = subprocess.check_output(["zstd", "-dc", str(ROOT / "large-output-probe.zst")])
    assert sha(executable) == identities["binaries_sha256"]["large-output"]
    assert large["verified"] and large["archive_zip64"] and large["descriptor"]
    assert large["compressed_bytes"] > 2**32 - 1
    assert large["input_bytes"] == 2**32 + 1
    assert large["output_bytes"] == python["archive_bytes"]
    assert large["input_crc32"] == python["entries"][0]["crc32"]
    assert large["compressed_bytes"] == python["entries"][0]["compressed_bytes"]
    assert large["input_bytes"] == python["entries"][0]["uncompressed_bytes"]
    timing("large-output.time.txt")
    timing("large-output-python.time.txt")
    producer, rust = read("python-producer.json"), read("candidate-python-readback.json")
    assert len(producer["entries"]) == len(rust["entries"]) == 4
    assert not rust["archive_zip64"] and rust["has_data_descriptors"]
    for left, right in zip(producer["entries"], rust["entries"]):
        for key in ("name", "compressed_bytes", "uncompressed_bytes", "crc32"):
            assert left[key] == right[key]
    assert rust["entries"][-1]["uncompressed_bytes"] == 2**32
    assert identities["control_refused_output_bytes"] == 0
    tests = (ROOT / "checks/final-tests.log").read_text()
    results = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", tests)
    assert len(results) == 41
    assert [sum(int(row[i]) for row in results) for i in range(3)] == [1222, 0, 2]
    for name in ("large-streaming-test", "independent-opc"):
        assert "test result: ok. 1 passed; 0 failed" in (ROOT / f"checks/{name}.log").read_text()
    for name in ("zip-fuzz", "opc-fuzz"):
        log = (ROOT / f"checks/{name}.log").read_text()
        assert "Done 1000 runs" in log and "ERROR: AddressSanitizer" not in log
    for name, expected in identities["profile_originals"].items():
        data = subprocess.check_output(["zstd", "-dc", str(ROOT / "profile" / f"{name}.zst")])
        assert sha(data) == expected["sha256"] and len(data) == expected["bytes"]
    print("OK: source/capture bindings, 64 guard processes, 18 size observations, physical ZIP64 output and independent OPC readback")


if __name__ == "__main__":
    main()
