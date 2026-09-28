"""Replay pinned historical report schemas without measuring or comparing speed."""
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
sys.path.insert(0, str(ROOT))
from tools import perf_allocation_schema as validator


def read(path):
    return json.loads(path.read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def descriptor(path):
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def derive():
    fixtures = read(P / "fixtures.json")
    assert fixtures["schema"] == "litchi.performance.0826.fixtures.v1"
    rows = fixtures["reports"]
    assert len(rows) == len({r["path"] for r in rows}) == 325
    seals = {packet: read(P.parent / f"change-{packet}/seal.json") for packet in ("0819", "0821", "0825")}
    counts = {}
    cases = set()
    failed_fixture = None
    for row in rows:
        path = ROOT / row["path"]
        actual = descriptor(path)
        assert actual == {key: row[key] for key in ("path", "bytes", "sha256")}
        assert seals[row["origin_packet"]]["files"][row["path"]] == row["sha256"]
        report = read(path)
        validator.validate_report(report, expected_case=row["case"], expected_samples=row["samples"],
                                  expected_warmup=row["warmup"], mode=row["mode"])
        key = row["origin_packet"] + "/" + row["lane"]
        count = counts.setdefault(key, {"reports": 0, "sample_envelopes": 0})
        count["reports"] += 1
        count["sample_envelopes"] += row["samples"]
        cases.add(row["case"])
        if row["origin_packet"] == "0825":
            failed_fixture = report
    assert counts == {
        "0819/qualification": {"reports": 12, "sample_envelopes": 12},
        "0819/native": {"reports": 72, "sample_envelopes": 2160},
        "0819/observer": {"reports": 24, "sample_envelopes": 72},
        "0821/qualification": {"reports": 24, "sample_envelopes": 24},
        "0821/native": {"reports": 144, "sample_envelopes": 4320},
        "0821/observer": {"reports": 48, "sample_envelopes": 144},
        "0825/failed-qualification": {"reports": 1, "sample_envelopes": 1},
    }
    assert cases == {f"{fmt}_real_file_ordinary_save_{phase}" for fmt in ("docx", "xlsx", "pptx")
                     for phase in ("lifecycle", "edit", "atomic_publish", "counting_publish")}
    old = P.parent / "change-0825/custody.py"
    assert seals["0825"]["files"][str(old.relative_to(ROOT))] == sha(old)
    spec = importlib.util.spec_from_file_location("frozen_0825_custody", old)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    try:
        module.check_alloc(failed_fixture["results"][0]["operation_metrics"]["allocation"], "regression")
    except AssertionError as error:
        assert str(error) == "regression.allocation_calls"
    else:
        raise AssertionError("original scalar checker no longer reproduces its failure")
    return {"schema": "litchi.performance.0826.preflight.v1", "status": "pass",
            "validator": descriptor(ROOT / "tools/perf_allocation_schema.py"),
            "driver": descriptor(Path(__file__)), "fixtures": descriptor(P / "fixtures.json"),
            "original_checker": descriptor(old), "original_rejection_reproduced": True,
            "counts": counts, "reports": len(rows), "sample_envelopes": sum(r["samples"] for r in rows),
            "case_count": len(cases), "new_measurements": 0, "performance_claim": None}


if __name__ == "__main__":
    assert sys.argv[1:] in (["--write"], ["--check"])
    value = derive()
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    out = P / "preflight.json"
    if sys.argv[1] == "--write":
        with out.open("x") as stream:
            stream.write(encoded)
    else:
        assert out.read_text() == encoded
    print("0826 schema preflight PASS: 325 pinned reports / 6,733 sample envelopes; no new timing")
