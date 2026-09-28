"""Replay the sealed historical allocation-report fixtures before freezing 0827.

This is a pure schema/custody preflight.  It reads the 325 report descriptors
from the completed 0826 packet and their historical seals, then invokes the
same custody report validator that live 0827 captures use.  It never runs
Cargo, a benchmark child, an artifact exporter, or a numerical reader, and it
does not reuse any historical timing value as a new measurement.
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path
from typing import Any

import custody as c


FIXTURES = c.ROOT / "docs/performance/results/change-0826/fixtures.json"
FIXTURE_SEAL = c.ROOT / "docs/performance/results/change-0826/seal.json"
HELPER_PATH = c.ROOT / "tools/perf_allocation_schema.py"
ORIGINS = ("0819", "0821", "0825")
EXPECTED_COUNTS = {
    "0819/qualification": {"reports": 12, "sample_envelopes": 12},
    "0819/native": {"reports": 72, "sample_envelopes": 2160},
    "0819/observer": {"reports": 24, "sample_envelopes": 72},
    "0821/qualification": {"reports": 24, "sample_envelopes": 24},
    "0821/native": {"reports": 144, "sample_envelopes": 4320},
    "0821/observer": {"reports": 48, "sample_envelopes": 144},
    "0825/failed-qualification": {"reports": 1, "sample_envelopes": 1},
}
EXPECTED_CASES = {
    f"{fmt}_real_file_ordinary_save_{phase}"
    for fmt in ("docx", "xlsx", "pptx")
    for phase in ("lifecycle", "edit", "atomic_publish", "counting_publish")
}
EXPECTED_0821_POLICIES = {"default", "file-only", "full", "no-sync"}


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def fixture_descriptor(path: Path) -> dict[str, Any]:
    return {
        "path": str(path.relative_to(c.ROOT)),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }


def historical_case(case_name: str) -> dict[str, str]:
    prefix, _, suffix = case_name.partition("_real_file_ordinary_save_")
    assert prefix in {"docx", "xlsx", "pptx"}
    assert suffix in {"lifecycle", "edit", "atomic_publish", "counting_publish"}
    input_by_format = {
        "docx": "test-data/ooxml/docx/documentProperties.docx",
        "xlsx": "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
        "pptx": "test-data/ooxml/pptx/shapes.pptx",
    }
    return {"case": case_name, "format": prefix, "input": input_by_format[prefix], "phase": suffix}


def derive() -> dict[str, Any]:
    assert FIXTURES.is_file() and not FIXTURES.is_symlink()
    assert FIXTURE_SEAL.is_file() and not FIXTURE_SEAL.is_symlink()
    fixture_seal = read(FIXTURE_SEAL)
    assert fixture_seal["schema"] == "litchi.performance.0826.seal.v1"
    fixture_seal_files = fixture_seal["files"]
    assert fixture_seal_files[str(FIXTURES.relative_to(c.ROOT))] == c.sha(FIXTURES)

    helper = c.allocation_schema_source()
    helper_name = "tools/perf_allocation_schema.py"
    assert fixture_seal_files[helper_name] == helper["sha256"], (
        "allocator helper differs from the sealed 0826 source"
    )
    assert helper["path"] == str(HELPER_PATH.resolve())

    fixtures = read(FIXTURES)
    assert fixtures.get("purpose") == "Historical reports used only to validate schema; no timing reuse"
    rows = fixtures.get("reports")
    assert isinstance(rows, list) and len(rows) == 325
    assert len({row.get("path") for row in rows}) == len(rows)

    seals = {}
    for origin in ORIGINS:
        path = c.ROOT / "docs/performance/results" / f"change-{origin}/seal.json"
        value = read(path)
        assert isinstance(value.get("files"), dict)
        seals[origin] = value

    counts: dict[str, dict[str, int]] = {}
    admission_path = c.ROOT / "docs/performance/results/change-0825/artifact-admission-before.json"
    assert seals["0825"]["files"][str(admission_path.relative_to(c.ROOT))] == c.sha(admission_path)
    admission = read(admission_path)
    assert admission["accepted"] is True and len(admission["selectors"]) == 12
    cases: set[str] = set()
    policy_counts: Counter[str] = Counter()
    for row in rows:
        required = {
            "bytes", "case", "lane", "mode", "origin_packet", "path",
            "samples", "sha256", "warmup",
        }
        assert set(row) == required
        origin = row["origin_packet"]
        assert origin in ORIGINS
        path = c.ROOT / row["path"]
        assert not Path(row["path"]).is_absolute()
        assert path.is_file() and not path.is_symlink()
        assert fixture_descriptor(path) == {
            "path": row["path"], "bytes": row["bytes"], "sha256": row["sha256"]
        }
        assert seals[origin]["files"].get(row["path"]) == row["sha256"]

        case = historical_case(row["case"])
        cases.add(row["case"])
        binary_name = "observer" if row["mode"] == "observer" else "native"
        c.validate_report(
            path, case, row["samples"], row["warmup"], binary_name, admission
        )

        key = f"{origin}/{row['lane']}"
        count = counts.setdefault(key, {"reports": 0, "sample_envelopes": 0})
        count["reports"] += 1
        count["sample_envelopes"] += row["samples"]
        if origin == "0821":
            name = Path(row["path"]).name
            assert "__" in name
            policy = name.rsplit("__", 1)[1].removesuffix(".json")
            assert policy in EXPECTED_0821_POLICIES
            policy_counts[policy] += 1

    assert counts == EXPECTED_COUNTS
    assert cases == EXPECTED_CASES
    assert policy_counts == Counter({
        "default": 54, "file-only": 54, "full": 54, "no-sync": 54,
    })
    return {
        "schema": "litchi.performance.0827.driver-preflight.v1",
        "status": "pass",
        "validator": helper,
        "validator_seal": {
            "path": str(FIXTURE_SEAL.resolve()),
            "sha256": c.sha(FIXTURE_SEAL),
            "helper_sha256": fixture_seal_files[helper_name],
        },
        "fixtures": fixture_descriptor(FIXTURES),
        "historical_artifact_admission": fixture_descriptor(admission_path),
        "historical_seals": {
            origin: {
                "path": str((c.ROOT / "docs/performance/results" / f"change-{origin}/seal.json").resolve()),
                "sha256": c.sha(c.ROOT / "docs/performance/results" / f"change-{origin}/seal.json"),
            }
            for origin in ORIGINS
        },
        "counts": counts,
        "reports": len(rows),
        "sample_envelopes": sum(row["samples"] for row in rows),
        "case_count": len(cases),
        "0821_policy_counts": dict(sorted(policy_counts.items())),
        "new_measurements": 0,
        "performance_claim": None,
    }


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    assert args in (["--write"], ["--check"])
    value = derive()
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    output = c.P / "preflight.json"
    if args == ["--write"]:
        assert not output.exists(), f"refusing to overwrite {output}"
        output.write_text(encoded, encoding="utf-8")
    else:
        assert output.read_text(encoding="utf-8") == encoded
    print("0827 driver preflight PASS: 325 sealed reports / 6,733 sample envelopes; no new timing")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
