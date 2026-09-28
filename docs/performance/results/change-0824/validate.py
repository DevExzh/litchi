"""Final offline validator for the 0824 opened XML/transaction trial packet.

The validator replays the packet readers and checks independent arithmetic,
profile, disposition, and cleanup witnesses.  It never runs Cargo, a probe,
profiler, decoder, or shell command.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import analysis
import fixture_basis
import profile_basis
import raw_audit


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise analysis.ReplayError(message)


def cleanup_witness(value: dict[str, Any], build_states: dict[str, Any]) -> None:
    require(value.get("schema") == "litchi.performance.0824.cleanup.v1"
            and value.get("target") == str(analysis.TARGET)
            and value.get("verified") is True
            and value.get("target_absent_after_removal") is True,
            "cleanup target witness changed")
    require(not analysis.TARGET.exists(), "owned target still exists after cleanup")
    require(isinstance(value.get("files"), int) and value["files"] >= 8
            and isinstance(value.get("logical_bytes"), int)
            and value["logical_bytes"] >= 0,
            "cleanup inventory is missing")
    expected = []
    for leg in ("before", "after"):
        for name, descriptor in build_states[leg]["binaries"].items():
            expected.append({"path": descriptor["path"], "bytes": descriptor["bytes"],
                             "sha256": descriptor["sha256"]})
    removed = value.get("binaries")
    require(isinstance(removed, list), "cleanup removed_binaries is missing")
    canonical = lambda rows: sorted(
        (row.get("path"), row.get("bytes"), row.get("sha256"))
        for row in rows if isinstance(row, dict))
    require(len(removed) == len(expected) == 8 and canonical(removed) == canonical(expected),
            "cleanup does not contain the exact eight binary witnesses")


def disposition_witness(value: dict[str, Any], analysis_value: dict[str, Any]) -> None:
    """Bind either policy disposition to the exact two-file archive selected."""
    require(value.get("schema") == "litchi.performance.0824.disposition.v1",
            "disposition schema changed")
    decision = value.get("decision")
    require(decision in {"adopt", "reject"}, "disposition decision is invalid")
    eligible = analysis_value["decision"].get("adoption_eligible")
    require(isinstance(eligible, bool), "analysis policy eligibility is malformed")
    if decision == "adopt":
        require(eligible, "adoption disposition is ineligible")
    elif eligible:
        reason = value.get("reason")
        require(isinstance(reason, str) and reason.strip(),
                "eligible rejection has no explicit reason")
    require(value.get("analysis_sha256") == analysis.sha(PACKET / "analysis.json"),
            "disposition analysis binding changed")
    for field, expected in (
        ("benefit_satisfied", analysis_value["decision"]["benefit_satisfied"]),
        ("latency_vetoes", len(analysis_value["decision"]["latency_vetoes"])),
        ("allocation_resource_violations",
         len(analysis_value["decision"]["allocation_resource_violations"])),
    ):
        if field in value:
            require(value[field] == expected, f"disposition {field} changed")
    reviews = analysis_value["decision"].get("rss_reviews", [])
    if reviews:
        review = value.get("rss_review")
        if isinstance(review, dict):
            require(any(isinstance(item, str) and item.strip() for item in review.values()),
                    "RSS review disposition is empty")
        else:
            require(isinstance(review, str) and review.strip(),
                    "RSS review disposition is missing")

    allowlist = analysis_value["source"]["changed_files"]
    require(allowlist == [
        "crates/litchi-pptx/src/opened/transaction.rs",
        "crates/litchi-pptx/src/opened/xml.rs",
    ], "disposition source allowlist changed")
    selected = value.get("selected_archive", value.get("selected_source"))
    selected_kind = "after" if decision == "adopt" else "before"
    expected = {
        source: f"candidate/{selected_kind}/{Path(source).name}"
        for source in allowlist
    }
    if isinstance(selected, str):
        if selected in {"before", "after"}:
            require(selected == selected_kind, "disposition selected archive disagrees with decision")
            selected = expected
        elif selected.startswith("candidate/"):
            selected = {source: selected for source in allowlist}
        else:
            raise analysis.ReplayError("disposition selected archive is malformed")
    require(isinstance(selected, dict), "disposition selected archive is missing")
    normalized = {}
    for source, archive in expected.items():
        chosen = selected.get(source)
        require(chosen == archive, f"disposition archive selection changed: {source}")
        path = (PACKET / chosen).resolve()
        require(path.is_file() and not path.is_symlink()
                and path == (PACKET / "candidate" / selected_kind / Path(source).name).resolve(),
                f"disposition archive missing: {source}")
        normalized[source] = path
    for source, path in normalized.items():
        production = ROOT / source
        require(production.is_file() and not production.is_symlink()
                and production.read_bytes() == path.read_bytes()
                and analysis.sha(production) == analysis.sha(path),
                f"production source does not match selected archive: {source}")


def failure_witness() -> None:
    """Every retained failed reader invocation must name its own log."""
    paths = (PACKET / "reader-failures.json", PACKET / "reader-failure-attestation.json")
    existing = [path for path in paths if path.is_file() and not path.is_symlink()]
    if not existing:
        return
    value = analysis.read(existing[0])
    failures = value.get("failures", value.get("attempts", []))
    require(isinstance(failures, list), "reader failure witness is malformed")
    seen = set()
    for index, row in enumerate(failures):
        require(isinstance(row, dict), f"reader failure {index} is malformed")
        log = row.get("log") or row.get("console_log")
        if isinstance(log, dict):
            path = analysis.artifact(log, f"reader failure {index} log")
        else:
            require(isinstance(log, str), f"reader failure {index} has no log")
            path = Path(log)
            if not path.is_absolute():
                path = (PACKET / path).resolve()
            require(path.is_file() and not path.is_symlink(),
                    f"reader failure {index} log is missing")
        require(str(path) not in seen, f"reader failure {index} reuses a log")
        seen.add(str(path))
        if "exit_code" in row:
            require(isinstance(row["exit_code"], int) and row["exit_code"] != 0,
                    f"reader failure {index} is not a failed invocation")


def independent_readers() -> None:
    raw = raw_audit.derive()
    raw_audit.compare(raw)
    raw_path = PACKET / "raw-audit.json"
    require(raw_path.is_file() and raw_path.read_text(encoding="utf-8")
            == json.dumps(raw, indent=2, sort_keys=True) + "\n",
            "raw-audit.json does not replay byte-for-byte")

    profile = profile_basis.derive()
    profile_path = PACKET / "profile-basis.json"
    require(profile_path.is_file() and profile_path.read_text(encoding="utf-8")
            == json.dumps(profile, indent=2, sort_keys=True) + "\n",
            "profile-basis.json does not replay byte-for-byte")

    fixture = fixture_basis.derive()
    fixture_path = PACKET / "fixture-basis.json"
    require(fixture_path.is_file() and fixture_path.read_text(encoding="utf-8")
            == json.dumps(fixture, indent=2, sort_keys=True) + "\n",
            "fixture-basis.json does not replay byte-for-byte")


def run(final: bool) -> dict[str, Any]:
    value = analysis.replay(check=True)
    require(value["counts"] == {"qualification_reports": 38,
                                 "qualification_samples": 38,
                                 "native_reports": 228,
                                 "native_samples": 6840,
                                 "allocation_reports": 76,
                                 "allocation_samples": 228,
                                 "reports": 342, "samples": 7106},
            "final cardinality changed")
    require(len(value["rows"]) == 19, "not every public workflow row is retained")
    require(value["verification"]["no_aggregate_row_hidden"] is True,
            "analysis row-visibility proof missing")
    require(value["verification"]["serial_chronology_checked"] is True,
            "analysis serial chronology proof missing")
    independent_readers()
    failure_witness()
    import failure_audit
    failure_audit.check()
    if final:
        disposition = analysis.read(PACKET / "disposition.json")
        disposition_witness(disposition, value)
        cleanup = PACKET / "cleanup.json"
        require(cleanup.is_file() and not cleanup.is_symlink(),
                "final validation requires cleanup.json")
        before = analysis.source_state("before")
        after = analysis.source_state("after")
        cleanup_witness(analysis.read(cleanup), {"before": before, "after": after})
    return value


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    require(args in ([], ["--final"]), "use no arguments or --final")
    value = run(final=args == ["--final"])
    print(json.dumps({"status": "accepted", "reports": value["counts"]["reports"],
                      "samples": value["counts"]["samples"], "final": args == ["--final"]},
                     sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (analysis.ReplayError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"0824 validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
