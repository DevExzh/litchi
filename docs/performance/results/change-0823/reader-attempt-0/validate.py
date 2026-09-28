"""Final offline validator for the 0823 scanner trial packet.

The validator replays ``analysis.py`` and checks the independent cleanup and
failure-log witnesses when they exist.  It never runs Cargo, a probe,
profiler, decoder, or shell command.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import analysis
import profile_basis
import raw_audit


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise analysis.ReplayError(message)


def cleanup_witness(value: dict[str, Any], build_states: dict[str, Any]) -> None:
    require(value.get("schema") == "litchi.performance.0823.cleanup.v1"
            and value.get("target") == str(analysis.TARGET)
            and value.get("verified") is True
            and value.get("target_absent_after_removal") is True,
            "cleanup target witness changed")
    require(not analysis.TARGET.exists(), "owned target still exists after cleanup")
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


def interrupted_native_witness() -> None:
    directory = PACKET / "native-interrupted-0"
    receipt = analysis.read(directory / "receipt.json")
    require(receipt.get("exit_code") == 1 and receipt.get("base") == analysis.BASE,
            "interrupted native receipt changed")
    for name, digest in receipt["files"].items():
        path = directory / name
        require(path.is_file() and analysis.sha(path) == digest,
                f"interrupted native evidence changed: {name}")
    log = (directory / "native-run.log").read_text()
    require("production source changed during capture" in log,
            "interrupted native failure log missing")
    require(not (directory / "native/complete.json").exists(),
            "interrupted native lane incorrectly marked complete")


def root_artifact(value: dict[str, Any], label: str) -> Path:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label}: artifact malformed")
    path = Path(value["path"])
    if not path.is_absolute():
        path = ROOT / path
    require(path.is_file() and not path.is_symlink(), f"{label}: artifact missing")
    require(path.stat().st_size == value.get("bytes")
            and analysis.sha(path) == value.get("sha256"),
            f"{label}: artifact identity changed")
    return path


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


def codegen_witness() -> None:
    value = analysis.read(PACKET / "codegen/result.json")
    require(value.get("schema") == "litchi.performance.0823.codegen.v1"
            and value.get("symbol") == "litchi_pptx::shape::reader::Scene::read_with",
            "codegen owner changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and [row.get("leg") for row in rows] == ["before", "after"],
            "codegen leg matrix changed")
    states = {leg: analysis.source_state(leg) for leg in ("before", "after")}
    expected_binaries = {leg: states[leg]["binaries"]["real-native"] for leg in states}
    for row in rows:
        leg = row["leg"]
        require(row.get("binary") == expected_binaries[leg], f"codegen {leg}: binary identity changed")
        matched = row.get("demangled_symbol")
        raw = row.get("raw_symbol")
        require(isinstance(matched, list) and len(matched) == 4
                and matched[3] == value["symbol"]
                and isinstance(raw, list) and len(raw) == 4
                and raw[:3] == matched[:3], f"codegen {leg}: symbol identity changed")
        assembly_path = root_artifact(row.get("assembly"), f"codegen {leg} assembly")
        text = assembly_path.read_text(encoding="utf-8")
        require(f"<{value['symbol']}>:" in text and len(text.splitlines()) > 20,
                f"codegen {leg}: assembly body missing")
        call_lines = [line.strip() for line in text.splitlines() if "call" in line]
        expected_calls = {
            name: [line for line in call_lines if name in line]
            for name in ("NsReader<R>::process_event", "NamespaceResolver::resolve_event",
                         "NamespaceResolver::push")
        }
        require(row.get("calls") == expected_calls,
                f"codegen {leg}: retained call counts/body differ")
        commands_path = PACKET / "codegen" / f"{leg}-commands.json"
        commands = analysis.read(commands_path)
        require(isinstance(commands, list) and len(commands) == 3,
                f"codegen {leg}: command receipt count changed")
        for index, command in enumerate(commands):
            require(command.get("exit_code") == 0,
                    f"codegen {leg}/{index}: command failed")
            for name in ("stderr", "stdout_gzip", "stdout"):
                if name in command:
                    root_artifact(command[name], f"codegen {leg}/{index} {name}")


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
    codegen_witness()
    failure_witness()
    interrupted_native_witness()
    if final:
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
        print(f"0823 validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
