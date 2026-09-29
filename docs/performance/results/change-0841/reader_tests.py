"""Offline custody and mutation checks for the 0841 evidence readers.

The harness consumes retained qualification artifacts and synthetic Callgrind
text only.  It never starts the probe, invokes a build, or runs a project test
command.  Mutations live under a temporary packet directory and disappear
before this script exits.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any, Callable

import readers as r


PACKET = Path(__file__).resolve().parent
RESULT = PACKET / "reader-tests.json"
QUALIFICATION_PREFIX = "qualification-00-"
EXPECTED_CASES = {(shape, mode) for shape in r.SHAPES for mode in r.MODES}


def fail(message: str) -> None:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def qualification_reports() -> list[tuple[Path, dict[str, Any]]]:
    rows: list[tuple[Path, dict[str, Any]]] = []
    for path in sorted((PACKET / "runs").glob(f"{QUALIFICATION_PREFIX}*/report.json")):
        report = read_json(path)
        require(isinstance(report, dict), f"qualification report is not an object: {path}")
        rows.append((path, report))
    require(len(rows) == 20, f"qualification report count changed: {len(rows)}")
    identities = {(row.get("shape"), row.get("mode")) for _, row in rows}
    require(identities == EXPECTED_CASES, "qualification shape/mode matrix changed")
    for case in EXPECTED_CASES:
        selected = [row for row in rows if (row[1].get("shape"), row[1].get("mode")) == case]
        require(len(selected) == 2, f"qualification leg count changed for {case}")
    return rows


def validate_qualification() -> tuple[list[dict[str, Any]], tuple[Path, dict[str, Any]]]:
    rows = qualification_reports()
    validated: list[dict[str, Any]] = []
    for path, report in rows:
        shape = report["shape"]
        mode = report["mode"]
        artifact = path.parent / "artifact.ppt"
        result = r.validate_report(path, shape, mode, 1, 0, artifact_path=artifact)
        validated.append({
            "path": str(path),
            "shape": shape,
            "mode": mode,
            "report_sha256": sha256_file(path),
            "artifact_sha256": result["artifact_identity"]["sha256"],
            "source": result["source"],
            "output": result["output"],
            "verification_ok": result["verification_ok"],
        })
    first = rows[0]
    return validated, first


def expect_reject(label: str, action: Callable[[], Any]) -> dict[str, str]:
    try:
        action()
    except Exception as error:
        return {"name": label, "rejected": True, "error": type(error).__name__}
    fail(f"mutation was accepted: {label}")


def write_report(path: Path, value: dict[str, Any]) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def report_mutations(
    temporary: Path, source_path: Path, source_report: dict[str, Any],
) -> list[dict[str, str]]:
    shape = source_report["shape"]
    mode = source_report["mode"]
    artifact = source_path.parent / "artifact.ppt"
    mutations: list[tuple[str, Callable[[dict[str, Any]], None]]] = []

    def change_schema(value: dict[str, Any]) -> None:
        value["schema"] = "litchi.ppt.profile-probe-0841.mutated"

    def change_owner(value: dict[str, Any]) -> None:
        value["profile_owner"] = "ppt_profile_probe::mutated"

    def change_source(value: dict[str, Any]) -> None:
        value["source"]["sha256"] = "0" * 64

    def change_output(value: dict[str, Any]) -> None:
        value["samples"][0]["output"]["sha256"] = "1" * 64

    def change_semantic(value: dict[str, Any]) -> None:
        value["samples"][0]["verification"]["semantic_check"] = False

    def change_actual_length(value: dict[str, Any]) -> None:
        value["samples"][0]["verification"]["semantic_text_bytes"] += 1

    def change_counter(value: dict[str, Any]) -> None:
        value["allocator"]["counter_revision"] = "mutated-counter"

    def change_allocator(value: dict[str, Any]) -> None:
        value["allocator"]["allocator"] = "mutated-allocator"

    def change_elapsed(value: dict[str, Any]) -> None:
        value["samples"][0]["elapsed_ns"] = 0

    mutations.extend([
        ("report-schema", change_schema),
        ("profile-owner", change_owner),
        ("source-digest", change_source),
        ("sample-output-digest", change_output),
        ("verification-boolean", change_semantic),
        ("verification-actual-length", change_actual_length),
        ("allocator-counter-identity", change_counter),
        ("allocator-identity", change_allocator),
        ("nonpositive-elapsed", change_elapsed),
    ])
    results: list[dict[str, str]] = []
    for name, mutate in mutations:
        value = copy.deepcopy(source_report)
        mutate(value)
        target = temporary / f"{name}.json"
        write_report(target, value)
        results.append(expect_reject(
            name,
            lambda target=target: r.validate_report(
                target, shape, mode, 1, 0,
                artifact_path=artifact,
            ),
        ))

    artifact_copy = temporary / "artifact-mutated.ppt"
    shutil.copyfile(artifact, artifact_copy)
    with artifact_copy.open("ab") as stream:
        stream.write(b"mutation")
    results.append(expect_reject(
        "artifact-bytes",
        lambda: r.validate_report(
            source_path, shape, mode, 1, 0,
            artifact_path=artifact_copy,
        ),
    ))
    return results


def synthetic_callgrind(temporary: Path) -> tuple[Path, Path]:
    owner = r.PROFILE_OWNER
    command = (
        "/usr/bin/valgrind --tool=callgrind --collect-atstart=no "
        f"--toggle-collect={owner} --zero-before={owner} "
        f"--dump-after={owner} /tmp/ppt-profile-probe"
    )
    positive = "\n".join([
        "version: 1",
        "creator: callgrind-3.21",
        f"cmd: {command}",
        "part: 1",
        f"desc: Trigger: --dump-after={owner}",
        "events: Ir",
        "summary: 10",
        "totals: 10",
        "fn=(1) main",
        f"cfn=(2) {owner}",
        "calls=1 1",
        "* 10",
        "fn=(2)",
        "10 10",
        "",
    ])
    terminal = "\n".join([
        "version: 1",
        "creator: callgrind-3.21",
        f"cmd: {command}",
        "part: 2",
        "desc: Trigger: Program termination",
        "events: Ir",
        "summary: 0",
        "totals: 0",
        "",
    ])
    numbered = temporary / "synthetic.callgrind.1"
    zero = temporary / "synthetic.callgrind"
    numbered.write_text(positive, encoding="utf-8")
    zero.write_text(terminal, encoding="utf-8")
    return numbered, zero


def callgrind_mutations(temporary: Path, numbered: Path, terminal: Path) -> list[dict[str, str]]:
    base_numbered = numbered.read_text(encoding="utf-8")
    base_terminal = terminal.read_text(encoding="utf-8")
    results: list[dict[str, str]] = []

    bad_events = temporary / "bad-events.callgrind.1"
    bad_events.write_text(base_numbered.replace("events: Ir", "events: Ir\nevents: Ir"),
                          encoding="utf-8")
    results.append(expect_reject("callgrind-duplicate-events",
                                lambda: r.parse_raw(bad_events)))

    bad_summary = temporary / "bad-summary.callgrind.1"
    bad_summary.write_text(base_numbered.replace("summary: 10", "summary: 0"),
                           encoding="utf-8")
    results.append(expect_reject(
        "callgrind-zero-summary",
        lambda: r.validate_profile(bad_summary, terminal, r.PROFILE_OWNER),
    ))

    bad_calls = temporary / "bad-calls.callgrind.1"
    bad_calls.write_text(base_numbered.replace("calls=1 1", "calls=2 1"),
                         encoding="utf-8")
    results.append(expect_reject(
        "callgrind-owner-call-count",
        lambda: r.validate_profile(bad_calls, terminal, r.PROFILE_OWNER),
    ))

    bad_command = temporary / "bad-command.callgrind.1"
    bad_command.write_text(base_numbered.replace("--zero-before=" + r.PROFILE_OWNER, ""),
                           encoding="utf-8")
    results.append(expect_reject(
        "callgrind-zero-before-command",
        lambda: r.validate_profile(bad_command, terminal, r.PROFILE_OWNER),
    ))

    bad_terminal = temporary / "bad-terminal.callgrind"
    bad_terminal.write_text(base_terminal.replace("summary: 0", "summary: 1"),
                            encoding="utf-8")
    results.append(expect_reject(
        "callgrind-terminal-nonzero",
        lambda: r.validate_profile(numbered, bad_terminal, r.PROFILE_OWNER),
    ))
    return results


def run() -> dict[str, Any]:
    validated, (source_path, source_report) = validate_qualification()
    with tempfile.TemporaryDirectory(prefix=".reader-tests-0841-", dir=PACKET) as raw:
        temporary = Path(raw)
        mutations = report_mutations(temporary, source_path, source_report)
        numbered, terminal = synthetic_callgrind(temporary)
        profile = r.validate_profile(numbered, terminal, r.PROFILE_OWNER)
        profile_mutations = callgrind_mutations(temporary, numbered, terminal)
    require(profile["summary_ir"] == 10, "synthetic Callgrind summary changed")
    reader_path = PACKET / "readers.py"
    return {
        "status": "pass",
        "reader": {"path": str(reader_path.resolve()),
                   "sha256": sha256_file(reader_path)},
        "reader_tests": {"path": str(Path(__file__).resolve()),
                         "sha256": sha256_file(Path(__file__).resolve())},
        "qualification": {
            "reports": len(validated),
            "cases": len(EXPECTED_CASES),
            "rows": validated,
        },
        "callgrind": {
            "synthetic_summary_ir": profile["summary_ir"],
            "owner": profile["owner"]["name"],
            "validation": profile["validation"],
        },
        "mutations": mutations + profile_mutations,
        "mutation_count": len(mutations) + len(profile_mutations),
    }


def check_existing() -> None:
    require(RESULT.is_file() and not RESULT.is_symlink(),
            f"missing reader test receipt: {RESULT}")
    value = read_json(RESULT)
    require(isinstance(value, dict) and value.get("status") == "pass",
            "reader test receipt is not a passing object")
    reader = value.get("reader")
    require(isinstance(reader, dict) and reader.get("path") == str((PACKET / "readers.py").resolve()),
            "reader receipt path changed")
    require(reader.get("sha256") == sha256_file(PACKET / "readers.py"),
            "reader receipt hash changed")
    reader_tests = value.get("reader_tests")
    require(isinstance(reader_tests, dict)
            and reader_tests.get("path") == str(Path(__file__).resolve()),
            "reader-tests receipt path changed")
    require(reader_tests.get("sha256") == sha256_file(Path(__file__).resolve()),
            "reader-tests receipt hash changed")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true",
                        help="check the existing receipt without rerunning mutations")
    args = parser.parse_args()
    if args.check:
        check_existing()
        print("reader tests PASS (existing receipt)")
        return
    result = run()
    try:
        with RESULT.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2)
            stream.write("\n")
    except FileExistsError as error:
        raise RuntimeError(
            f"reader test receipt already exists; remove it before a fresh run: {RESULT}"
        ) from error
    print(f"reader tests PASS ({result['mutation_count']} mutations)")


if __name__ == "__main__":
    main()
