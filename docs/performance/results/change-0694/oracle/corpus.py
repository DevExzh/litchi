#!/usr/bin/env python3
"""Build and run a bounded exact-output MCE differential corpus.

The driver keeps corpus construction separate from the native probe.  It
extracts a small, deterministic set of XML members from checked-in DOCX/XLSX/
PPTX archives, adds deterministic malformed and resource-boundary cases, and
invokes the baseline and candidate binaries in the same order.  Equality is
byte-for-byte equality of the probe's stdout record.  The probe record itself
contains an output SHA-256/length/report or the typed Rust error Debug value;
the driver never canonicalizes XML.

usage:
    corpus.py BASELINE_BIN CANDIDATE_BIN TEST_DATA_ROOT OUTPUT_DIR
             [--timing] [--archives-per-kind N] [--members-per-archive N]
             [--mutants-per-seed N] [--timeout SECONDS]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import time
import zipfile
from pathlib import Path
from typing import Any


PROBE_ID = "0694-mce-oracle"
PROFILES = ("baseline", "opaque", "opaque-small", "opaque-large", "opaque-many")
ARCHIVE_SUFFIXES = (".docx", ".xlsx", ".pptx")
XML_SUFFIXES = (".xml", ".rels")
MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"
OPAQUE_NAMESPACE = "urn:litchi:oracle:opaque"
MAX_MEMBER_BYTES = 128 * 1024

PREFERRED_MEMBERS = {
    ".docx": (
        "word/document.xml",
        "word/styles.xml",
        "word/settings.xml",
        "word/numbering.xml",
        "word/comments.xml",
    ),
    ".xlsx": (
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "xl/theme/theme1.xml",
    ),
    ".pptx": (
        "ppt/presentation.xml",
        "ppt/slides/slide1.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/theme/theme1.xml",
    ),
}


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=True,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=True, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
    )


def relpath(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def archive_paths(root: Path) -> dict[str, list[Path]]:
    result = {suffix: [] for suffix in ARCHIVE_SUFFIXES}
    for path in sorted(
        (candidate for candidate in root.rglob("*") if candidate.is_file()),
        key=lambda candidate: candidate.relative_to(root).as_posix(),
    ):
        suffix = path.suffix.lower()
        if suffix in result:
            result[suffix].append(path)
    return result


def member_candidates(archive: zipfile.ZipFile, suffix: str) -> list[str]:
    names = sorted(
        info.filename
        for info in archive.infolist()
        if info.filename.lower().endswith(XML_SUFFIXES)
        and not info.is_dir()
        and info.file_size <= MAX_MEMBER_BYTES
    )
    preferred = [name for name in PREFERRED_MEMBERS[suffix] if name in names]
    preferred_set = set(preferred)

    def score(name: str) -> tuple[int, int, str]:
        lower = name.lower()
        # Main document parts first, then other XML members.  MCE-bearing
        # members are promoted after reading below, while this ordering keeps
        # selection stable when all members are ordinary XML.
        keyword = int(any(token in lower for token in ("document", "workbook", "slide", "sheet", "styles")))
        rel = int(lower.endswith(".rels"))
        return (-keyword, rel, name)

    remainder = sorted((name for name in names if name not in preferred_set), key=score)
    return preferred + remainder


def extract_members(
    root: Path,
    per_kind: int,
    per_archive: int,
) -> list[dict[str, Any]]:
    selected: list[dict[str, Any]] = []
    paths_by_suffix = archive_paths(root)
    for suffix in ARCHIVE_SUFFIXES:
        selected_archives = 0
        for archive_path in paths_by_suffix[suffix]:
            if selected_archives >= per_kind:
                break
            try:
                archive = zipfile.ZipFile(archive_path)
            except (OSError, zipfile.BadZipFile):
                continue
            with archive:
                archive_members: list[tuple[str, bytes]] = []
                for member in member_candidates(archive, suffix):
                    try:
                        data = archive.read(member)
                    except (KeyError, OSError, RuntimeError, zipfile.BadZipFile):
                        continue
                    if len(data) <= MAX_MEMBER_BYTES:
                        archive_members.append((member, data))
                if not archive_members:
                    continue

                # A member carrying MCE markup is more useful for this probe,
                # but ordinary parts remain in the fixed per-archive budget.
                archive_members.sort(
                    key=lambda item: (
                        0 if MC_NAMESPACE.encode() in item[1] else 1,
                        item[0] not in PREFERRED_MEMBERS[suffix],
                        item[0],
                    )
                )
                archive_hash = sha256_file(archive_path)
                for member, data in archive_members[:per_archive]:
                    selected.append(
                        {
                            "kind": suffix[1:],
                            "archive": relpath(archive_path, root),
                            "archive_sha256": archive_hash,
                            "member": member,
                            "data": data,
                        }
                    )
                selected_archives += 1
    selected.sort(key=lambda item: (item["kind"], item["archive"], item["member"]))
    return selected


def first_start_tag(data: bytes) -> tuple[int, int] | None:
    """Return the byte range of the first ordinary XML start tag."""

    for match in re.finditer(rb"<([A-Za-z_][A-Za-z0-9_.:-]*)\b", data):
        start = match.start()
        end = data.find(b">", match.end())
        if end >= 0:
            return start, end
    return None


def inject_in_start_tag(data: bytes, fragment: bytes) -> bytes:
    span = first_start_tag(data)
    if span is None:
        return data + fragment
    start, end = span
    if data[end - 1 : end] == b"/":
        return data[: end - 1] + fragment + data[end - 1 :]
    return data[:end] + fragment + data[end:]


def mutation_qname(data: bytes) -> bytes:
    return inject_in_start_tag(
        data,
        b' xmlns:oracle="urn:litchi:oracle:qname" oracle:marker="qname"',
    )


def mutation_duplicate(data: bytes) -> bytes:
    return inject_in_start_tag(data, b' duplicate="first" duplicate="second"')


def mutation_unbound(data: bytes) -> bytes:
    return inject_in_start_tag(data, b' missing:marker="unbound"')


def mutation_directives(data: bytes) -> bytes:
    return inject_in_start_tag(
        data,
        (
            f' xmlns:mc="{MC_NAMESPACE}" xmlns:oracle="urn:litchi:oracle:directive"'
            ' mc:Ignorable="oracle"'
            ' mc:PreserveAttributes="oracle:a oracle:b oracle:c oracle:d"'
            ' oracle:a="kept"'
        ).encode(),
    )


def mutation_limits(data: bytes) -> bytes:
    marker = sha256_bytes(data)[:16].encode()
    nested = b"<n>" * 40 + marker + b"</n>" * 40
    return (
        f'<r xmlns:mc="{MC_NAMESPACE}">'.encode()
        + nested
        + b"</r>"
    )


def synthetic_cases() -> list[tuple[str, bytes]]:
    mc = MC_NAMESPACE
    return [
        ("plain", b"<r><a/></r>"),
        (
            "alternate",
            (
                f'<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent>'
                '<mc:Choice Requires="x"><x:choice/></mc:Choice>'
                "<mc:Fallback><fallback/></mc:Fallback>"
                "</mc:AlternateContent></r>"
            ).encode(),
        ),
        (
            "qname",
            (
                f'<r xmlns:mc="{mc}" xmlns:x="urn:x" mc:Ignorable="x" '
                'mc:PreserveElements="x:keep"><x:keep/><x:drop/></r>'
            ).encode(),
        ),
        (
            "duplicate",
            f'<r xmlns:mc="{mc}" a="one" a="two"/>'.encode(),
        ),
        (
            "unbound",
            f'<r xmlns:mc="{mc}" missing:value="x"/>'.encode(),
        ),
        (
            "directives",
            (
                f'<r xmlns:mc="{mc}" xmlns:x="urn:x" mc:Ignorable="x" '
                'mc:PreserveAttributes="x:a x:b x:c x:d"><x:item x:a="1"/></r>'
            ).encode(),
        ),
        (
            "opaque-inactive",
            (
                f'<r xmlns:mc="{mc}" xmlns:ext="{OPAQUE_NAMESPACE}" '
                'mc:Ignorable="ext"><ext:payload><mc:AlternateContent>'
                '<mc:Choice Requires="missing"><bad/></mc:Choice>'
                "<mc:Fallback><kept/></mc:Fallback>"
                "</mc:AlternateContent></ext:payload></r>"
            ).encode(),
        ),
        (
            "opaque-valid",
            (
                f'<r xmlns:mc="{mc}" xmlns:ext="{OPAQUE_NAMESPACE}" '
                'mc:Ignorable="ext"><ext:payload><leaf>opaque</leaf></ext:payload></r>'
            ).encode(),
        ),
        (
            "opaque-invalid-qname",
            (
                f'<r xmlns:mc="{mc}" xmlns:ext="{OPAQUE_NAMESPACE}" '
                'mc:Ignorable="ext"><ext:payload><bad:too:many/></ext:payload></r>'
            ).encode(),
        ),
        (
            "limits-depth",
            mutation_limits(b"limits-depth"),
        ),
        (
            "limits-directives",
            (
                f'<r xmlns:mc="{mc}" xmlns:x="urn:x" mc:Ignorable="x" '
                + 'mc:PreserveAttributes="'
                + " ".join(f"x:n{index}" for index in range(24))
                + '">'
                + '<x:item x:n0="0"/></r>'
            ).encode(),
        ),
        (
            "limits-choices",
            (
                f'<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent>'
                + "".join(
                    f'<mc:Choice Requires="x"><x:c{index}/></mc:Choice>'
                    for index in range(10)
                )
                + "<mc:Fallback><fallback/></mc:Fallback>"
                + "</mc:AlternateContent></r>"
            ).encode(),
        ),
    ]


def add_case(
    cases: list[dict[str, Any]],
    output: Path,
    case_id: str,
    data: bytes,
    origin: dict[str, Any],
    mutation: dict[str, Any],
) -> None:
    relative = Path("cases") / f"{case_id}.xml"
    path = output / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    cases.append(
        {
            "id": case_id,
            "path": relative.as_posix(),
            "input_len": len(data),
            "input_sha256": sha256_bytes(data),
            "origin": origin,
            "mutation": mutation,
        }
    )


def build_corpus(
    test_root: Path,
    output: Path,
    archives_per_kind: int,
    members_per_archive: int,
    mutants_per_seed: int,
) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []
    real_members = extract_members(test_root, archives_per_kind, members_per_archive)
    seed_cases: list[dict[str, Any]] = []
    for index, item in enumerate(real_members):
        case_id = f"real-{index:04d}"
        origin = {
            "kind": item["kind"],
            "archive": item["archive"],
            "archive_sha256": item["archive_sha256"],
            "member": item["member"],
            "source_len": len(item["data"]),
            "source_sha256": sha256_bytes(item["data"]),
        }
        add_case(cases, output, case_id, item["data"], origin, {"kind": "identity"})
        seed_cases.append(cases[-1])

    mutation_functions = (
        ("qname", mutation_qname),
        ("duplicate", mutation_duplicate),
        ("unbound", mutation_unbound),
        ("directives", mutation_directives),
        ("limits", mutation_limits),
    )[:mutants_per_seed]
    for seed in seed_cases:
        seed_data = (output / seed["path"]).read_bytes()
        for kind, mutate in mutation_functions:
            case_id = f"{seed['id']}-{kind}"
            add_case(
                cases,
                output,
                case_id,
                mutate(seed_data),
                dict(seed["origin"]),
                {
                    "kind": kind,
                    "parent": seed["id"],
                    "parent_sha256": seed["input_sha256"],
                },
            )

    for index, (name, data) in enumerate(synthetic_cases()):
        case_id = f"synthetic-{index:02d}-{name}"
        add_case(
            cases,
            output,
            case_id,
            data,
            {"kind": "synthetic", "name": name},
            {"kind": "synthetic", "name": name},
        )

    cases.sort(key=lambda case: case["id"])
    for index, case in enumerate(cases):
        case["ordinal"] = index
    return cases


def invoke(
    binary: Path,
    side: str,
    profile: str,
    case: dict[str, Any],
    output: Path,
    timeout: float,
) -> tuple[dict[str, Any], bytes, bytes]:
    case_path = output / case["path"]
    command = [str(binary), "run", profile, str(case_path)]
    started = time.perf_counter_ns()
    try:
        completed = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=timeout,
        )
        stdout = completed.stdout
        stderr = completed.stderr
        return_code = completed.returncode
        status = "completed"
    except subprocess.TimeoutExpired as error:
        stdout = error.stdout or b""
        stderr = error.stderr or b""
        return_code = None
        status = "timeout"
    elapsed = time.perf_counter_ns() - started
    receipt = {
        "probe": PROBE_ID,
        "side": side,
        "profile": profile,
        "case": case["id"],
        "case_sha256": case["input_sha256"],
        "command": ["run", profile, case["path"]],
        "binary_sha256": sha256_file(binary),
        "status": status,
        "returncode": return_code,
        "elapsed_ns": elapsed,
        "stdout_sha256": sha256_bytes(stdout),
        "stderr_sha256": sha256_bytes(stderr),
        "stdout": stdout.decode("utf-8", "replace"),
        "stderr": stderr.decode("utf-8", "replace"),
    }
    return receipt, stdout, stderr


def valid_probe_record(stdout: bytes) -> bool:
    lines = stdout.decode("utf-8", "replace").splitlines()
    return len(lines) == 1 and lines[0].startswith(("OK\t", "ERR\t"))


def run_differential(
    baseline: Path,
    candidate: Path,
    output: Path,
    cases: list[dict[str, Any]],
    timeout: float,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    receipts: list[dict[str, Any]] = []
    mismatches: list[dict[str, Any]] = []
    binaries = (("baseline", baseline), ("candidate", candidate))
    receipt_path = output / "invocations.jsonl"
    with receipt_path.open("w", encoding="utf-8") as receipt_stream:
        for case in cases:
            for profile in PROFILES:
                results: dict[str, tuple[dict[str, Any], bytes, bytes]] = {}
                for side, binary in binaries:
                    receipt, stdout, stderr = invoke(
                        binary, side, profile, case, output, timeout
                    )
                    results[side] = (receipt, stdout, stderr)
                    receipts.append(receipt)
                    receipt_stream.write(json.dumps(receipt, sort_keys=True) + "\n")
                baseline_receipt, baseline_stdout, _ = results["baseline"]
                candidate_receipt, candidate_stdout, _ = results["candidate"]
                baseline_valid = (
                    baseline_receipt["status"] == "completed"
                    and baseline_receipt["returncode"] == 0
                    and valid_probe_record(baseline_stdout)
                )
                candidate_valid = (
                    candidate_receipt["status"] == "completed"
                    and candidate_receipt["returncode"] == 0
                    and valid_probe_record(candidate_stdout)
                )
                reason = None
                if not baseline_valid or not candidate_valid:
                    reason = "probe invocation failed or emitted no typed record"
                elif baseline_stdout != candidate_stdout:
                    reason = "exact probe records differ"
                if (
                    reason is not None
                ):
                    mismatches.append(
                        {
                            "case": case["id"],
                            "profile": profile,
                            "reason": reason,
                            "baseline_valid": baseline_valid,
                            "candidate_valid": candidate_valid,
                            "baseline_stdout": baseline_stdout.decode("utf-8", "replace"),
                            "candidate_stdout": candidate_stdout.decode("utf-8", "replace"),
                            "baseline_status": baseline_receipt["status"],
                            "candidate_status": candidate_receipt["status"],
                            "baseline_returncode": baseline_receipt["returncode"],
                            "candidate_returncode": candidate_receipt["returncode"],
                        }
                    )
    return receipts, mismatches


def run_timing(
    baseline: Path,
    candidate: Path,
    output: Path,
    cases: list[dict[str, Any]],
    warmups: int,
    samples: int,
    timeout: float,
) -> list[dict[str, Any]]:
    preferred = next(
        (case for case in cases if case["id"].endswith("-opaque-valid")),
        cases[0],
    )
    records: list[dict[str, Any]] = []
    profiles = ("baseline", "opaque-small", "opaque-large")
    for case in (preferred,):
        for profile in profiles:
            for side, binary in (("baseline", baseline), ("candidate", candidate)):
                command = [
                    str(binary),
                    "time",
                    profile,
                    str(output / case["path"]),
                    str(warmups),
                    str(samples),
                ]
                started = time.perf_counter_ns()
                try:
                    completed = subprocess.run(
                        command,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        check=False,
                        timeout=timeout,
                    )
                    stdout = completed.stdout
                    stderr = completed.stderr
                    returncode = completed.returncode
                    status = "completed"
                except subprocess.TimeoutExpired as error:
                    stdout = error.stdout or b""
                    stderr = error.stderr or b""
                    returncode = None
                    status = "timeout"
                records.append(
                    {
                        "probe": PROBE_ID,
                        "side": side,
                        "profile": profile,
                        "case": case["id"],
                        "case_sha256": case["input_sha256"],
                        "command": ["time", profile, case["path"], str(warmups), str(samples)],
                        "binary_sha256": sha256_file(binary),
                        "status": status,
                        "returncode": returncode,
                        "elapsed_ns": time.perf_counter_ns() - started,
                        "stdout_sha256": sha256_bytes(stdout),
                        "stderr_sha256": sha256_bytes(stderr),
                        "stdout": stdout.decode("utf-8", "replace"),
                        "stderr": stderr.decode("utf-8", "replace"),
                    }
                )
    return records


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline_bin", type=Path)
    parser.add_argument("candidate_bin", type=Path)
    parser.add_argument("test_data_root", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--timing", action="store_true")
    parser.add_argument("--timing-warmups", type=int, default=5)
    parser.add_argument("--timing-samples", type=int, default=20)
    parser.add_argument("--archives-per-kind", type=int, default=4)
    parser.add_argument("--members-per-archive", type=int, default=3)
    parser.add_argument("--mutants-per-seed", type=int, default=4)
    parser.add_argument("--timeout", type=float, default=60.0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.test_data_root.is_dir():
        raise SystemExit(f"test-data root is not a directory: {args.test_data_root}")
    if not args.baseline_bin.is_file() or not args.candidate_bin.is_file():
        raise SystemExit("both probe binaries must exist before running the driver")
    if args.output_dir.exists() and any(args.output_dir.iterdir()):
        raise SystemExit(f"output directory must be new or empty: {args.output_dir}")
    if args.archives_per_kind < 1 or args.members_per_archive < 1:
        raise SystemExit("archive and member budgets must be positive")
    if not 0 <= args.mutants_per_seed <= 5:
        raise SystemExit("mutants-per-seed must be between 0 and 5")
    if args.timing_warmups < 0 or args.timing_samples < 1:
        raise SystemExit("timing warmups must be nonnegative and samples positive")
    args.output_dir.mkdir(parents=True, exist_ok=True)

    cases = build_corpus(
        args.test_data_root,
        args.output_dir,
        args.archives_per_kind,
        args.members_per_archive,
        args.mutants_per_seed,
    )
    if len(cases) < 10:
        raise SystemExit(f"corpus unexpectedly small: {len(cases)} cases")
    corpus_rows = [
        {
            "id": case["id"],
            "path": case["path"],
            "input_len": case["input_len"],
            "input_sha256": case["input_sha256"],
            "origin": case["origin"],
            "mutation": case["mutation"],
        }
        for case in cases
    ]
    corpus_sha256 = sha256_bytes(canonical_json(corpus_rows))
    write_json(
        args.output_dir / "corpus.json",
        {
            "probe": PROBE_ID,
            "schema": 1,
            "test_data_root_label": args.test_data_root.name,
            "case_count": len(cases),
            "profiles": list(PROFILES),
            "corpus_sha256": corpus_sha256,
            "cases": cases,
        },
    )

    receipts, mismatches = run_differential(
        args.baseline_bin,
        args.candidate_bin,
        args.output_dir,
        cases,
        args.timeout,
    )
    result: dict[str, Any] = {
        "probe": PROBE_ID,
        "schema": 1,
        "baseline_binary_sha256": sha256_file(args.baseline_bin),
        "candidate_binary_sha256": sha256_file(args.candidate_bin),
        "case_count": len(cases),
        "profile_count": len(PROFILES),
        "invocation_count": len(receipts),
        "mismatch_count": len(mismatches),
        "corpus_sha256": corpus_sha256,
        "mismatches": mismatches,
    }
    write_json(args.output_dir / "results.json", result)
    if args.timing:
        timing = run_timing(
            args.baseline_bin,
            args.candidate_bin,
            args.output_dir,
            cases,
            args.timing_warmups,
            args.timing_samples,
            args.timeout,
        )
        write_json(args.output_dir / "timings.json", {"probe": PROBE_ID, "records": timing})

    print(
        f"probe={PROBE_ID} cases={len(cases)} profiles={len(PROFILES)} "
        f"invocations={len(receipts)} mismatches={len(mismatches)} "
        f"corpus_sha256={corpus_sha256}"
    )
    return 0 if not mismatches else 1


if __name__ == "__main__":
    raise SystemExit(main())
