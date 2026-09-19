#!/usr/bin/env python3
"""Prepare the 0695 per-member MCE attribution corpus.

The member order is taken from the first real, count-one capture in the
0693 diagnostic trace.  This script does not build or run a probe.  It
extracts the fourteen source XML members, writes the ordered sequence files
used by the standalone probe, and binds the trace, archive, member, and
current-0694 source identities in ``corpus/manifest.json``.
"""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import zipfile
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]

SOURCE_REL = Path("test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx")
TRACE_REL = Path("docs/performance/results/change-0693/trace-runs/20260919T195826994463Z-3152718")
TRACE_STDERR_REL = TRACE_REL / "real-capture-1.stderr"
TRACE_STDOUT_REL = TRACE_REL / "real-capture-1.stdout"
TRACE_SUMMARY_REL = Path("docs/performance/results/change-0693/trace-summary.json")
TRACE_MANIFEST_REL = TRACE_REL / "manifest.json"
TRACE_PROBES_REL = TRACE_REL / "probe-runs.json"
TRACE_SOURCE_HASHES_REL = TRACE_REL / "source-hashes.json"
TRACE_PATCH_REL = TRACE_REL / "trace.patch"

# The source files below are the production bytes at the 0694 commit.  The
# trace itself was generated with temporary 0693 instrumentation, so these
# hashes intentionally differ from the trace's ``after_sha256`` values.
CURRENT_SOURCE = {
    "crates/litchi-ooxml-common/src/mce/codec.rs": {
        "sha256": "266eee56f526622f96d46adbb08026dbcbc434e260a4f59efcc9253d982c4674",
        "markers": [
            "pub fn process_markup_compatibility<'a>(",
            "let (namespace, local) = expand_parts(q, &c.ns, true)?;",
            "pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {",
            "process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml)",
            "pub fn process_part(part: &dyn litchi_opc::Part) -> R<Cow<'_, [u8]>>",
            "process_ooxml(part.blob())",
        ],
    },
    "crates/litchi-pptx/src/parts/mod.rs": {
        "sha256": "2a3eaa07274e992a13b56fe2aa9dbc4ad24d679f6adc893029d59506f562d94c",
        "markers": [
            "pub(crate) fn processed_xml_with_source(part: &dyn Part) -> Result<ProcessedXml<'_>>",
            "let processed = litchi_ooxml_common::mce::process_ooxml(source)?;",
            "pub(crate) fn processed_xml(part: &dyn Part) -> Result<Cow<'_, [u8]>>",
            "Ok(process_part(part)?)",
        ],
    },
    "crates/litchi-pptx/src/opened/model.rs": {
        "sha256": "5cb5a59ba6bed9e9ca626e7e9fb60f2aad182db2f44523f642d50064cf5e971d",
        "markers": [
            "fn capture_internal(",
            "let captured = view.capture_slides()?;",
            "let _notes = match slide_root_proofs.as_deref()",
        ],
    },
}

PART_RE = re.compile(
    r"LITCHI0693_PART profile=(?P<profile>\S+) stage=(?P<stage>\S+) "
    r"uri=(?P<uri>\S+) raw_ptr=(?P<ptr>0x[0-9a-f]+) raw_len=(?P<raw_len>\d+)"
)
CALL_RE = re.compile(
    r"LITCHI0693_CODEC call=(?P<call>\d+) profile=(?P<profile>\S+) "
    r"stage=(?P<stage>\S+) input_ptr=(?P<input_ptr>0x[0-9a-f]+) "
    r"input_len=(?P<input_len>\d+) output_ptr=(?P<output_ptr>0x[0-9a-f]+) "
    r"output_len=(?P<output_len>\d+) mode=(?P<mode>\S+) status=(?P<status>\S+)"
)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha_file(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def read_json(path: Path):
    return json.loads(path.read_text())


def line_number(text: str, marker: str) -> int:
    for number, line in enumerate(text.splitlines(), 1):
        if marker in line:
            return number
    raise AssertionError(f"missing marker {marker!r}")


def parse_stdout(path: Path) -> dict[str, str]:
    values = {}
    for line in path.read_text().splitlines():
        if "\t" in line:
            key, value = line.split("\t", 1)
            values[key] = value
    return values


def write_if_same_or_absent(path: Path, data: bytes) -> None:
    if path.exists():
        assert path.read_bytes() == data, f"refusing to overwrite changed {path}"
    else:
        path.write_bytes(data)


def main() -> None:
    source = ROOT / SOURCE_REL
    trace_stderr = ROOT / TRACE_STDERR_REL
    trace_stdout = ROOT / TRACE_STDOUT_REL
    trace_summary_path = ROOT / TRACE_SUMMARY_REL
    trace_manifest_path = ROOT / TRACE_MANIFEST_REL
    trace_probes_path = ROOT / TRACE_PROBES_REL
    trace_source_hashes_path = ROOT / TRACE_SOURCE_HASHES_REL
    trace_patch_path = ROOT / TRACE_PATCH_REL

    for path in (
        source,
        trace_stderr,
        trace_stdout,
        trace_summary_path,
        trace_manifest_path,
        trace_probes_path,
        trace_source_hashes_path,
        trace_patch_path,
    ):
        assert path.is_file(), path

    head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    # Keep this assertion explicit: this packet is bound to the already
    # committed 0694 production bytes.  The packet's own untracked files do
    # not alter HEAD, so rerunning preparation remains idempotent.
    assert head == "df44030d613d7d7e3856cc0494834e414cdb788a", head

    source_bytes = source.read_bytes()
    source_sha = sha_bytes(source_bytes)
    assert source_sha == "aebee5a724a9b5b8020c41fe789df9b3805265f20bf9a87dd28d8b0ed556deb0"

    # Verify that this is precisely the first real count-one capture, rather
    # than the later apply capture or one of the repeated intervals.
    trace_manifest = read_json(trace_manifest_path)
    assert trace_manifest["status"] == "completed"
    assert trace_manifest["profile"] == "candidate"
    assert trace_manifest["diagnostic_only"] is True
    matching_probe_runs = [
        row
        for row in trace_manifest["probe_runs"]
        if row["case"] == "real" and row["operation"] == "capture" and row["count"] == 1
    ]
    assert len(matching_probe_runs) == 1
    probe_run = matching_probe_runs[0]
    historical_source = probe_run["source"]
    assert historical_source.endswith("/" + SOURCE_REL.as_posix())
    assert probe_run["stderr_sha256"] == sha_file(trace_stderr)
    assert probe_run["stdout_sha256"] == sha_file(trace_stdout)

    stdout = parse_stdout(trace_stdout)
    assert stdout == {
        "probe": "0693",
        "mode": "prefix",
        "source": historical_source,
        "source_archive_bytes": str(len(source_bytes)),
        "source_archive_sha256": source_sha,
        "stage": "capture",
        "iterations": "1",
        "completed": "1",
    }

    trace_summary = read_json(trace_summary_path)
    assert trace_summary["status"] == "pass"
    assert trace_summary["schema"] == "litchi-0693-trace-summary-v1"
    summary_run = trace_summary["runs"][0]
    assert summary_run["profile"] == "candidate"
    assert summary_run["verification"] == "pass"
    summary_legs = [
        leg
        for leg in summary_run["legs"]
        if leg["case"] == "real" and leg["operation"] == "capture" and leg["count"] == 1
    ]
    assert len(summary_legs) == 1
    summary_leg = summary_legs[0]
    assert summary_leg["source"] == historical_source
    assert summary_leg["intervals"] == [
        {
            **summary_leg["intervals"][0],
            "call_first": 1,
            "call_last": 18,
        }
    ]
    interval = summary_leg["intervals"][0]
    assert interval["call_first"] == 1 and interval["call_last"] == 18
    assert interval["part_records"] == 16
    assert interval["mapped_part_records"] == 16
    assert interval["processed_xml_part_records"] == 3
    assert interval["owned_calls"] == 18
    assert interval["borrowed_calls"] == 0

    # The summary and the raw run receipt must agree on the selected source
    # and on the raw files.  The trace process has one extra call (call 0)
    # before capture; it is intentionally outside the 18-call attribution
    # sequence.
    assert summary_leg["stderr"] == trace_stderr.name
    assert len(trace_manifest["probe_runs"]) == 6
    probe_receipt = read_json(trace_probes_path)
    assert len(probe_receipt) == 6
    assert next(row for row in probe_receipt if row["case"] == "real" and row["operation"] == "capture")["count"] == 1

    part_records: list[dict[str, str]] = []
    pointer_to_part: dict[str, dict[str, str]] = {}
    calls: list[dict[str, str]] = []
    for line_number_, line in enumerate(trace_stderr.read_text().splitlines(), 1):
        part_match = PART_RE.fullmatch(line)
        if part_match:
            part = part_match.groupdict()
            assert part["profile"] == "candidate"
            part["line"] = line_number_
            part_records.append(part)
            pointer_to_part[part["ptr"]] = part
            continue
        call_match = CALL_RE.fullmatch(line)
        if call_match:
            call = call_match.groupdict()
            call["line"] = line_number_
            calls.append(call)

    assert [int(row["call"]) for row in calls] == list(range(19))
    assert len(part_records) == 17
    assert len(pointer_to_part) == 14
    assert calls[0]["call"] == "0"
    assert pointer_to_part[calls[0]["input_ptr"]]["uri"] == "/ppt/presentation.xml"
    selected = calls[1:]
    assert len(selected) == 18
    assert all(row["profile"] == "candidate" for row in selected)
    assert all(row["stage"] == "prefix:real:capture" for row in selected)
    assert all(row["status"] == "ok" for row in selected)
    assert all(row["mode"] == "owned" for row in selected)

    expected_uris = ["/ppt/presentation.xml"] * 3 + [
        f"/ppt/slides/slide{i}.xml" for i in range(1, 14)
    ] + ["/ppt/presentation.xml"] * 2
    actual_uris = [pointer_to_part[row["input_ptr"]]["uri"] for row in selected]
    assert actual_uris == expected_uris, actual_uris

    # The 0693 trace source map and patch are retained as provenance.  Their
    # bytes are not silently treated as current 0694 production sources.
    trace_source_hashes = read_json(trace_source_hashes_path)
    trace_after = {row["path"]: row["after_sha256"] for row in trace_source_hashes["files"]}
    trace_before = {row["path"]: row["before_sha256"] for row in trace_source_hashes["files"]}
    assert set(trace_after) == set(CURRENT_SOURCE)

    current_source_identity = {}
    for name, expected in CURRENT_SOURCE.items():
        path = ROOT / name
        data = path.read_bytes()
        text = data.decode()
        actual_sha = sha_bytes(data)
        assert actual_sha == expected["sha256"], (name, actual_sha)
        marker_lines = {marker: line_number(text, marker) for marker in expected["markers"]}
        trace_digest = trace_after[name]
        assert actual_sha != trace_digest, (name, "trace instrumentation unexpectedly current")
        current_source_identity[name] = {
            "sha256": actual_sha,
            "trace_after_sha256": trace_digest,
            "trace_before_sha256": trace_before[name],
            "same_bytes_as_trace_after": False,
            "required_marker_lines": marker_lines,
        }

    # Extract the exact fourteen members named by the trace.  Keep the source
    # member names stable and simple so the sequence runner can resolve them
    # from its own packet-relative directory.
    member_uris = ["/ppt/presentation.xml"] + [
        f"/ppt/slides/slide{i}.xml" for i in range(1, 14)
    ]
    member_paths = [uri.lstrip("/") for uri in member_uris]
    corpus = P / "corpus"
    sequences = P / "sequences"
    corpus.mkdir(exist_ok=True)
    sequences.mkdir(exist_ok=True)
    member_rows = []
    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        assert len(names) == 103
        ppt_xml_names = [name for name in names if name.startswith("ppt/") and name.endswith(".xml")]
        assert len(ppt_xml_names) == 55
        archive_name_bytes = "\n".join(names).encode()
        archive_name_sha = sha_bytes(archive_name_bytes)
        for order, (uri, member_path) in enumerate(zip(member_uris, member_paths)):
            info = archive.getinfo(member_path)
            data = archive.read(member_path)
            target = corpus / Path(member_path).name
            write_if_same_or_absent(target, data)
            assert info.file_size == len(data)
            member_rows.append(
                {
                    "order": order,
                    "uri": uri,
                    "zip_member": member_path,
                    "corpus_path": relative(target),
                    "bytes": len(data),
                    "sha256": sha_bytes(data),
                    "crc32": info.CRC,
                    "compress_size": info.compress_size,
                    "compress_type": info.compress_type,
                    "header_offset": info.header_offset,
                }
            )

    member_by_uri = {row["uri"]: row for row in member_rows}
    calls_rows = []
    for order, call in enumerate(selected):
        part = pointer_to_part[call["input_ptr"]]
        member = member_by_uri[part["uri"]]
        assert int(call["input_len"]) == member["bytes"]
        calls_rows.append(
            {
                "order": order,
                "call": int(call["call"]),
                "trace_line": int(call["line"]),
                "uri": part["uri"],
                "relative_path": member["corpus_path"],
                "trace_part_stage": part["stage"],
                "trace_part_line": int(part["line"]),
                "trace_input_ptr": call["input_ptr"],
                "trace_input_len": int(call["input_len"]),
                "trace_output_ptr": call["output_ptr"],
                "trace_output_len": int(call["output_len"]),
                "mode": call["mode"],
                "status": call["status"],
                "input_sha256": member["sha256"],
            }
        )

    assert [row["uri"] for row in calls_rows] == expected_uris
    assert [row["trace_part_stage"] for row in calls_rows[:3]] == ["processed_xml"] * 3
    assert [row["trace_part_stage"] for row in calls_rows[3:16]] == ["source_bound"] * 13
    assert [row["trace_part_stage"] for row in calls_rows[16:]] == ["processed_xml"] * 2

    # The standalone probe reads paths relative to each sequence file.  The
    # all sequence preserves the exact 18-call order; presentation/slides and
    # single-member sequences support attribution and profile controls.
    presentation_path = "../corpus/presentation.xml"
    slide_paths = [f"../corpus/slide{i}.xml" for i in range(1, 14)]
    all_paths = [presentation_path] * 3 + slide_paths + [presentation_path] * 2
    assert len(all_paths) == 18
    sequence_rows = {
        "all": all_paths,
        "presentation": [presentation_path] * 5,
        "slides": slide_paths,
    }
    for index, path in enumerate(slide_paths, 1):
        sequence_rows[f"slide{index}"] = [path]
    for name, rows in sequence_rows.items():
        payload = ("\n".join(rows) + "\n").encode()
        write_if_same_or_absent(sequences / f"{name}.txt", payload)

    manifest = {
        "schema": "litchi-0695-mce-attribution-v1",
        "status": "prepared",
        "purpose": "price the five presentation passes and thirteen slide passes using current default MCE entry code",
        "production_revision": {
            "change": "0694",
            "commit": head,
            "commit_subject": "perf(ooxml): defer expanded MCE element name ownership",
            "source_identity": current_source_identity,
        },
        "source_archive": {
            "path": relative(source),
            "bytes": len(source_bytes),
            "sha256": source_sha,
        },
        "trace_provenance": {
            "change": "0693",
            "trace_dir": relative(ROOT / TRACE_REL),
            "stderr": relative(trace_stderr),
            "stderr_sha256": sha_file(trace_stderr),
            "stdout": relative(trace_stdout),
            "stdout_sha256": sha_file(trace_stdout),
            "summary": relative(trace_summary_path),
            "summary_sha256": sha_file(trace_summary_path),
            "manifest": relative(trace_manifest_path),
            "manifest_sha256": sha_file(trace_manifest_path),
            "probe_runs": relative(trace_probes_path),
            "probe_runs_sha256": sha_file(trace_probes_path),
            "source_hashes": relative(trace_source_hashes_path),
            "source_hashes_sha256": sha_file(trace_source_hashes_path),
            "trace_patch": relative(trace_patch_path),
            "trace_patch_sha256": sha_file(trace_patch_path),
            "selection": {
                "case": "real",
                "operation": "capture",
                "count": 1,
                "profile": "candidate",
                "selected_calls": "1..18",
                "excluded_call": 0,
                "pointer_scope": "one probe process; do not compare addresses across stderr files",
            },
            "source_identity_note": "0693 raw trace uses temporary instrumentation; current 0694 source hashes and callsites are separately verified above",
        },
        "archive_inventory": {
            "zip_member_count": 103,
            "ppt_xml_member_count": 55,
            "selected_member_count": 14,
            "zip_namelist_sha256": archive_name_sha,
        },
        "members": member_rows,
        "calls": calls_rows,
        "groups": {
            "presentation": {"calls": [1, 2, 3, 17, 18], "count": 5},
            "slides": {"calls": list(range(4, 17)), "count": 13},
            "all": {"calls": list(range(1, 19)), "count": 18},
        },
        "sequences": {
            name: {
                "path": relative(sequences / f"{name}.txt"),
                "lines": len(rows),
                "sha256": sha_file(sequences / f"{name}.txt"),
                "relative_paths": rows,
            }
            for name, rows in sequence_rows.items()
        },
    }
    manifest_path = corpus / "manifest.json"
    write_if_same_or_absent(
        manifest_path,
        (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode(),
    )

    # Keep a plainly consumable exact-order file next to the manifest for
    # reviewers and simple shell runners.  The probe's authoritative files
    # are under sequences/; this copy is still bound and checked below.
    sequence_path = corpus / "sequence.txt"
    write_if_same_or_absent(
        sequence_path,
        ("\n".join(member["uri"].lstrip("/") for member in calls_rows) + "\n").encode(),
    )
    assert sha_file(sequence_path) == sha_bytes(
        ("\n".join(member["uri"].lstrip("/") for member in calls_rows) + "\n").encode()
    )

    print(
        f"Prepared 0695 corpus: {len(member_rows)} members, "
        f"{len(calls_rows)} ordered calls, {len(sequence_rows)} sequences.",
        flush=True,
    )


if __name__ == "__main__":
    main()
