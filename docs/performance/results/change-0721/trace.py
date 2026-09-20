#!/usr/bin/env python3
"""Apply and restore the source-bound 0721 DOCX differential tracer.

The recipe is intentionally a patcher, rather than a runner.  The coordinator
builds and runs the public oracle while this exact-source patch is installed,
then invokes restore before changing lanes.  restore refuses to write over a
production file whose bytes are neither the recorded transformed bytes nor the
original bytes.

The baseline lane instruments the independent alt and range readers.  The
candidate lane instruments the fused reader and its range observer.  Observer
callbacks have their own counter and never increment reader_reads.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
FRAGMENT = PACKET / "trace.fragment"
SCRATCH = Path("/home/zhuhe/code/litchi-scratch-0721-trace")
MODULE = SCRATCH / "scan_trace_0721.rs"

LIB_RELATIVE = "crates/litchi-docx/src/lib.rs"
BASELINE_LIB_SHA256 = "ba9965127f1b6ddfdc2cd5fe5e4684479e2c7b6848d48b945489e30296d87225"

COMMON_PATCHED = (
    LIB_RELATIVE,
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/writer/doc/package.rs",
)
BASELINE_PATCHED = COMMON_PATCHED + ("crates/litchi-docx/src/parts/document_part.rs",)


class RecipeError(RuntimeError):
    """A refusal caused by stale or unexpected source state."""


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def root_path(relative: str) -> Path:
    return ROOT / relative


def snapshot_path(lane: str, relative: str) -> Path:
    return PACKET / lane / relative


def patched_relatives(lane: str) -> tuple[str, ...]:
    if lane not in {"baseline", "candidate"}:
        raise RecipeError(f"unknown lane {lane!r}")
    return BASELINE_PATCHED if lane == "baseline" else COMMON_PATCHED


def expected_source(lane: str, relative: str) -> bytes:
    if relative == LIB_RELATIVE:
        path = root_path(relative)
        data = path.read_bytes()
        if digest(data) != BASELINE_LIB_SHA256:
            raise RecipeError(
                f"{relative}: expected the unchanged library source before tracing; "
                f"found {digest(data)}"
            )
        return data
    path = snapshot_path(lane, relative)
    if not path.is_file():
        raise RecipeError(f"missing {lane} source snapshot: {path}")
    return path.read_bytes()


def replace_once(source: bytes, needle: bytes, replacement: bytes, label: str) -> bytes:
    count = source.count(needle)
    if count != 1:
        raise RecipeError(f"{label}: expected one source occurrence, found {count}")
    return source.replace(needle, replacement, 1)


def replace_all(source: bytes, needle: bytes, replacement: bytes, label: str) -> bytes:
    count = source.count(needle)
    if count < 1:
        raise RecipeError(f"{label}: expected at least one source occurrence")
    return source.replace(needle, replacement)


def transform_lib(source: bytes) -> bytes:
    marker = b"mod scan_trace_0721;"
    if marker in source:
        raise RecipeError(f"{LIB_RELATIVE}: trace module declaration already present")
    if not source.endswith(b"\n"):
        raise RecipeError(f"{LIB_RELATIVE}: expected a final newline")
    declaration = (
        f'#[path = "{MODULE.as_posix()}"]\n'
        "mod scan_trace_0721;\n"
    ).encode()
    return source + b"\n" + declaration


def transform_active(source: bytes) -> bytes:
    source = replace_once(
        source,
        b"""pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    validate_xml(xml)?;""",
        b"""pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    if let Err(error) = validate_xml(xml) {
        crate::scan_trace_0721::active_error(xml, offsets, &error, "validate");
        return Err(error);
    }""",
        "alt active validation boundary",
    )
    return replace_once(
        source,
        b"""    litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map_err(Error::from)
""",
        b"""    let result = litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map_err(Error::from);
    crate::scan_trace_0721::active_result(xml, offsets, &result, "mce");
    result
""",
        "alt active MCE result",
    )


def transform_alt(source: bytes, lane: str) -> bytes:
    source = transform_active(source)
    source = replace_all(
        source,
        b"""        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
""",
        b"""        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        crate::scan_trace_0721::reader_read(
            "alt",
            u64::from(event_start),
            reader.buffer_position(),
        );
""",
        "alt XML reader calls",
    )

    if lane == "baseline":
        source = replace_once(
            source,
            b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
""",
            b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
    let _trace = crate::scan_trace_0721::alt_scan_scope(xml);
""",
            "baseline alt scan scope",
        )
        source = replace_once(
            source,
            b"""    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;""",
            b"""    crate::scan_trace_0721::alt_stage("raw_chunks");
    crate::scan_trace_0721::alt_raw_chunks(&chunks);
    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;""",
            "baseline raw chunk boundary",
        )
        source = replace_once(
            source,
            b"""    });
    Ok(chunks)
}""",
            b"""    });
    crate::scan_trace_0721::alt_stage("selected_chunks");
    crate::scan_trace_0721::alt_selected_chunks(&chunks);
    crate::scan_trace_0721::alt_complete();
    Ok(chunks)
}""",
            "baseline selected chunk boundary",
        )
    else:
        source = replace_once(
            source,
            b"""pub(crate) fn scan_with_block_ranges(
    xml: &[u8],
) -> Result<(BTreeMap<u32, Chunk>, Vec<(usize, u32, u32)>)> {
    validate_xml(xml)?;""",
            b"""pub(crate) fn scan_with_block_ranges(
    xml: &[u8],
) -> Result<(BTreeMap<u32, Chunk>, Vec<(usize, u32, u32)>)> {
    validate_xml(xml)?;
    let _trace_alt = crate::scan_trace_0721::alt_scan_scope(xml);
    let _trace_range = crate::scan_trace_0721::range_scan_scope(xml);""",
            "candidate fused scan scopes",
        )
        active_binding = (
            b"    let active_offsets = active(xml, &offsets)?;"
            if b"let active_offsets = active(xml, &offsets)?;" in source
            else b"    let active = active(xml, &offsets)?;"
        )
        source = replace_once(
            source,
            b"    let offsets = alt_scanner.chunks.keys().copied().collect::<Vec<_>>();\n" + active_binding,
            b"""    crate::scan_trace_0721::alt_stage("raw_chunks");
    crate::scan_trace_0721::alt_raw_chunks(&alt_scanner.chunks);
    let offsets = alt_scanner.chunks.keys().copied().collect::<Vec<_>>();
""" + active_binding,
            "candidate raw chunk boundary",
        )
        source = replace_once(
            source,
            b"""    });

    // The old range walk completed before its second MCE call.""",
            b"""    });
    crate::scan_trace_0721::alt_stage("selected_chunks");
    crate::scan_trace_0721::alt_selected_chunks(&alt_scanner.chunks);

    // The old range walk completed before its second MCE call.""",
            "candidate selected chunk boundary",
        )
        source = replace_once(
            source,
            b"""    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok((alt_scanner.chunks, ranges))""",
            b"""    crate::scan_trace_0721::range_stage("active");
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    crate::scan_trace_0721::range_stage("filter");
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    crate::scan_trace_0721::range_selected_ranges(&ranges);
    crate::scan_trace_0721::range_complete();
    crate::scan_trace_0721::alt_complete();
    Ok((alt_scanner.chunks, ranges))""",
            "candidate selected range boundary",
        )
        source = replace_once(
            source,
            b"""    if let Some(error) = range_error {
        return Err(error);
    }
    let mut starts = Vec::new();""",
            b"""    if let Some(error) = range_error {
        return Err(error);
    }
    crate::scan_trace_0721::range_stage("raw_ranges");
    crate::scan_trace_0721::range_raw_ranges(&ranges);
    let mut starts = Vec::new();""",
            "candidate deferred raw range boundary",
        )
        source = replace_once(
            source,
            b"""        let eof = alt_scanner.observe(event_start, decoder, &resolver, &namespace, &event)?;""",
            b"""        let eof = alt_scanner.observe(event_start, decoder, &resolver, &namespace, &event)?;
        crate::scan_trace_0721::range_observer(
            u64::from(event_start),
            reader.buffer_position(),
        );""",
            "candidate range observer boundary",
        )
    return source


def transform_namespace(source: bytes, lane: str) -> bytes:
    if lane == "baseline":
        return replace_once(
            source,
            b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;

        match event {""",
            b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        crate::scan_trace_0721::reader_read(
            "range",
            u64::try_from(event_start).unwrap_or(u64::MAX),
            u64::try_from(event_end).unwrap_or(u64::MAX),
        );
        crate::scan_trace_0721::range_observer(
            u64::try_from(event_start).unwrap_or(u64::MAX),
            u64::try_from(event_end).unwrap_or(u64::MAX),
        );

        match event {""",
            "baseline range XML reader boundary",
        )
    return replace_once(
        source,
        b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let eof = matches!(scan_event, WordElementScanEvent::Eof);""",
        b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        crate::scan_trace_0721::reader_read(
            "range",
            u64::try_from(event_start).unwrap_or(u64::MAX),
            u64::try_from(event_end).unwrap_or(u64::MAX),
        );
        crate::scan_trace_0721::range_observer(
            u64::try_from(event_start).unwrap_or(u64::MAX),
            u64::try_from(event_end).unwrap_or(u64::MAX),
        );
        let eof = matches!(scan_event, WordElementScanEvent::Eof);""",
        "candidate standalone range reader boundary",
    )


def transform_document_part(source: bytes) -> bytes:
    source = replace_once(
        source,
        b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let mut ranges = Vec::new();""",
        b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let _trace = crate::scan_trace_0721::range_scan_scope(xml);
    let mut ranges = Vec::new();""",
        "baseline range pass scope",
    )
    source = replace_once(
        source,
        b"""    )?;
    let mut starts = Vec::new();""",
        b"""    )?;
    crate::scan_trace_0721::range_stage("raw_ranges");
    crate::scan_trace_0721::range_raw_ranges(&ranges);
    let mut starts = Vec::new();""",
        "baseline raw range boundary",
    )
    return replace_once(
        source,
        b"""    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)""",
        b"""    crate::scan_trace_0721::range_stage("active");
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    crate::scan_trace_0721::range_stage("filter");
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    crate::scan_trace_0721::range_selected_ranges(&ranges);
    crate::scan_trace_0721::range_complete();
    Ok(ranges)""",
        "baseline selected range boundary",
    )


def transform_package(source: bytes) -> bytes:
    return replace_once(
        source,
        b"""        let bytes = xml.as_bytes();""",
        b"""        let _trace = crate::scan_trace_0721::document_scope(xml.as_bytes());
        let bytes = xml.as_bytes();""",
        "DocumentBody boundary scope",
    )


def transform(relative: str, source: bytes, lane: str) -> bytes:
    if relative == LIB_RELATIVE:
        return transform_lib(source)
    if relative.endswith("/alt/codec.rs"):
        return transform_alt(source, lane)
    if relative.endswith("/namespace.rs"):
        return transform_namespace(source, lane)
    if relative.endswith("/parts/document_part.rs"):
        return transform_document_part(source)
    if relative.endswith("/writer/doc/package.rs"):
        return transform_package(source)
    raise RecipeError(f"no transform for {relative}")


def manifest_path() -> Path:
    return SCRATCH / "manifest.json"


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def read_manifest() -> dict[str, object]:
    path = manifest_path()
    if not path.is_file():
        raise RecipeError(f"missing trace manifest: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RecipeError(f"cannot read trace manifest: {error}") from error
    if not isinstance(value, dict) or value.get("version") != 1:
        raise RecipeError("trace manifest has an unexpected format")
    return value


def apply(lane: str) -> None:
    if SCRATCH.exists():
        raise RecipeError(
            f"owned scratch path already exists: {SCRATCH}; restore it before reapplying"
        )
    relatives = patched_relatives(lane)
    originals: dict[str, bytes] = {}
    for relative in relatives:
        path = root_path(relative)
        expected = expected_source(lane, relative)
        if not path.is_file():
            raise RecipeError(f"missing production source: {path}")
        actual = path.read_bytes()
        if actual != expected:
            raise RecipeError(
                f"{relative}: source does not match the {lane} staging snapshot; "
                f"expected {digest(expected)}, found {digest(actual)}"
            )
        originals[relative] = actual

    fragment = FRAGMENT.read_bytes()
    if not fragment or b"TRACE0721" not in fragment or b"document_scope" not in fragment:
        raise RecipeError(f"trace fragment is empty or missing TRACE0721 hooks: {FRAGMENT}")
    transformed = {relative: transform(relative, data, lane) for relative, data in originals.items()}
    entries = [
        {
            "relative": relative,
            "original_sha256": digest(originals[relative]),
            "transformed_sha256": digest(transformed[relative]),
            "original_bytes": len(originals[relative]),
            "transformed_bytes": len(transformed[relative]),
            "backup": f"originals/{relative}.orig",
        }
        for relative in relatives
    ]
    manifest: dict[str, object] = {
        "version": 1,
        "status": "prepared",
        "lane": lane,
        "repo_root": str(ROOT),
        "scratch": str(SCRATCH),
        "module": str(MODULE),
        "fragment_sha256": digest(fragment),
        "sources": entries,
    }

    SCRATCH.mkdir()
    changed: list[str] = []
    try:
        for entry in entries:
            backup = SCRATCH / str(entry["backup"])
            backup.parent.mkdir(parents=True, exist_ok=True)
            backup.write_bytes(originals[str(entry["relative"])])
        MODULE.write_bytes(fragment)
        write_json(manifest_path(), manifest)
        for relative in relatives:
            path = root_path(relative)
            path.write_bytes(transformed[relative])
            actual = digest(path.read_bytes())
            expected = digest(transformed[relative])
            if actual != expected:
                raise RecipeError(f"{relative}: transformed hash verification failed")
            changed.append(relative)
        manifest["status"] = "applied"
        write_json(manifest_path(), manifest)
    except BaseException:
        for relative in reversed(changed):
            root_path(relative).write_bytes(originals[relative])
        raise
    print(f"applied TRACE0721 {lane} instrumentation: {SCRATCH}")


def restore() -> None:
    manifest = read_manifest()
    lane = manifest.get("lane")
    if lane not in {"baseline", "candidate"}:
        raise RecipeError("trace manifest has no valid lane")
    values = manifest.get("sources")
    if not isinstance(values, list):
        raise RecipeError("trace manifest has no source list")
    entries: dict[str, dict[str, object]] = {}
    for value in values:
        if not isinstance(value, dict) or not isinstance(value.get("relative"), str):
            raise RecipeError("trace manifest contains an invalid source entry")
        entries[str(value["relative"])] = value
    if set(entries) != set(patched_relatives(str(lane))):
        raise RecipeError("trace manifest source set does not match the lane recipe")

    current: dict[str, bytes] = {}
    for relative, entry in entries.items():
        path = root_path(relative)
        data = path.read_bytes()
        current[relative] = data
        original = str(entry.get("original_sha256"))
        transformed = str(entry.get("transformed_sha256"))
        if digest(data) not in {original, transformed}:
            raise RecipeError(
                f"{relative}: source changed after apply; refusing to overwrite {digest(data)}"
            )
        backup = SCRATCH / str(entry.get("backup"))
        if not backup.is_file() or digest(backup.read_bytes()) != original:
            raise RecipeError(f"{relative}: original backup is missing or has the wrong hash")

    for relative, entry in entries.items():
        if digest(current[relative]) == str(entry["transformed_sha256"]):
            root_path(relative).write_bytes((SCRATCH / str(entry["backup"])).read_bytes())
        if digest(root_path(relative).read_bytes()) != str(entry["original_sha256"]):
            raise RecipeError(f"{relative}: restore hash verification failed")

    shutil.rmtree(SCRATCH)
    print(f"restored TRACE0721 {lane} instrumentation and removed {SCRATCH}")


def check() -> None:
    if SCRATCH.exists():
        manifest = read_manifest()
        print(f"scratch present: {SCRATCH} (status={manifest.get('status')!r})")
        return
    print("no TRACE0721 instrumentation scratch directory is present")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("apply", "restore", "check"))
    parser.add_argument("lane", nargs="?", choices=("baseline", "candidate"))
    args = parser.parse_args()
    try:
        if args.action == "apply":
            if args.lane is None:
                raise RecipeError("apply requires a lane: baseline or candidate")
            apply(args.lane)
        elif args.action == "restore":
            if args.lane is not None:
                raise RecipeError("restore takes no lane; the manifest records it")
            restore()
        else:
            if args.lane is not None:
                raise RecipeError("check takes no lane")
            check()
    except (OSError, RecipeError) as error:
        print(f"trace.py: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
