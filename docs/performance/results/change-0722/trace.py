#!/usr/bin/env python3
"""Apply and restore the source-bound 0722 DOCX differential tracer.

The recipe is a guarded source swap. It accepts only packet snapshots, keeps
its Rust module outside the checkout, and refuses to restore a file changed
while tracing. Reader calls and range-state callbacks are separate signals.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
FRAGMENT = PACKET / "trace.fragment"
SCRATCH = Path("/home/zhuhe/code/litchi-scratch-0722-trace")
MODULE = SCRATCH / "scan_trace_0722.rs"

LIB_RELATIVE = "crates/litchi-docx/src/lib.rs"
BASELINE_LIB_SHA256 = "ba9965127f1b6ddfdc2cd5fe5e4684479e2c7b6848d48b945489e30296d87225"
PACKAGE_RELATIVE = "crates/litchi-docx/src/writer/doc/package.rs"
BASELINE_PATCHED = (
    LIB_RELATIVE,
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/parts/document_part.rs",
    PACKAGE_RELATIVE,
)


class RecipeError(RuntimeError):
    """A refusal caused by stale or unexpected source state."""


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def root_path(relative: str) -> Path:
    return ROOT / relative


def snapshot_path(lane: str, relative: str) -> Path:
    return PACKET / lane / relative


def source_manifest(lane: str) -> dict[str, str]:
    path = PACKET / f"source-{lane}.json"
    if not path.is_file():
        raise RecipeError(f"missing {lane} source manifest: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RecipeError(f"cannot read {path}: {error}") from error
    if not isinstance(value, dict) or not all(
        isinstance(name, str) and isinstance(item, str) for name, item in value.items()
    ):
        raise RecipeError(f"{path}: source manifest is not a hash map")
    return value


def candidate_trace_relatives() -> tuple[str, ...]:
    """Select the public codec and the private fused implementation file."""
    source_root = PACKET / "candidate" / "crates/litchi-docx/src"
    if not source_root.is_dir():
        raise RecipeError(f"missing candidate source tree: {source_root}")
    selected: set[str] = {LIB_RELATIVE, PACKAGE_RELATIVE}
    for path in sorted(source_root.rglob("*.rs")):
        relative = str(path.relative_to(PACKET / "candidate"))
        data = path.read_bytes()
        # codec.rs remains the public owner in 0722; a child may hold only the
        # new private writer scan and is selected by its fused symbol.
        if path.name in {"codec.rs", "document_scan.rs"} and b"pub fn scan(" in data:
            selected.add(relative)
        if (
            b"pub(crate) fn scan_with_block_ranges" in data
            or b"fn scan_alt_and_block_ranges(\n" in data
        ):
            selected.add(relative)
    return tuple(sorted(selected))


def patched_relatives(lane: str) -> tuple[str, ...]:
    if lane == "baseline":
        return BASELINE_PATCHED
    if lane == "candidate":
        return candidate_trace_relatives()
    raise RecipeError(f"unknown lane {lane!r}")


def expected_source(lane: str, relative: str) -> bytes:
    path = snapshot_path(lane, relative)
    if path.is_file():
        return path.read_bytes()
    if relative == LIB_RELATIVE:
        data = root_path(relative).read_bytes()
        if digest(data) != BASELINE_LIB_SHA256:
            raise RecipeError(
                f"{relative}: expected unchanged library source before tracing; "
                f"found {digest(data)}"
            )
        return data
    raise RecipeError(f"missing {lane} source snapshot: {path}")


def replace_once(source: bytes, needle: bytes, replacement: bytes, label: str) -> bytes:
    count = source.count(needle)
    if count != 1:
        raise RecipeError(f"{label}: expected one source occurrence, found {count}")
    return source.replace(needle, replacement, 1)


def transform_lib(source: bytes) -> bytes:
    if b"mod scan_trace_0722;" in source:
        raise RecipeError(f"{LIB_RELATIVE}: trace module declaration already present")
    if not source.endswith(b"\n"):
        raise RecipeError(f"{LIB_RELATIVE}: expected a final newline")
    declaration = (
        f'#[path = "{MODULE.as_posix()}"]\n'
        "mod scan_trace_0722;\n"
    ).encode()
    return source + b"\n" + declaration


def transform_active(source: bytes) -> bytes:
    active_start = source.find(b"pub fn active(")
    if active_start < 0:
        raise RecipeError("active hook: public active function is missing")
    active_end = source.find(b"\n///", active_start)
    if active_end < 0:
        active_end = len(source)
    section = source[active_start:active_end]
    section = replace_once(
        section,
        b"    validate_xml(xml)?;",
        b"""    if let Err(error) = validate_xml(xml) {
        crate::scan_trace_0722::active_error(xml, offsets, &error, "validate");
        return Err(error);
    }""",
        "active validation boundary",
    )
    pattern = rb"(?ms)^    litchi_ooxml_common::mce::active_offsets\(.*?^    \.map_err\(Error::from\)\n"
    matches = list(re.finditer(pattern, section))
    if len(matches) != 1:
        raise RecipeError(f"active MCE boundary: expected one source match, found {len(matches)}")
    match = matches[0]
    call = match.group(0)
    replacement = (
        b"    let result = "
        + call[4:].rstrip(b"\n")
        + b";\n"
        + b"    crate::scan_trace_0722::active_result(xml, offsets, &result, \"mce\");\n"
        + b"    result\n"
    )
    section = section[:match.start()] + replacement + section[match.end():]
    return source[:active_start] + section + source[active_end:]


def transform_alt_scanner(source: bytes) -> bytes:
    source = transform_active(source)
    if b"TRACE0722" in source:
        raise RecipeError("alt scanner is already instrumented")
    source = replace_once(
        source,
        b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;""",
        b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
    let _trace = crate::scan_trace_0722::alt_scan_scope(xml);""",
        "alt scan scope",
    )
    read_event = b"""        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();"""
    source = replace_once(
        source,
        read_event,
        read_event
        + b"""
        crate::scan_trace_0722::reader_read(
            "alt",
            u64::from(event_start),
            Some(u64::try_from(reader.buffer_position()).unwrap_or(u64::MAX)),
        );""",
        "alt successful XML reader boundary",
    )
    source = replace_once(
        source,
        b"""    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;""",
        b"""    crate::scan_trace_0722::alt_stage("raw_chunks");
    crate::scan_trace_0722::alt_raw_chunks(&chunks);
    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;""",
        "alt raw chunk boundary",
    )
    source = replace_once(
        source,
        b"""    Ok(chunks)
}""",
        b"""    crate::scan_trace_0722::alt_stage("selected_chunks");
    crate::scan_trace_0722::alt_selected_chunks(&chunks);
    crate::scan_trace_0722::alt_complete();
    Ok(chunks)
}""",
        "alt selected chunk boundary",
    )
    return source


def transform_namespace(source: bytes) -> bytes:
    if b"TRACE0722" in source:
        raise RecipeError("namespace scanner is already instrumented")
    start = source.find(b"pub(crate) fn scan_word_element_ranges(")
    if start < 0:
        raise RecipeError("range scanner function is missing")
    end = source.find(b"\npub(crate) fn ", start + 1)
    if end < 0:
        end = len(source)
    section = source[start:end]
    read_resolved = b"""            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;"""
    section = replace_once(
        section,
        read_resolved,
        read_resolved
        + b"""
            // Count the successful read before node/depth admission guards.
            crate::scan_trace_0722::reader_read(
                "range",
                u64::try_from(event_start).unwrap_or(u64::MAX),
                None,
            );""",
        "baseline range reader boundary",
    )
    event_end = b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;"""
    section = replace_once(
        section,
        event_end,
        event_end
        + b"""
        crate::scan_trace_0722::range_observer(
            u64::try_from(event_start).unwrap_or(u64::MAX),
            Some(u64::try_from(event_end).unwrap_or(u64::MAX)),
        );""",
        "baseline range observer boundary",
    )
    return source[:start] + section + source[end:]


def transform_document_part(source: bytes) -> bytes:
    if b"TRACE0722" in source:
        raise RecipeError("document range scanner is already instrumented")
    start = source.find(b"pub(crate) fn active_block_ranges(")
    if start < 0:
        raise RecipeError("baseline active range scanner is missing")
    end = source.find(b"\npub(crate) fn ", start + 1)
    if end < 0:
        end = len(source)
    section = source[start:end]
    section = replace_once(
        section,
        b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let mut ranges = Vec::new();""",
        b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let _trace = crate::scan_trace_0722::range_scan_scope(xml);
    let mut ranges = Vec::new();""",
        "baseline range scan scope",
    )
    section = replace_once(
        section,
        b"""    )?;
    let mut starts = Vec::new();""",
        b"""    )?;
    crate::scan_trace_0722::range_stage("raw_ranges");
    crate::scan_trace_0722::range_raw_ranges(&ranges);
    let mut starts = Vec::new();""",
        "baseline raw range boundary",
    )
    section = replace_once(
        section,
        b"""    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)""",
        b"""    crate::scan_trace_0722::range_stage("active");
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    crate::scan_trace_0722::range_stage("filter");
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    crate::scan_trace_0722::range_selected_ranges(&ranges);
    crate::scan_trace_0722::range_complete();
    Ok(ranges)""",
        "baseline selected range boundary",
    )
    return source[:start] + section + source[end:]


def transform_package(source: bytes, lane: str) -> bytes:
    source = replace_once(
        source,
        b"        let bytes = xml.as_bytes();",
        b"""        let _trace_document = crate::scan_trace_0722::document_scope(xml.as_bytes());
        let bytes = xml.as_bytes();""",
        "DocumentBody boundary scope",
    )
    # In 0722 the fused implementation belongs to a private child of codec.rs;
    # the package adapter is only the document boundary.
    return source


def transform_fused_owner(source: bytes) -> bytes:
    markers = (b"scan_alt_and_block_ranges", b"scan_with_block_ranges")
    if not any(marker in source for marker in markers):
        raise RecipeError("candidate writer-local fused scan helper is missing")
    if b"TRACE0722" in source:
        raise RecipeError("candidate fused owner is already instrumented")
    source = replace_once(
        source,
        b"""    validate_xml(xml)?;
    let targets =""",
        b"""    validate_xml(xml)?;
    let _trace_alt = crate::scan_trace_0722::alt_scan_scope(xml);
    let _trace_range = crate::scan_trace_0722::range_scan_scope(xml);
    let targets =""",
        "candidate fused scan scopes",
    )
    read_pattern = rb"(?ms)(?P<indent>\s*)let event = reader\s*\.read_(?:resolved_)?event\(\)\s*\.map_err\(.*?\)\?\s*\.into_owned\(\);"
    matches = list(re.finditer(read_pattern, source))
    if not matches:
        raise RecipeError("candidate fused owner has no instrumentable XML reader")
    match = matches[0]
    indent = match.group("indent")
    replacement = (
        match.group(0)
        + b"\n"
        + indent
        + b"crate::scan_trace_0722::reader_read(\n"
        + indent
        + b"    \"alt\",\n"
        + indent
        + b"    u64::try_from(event_start).unwrap_or(u64::MAX),\n"
        + indent
        + b"    Some(u64::try_from(reader.buffer_position()).unwrap_or(u64::MAX)),\n"
        + indent
        + b");"
    )
    source = source[:match.start()] + replacement + source[match.end():]
    classify = b"""            let scan_event = match range_scanner.classify(&namespace, &event) {"""
    source = replace_once(
        source,
        classify,
        b"""            crate::scan_trace_0722::range_observer(
                u64::try_from(event_start).unwrap_or(u64::MAX),
                Some(u64::try_from(reader.buffer_position()).unwrap_or(u64::MAX)),
            );
            let scan_event = match range_scanner.classify(&namespace, &event) {""",
        "candidate range observer boundary",
    )
    raw_chunks = next(
        (
            needle
            for needle in (
                b"    let offsets = alt_scanner.chunks.keys().copied().collect::<Vec<_>>();",
                b"    let offsets = chunks.keys().copied().collect::<Vec<_>>();",
            )
            if source.count(needle) == 1
        ),
        None,
    )
    if raw_chunks is not None:
        source = source.replace(
            raw_chunks,
            b"""    crate::scan_trace_0722::alt_stage("raw_chunks");
    crate::scan_trace_0722::alt_raw_chunks(&alt_scanner.chunks);
    let offsets = alt_scanner.chunks.keys().copied().collect::<Vec<_>>();""",
            1,
        )
    elif b"chunks.keys" in source:
        raise RecipeError("candidate fused raw chunk boundary is ambiguous")
    else:
        raise RecipeError("candidate fused owner has no raw chunk boundary")
    deferred = b"""    if let Some(error) = range_error {
        return Err(error);
    }"""
    if source.count(deferred) != 1:
        raise RecipeError("candidate deferred range error boundary is missing or ambiguous")
    # The first MCE decision belongs to the old alt scan and remains visible
    # even when the deferred range state later refuses.  Raw ranges are
    # intentionally recorded only after this refusal boundary.
    source = source.replace(
        deferred,
        b"""    crate::scan_trace_0722::alt_stage("selected_chunks");
    crate::scan_trace_0722::alt_selected_chunks(&alt_scanner.chunks);
    crate::scan_trace_0722::alt_complete();
""" + deferred
        + b"""
    crate::scan_trace_0722::range_stage("raw_ranges");
    crate::scan_trace_0722::range_raw_ranges(&ranges);""",
        1,
    )
    selected = b"    ranges.retain(|&(_, start, _)| selected.contains(&start));"
    if source.count(selected) != 1:
        raise RecipeError("candidate fused selected range boundary is missing or ambiguous")
    source = source.replace(
        selected,
        selected
        + b"""
    crate::scan_trace_0722::range_stage("filter");
    crate::scan_trace_0722::range_selected_ranges(&ranges);""",
        1,
    )
    completion = b"    Ok((alt_scanner.chunks, ranges))"
    if source.count(completion) != 1:
        raise RecipeError("candidate fused completion boundary is missing or ambiguous")
    source = source.replace(
        completion,
        b"""    crate::scan_trace_0722::range_complete();
    Ok((alt_scanner.chunks, ranges))""",
        1,
    )
    return source


def transform(relative: str, source: bytes, lane: str) -> bytes:
    if relative == LIB_RELATIVE:
        return transform_lib(source)
    if relative.endswith("/namespace.rs"):
        if lane != "baseline":
            raise RecipeError("candidate namespace scanner must remain unchanged")
        return transform_namespace(source)
    if relative.endswith("/parts/document_part.rs"):
        if lane != "baseline":
            raise RecipeError("candidate document-part range scanner must remain unchanged")
        return transform_document_part(source)
    if relative == PACKAGE_RELATIVE:
        return transform_package(source, lane)
    if lane == "baseline" and relative.endswith("/alt/codec.rs"):
        return transform_alt_scanner(source)
    if lane == "candidate" and (
        relative.endswith("/alt/codec.rs") or relative.endswith("/alt/codec/document_scan.rs")
    ) and b"pub fn scan(" in source:
        return transform_alt_scanner(source)
    if lane == "candidate" and (
        b"scan_alt_and_block_ranges" in source or b"scan_with_block_ranges" in source
    ):
        return transform_fused_owner(source)
    raise RecipeError(f"no transform for {lane} source {relative}")


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
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RecipeError(f"cannot read trace manifest: {error}") from error
    if not isinstance(value, dict) or value.get("version") != 2:
        raise RecipeError("trace manifest has an unexpected format")
    return value


def apply(lane: str) -> None:
    if SCRATCH.exists():
        raise RecipeError(f"owned scratch path already exists: {SCRATCH}; restore it first")
    relatives = patched_relatives(lane)
    hashes = source_manifest(lane)
    originals: dict[str, bytes] = {}
    for relative in relatives:
        expected = expected_source(lane, relative)
        if hashes.get(relative) != digest(expected):
            raise RecipeError(f"{relative}: staging bytes do not match source-{lane}.json")
        path = root_path(relative)
        if not path.is_file():
            raise RecipeError(f"missing production source: {path}")
        actual = path.read_bytes()
        if actual != expected:
            raise RecipeError(
                f"{relative}: source does not match {lane} staging; "
                f"expected {digest(expected)}, found {digest(actual)}"
            )
        originals[relative] = actual
    fragment = FRAGMENT.read_bytes()
    if not fragment or b"TRACE0722" not in fragment or b"document_scope" not in fragment:
        raise RecipeError(f"trace fragment is empty or missing TRACE0722 hooks: {FRAGMENT}")
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
        "version": 2,
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
            if digest(path.read_bytes()) != digest(transformed[relative]):
                raise RecipeError(f"{relative}: transformed hash verification failed")
            changed.append(relative)
        manifest["status"] = "applied"
        write_json(manifest_path(), manifest)
    except BaseException:
        for relative in reversed(changed):
            root_path(relative).write_bytes(originals[relative])
        shutil.rmtree(SCRATCH, ignore_errors=True)
        raise
    print(f"applied TRACE0722 {lane} instrumentation: {SCRATCH}")


def restore() -> None:
    manifest = read_manifest()
    lane = manifest.get("lane")
    if lane not in {"baseline", "candidate"}:
        raise RecipeError("trace manifest has no valid lane")
    values = manifest.get("sources")
    if not isinstance(values, list):
        raise RecipeError("trace manifest has no source list")
    expected = set(patched_relatives(str(lane)))
    entries: dict[str, dict[str, object]] = {}
    for value in values:
        if not isinstance(value, dict) or not isinstance(value.get("relative"), str):
            raise RecipeError("trace manifest contains an invalid source entry")
        entries[str(value["relative"])] = value
    if set(entries) != expected:
        raise RecipeError("trace manifest source set does not match the lane recipe")
    current: dict[str, bytes] = {}
    for relative, entry in entries.items():
        path = root_path(relative)
        data = path.read_bytes()
        current[relative] = data
        original = str(entry.get("original_sha256"))
        transformed = str(entry.get("transformed_sha256"))
        if digest(data) not in {original, transformed}:
            raise RecipeError(f"{relative}: source changed after apply; refusing restore")
        backup = SCRATCH / str(entry.get("backup"))
        if not backup.is_file() or digest(backup.read_bytes()) != original:
            raise RecipeError(f"{relative}: original backup is missing or has the wrong hash")
    for relative, entry in entries.items():
        if digest(current[relative]) == str(entry["transformed_sha256"]):
            root_path(relative).write_bytes((SCRATCH / str(entry["backup"])).read_bytes())
        if digest(root_path(relative).read_bytes()) != str(entry["original_sha256"]):
            raise RecipeError(f"{relative}: restore hash verification failed")
    shutil.rmtree(SCRATCH)
    print(f"restored TRACE0722 {lane} instrumentation and removed {SCRATCH}")


def check() -> None:
    if SCRATCH.exists():
        manifest = read_manifest()
        print(f"scratch present: {SCRATCH} (status={manifest.get('status')!r})")
    else:
        print("no TRACE0722 instrumentation scratch directory is present")


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
                raise RecipeError("restore takes no lane; manifest records it")
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
