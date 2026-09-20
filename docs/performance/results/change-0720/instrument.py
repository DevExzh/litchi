#!/usr/bin/env python3
"""Apply and restore the narrowly scoped TRACE0720 diagnostic instrumentation.

The recipe is intentionally source exact. It is for one controlled diagnostic
run, not for a permanent feature: apply takes byte-for-byte backups in the
owned scratch directory, and restore refuses to overwrite a source that no
longer has the hash written by apply.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[4]
SCRATCH = Path("/home/zhuhe/code/litchi-scratch-0720")
FRAGMENT = Path(__file__).with_name("trace.fragment")
MODULE = SCRATCH / "scan_trace_0720.rs"
MODULE_LINE = (
    f'#[path = "{MODULE.as_posix()}"]\n'
    "mod scan_trace_0720;\n"
).encode()

SOURCE_SHA256 = {
    "crates/litchi-docx/src/lib.rs": "ba9965127f1b6ddfdc2cd5fe5e4684479e2c7b6848d48b945489e30296d87225",
    "crates/litchi-docx/src/namespace.rs": "1b50df0db904a6c9f39cd8780562b43d1fa45f9959989802f3b98dc5f09b527f",
    "crates/litchi-docx/src/alt/codec.rs": "4c60069945e369657a0fd3bcf968f317bf3c0d42fb7dce3795321c41a1aa9725",
    "crates/litchi-docx/src/parts/document_part.rs": "847caf3dba89a7bf948d07c3b70a4729b4e6d742dcc48dae2fdcd25ac26f65a9",
    "crates/litchi-docx/src/writer/doc/package.rs": "f5670594685a8035952169487777381e39c950515c59fcc9d176ebedf0e42e92",
}


class RecipeError(RuntimeError):
    """A refusal caused by stale or unexpected source state."""


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def source_path(relative: str) -> Path:
    return ROOT / relative


def replace_once(source: bytes, needle: bytes, replacement: bytes, label: str) -> bytes:
    count = source.count(needle)
    if count != 1:
        raise RecipeError(f"{label}: expected one source occurrence, found {count}")
    return source.replace(needle, replacement, 1)


def transform(relative: str, source: bytes) -> bytes:
    if relative == "crates/litchi-docx/src/lib.rs":
        if MODULE_LINE in source:
            raise RecipeError(f"{relative}: TRACE0720 module declaration already present")
        if not source.endswith(b"\n"):
            raise RecipeError(f"{relative}: expected a final newline")
        return source + b"\n" + MODULE_LINE

    if relative == "crates/litchi-docx/src/alt/codec.rs":
        source = replace_once(
            source,
            b"""pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    validate_xml(xml)?;""",
            b"""pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    if let Err(error) = validate_xml(xml) {
        crate::scan_trace_0720::record_active_error(xml, offsets, &error, "validate");
        return Err(error);
    }""",
            f"{relative}: active validation boundary",
        )
        source = replace_once(
            source,
            b"""    litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map_err(Error::from)
}""",
            b"""    let result = litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map_err(Error::from);
    crate::scan_trace_0720::record_active_result(xml, offsets, &result, "mce");
    result
}""",
            f"{relative}: active MCE result",
        )
        source = replace_once(
            source,
            b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
""",
            b"""pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    let _trace = crate::scan_trace_0720::alt_scan_scope(xml);
    validate_xml(xml)?;
    crate::scan_trace_0720::alt_scan_stage("events");
""",
            f"{relative}: alt scan scope",
        )
        source = replace_once(
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
        crate::scan_trace_0720::record_alt_scan_event(event_start, reader.buffer_position());
""",
            f"{relative}: alt scan event counter",
        )
        source = replace_once(
            source,
            b"""    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;
""",
            b"""    crate::scan_trace_0720::alt_scan_stage("raw_chunks");
    crate::scan_trace_0720::record_alt_scan_raw_chunks(&chunks);
    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    crate::scan_trace_0720::alt_scan_stage("active");
    let active = active(xml, &offsets)?;
    crate::scan_trace_0720::alt_scan_stage("filter");
""",
            f"{relative}: alt scan raw and active boundary",
        )
        source = replace_once(
            source,
            b"""    });
    Ok(chunks)
}

fn validate_xml""",
            b"""    });
    crate::scan_trace_0720::record_alt_scan_selected_chunks(&chunks);
    crate::scan_trace_0720::alt_scan_complete();
    Ok(chunks)
}

fn validate_xml""",
            f"{relative}: alt scan completion",
        )
        return source

    if relative == "crates/litchi-docx/src/namespace.rs":
        return replace_once(
            source,
            b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;

        match event {""",
            b"""        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        crate::scan_trace_0720::record_range_event(event_start, event_end);

        match event {""",
            f"{relative}: structural event counter",
        )

    if relative == "crates/litchi-docx/src/parts/document_part.rs":
        source = replace_once(
            source,
            b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let mut ranges = Vec::new();""",
            b"""pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let _trace = crate::scan_trace_0720::range_pass_scope(xml);
    let mut ranges = Vec::new();""",
            f"{relative}: active range pass scope",
        )
        source = replace_once(
            source,
            b"""    )?;
    let mut starts = Vec::new();""",
            b"""    )?;
    crate::scan_trace_0720::range_pass_stage("raw_ranges");
    crate::scan_trace_0720::record_raw_ranges(&ranges);
    let mut starts = Vec::new();""",
            f"{relative}: raw range capture",
        )
        source = replace_once(
            source,
            b"""    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)""",
            b"""    crate::scan_trace_0720::range_pass_stage("active");
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    crate::scan_trace_0720::range_pass_stage("filter");
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    crate::scan_trace_0720::record_selected_ranges(&ranges);
    crate::scan_trace_0720::range_pass_complete();
    Ok(ranges)""",
            f"{relative}: selected range capture",
        )
        return source

    if relative == "crates/litchi-docx/src/writer/doc/package.rs":
        return replace_once(
            source,
            b"""        };
        let bytes = xml.as_bytes();
        let mut chunks = scan(bytes)?;""",
            b"""        };
        let _trace = crate::scan_trace_0720::document_scope(xml.as_bytes());
        let bytes = xml.as_bytes();
        let mut chunks = scan(bytes)?;""",
            f"{relative}: BOM-stripped DocumentBody boundary",
        )

    raise RecipeError(f"unknown source plan for {relative}")


def checked_sources() -> dict[str, bytes]:
    sources: dict[str, bytes] = {}
    for relative, expected in SOURCE_SHA256.items():
        path = source_path(relative)
        if not path.is_file():
            raise RecipeError(f"missing source file: {path}")
        data = path.read_bytes()
        actual = digest(data)
        if actual != expected:
            raise RecipeError(
                f"{relative}: original SHA-256 mismatch; expected {expected}, found {actual}"
            )
        sources[relative] = data
    return sources


def manifest_path() -> Path:
    return SCRATCH / "manifest.json"


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def read_manifest() -> dict[str, object]:
    path = manifest_path()
    if not path.is_file():
        raise RecipeError(f"missing instrumentation manifest: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RecipeError(f"cannot read instrumentation manifest: {error}") from error
    if not isinstance(value, dict) or value.get("version") != 1:
        raise RecipeError("instrumentation manifest has an unexpected format")
    return value


def apply() -> None:
    if SCRATCH.exists():
        raise RecipeError(
            f"owned scratch path already exists: {SCRATCH}; restore or remove it after inspection"
        )
    sources = checked_sources()
    fragment = FRAGMENT.read_bytes()
    if not fragment:
        raise RecipeError(f"empty trace fragment: {FRAGMENT}")
    if b"TRACE0720" not in fragment or b"document_scope" not in fragment:
        raise RecipeError(
            f"trace fragment does not contain the expected TRACE0720 module: {FRAGMENT}"
        )

    transformed = {relative: transform(relative, data) for relative, data in sources.items()}
    entries = []
    for relative in SOURCE_SHA256:
        entries.append(
            {
                "relative": relative,
                "original_sha256": digest(sources[relative]),
                "transformed_sha256": digest(transformed[relative]),
                "original_bytes": len(sources[relative]),
                "transformed_bytes": len(transformed[relative]),
                "backup": f"originals/{relative}.orig",
            }
        )
    manifest = {
        "version": 1,
        "status": "prepared",
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
            backup.write_bytes(sources[str(entry["relative"])])
        MODULE.write_bytes(fragment)
        write_json(manifest_path(), manifest)
        for relative in SOURCE_SHA256:
            path = source_path(relative)
            path.write_bytes(transformed[relative])
            actual = digest(path.read_bytes())
            expected = next(
                str(entry["transformed_sha256"])
                for entry in entries
                if entry["relative"] == relative
            )
            if actual != expected:
                raise RecipeError(f"{relative}: transformed hash verification failed")
            changed.append(relative)
        manifest["status"] = "applied"
        write_json(manifest_path(), manifest)
    except BaseException:
        for relative in reversed(changed):
            source_path(relative).write_bytes(sources[relative])
        raise

    print(f"applied TRACE0720 instrumentation; backups and hashes: {SCRATCH}")


def restore() -> None:
    manifest = read_manifest()
    source_entries = manifest.get("sources")
    if not isinstance(source_entries, list):
        raise RecipeError("instrumentation manifest has no source entries")

    entries: dict[str, dict[str, object]] = {}
    for value in source_entries:
        if not isinstance(value, dict) or not isinstance(value.get("relative"), str):
            raise RecipeError("instrumentation manifest contains an invalid source entry")
        entries[str(value["relative"])] = value
    if set(entries) != set(SOURCE_SHA256):
        raise RecipeError("instrumentation manifest source set does not match this recipe")

    current: dict[str, bytes] = {}
    for relative, entry in entries.items():
        path = source_path(relative)
        data = path.read_bytes()
        current[relative] = data
        original = str(entry.get("original_sha256"))
        transformed = str(entry.get("transformed_sha256"))
        if digest(data) not in {original, transformed}:
            raise RecipeError(
                f"{relative}: source changed after apply; refusing to overwrite SHA-256 {digest(data)}"
            )
        backup = SCRATCH / str(entry.get("backup"))
        if not backup.is_file() or digest(backup.read_bytes()) != original:
            raise RecipeError(f"{relative}: original backup is missing or has the wrong hash")

    for relative, entry in entries.items():
        if digest(current[relative]) == str(entry["transformed_sha256"]):
            source_path(relative).write_bytes((SCRATCH / str(entry["backup"])).read_bytes())
        if digest(source_path(relative).read_bytes()) != str(entry["original_sha256"]):
            raise RecipeError(f"{relative}: restore hash verification failed")

    shutil.rmtree(SCRATCH)
    print("restored TRACE0720 instrumentation and removed the owned scratch directory")


def check() -> None:
    if SCRATCH.exists():
        manifest = read_manifest()
        print(f"scratch present: {SCRATCH} (status={manifest.get('status')!r})")
        return
    checked_sources()
    print("working tree has the exact uninstrumented source hashes")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("apply", "restore", "check"))
    args = parser.parse_args()
    try:
        {"apply": apply, "restore": restore, "check": check}[args.action]()
    except (OSError, RecipeError) as error:
        print(f"instrument.py: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
