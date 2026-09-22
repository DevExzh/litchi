//! Bounded current-source DOC/PPT save and common-container probes.
//!
//! This probe deliberately keeps the two scopes separate.  The `format`
//! operation measures the public format owner from open through commit.  The
//! `container` operation opens the common OLE editor, stages the exact stream
//! replacements produced by one untimed public edit, and finishes it.  The
//! latter is a logical container control; it is not a claim about the nested
//! PPT editor's internal save route.

use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_cfb::{DirectoryEntry, OleFile, SectorLayoutPolicy, validate_source};
use litchi_core::{OwnedSource, Position};
use litchi_ole_common::object::{Editor, Limits, Targets};
use serde::Serialize;
use sha2::{Digest, Sha256};

pub mod alloc_metrics;

pub type BoxError = Box<dyn Error>;

#[derive(Debug)]
struct ProbeError(String);

impl Display for ProbeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for ProbeError {}

fn failure<T>(message: impl Into<String>) -> Result<T, BoxError> {
    Err(Box::new(ProbeError(message.into())))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Case {
    DocFloat,
    DocNoHf,
    Ppt45543,
}

impl Case {
    fn name(self) -> &'static str {
        match self {
            Self::DocFloat => "docfloat",
            Self::DocNoHf => "docnohf",
            Self::Ppt45543 => "ppt45543",
        }
    }

    fn format_name(self) -> &'static str {
        match self {
            Self::DocFloat | Self::DocNoHf => "doc",
            Self::Ppt45543 => "ppt",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Format,
    Container,
}

impl Operation {
    fn name(self) -> &'static str {
        match self {
            Self::Format => "format",
            Self::Container => "container",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Policy {
    Reuse,
    Rewrite,
}

impl Policy {
    fn name(self) -> &'static str {
        match self {
            Self::Reuse => "reuse",
            Self::Rewrite => "rewrite",
        }
    }

    fn cfb(self) -> SectorLayoutPolicy {
        match self {
            Self::Reuse => SectorLayoutPolicy::Reuse,
            Self::Rewrite => SectorLayoutPolicy::Rewrite,
        }
    }
}

struct Args {
    case: Case,
    input: PathBuf,
    operation: Operation,
    policy: Policy,
    warmups: usize,
    samples: usize,
    text: String,
}

fn parse_count(flag: &str, value: &str) -> Result<usize, BoxError> {
    let count = value.parse::<usize>().map_err(|error| {
        Box::new(ProbeError(format!(
            "invalid {flag} value {value:?}: {error}"
        ))) as BoxError
    })?;
    if count > 10_000 {
        return failure(format!("{flag} exceeds probe bound 10000"));
    }
    Ok(count)
}

fn parse_args() -> Result<Args, BoxError> {
    let mut case = None;
    let mut input = None;
    let mut operation = Operation::Format;
    let mut policy = Policy::Reuse;
    let mut warmups = 1usize;
    let mut samples = 1usize;
    // This is the same kind of length-changing replacement used by the 0617
    // public-shape probe.  It is intentionally deterministic and overridable
    // for a repeated measurement with the exact same source fixture.
    let mut text = String::from("litchi copy-through baseline replacement text");

    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut next_value = || {
            arguments.next().ok_or_else(|| {
                Box::new(ProbeError(format!("missing value for {flag}"))) as BoxError
            })
        };
        match flag.as_str() {
            "--case" => {
                case = Some(match next_value()?.as_str() {
                    "docfloat" => Case::DocFloat,
                    "docnohf" => Case::DocNoHf,
                    "ppt45543" => Case::Ppt45543,
                    other => return failure(format!("unknown --case {other:?}")),
                });
            },
            "--input" => input = Some(PathBuf::from(next_value()?)),
            "--operation" => {
                operation = match next_value()?.as_str() {
                    "format" => Operation::Format,
                    "container" => Operation::Container,
                    other => return failure(format!("unknown --operation {other:?}")),
                };
            },
            "--policy" => {
                policy = match next_value()?.as_str() {
                    "reuse" => Policy::Reuse,
                    "rewrite" => Policy::Rewrite,
                    other => return failure(format!("unknown --policy {other:?}")),
                };
            },
            "--warmups" => warmups = parse_count("--warmups", &next_value()?)?,
            "--samples" => samples = parse_count("--samples", &next_value()?)?,
            "--text" => text = next_value()?,
            "--oracle-only" => {
                warmups = 0;
                samples = 0;
            },
            other => return failure(format!("unknown flag {other:?}")),
        }
    }

    Ok(Args {
        case: case.ok_or_else(|| Box::new(ProbeError("missing --case".into())) as BoxError)?,
        input: input.ok_or_else(|| Box::new(ProbeError("missing --input".into())) as BoxError)?,
        operation,
        policy,
        warmups,
        samples,
        text,
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

fn update_len_prefixed(hasher: &mut Sha256, bytes: &[u8]) {
    hasher.update((bytes.len() as u64).to_le_bytes());
    hasher.update(bytes);
}

#[derive(Clone, Debug, Serialize)]
struct StreamSummary {
    path: Vec<String>,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct StorageSummary {
    path: Vec<String>,
    clsid: String,
}

#[derive(Clone, Debug, Serialize)]
struct InventorySummary {
    file_bytes: usize,
    sector_size: usize,
    root_clsid: String,
    streams: Vec<StreamSummary>,
    storages: Vec<StorageSummary>,
    directory_entries: Vec<DirectoryEntrySummary>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct DirectoryEntrySummary {
    path: Vec<String>,
    entry_type: u8,
    clsid: String,
    bytes: usize,
    start_sector: u32,
    is_minifat: bool,
}

#[derive(Clone, Debug)]
struct Inventory {
    summary: InventorySummary,
    stream_bytes: BTreeMap<Vec<String>, Vec<u8>>,
    storage_clsids: BTreeMap<Vec<String>, String>,
    directory_entries: BTreeMap<Vec<String>, DirectoryEntrySummary>,
}

fn collect_storages(
    entry: &DirectoryEntry,
    prefix: &mut Vec<String>,
    output: &mut BTreeMap<Vec<String>, String>,
) {
    for child in &entry.children {
        if child.entry_type == litchi_cfb::consts::STGTY_STORAGE {
            prefix.push(child.name.clone());
            output.insert(prefix.clone(), child.clsid.clone());
            collect_storages(child, prefix, output);
            prefix.pop();
        }
    }
}

fn collect_directory_entries(
    entry: &DirectoryEntry,
    path: &mut Vec<String>,
    output: &mut BTreeMap<Vec<String>, DirectoryEntrySummary>,
) {
    output.insert(
        path.clone(),
        DirectoryEntrySummary {
            path: path.clone(),
            entry_type: entry.entry_type,
            clsid: entry.clsid.clone(),
            bytes: usize::try_from(entry.size).unwrap_or(usize::MAX),
            start_sector: entry.start_sector,
            is_minifat: entry.is_minifat,
        },
    );
    for child in &entry.children {
        path.push(child.name.clone());
        collect_directory_entries(child, path, output);
        path.pop();
    }
}

fn inventory(bytes: &[u8]) -> Result<Inventory, BoxError> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec()))?;
    let sector_size = ole.sector_size();
    let root = ole.root_entry().cloned().ok_or_else(|| {
        Box::new(ProbeError("CFB has no root directory entry".into())) as BoxError
    })?;
    let mut storage_clsids = BTreeMap::new();
    let mut prefix = Vec::new();
    collect_storages(&root, &mut prefix, &mut storage_clsids);
    let mut directory_entries = BTreeMap::new();
    let mut directory_path = Vec::new();
    collect_directory_entries(&root, &mut directory_path, &mut directory_entries);

    let mut paths = ole.list_streams();
    paths.sort();
    let mut stream_bytes = BTreeMap::new();
    let mut streams = Vec::with_capacity(paths.len());
    for path in paths {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = ole.open_stream(&refs)?;
        streams.push(StreamSummary {
            path: path.clone(),
            bytes: data.len(),
            sha256: sha256_hex(&data),
        });
        stream_bytes.insert(path, data);
    }
    let storages = storage_clsids
        .iter()
        .map(|(path, clsid)| StorageSummary {
            path: path.clone(),
            clsid: clsid.clone(),
        })
        .collect();
    let directory_entries_json = directory_entries.values().cloned().collect();
    Ok(Inventory {
        summary: InventorySummary {
            file_bytes: bytes.len(),
            sector_size,
            root_clsid: root.clsid,
            streams,
            storages,
            directory_entries: directory_entries_json,
        },
        stream_bytes,
        storage_clsids,
        directory_entries,
    })
}

fn structural_valid(bytes: &[u8]) -> bool {
    validate_source(Arc::new(OwnedSource::new(bytes.to_vec())))
        .map(|report| report.is_complete() && !report.has_errors())
        .unwrap_or(false)
}

#[derive(Clone, Debug, Serialize)]
struct ChangedStream {
    path: Vec<String>,
    before_bytes: usize,
    after_bytes: usize,
    before_sha256: String,
    after_sha256: String,
    length_changed: bool,
}

#[derive(Clone, Debug, Serialize)]
struct ChangedLengthProof {
    changed_streams: Vec<ChangedStream>,
    changed_stream_count: usize,
    changed_bytes_before: usize,
    changed_bytes_after: usize,
    output_file_bytes_delta: i64,
    any_stream_length_changed: bool,
    output_file_length_changed: bool,
    logical_stream_length_change_proven: bool,
}

#[derive(Clone, Debug, Serialize)]
struct DirectoryMetadataDifference {
    path: Vec<String>,
    field: String,
    expected: String,
    actual: String,
}

#[derive(Clone, Debug, Serialize)]
struct ReplacementSummary {
    path: Vec<String>,
    before_bytes: usize,
    after_bytes: usize,
    before_sha256: String,
    after_sha256: String,
}

#[derive(Clone, Debug)]
struct Replacement {
    path: Vec<String>,
    data: Arc<[u8]>,
}

fn signed_delta(after: usize, before: usize) -> i64 {
    if after >= before {
        i64::try_from(after - before).unwrap_or(i64::MAX)
    } else {
        -i64::try_from(before - after).unwrap_or(i64::MAX)
    }
}

fn changed_length_proof(source: &Inventory, expected: &Inventory) -> ChangedLengthProof {
    let mut changed_streams = Vec::new();
    let paths = source
        .stream_bytes
        .keys()
        .chain(expected.stream_bytes.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    for path in paths {
        let before = source.stream_bytes.get(&path);
        let after = expected.stream_bytes.get(&path);
        if before == after {
            continue;
        }
        let before_data = before.map_or(&[][..], Vec::as_slice);
        let after_data = after.map_or(&[][..], Vec::as_slice);
        changed_streams.push(ChangedStream {
            path,
            before_bytes: before_data.len(),
            after_bytes: after_data.len(),
            before_sha256: sha256_hex(before_data),
            after_sha256: sha256_hex(after_data),
            length_changed: before_data.len() != after_data.len(),
        });
    }
    let changed_bytes_before = changed_streams.iter().map(|item| item.before_bytes).sum();
    let changed_bytes_after = changed_streams.iter().map(|item| item.after_bytes).sum();
    let any_stream_length_changed = changed_streams.iter().any(|item| item.length_changed);
    let output_file_bytes_delta =
        signed_delta(expected.summary.file_bytes, source.summary.file_bytes);
    let output_file_length_changed = output_file_bytes_delta != 0;
    ChangedLengthProof {
        changed_stream_count: changed_streams.len(),
        changed_streams,
        changed_bytes_before,
        changed_bytes_after,
        output_file_bytes_delta,
        any_stream_length_changed,
        output_file_length_changed,
        logical_stream_length_change_proven: any_stream_length_changed,
    }
}

fn directory_metadata_differences(
    expected: &Inventory,
    actual: &Inventory,
) -> Vec<DirectoryMetadataDifference> {
    let paths = expected
        .directory_entries
        .keys()
        .chain(actual.directory_entries.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let mut differences = Vec::new();
    for path in paths {
        let expected_entry = expected.directory_entries.get(&path);
        let actual_entry = actual.directory_entries.get(&path);
        match (expected_entry, actual_entry) {
            (Some(expected_entry), Some(actual_entry)) => {
                if expected_entry.entry_type != actual_entry.entry_type {
                    differences.push(DirectoryMetadataDifference {
                        path: path.clone(),
                        field: "entry_type".into(),
                        expected: expected_entry.entry_type.to_string(),
                        actual: actual_entry.entry_type.to_string(),
                    });
                }
                if expected_entry.clsid != actual_entry.clsid {
                    differences.push(DirectoryMetadataDifference {
                        path: path.clone(),
                        field: "clsid".into(),
                        expected: expected_entry.clsid.clone(),
                        actual: actual_entry.clsid.clone(),
                    });
                }
                if expected_entry.bytes != actual_entry.bytes {
                    differences.push(DirectoryMetadataDifference {
                        path: path.clone(),
                        field: "bytes".into(),
                        expected: expected_entry.bytes.to_string(),
                        actual: actual_entry.bytes.to_string(),
                    });
                }
                if expected_entry.start_sector != actual_entry.start_sector {
                    differences.push(DirectoryMetadataDifference {
                        path: path.clone(),
                        field: "start_sector".into(),
                        expected: expected_entry.start_sector.to_string(),
                        actual: actual_entry.start_sector.to_string(),
                    });
                }
                if expected_entry.is_minifat != actual_entry.is_minifat {
                    differences.push(DirectoryMetadataDifference {
                        path: path.clone(),
                        field: "is_minifat".into(),
                        expected: expected_entry.is_minifat.to_string(),
                        actual: actual_entry.is_minifat.to_string(),
                    });
                }
            },
            (Some(expected_entry), None) => differences.push(DirectoryMetadataDifference {
                path: path.clone(),
                field: "directory_entry".into(),
                expected: format!("present type {}", expected_entry.entry_type),
                actual: "missing".into(),
            }),
            (None, Some(actual_entry)) => differences.push(DirectoryMetadataDifference {
                path: path.clone(),
                field: "directory_entry".into(),
                expected: "missing".into(),
                actual: format!("present type {}", actual_entry.entry_type),
            }),
            (None, None) => {},
        }
    }
    differences
}

fn semantic_directory_metadata_matches(expected: &Inventory, actual: &Inventory) -> bool {
    let expected_paths = expected.directory_entries.keys().collect::<BTreeSet<_>>();
    let actual_paths = actual.directory_entries.keys().collect::<BTreeSet<_>>();
    expected_paths == actual_paths
        && expected.directory_entries.iter().all(|(path, entry)| {
            actual
                .directory_entries
                .get(path)
                .is_some_and(|actual_entry| {
                    entry.entry_type == actual_entry.entry_type
                        && entry.clsid == actual_entry.clsid
                        && entry.bytes == actual_entry.bytes
                        && entry.is_minifat == actual_entry.is_minifat
                })
        })
}

fn exact_output_streams(expected: &Inventory, actual: &Inventory) -> (bool, bool) {
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    let actual_paths = actual.stream_bytes.keys().collect::<BTreeSet<_>>();
    let paths_match = expected_paths == actual_paths;
    let bytes_match = paths_match
        && expected
            .stream_bytes
            .iter()
            .all(|(path, bytes)| actual.stream_bytes.get(path) == Some(bytes));
    (paths_match, bytes_match)
}

fn preserved_clsids(source: &Inventory, expected: &Inventory, actual: &Inventory) -> (bool, bool) {
    (
        source.summary.root_clsid == expected.summary.root_clsid
            && actual.summary.root_clsid == source.summary.root_clsid,
        source.storage_clsids == expected.storage_clsids
            && actual.storage_clsids == source.storage_clsids,
    )
}

fn doc_text_projection_matches(source: &[String], output: &[String], replacement: &str) -> bool {
    source.len() == output.len()
        && source.iter().enumerate().all(|(index, before)| {
            let expected = if index == 0 {
                replacement
            } else {
                before.as_str()
            };
            output.get(index).is_some_and(|after| after == expected)
        })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct PptSlideIdentity {
    slide_id: u32,
    persist_id: u32,
    flags: u32,
    text_count: u32,
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, String> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| "PPT u32 offset overflow".to_string())?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| "PPT record is truncated".to_string())?;
    Ok(u32::from_le_bytes(
        bytes
            .try_into()
            .map_err(|_| "PPT u32 width mismatch".to_string())?,
    ))
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, String> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| "PPT u16 offset overflow".to_string())?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| "PPT record is truncated".to_string())?;
    Ok(u16::from_le_bytes(
        bytes
            .try_into()
            .map_err(|_| "PPT u16 width mismatch".to_string())?,
    ))
}

fn record_slice(data: &[u8], offset: usize) -> Result<&[u8], String> {
    let length_offset = offset
        .checked_add(4)
        .ok_or_else(|| "PPT record length offset overflow".to_string())?;
    let length = usize::try_from(read_u32(data, length_offset)?)
        .map_err(|_| "PPT record length does not fit usize".to_string())?;
    let end = offset
        .checked_add(8)
        .and_then(|value| value.checked_add(length))
        .ok_or_else(|| "PPT record end overflow".to_string())?;
    data.get(offset..end)
        .ok_or_else(|| "PPT record is truncated".to_string())
}

fn named_ppt_stream(ole: &mut OleFile<Cursor<Vec<u8>>>, name: &str) -> Result<Vec<u8>, String> {
    let path = ole
        .list_streams()
        .into_iter()
        .find(|path| path.last().is_some_and(|component| component == name))
        .ok_or_else(|| format!("PPT stream {name:?} was not found"))?;
    let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
    ole.open_stream(&refs).map_err(|error| error.to_string())
}

fn merge_persist_directory(
    mapping: &mut BTreeMap<u32, u32>,
    directory: &[u8],
) -> Result<(), String> {
    let record_type = read_u16(directory, 2)?;
    if record_type != 6001 && record_type != 6002 {
        return Err("PPT persist directory has an unexpected record type".into());
    }
    let mut offset = 8usize;
    while offset < directory.len() {
        let info = read_u32(directory, offset)?;
        offset += 4;
        let base = info & 0x000F_FFFF;
        let count = info >> 20;
        if count == 0 || count > 4_096 {
            return Err("PPT persist directory run exceeds the probe bound".into());
        }
        for index in 0..count {
            let value = read_u32(directory, offset)?;
            offset += 4;
            let persist_id = base
                .checked_add(index)
                .ok_or_else(|| "PPT persist identifier overflow".to_string())?;
            mapping.entry(persist_id).or_insert(value);
        }
    }
    Ok(())
}

fn ppt_slide_projection(bytes: &[u8]) -> Result<Vec<PptSlideIdentity>, String> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec())).map_err(|error| error.to_string())?;
    let document = named_ppt_stream(&mut ole, "PowerPoint Document")?;
    let current_user = named_ppt_stream(&mut ole, "Current User")?;
    let mut edit_offset = usize::try_from(read_u32(&current_user, 16)?)
        .map_err(|_| "PPT current edit offset does not fit usize".to_string())?;
    let mut mapping = BTreeMap::new();
    let mut seen = BTreeSet::new();
    let mut document_id = 0u32;
    while edit_offset != 0 {
        if !seen.insert(edit_offset) || seen.len() > 4_096 {
            return Err("PPT UserEdit chain is cyclic or excessive".into());
        }
        let record = record_slice(&document, edit_offset)?;
        if read_u16(record, 2)? != 4085 || record.len() < 36 {
            return Err("PPT UserEdit record is invalid".into());
        }
        let data = &record[8..];
        if document_id == 0 {
            document_id = read_u32(data, 16)?;
        }
        let directory_offset = usize::try_from(read_u32(data, 12)?)
            .map_err(|_| "PPT persist directory offset does not fit usize".to_string())?;
        let directory = record_slice(&document, directory_offset)?;
        merge_persist_directory(&mut mapping, directory)?;
        edit_offset = usize::try_from(read_u32(data, 8)?)
            .map_err(|_| "PPT previous edit offset does not fit usize".to_string())?;
    }
    if document_id == 0 {
        return Err("PPT UserEdit chain has no document persist ID".into());
    }
    let live_offset = mapping
        .get(&document_id)
        .copied()
        .ok_or_else(|| "PPT live document persist mapping is missing".to_string())?;
    let live = record_slice(
        &document,
        usize::try_from(live_offset)
            .map_err(|_| "PPT live document offset does not fit usize".to_string())?,
    )?;
    let snapshot =
        litchi_ppt::document_structure::Snapshot::parse(live).map_err(|error| error.to_string())?;
    Ok(snapshot
        .slides()
        .iter()
        .map(|slide| PptSlideIdentity {
            slide_id: slide.slide_id(),
            persist_id: slide.persist_id(),
            flags: slide.flags(),
            text_count: slide.text_count(),
        })
        .collect())
}

fn ppt_slide_projection_matches(source: &[PptSlideIdentity], output: &[PptSlideIdentity]) -> bool {
    source.len() > 1
        && output.len() == source.len() - 1
        && output
            == source
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != 1)
                .map(|(_, slide)| *slide)
                .collect::<Vec<_>>()
}

fn replacement_digest(replacements: &[Replacement]) -> String {
    let mut hasher = Sha256::new();
    for replacement in replacements {
        hasher.update((replacement.path.len() as u64).to_le_bytes());
        for component in &replacement.path {
            update_len_prefixed(&mut hasher, component.as_bytes());
        }
        update_len_prefixed(&mut hasher, replacement.data.as_ref());
    }
    let digest = hasher.finalize();
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

fn allowed_changed_stream(case: Case, path: &[String]) -> bool {
    let Some(name) = path.last() else {
        return false;
    };
    match case {
        Case::DocFloat | Case::DocNoHf => {
            name.eq_ignore_ascii_case("WordDocument")
                || name.eq_ignore_ascii_case("0Table")
                || name.eq_ignore_ascii_case("1Table")
        },
        Case::Ppt45543 => {
            name.eq_ignore_ascii_case("PowerPoint Document")
                || name.eq_ignore_ascii_case("Current User")
        },
    }
}

fn changed_stream_paths_allowed(case: Case, source: &Inventory, expected: &Inventory) -> bool {
    source
        .stream_bytes
        .keys()
        .filter(|path| source.stream_bytes.get(*path) != expected.stream_bytes.get(*path))
        .all(|path| allowed_changed_stream(case, path))
}

fn unchanged_stream_bytes_match(source: &Inventory, expected: &Inventory) -> bool {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    if source_paths != expected_paths {
        return false;
    }
    let mut changed_paths = BTreeSet::new();
    for (path, before) in &source.stream_bytes {
        if expected.stream_bytes.get(path) != Some(before) {
            changed_paths.insert(path.clone());
        }
    }
    for (path, before) in &source.stream_bytes {
        if !changed_paths.contains(path) && expected.stream_bytes.get(path) != Some(before) {
            return false;
        }
    }
    true
}

fn derive_replacements(
    source: &Inventory,
    expected: &Inventory,
    case: Case,
) -> Result<Vec<Replacement>, BoxError> {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    if source_paths != expected_paths {
        return failure(
            "public edit added or deleted CFB streams; no explicit stream mapping is available",
        );
    }
    if !changed_stream_paths_allowed(case, source, expected) {
        return failure(
            "public edit changed an unapproved stream; explicit stream mapping is required",
        );
    }
    let mut replacements = Vec::new();
    for path in source.stream_bytes.keys() {
        let before = source
            .stream_bytes
            .get(path)
            .ok_or_else(|| Box::new(ProbeError("source stream disappeared".into())) as BoxError)?;
        let after = expected.stream_bytes.get(path).ok_or_else(|| {
            Box::new(ProbeError("expected stream disappeared".into())) as BoxError
        })?;
        if before != after {
            replacements.push(Replacement {
                path: path.clone(),
                data: Arc::from(after.clone().into_boxed_slice()),
            });
        }
    }
    if replacements.is_empty() {
        return failure("public edit produced no changed stream replacement");
    }
    Ok(replacements)
}

fn replacement_summaries(
    source: &Inventory,
    expected: &Inventory,
    replacements: &[Replacement],
) -> Vec<ReplacementSummary> {
    replacements
        .iter()
        .map(|replacement| {
            let before = source
                .stream_bytes
                .get(&replacement.path)
                .map_or(&[][..], Vec::as_slice);
            let after = expected
                .stream_bytes
                .get(&replacement.path)
                .map_or(&[][..], Vec::as_slice);
            ReplacementSummary {
                path: replacement.path.clone(),
                before_bytes: before.len(),
                after_bytes: after.len(),
                before_sha256: sha256_hex(before),
                after_sha256: sha256_hex(after),
            }
        })
        .collect()
}

#[derive(Clone, Debug, Serialize)]
struct Oracle {
    source_structurally_valid: bool,
    expected_structurally_valid: bool,
    output_structurally_valid: bool,
    semantic_reopen_ok: bool,
    semantic_directory_metadata_matches_expected: bool,
    directory_metadata_differences: Vec<DirectoryMetadataDifference>,
    source_stream_paths_match_expected: bool,
    output_stream_paths_match_expected: bool,
    output_stream_bytes_match_expected: bool,
    unchanged_stream_bytes_match_expected: bool,
    changed_stream_paths_allowed: bool,
    root_clsid_preserved: bool,
    storage_clsids_preserved: bool,
    no_deleted_or_new_streams: bool,
    oracle_ok: bool,
    failure_reasons: Vec<String>,
}

fn semantic_reopen_check(
    case: Case,
    source_bytes: &[u8],
    output_bytes: &[u8],
    text: &str,
) -> (bool, String) {
    match case {
        Case::DocFloat | Case::DocNoHf => {
            let limits = litchi_doc::tracked_revision::Limits::default();
            let source = match litchi_doc::body_text::Snapshot::open(source_bytes.to_vec(), limits)
            {
                Ok(snapshot) => snapshot,
                Err(error) => return (false, format!("DOC source public reopen failed: {error}")),
            };
            let output = match litchi_doc::body_text::Snapshot::open(output_bytes.to_vec(), limits)
            {
                Ok(snapshot) => snapshot,
                Err(error) => return (false, format!("DOC output public reopen failed: {error}")),
            };
            let source_paragraphs = match source.paragraphs(litchi_doc::body_text::Projection::All)
            {
                Ok(paragraphs) => paragraphs,
                Err(error) => return (false, format!("DOC source paragraph read failed: {error}")),
            };
            let output_paragraphs = match output.paragraphs(litchi_doc::body_text::Projection::All)
            {
                Ok(paragraphs) => paragraphs,
                Err(error) => return (false, format!("DOC output paragraph read failed: {error}")),
            };
            if source_paragraphs.len() != output_paragraphs.len() {
                return (false, "DOC paragraph count changed".into());
            }
            let source_text = source_paragraphs
                .iter()
                .map(|paragraph| paragraph.text().to_string())
                .collect::<Vec<_>>();
            let output_text = output_paragraphs
                .iter()
                .map(|paragraph| paragraph.text().to_string())
                .collect::<Vec<_>>();
            if !doc_text_projection_matches(&source_text, &output_text, text) {
                return (
                    false,
                    "DOC paragraph projection differs from the public edit oracle".into(),
                );
            }
            (true, String::new())
        },
        Case::Ppt45543 => {
            let source = match litchi_ppt::slide_order::Snapshot::from_bytes(source_bytes.to_vec())
            {
                Ok(snapshot) => snapshot,
                Err(error) => return (false, format!("PPT source public reopen failed: {error}")),
            };
            let output = match litchi_ppt::slide_order::Snapshot::from_bytes(output_bytes.to_vec())
            {
                Ok(snapshot) => snapshot,
                Err(error) => return (false, format!("PPT output public reopen failed: {error}")),
            };
            let expected_count = source.slide_count().saturating_sub(1);
            if output.slide_count() != expected_count {
                return (
                    false,
                    format!(
                        "PPT slide count {} does not equal source count minus one {expected_count}",
                        output.slide_count()
                    ),
                );
            }
            let source_projection = match ppt_slide_projection(source_bytes) {
                Ok(projection) => projection,
                Err(error) => {
                    return (
                        false,
                        format!("PPT source identity projection failed: {error}"),
                    );
                },
            };
            let output_projection = match ppt_slide_projection(output_bytes) {
                Ok(projection) => projection,
                Err(error) => {
                    return (
                        false,
                        format!("PPT output identity projection failed: {error}"),
                    );
                },
            };
            if !ppt_slide_projection_matches(&source_projection, &output_projection) {
                return (
                    false,
                    "PPT output slide identities are not source order with slide 1 removed".into(),
                );
            }
            (true, String::new())
        },
    }
}

fn oracle_for_output(
    args: &Args,
    source: &Inventory,
    expected: &Inventory,
    output: &Inventory,
    source_bytes: &[u8],
    expected_bytes: &[u8],
    output_bytes: &[u8],
) -> Oracle {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    let source_stream_paths_match_expected = source_paths == expected_paths;
    let (output_stream_paths_match_expected, output_stream_bytes_match_expected) =
        exact_output_streams(expected, output);
    let unchanged_stream_bytes_match_expected = unchanged_stream_bytes_match(source, expected);
    let changed_stream_paths_allowed = changed_stream_paths_allowed(args.case, source, expected);
    let (root_clsid_preserved, storage_clsids_preserved) =
        preserved_clsids(source, expected, output);
    let (semantic_reopen_ok, semantic_failure) =
        semantic_reopen_check(args.case, source_bytes, output_bytes, &args.text);
    let directory_metadata_differences = directory_metadata_differences(expected, output);
    let semantic_directory_metadata_matches_expected =
        semantic_directory_metadata_matches(expected, output);
    let source_structurally_valid = structural_valid(source_bytes);
    let expected_structurally_valid = structural_valid(expected_bytes);
    let output_structurally_valid = structural_valid(output_bytes);
    let no_deleted_or_new_streams =
        source_stream_paths_match_expected && output_stream_paths_match_expected;
    let mut failure_reasons = Vec::new();
    if !source_structurally_valid {
        failure_reasons.push("source CFB structural validation failed".into());
    }
    if !expected_structurally_valid {
        failure_reasons.push("expected format output CFB structural validation failed".into());
    }
    if !output_structurally_valid {
        failure_reasons.push("output CFB structural validation failed".into());
    }
    if !source_stream_paths_match_expected {
        failure_reasons.push("public edit changed the stream path set".into());
    }
    if !output_stream_paths_match_expected {
        failure_reasons.push("output stream path set differs from expected format output".into());
    }
    if !output_stream_bytes_match_expected {
        failure_reasons.push("one or more output stream bytes differ from expected".into());
    }
    if !unchanged_stream_bytes_match_expected {
        failure_reasons
            .push("one or more source streams changed outside the reported replacement set".into());
    }
    if !changed_stream_paths_allowed {
        failure_reasons
            .push("a changed stream is outside the allowed public-edit stream set".into());
    }
    if !semantic_reopen_ok {
        failure_reasons.push(semantic_failure);
    }
    if !semantic_directory_metadata_matches_expected {
        failure_reasons.push("semantic CFB directory metadata differs from expected".into());
    }
    if !root_clsid_preserved {
        failure_reasons.push("root CLSID was not preserved from the source".into());
    }
    if !storage_clsids_preserved {
        failure_reasons.push("storage CLSIDs were not preserved from the source".into());
    }
    let oracle_ok = source_structurally_valid
        && expected_structurally_valid
        && output_structurally_valid
        && no_deleted_or_new_streams
        && output_stream_bytes_match_expected
        && unchanged_stream_bytes_match_expected
        && changed_stream_paths_allowed
        && root_clsid_preserved
        && storage_clsids_preserved
        && semantic_reopen_ok
        && semantic_directory_metadata_matches_expected;
    Oracle {
        source_structurally_valid,
        expected_structurally_valid,
        output_structurally_valid,
        semantic_reopen_ok,
        semantic_directory_metadata_matches_expected,
        directory_metadata_differences,
        source_stream_paths_match_expected,
        output_stream_paths_match_expected,
        output_stream_bytes_match_expected,
        unchanged_stream_bytes_match_expected,
        changed_stream_paths_allowed,
        root_clsid_preserved,
        storage_clsids_preserved,
        no_deleted_or_new_streams,
        oracle_ok,
        failure_reasons,
    }
}

fn public_format_edit(source: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    match args.case {
        Case::DocFloat | Case::DocNoHf => {
            let snapshot = litchi_doc::body_text::Snapshot::open(
                source.to_vec(),
                litchi_doc::tracked_revision::Limits::default(),
            )?;
            let mut edit = snapshot.edit()?;
            edit.replace_paragraph(Position::new(0), &args.text)?;
            let commit = edit.commit()?;
            Ok(commit.snapshot().bytes().to_vec())
        },
        Case::Ppt45543 => {
            let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
            if snapshot.slide_count() <= 1 {
                return failure("ppt45543 requires at least two slides to remove slide 1");
            }
            let mut edit = snapshot.edit()?;
            edit.remove_slide(Position::new(1))?;
            let commit = edit.commit()?;
            Ok(commit.snapshot().bytes().to_vec())
        },
    }
}

struct TimedFormat {
    output: Vec<u8>,
    whole_ns: u128,
}

struct TimedContainer {
    output: Vec<u8>,
    open_ns: u128,
    stage_ns: u128,
    finish_ns: u128,
    whole_ns: u128,
}

fn timed_format(source: &[u8], args: &Args) -> Result<TimedFormat, BoxError> {
    let whole_start = Instant::now();
    let output = public_format_edit(source, args)?;
    let whole_ns = whole_start.elapsed().as_nanos();
    black_box(output.len());
    Ok(TimedFormat { output, whole_ns })
}

fn timed_container(
    source: &[u8],
    replacements: &[Replacement],
    policy: Policy,
) -> Result<TimedContainer, BoxError> {
    let whole_start = Instant::now();
    let open_start = Instant::now();
    let mut editor = Editor::open(source.to_vec(), Targets::default(), Limits::default())?;
    let open_ns = open_start.elapsed().as_nanos();

    let stage_start = Instant::now();
    editor.set_sector_layout_policy(policy.cfb());
    editor.put_streams_shared(
        replacements
            .iter()
            .map(|replacement| (replacement.path.as_slice(), Arc::clone(&replacement.data))),
    )?;
    let stage_ns = stage_start.elapsed().as_nanos();

    let finish_start = Instant::now();
    let output = editor.finish()?;
    let finish_ns = finish_start.elapsed().as_nanos();
    let whole_ns = whole_start.elapsed().as_nanos();
    black_box(output.len());
    Ok(TimedContainer {
        output,
        open_ns,
        stage_ns,
        finish_ns,
        whole_ns,
    })
}

#[derive(Clone, Debug, Serialize)]
struct PhaseTimes {
    #[serde(skip_serializing_if = "Option::is_none")]
    open_ns: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    stage_ns: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    finish_ns: Option<u128>,
    whole_ns: u128,
}

impl PhaseTimes {
    fn format(whole_ns: u128) -> Self {
        Self {
            open_ns: None,
            stage_ns: None,
            finish_ns: None,
            whole_ns,
        }
    }

    fn container(result: &TimedContainer) -> Self {
        Self {
            open_ns: Some(result.open_ns),
            stage_ns: Some(result.stage_ns),
            finish_ns: Some(result.finish_ns),
            whole_ns: result.whole_ns,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct AllocationRegion {
    allocated_bytes: u64,
    deallocated_bytes: u64,
    allocation_calls: u64,
    peak_live_bytes: u64,
    retained_bytes: i64,
}

impl From<alloc_metrics::Region> for AllocationRegion {
    fn from(region: alloc_metrics::Region) -> Self {
        Self {
            allocated_bytes: region.allocated_bytes,
            deallocated_bytes: region.deallocated_bytes,
            allocation_calls: region.allocation_calls,
            peak_live_bytes: region.peak_live_bytes,
            retained_bytes: region.retained_bytes,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct AllocationSample {
    #[serde(skip_serializing_if = "Option::is_none")]
    whole: Option<AllocationRegion>,
    #[serde(skip_serializing_if = "Option::is_none")]
    open: Option<AllocationRegion>,
    #[serde(skip_serializing_if = "Option::is_none")]
    stage: Option<AllocationRegion>,
    #[serde(skip_serializing_if = "Option::is_none")]
    finish: Option<AllocationRegion>,
}

#[derive(Clone, Debug, Serialize)]
struct Sample {
    index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    phase_ns: Option<PhaseTimes>,
    output_sha256: String,
    output_inventory: InventorySummary,
    oracle: Oracle,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocations: Option<AllocationSample>,
}

#[derive(Debug, Serialize)]
struct ProbeOutput {
    schema_version: u32,
    case: String,
    format: String,
    operation: String,
    scope: String,
    input: String,
    policy: String,
    policy_applied: bool,
    policy_application_scope: String,
    policy_contract: String,
    timing_claim: bool,
    allocator_instrumented: bool,
    directory_metadata_fields: Vec<String>,
    warmups: usize,
    samples_requested: usize,
    source_sha256: String,
    expected_output_sha256: String,
    replacements_sha256: String,
    source_inventory: InventorySummary,
    expected_output_inventory: InventorySummary,
    replacements: Vec<ReplacementSummary>,
    changed_length_proof: ChangedLengthProof,
    expected_oracle: Oracle,
    samples: Vec<Sample>,
}

fn allocation_format(
    source: &[u8],
    args: &Args,
) -> Result<(Vec<u8>, alloc_metrics::Region), BoxError> {
    let mut output = None;
    let region = alloc_metrics::region(|| {
        output = Some(public_format_edit(source, args)?);
        Ok::<(), BoxError>(())
    })?;
    let output = output
        .ok_or_else(|| Box::new(ProbeError("format produced no output".into())) as BoxError)?;
    Ok((output, region))
}

fn allocation_container(
    source: &[u8],
    replacements: &[Replacement],
    policy: Policy,
) -> Result<(Vec<u8>, AllocationSample), BoxError> {
    let mut editor = None;
    let open = alloc_metrics::region(|| {
        editor = Some(Editor::open(
            source.to_vec(),
            Targets::default(),
            Limits::default(),
        )?);
        Ok::<(), litchi_cfb::OleError>(())
    })?;
    let stage = alloc_metrics::region(|| {
        let current = editor
            .as_mut()
            .ok_or_else(|| litchi_cfb::OleError::InvalidFormat("missing opened editor".into()))?;
        current.set_sector_layout_policy(policy.cfb());
        current.put_streams_shared(
            replacements
                .iter()
                .map(|replacement| (replacement.path.as_slice(), Arc::clone(&replacement.data))),
        )?;
        Ok::<(), litchi_cfb::OleError>(())
    })?;
    let mut output = None;
    let finish = alloc_metrics::region(|| {
        let current = editor
            .take()
            .ok_or_else(|| litchi_cfb::OleError::InvalidFormat("missing staged editor".into()))?;
        output = Some(current.finish()?);
        Ok::<(), litchi_cfb::OleError>(())
    })?;
    let output = output
        .ok_or_else(|| Box::new(ProbeError("container produced no output".into())) as BoxError)?;
    Ok((
        output,
        AllocationSample {
            whole: None,
            open: Some(open.into()),
            stage: Some(stage.into()),
            finish: Some(finish.into()),
        },
    ))
}

fn output_sample(
    args: &Args,
    index: usize,
    output: Vec<u8>,
    phase_ns: Option<PhaseTimes>,
    allocations: Option<AllocationSample>,
    source: &Inventory,
    expected: &Inventory,
    source_bytes: &[u8],
    expected_bytes: &[u8],
    expected_length_proven: bool,
) -> Result<Sample, BoxError> {
    let output_sha256 = sha256_hex(&output);
    let output_inventory = inventory(&output)?;
    let mut oracle = oracle_for_output(
        args,
        source,
        expected,
        &output_inventory,
        source_bytes,
        expected_bytes,
        &output,
    );
    if !expected_length_proven {
        oracle
            .failure_reasons
            .push("public edit did not prove a length-changing stream or file".into());
        oracle.oracle_ok = false;
    }
    Ok(Sample {
        index,
        phase_ns,
        output_sha256,
        output_inventory: output_inventory.summary,
        oracle,
        allocations,
    })
}

pub fn run(timing_claim: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    let source_bytes = std::fs::read(&args.input)?;
    let source = inventory(&source_bytes)?;

    // The expected stream bytes are produced by one real public format edit,
    // outside every timed or allocator-counted interval.  This also makes the
    // common-container operation a direct exact-replacement control.
    let expected_bytes = public_format_edit(&source_bytes, &args)?;
    let expected = inventory(&expected_bytes)?;
    let replacements = derive_replacements(&source, &expected, args.case)?;
    let length_proof = changed_length_proof(&source, &expected);
    let replacements_json = replacement_summaries(&source, &expected, &replacements);
    let expected_oracle = oracle_for_output(
        &args,
        &source,
        &expected,
        &expected,
        &source_bytes,
        &expected_bytes,
        &expected_bytes,
    );
    if !expected_oracle.oracle_ok || !length_proof.logical_stream_length_change_proven {
        return failure(format!(
            "public format oracle failed before measurement: {:?}; length_proof={:?}",
            expected_oracle.failure_reasons, length_proof
        ));
    }

    let policy_applied = args.operation == Operation::Container;
    let operation_name = args.operation.name().to_string();
    let mut samples = Vec::with_capacity(args.samples);

    for _ in 0..args.warmups {
        if timing_claim {
            match args.operation {
                Operation::Format => {
                    let result = timed_format(&source_bytes, &args)?;
                    drop(result.output);
                },
                Operation::Container => {
                    let result = timed_container(&source_bytes, &replacements, args.policy)?;
                    drop(result.output);
                },
            }
        } else {
            match args.operation {
                Operation::Format => {
                    let (output, _) = allocation_format(&source_bytes, &args)?;
                    drop(output);
                },
                Operation::Container => {
                    let (output, _) =
                        allocation_container(&source_bytes, &replacements, args.policy)?;
                    drop(output);
                },
            }
        }
    }

    for index in 0..args.samples {
        if timing_claim {
            match args.operation {
                Operation::Format => {
                    let result = timed_format(&source_bytes, &args)?;
                    let phase_ns = PhaseTimes::format(result.whole_ns);
                    samples.push(output_sample(
                        &args,
                        index,
                        result.output,
                        Some(phase_ns),
                        None,
                        &source,
                        &expected,
                        &source_bytes,
                        &expected_bytes,
                        length_proof.logical_stream_length_change_proven,
                    )?);
                },
                Operation::Container => {
                    let result = timed_container(&source_bytes, &replacements, args.policy)?;
                    let phase_ns = PhaseTimes::container(&result);
                    samples.push(output_sample(
                        &args,
                        index,
                        result.output,
                        Some(phase_ns),
                        None,
                        &source,
                        &expected,
                        &source_bytes,
                        &expected_bytes,
                        length_proof.logical_stream_length_change_proven,
                    )?);
                },
            }
        } else {
            match args.operation {
                Operation::Format => {
                    let (output, allocation) = allocation_format(&source_bytes, &args)?;
                    // Allocation regions intentionally omit format output
                    // validation; it occurs after the measured region.
                    samples.push(output_sample(
                        &args,
                        index,
                        output,
                        None,
                        Some(AllocationSample {
                            whole: Some(allocation.into()),
                            open: None,
                            stage: None,
                            finish: None,
                        }),
                        &source,
                        &expected,
                        &source_bytes,
                        &expected_bytes,
                        length_proof.logical_stream_length_change_proven,
                    )?);
                },
                Operation::Container => {
                    let (output, allocations) =
                        allocation_container(&source_bytes, &replacements, args.policy)?;
                    samples.push(output_sample(
                        &args,
                        index,
                        output,
                        None,
                        Some(allocations),
                        &source,
                        &expected,
                        &source_bytes,
                        &expected_bytes,
                        length_proof.logical_stream_length_change_proven,
                    )?);
                },
            }
        }
    }

    let result = ProbeOutput {
        schema_version: 1,
        case: args.case.name().into(),
        format: args.case.format_name().into(),
        operation: operation_name.clone(),
        scope: if args.operation == Operation::Format {
            "public_format_open_edit_commit".into()
        } else {
            "common_container_open_stage_finish_control".into()
        },
        input: args.input.display().to_string(),
        policy: args.policy.name().into(),
        policy_applied,
        policy_application_scope: if policy_applied {
            "common_container_editor".into()
        } else {
            "not_applied_public_format_route".into()
        },
        policy_contract: "Reuse and Rewrite must preserve logical streams, semantic directory metadata, root/storage CLSIDs, and public edit meaning; physical sector assignments may differ and are reported in directory_metadata_differences".into(),
        timing_claim,
        allocator_instrumented: alloc_metrics::instrumented(),
        directory_metadata_fields: vec![
            "entry_type".into(),
            "clsid".into(),
            "bytes".into(),
            "start_sector".into(),
            "is_minifat".into(),
        ],
        warmups: args.warmups,
        samples_requested: args.samples,
        source_sha256: sha256_hex(&source_bytes),
        expected_output_sha256: sha256_hex(&expected_bytes),
        replacements_sha256: replacement_digest(&replacements),
        source_inventory: source.summary,
        expected_output_inventory: expected.summary,
        replacements: replacements_json,
        changed_length_proof: length_proof,
        expected_oracle,
        samples,
    };
    println!("{}", serde_json::to_string(&result)?);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fake_inventory(streams: &[(&str, &[u8])], root_clsid: &str) -> Inventory {
        let mut stream_bytes = BTreeMap::new();
        let mut summaries = Vec::new();
        for (path, bytes) in streams {
            let path = vec![(*path).to_string()];
            stream_bytes.insert(path.clone(), bytes.to_vec());
            summaries.push(StreamSummary {
                path,
                bytes: bytes.len(),
                sha256: sha256_hex(bytes),
            });
        }
        summaries.sort_by(|left, right| left.path.cmp(&right.path));
        Inventory {
            summary: InventorySummary {
                file_bytes: 0,
                sector_size: 512,
                root_clsid: root_clsid.into(),
                streams: summaries,
                storages: Vec::new(),
                directory_entries: BTreeMap::new(),
            },
            stream_bytes,
            storage_clsids: BTreeMap::new(),
            directory_entries: BTreeMap::new(),
        }
    }

    #[test]
    fn oracle_control_rejects_missing_stream() {
        let expected = fake_inventory(
            &[
                ("PowerPoint Document", b"document"),
                ("Current User", b"user"),
            ],
            "root",
        );
        let actual = fake_inventory(&[("PowerPoint Document", b"document")], "root");
        let (paths_match, bytes_match) = exact_output_streams(&expected, &actual);
        assert!(!paths_match);
        assert!(!bytes_match);
    }

    #[test]
    fn oracle_control_rejects_wrong_doc_target() {
        let source = vec!["first".to_string(), "second".to_string()];
        let wrong_target = vec!["first".to_string(), "replacement".to_string()];
        assert!(!doc_text_projection_matches(
            &source,
            &wrong_target,
            "replacement"
        ));
    }

    #[test]
    fn oracle_control_rejects_wrong_ppt_slide_identity() {
        let source = vec![
            PptSlideIdentity {
                slide_id: 10,
                persist_id: 100,
                flags: 0,
                text_count: 1,
            },
            PptSlideIdentity {
                slide_id: 20,
                persist_id: 200,
                flags: 0,
                text_count: 2,
            },
            PptSlideIdentity {
                slide_id: 30,
                persist_id: 300,
                flags: 0,
                text_count: 3,
            },
        ];
        let removed_wrong_slide = vec![source[0], source[1]];
        assert!(!ppt_slide_projection_matches(&source, &removed_wrong_slide));
    }

    #[test]
    fn oracle_control_rejects_metadata_change() {
        let source = fake_inventory(&[("WordDocument", b"before")], "root-a");
        let expected = fake_inventory(&[("WordDocument", b"after")], "root-a");
        let output = fake_inventory(&[("WordDocument", b"after")], "root-b");
        let (root_preserved, storage_preserved) = preserved_clsids(&source, &expected, &output);
        assert!(!root_preserved);
        assert!(storage_preserved);
    }
}
