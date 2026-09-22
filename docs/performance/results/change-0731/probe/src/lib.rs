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

const CFB_END_OF_CHAIN: u32 = 0xffff_fffe;
const CFB_FREE_SECTOR: u32 = 0xffff_ffff;
const CFB_HEADER_SIZE: usize = 512;
const CFB_HEADER_DIFAT_OFFSET: usize = 0x4c;
const CFB_HEADER_DIFAT_ENTRIES: usize = 109;
const CFB_DIRECTORY_ENTRY_SIZE: usize = 128;
const CFB_ENTRY_START_SECTOR_OFFSET: usize = 0x74;
const CFB_ENTRY_STREAM_SIZE_OFFSET: usize = 0x78;
const CFB_MAX_DIRECTORY_SECTORS: usize = 1_000_000;

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

#[derive(Clone)]
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
    raw_directory_image_bytes: usize,
    raw_directory_image_sha256: String,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct DirectoryEntrySummary {
    path: Vec<String>,
    sid: u32,
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
    raw_directory_image: RawDirectoryImage,
}

#[derive(Clone, Debug)]
struct RawDirectoryImage {
    bytes: Vec<u8>,
    sector_size: usize,
    cfb_version: u16,
}

fn raw_u16(bytes: &[u8], offset: usize, field: &str) -> Result<u16, BoxError> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| Box::new(ProbeError(format!("{field} offset overflow"))) as BoxError)?;
    let data = bytes.get(offset..end).ok_or_else(|| {
        Box::new(ProbeError(format!("CFB header is truncated at {field}"))) as BoxError
    })?;
    Ok(u16::from_le_bytes(data.try_into().map_err(|_| {
        Box::new(ProbeError(format!("{field} width mismatch"))) as BoxError
    })?))
}

fn raw_u32(bytes: &[u8], offset: usize, field: &str) -> Result<u32, BoxError> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| Box::new(ProbeError(format!("{field} offset overflow"))) as BoxError)?;
    let data = bytes.get(offset..end).ok_or_else(|| {
        Box::new(ProbeError(format!("CFB header is truncated at {field}"))) as BoxError
    })?;
    Ok(u32::from_le_bytes(data.try_into().map_err(|_| {
        Box::new(ProbeError(format!("{field} width mismatch"))) as BoxError
    })?))
}

fn raw_sector(bytes: &[u8], sector_size: usize, sector: u32) -> Result<&[u8], BoxError> {
    let sector = usize::try_from(sector)
        .map_err(|_| Box::new(ProbeError("CFB sector does not fit usize".into())) as BoxError)?;
    let start = sector
        .checked_add(1)
        .and_then(|value| value.checked_mul(sector_size))
        .ok_or_else(|| Box::new(ProbeError("CFB sector offset overflow".into())) as BoxError)?;
    let end = start
        .checked_add(sector_size)
        .ok_or_else(|| Box::new(ProbeError("CFB sector end overflow".into())) as BoxError)?;
    bytes.get(start..end).ok_or_else(|| {
        Box::new(ProbeError("CFB sector lies outside the input bytes".into())) as BoxError
    })
}

fn raw_fat(bytes: &[u8], sector_size: usize) -> Result<Vec<u32>, BoxError> {
    let fat_count = usize::try_from(raw_u32(bytes, 0x2c, "FAT sector count")?).map_err(|_| {
        Box::new(ProbeError("FAT sector count does not fit usize".into())) as BoxError
    })?;
    let difat_count =
        usize::try_from(raw_u32(bytes, 0x48, "DIFAT sector count")?).map_err(|_| {
            Box::new(ProbeError("DIFAT sector count does not fit usize".into())) as BoxError
        })?;
    if fat_count == 0 || fat_count > CFB_MAX_DIRECTORY_SECTORS {
        return failure("CFB FAT sector count is outside the probe bound");
    }
    let mut fat_sectors = Vec::with_capacity(fat_count.min(CFB_HEADER_DIFAT_ENTRIES));
    for index in 0..CFB_HEADER_DIFAT_ENTRIES {
        let sector = raw_u32(
            bytes,
            CFB_HEADER_DIFAT_OFFSET + index * 4,
            "header DIFAT entry",
        )?;
        if sector != CFB_FREE_SECTOR {
            fat_sectors.push(sector);
            if fat_sectors.len() == fat_count {
                break;
            }
        }
    }
    let difat_entries_per_sector = sector_size / 4 - 1;
    let mut next_difat = raw_u32(bytes, 0x44, "first DIFAT sector")?;
    let mut seen_difat = BTreeSet::new();
    while fat_sectors.len() < fat_count && next_difat != CFB_END_OF_CHAIN {
        if difat_count == 0 || !seen_difat.insert(next_difat) || seen_difat.len() > difat_count {
            return failure("CFB DIFAT chain is cyclic or exceeds its declared bound");
        }
        let sector = raw_sector(bytes, sector_size, next_difat)?;
        for index in 0..difat_entries_per_sector {
            let entry =
                u32::from_le_bytes(sector[index * 4..index * 4 + 4].try_into().map_err(|_| {
                    Box::new(ProbeError("DIFAT entry width mismatch".into())) as BoxError
                })?);
            if entry != CFB_FREE_SECTOR {
                fat_sectors.push(entry);
                if fat_sectors.len() == fat_count {
                    break;
                }
            }
        }
        next_difat = u32::from_le_bytes(
            sector[difat_entries_per_sector * 4..difat_entries_per_sector * 4 + 4]
                .try_into()
                .map_err(|_| {
                    Box::new(ProbeError("DIFAT next pointer width mismatch".into())) as BoxError
                })?,
        );
    }
    if fat_sectors.len() != fat_count {
        return failure("CFB DIFAT does not enumerate every declared FAT sector");
    }
    let mut fat = Vec::new();
    fat.try_reserve(fat_count.checked_mul(sector_size / 4).ok_or_else(|| {
        Box::new(ProbeError("CFB FAT allocation size overflow".into())) as BoxError
    })?)
    .map_err(|_| {
        Box::new(ProbeError(
            "CFB FAT exceeds the probe allocation bound".into(),
        )) as BoxError
    })?;
    for sector in fat_sectors {
        for word in raw_sector(bytes, sector_size, sector)?.chunks_exact(4) {
            fat.push(u32::from_le_bytes(word.try_into().map_err(|_| {
                Box::new(ProbeError("FAT word width mismatch".into())) as BoxError
            })?));
        }
    }
    Ok(fat)
}

fn raw_directory_image(bytes: &[u8]) -> Result<RawDirectoryImage, BoxError> {
    if bytes.len() < CFB_HEADER_SIZE {
        return failure("CFB header is shorter than one sector");
    }
    let version = raw_u16(bytes, 0x1a, "CFB major version")?;
    let sector_shift = raw_u16(bytes, 0x1e, "CFB sector shift")?;
    if version != 3 && version != 4 {
        return failure("CFB directory image has an unsupported version");
    }
    let sector_size = match sector_shift {
        9 => 512,
        12 => 4096,
        _ => return failure("CFB directory image has an unsupported sector size"),
    };
    let fat = raw_fat(bytes, sector_size)?;
    let mut sector = raw_u32(bytes, 0x30, "first directory sector")?;
    let mut seen = BTreeSet::new();
    let mut image = Vec::new();
    while sector != CFB_END_OF_CHAIN {
        if !seen.insert(sector) || seen.len() > CFB_MAX_DIRECTORY_SECTORS {
            return failure("CFB directory sector chain is cyclic or exceeds the probe bound");
        }
        image.extend_from_slice(raw_sector(bytes, sector_size, sector)?);
        let index = usize::try_from(sector).map_err(|_| {
            Box::new(ProbeError("CFB directory sector does not fit usize".into())) as BoxError
        })?;
        sector = *fat.get(index).ok_or_else(|| {
            Box::new(ProbeError("CFB directory sector is outside FAT".into())) as BoxError
        })?;
    }
    if image.is_empty() || image.len() % CFB_DIRECTORY_ENTRY_SIZE != 0 {
        return failure("CFB directory image is not entry aligned");
    }
    Ok(RawDirectoryImage {
        bytes: image,
        sector_size,
        cfb_version: version,
    })
}

fn raw_directory_entry_file_offset(bytes: &[u8], sid: u32) -> Result<usize, BoxError> {
    let sector_shift = raw_u16(bytes, 0x1e, "CFB sector shift")?;
    let sector_size = match sector_shift {
        9 => 512,
        12 => 4096,
        _ => return failure("CFB directory mutation has an unsupported sector size"),
    };
    let fat = raw_fat(bytes, sector_size)?;
    let target = usize::try_from(sid).map_err(|_| {
        Box::new(ProbeError("CFB directory SID does not fit usize".into())) as BoxError
    })?;
    let entry_byte = target
        .checked_mul(CFB_DIRECTORY_ENTRY_SIZE)
        .ok_or_else(|| {
            Box::new(ProbeError("CFB directory SID offset overflow".into())) as BoxError
        })?;
    let directory_sector_index = entry_byte / sector_size;
    let intra_sector = entry_byte % sector_size;
    let mut sector = raw_u32(bytes, 0x30, "first directory sector")?;
    let mut seen = BTreeSet::new();
    for _ in 0..=directory_sector_index {
        if sector == CFB_END_OF_CHAIN || !seen.insert(sector) {
            return failure("CFB directory mutation cannot reach the requested SID");
        }
        if seen.len() == directory_sector_index + 1 {
            let sector = usize::try_from(sector).map_err(|_| {
                Box::new(ProbeError("CFB directory sector does not fit usize".into())) as BoxError
            })?;
            return sector
                .checked_add(1)
                .and_then(|value| value.checked_mul(sector_size))
                .and_then(|value| value.checked_add(intra_sector))
                .ok_or_else(|| {
                    Box::new(ProbeError("CFB directory byte offset overflow".into())) as BoxError
                });
        }
        let index = usize::try_from(sector).map_err(|_| {
            Box::new(ProbeError("CFB directory sector does not fit usize".into())) as BoxError
        })?;
        sector = *fat.get(index).ok_or_else(|| {
            Box::new(ProbeError("CFB directory sector is outside FAT".into())) as BoxError
        })?;
    }
    failure("CFB directory mutation exceeded its bounded sector walk")
}

fn mutate_directory_byte(
    source: &[u8],
    inventory: &Inventory,
    path: &[String],
    field_offset: usize,
) -> Result<Vec<u8>, BoxError> {
    let entry = inventory.directory_entries.get(path).ok_or_else(|| {
        Box::new(ProbeError("directory mutation target is absent".into())) as BoxError
    })?;
    let offset = raw_directory_entry_file_offset(source, entry.sid)?
        .checked_add(field_offset)
        .ok_or_else(|| {
            Box::new(ProbeError("directory mutation offset overflow".into())) as BoxError
        })?;
    let mut output = source.to_vec();
    let byte = output.get_mut(offset).ok_or_else(|| {
        Box::new(ProbeError("directory mutation lies outside source".into())) as BoxError
    })?;
    *byte ^= 1;
    Ok(output)
}

fn collect_directory_entry(
    ole: &OleFile<Cursor<Vec<u8>>>,
    path: &[String],
    entry: &DirectoryEntry,
    output: &mut BTreeMap<Vec<String>, DirectoryEntrySummary>,
) -> Result<(), BoxError> {
    output.insert(
        path.to_vec(),
        DirectoryEntrySummary {
            path: path.to_vec(),
            sid: entry.sid,
            entry_type: entry.entry_type,
            clsid: entry.clsid.clone(),
            bytes: usize::try_from(entry.size).unwrap_or(usize::MAX),
            start_sector: entry.start_sector,
            is_minifat: entry.is_minifat,
        },
    );
    if entry.entry_type == litchi_cfb::consts::STGTY_STORAGE
        || entry.entry_type == litchi_cfb::consts::STGTY_ROOT
    {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let children = ole.list_directory_entries(&refs)?;
        for child in children {
            let mut child_path = path.to_vec();
            child_path.push(child.name.clone());
            collect_directory_entry(ole, &child_path, child, output)?;
        }
    }
    Ok(())
}

fn inventory(bytes: &[u8]) -> Result<Inventory, BoxError> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec()))?;
    let sector_size = ole.sector_size();
    let raw_directory = raw_directory_image(bytes)?;
    let root = ole.root_entry().cloned().ok_or_else(|| {
        Box::new(ProbeError("CFB has no root directory entry".into())) as BoxError
    })?;
    let mut storage_clsids = BTreeMap::new();
    let mut directory_entries = BTreeMap::new();
    collect_directory_entry(&ole, &[], &root, &mut directory_entries)?;
    for (path, entry) in &directory_entries {
        if entry.entry_type == litchi_cfb::consts::STGTY_STORAGE {
            storage_clsids.insert(path.clone(), entry.clsid.clone());
        }
    }

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
            raw_directory_image_bytes: raw_directory.bytes.len(),
            raw_directory_image_sha256: sha256_hex(&raw_directory.bytes),
        },
        stream_bytes,
        storage_clsids,
        directory_entries,
        raw_directory_image: raw_directory,
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
    format_specific_semantic_length_proven: bool,
    doc_source_target_utf16_units: Option<usize>,
    doc_requested_target_utf16_units: Option<usize>,
    ppt_source_live_slide_count: Option<usize>,
    ppt_output_live_slide_count: Option<usize>,
    ppt_removed_slide_id: Option<u32>,
    ppt_removed_slide_persist_id: Option<u32>,
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
        logical_stream_length_change_proven: false,
        format_specific_semantic_length_proven: false,
        doc_source_target_utf16_units: None,
        doc_requested_target_utf16_units: None,
        ppt_source_live_slide_count: None,
        ppt_output_live_slide_count: None,
        ppt_removed_slide_id: None,
        ppt_removed_slide_persist_id: None,
    }
}

fn directory_shape_matches(source: &Inventory, expected: &Inventory) -> bool {
    source.directory_entries.keys().collect::<BTreeSet<_>>()
        == expected.directory_entries.keys().collect::<BTreeSet<_>>()
        && source.directory_entries.iter().all(|(path, entry)| {
            expected.directory_entries.get(path).is_some_and(|other| {
                entry.entry_type == other.entry_type && entry.clsid == other.clsid
            })
        })
}

fn allocation_sids(left: &Inventory, right: &Inventory) -> (BTreeSet<u32>, BTreeSet<u32>) {
    let paths = left
        .directory_entries
        .keys()
        .chain(right.directory_entries.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let mut left_sids = BTreeSet::new();
    let mut right_sids = BTreeSet::new();
    if let Some(entry) = left.directory_entries.get(&Vec::new()) {
        left_sids.insert(entry.sid);
    }
    if let Some(entry) = right.directory_entries.get(&Vec::new()) {
        right_sids.insert(entry.sid);
    }
    for path in paths {
        let left_entry = left.directory_entries.get(&path);
        let right_entry = right.directory_entries.get(&path);
        let changed = match (left_entry, right_entry) {
            (Some(left_entry), Some(right_entry)) => {
                left_entry.bytes != right_entry.bytes
                    || left_entry.start_sector != right_entry.start_sector
                    || left_entry.is_minifat != right_entry.is_minifat
            },
            _ => true,
        };
        if changed {
            if let Some(entry) = left_entry {
                left_sids.insert(entry.sid);
            }
            if let Some(entry) = right_entry {
                right_sids.insert(entry.sid);
            }
        }
    }
    (left_sids, right_sids)
}

fn v3_size_sids(inventory: &Inventory) -> BTreeSet<u32> {
    if inventory.raw_directory_image.cfb_version != 3 {
        return BTreeSet::new();
    }
    inventory
        .directory_entries
        .values()
        .filter(|entry| entry.entry_type == litchi_cfb::consts::STGTY_STREAM)
        .map(|entry| entry.sid)
        .collect()
}

fn normalized_directory_bytes(inventory: &Inventory, allocation: &BTreeSet<u32>) -> Vec<u8> {
    let mut image = inventory.raw_directory_image.bytes.clone();
    let high_size_sids = v3_size_sids(inventory);
    for sid in allocation {
        let Ok(sid) = usize::try_from(*sid) else {
            continue;
        };
        let Some(start) = sid.checked_mul(CFB_DIRECTORY_ENTRY_SIZE) else {
            continue;
        };
        let Some(end) = start.checked_add(CFB_DIRECTORY_ENTRY_SIZE) else {
            continue;
        };
        let Some(entry) = image.get_mut(start..end) else {
            continue;
        };
        entry[CFB_ENTRY_START_SECTOR_OFFSET..CFB_ENTRY_STREAM_SIZE_OFFSET + 8].fill(0);
    }
    for sid in high_size_sids {
        if allocation.contains(&sid) {
            continue;
        }
        let Ok(sid) = usize::try_from(sid) else {
            continue;
        };
        let Some(start) = sid
            .checked_mul(CFB_DIRECTORY_ENTRY_SIZE)
            .and_then(|value| value.checked_add(CFB_ENTRY_STREAM_SIZE_OFFSET + 4))
        else {
            continue;
        };
        let Some(end) = start.checked_add(4) else {
            continue;
        };
        if let Some(field) = image.get_mut(start..end) {
            field.fill(0);
        }
    }
    image
}

fn normalized_directory_comparison(left: &Inventory, right: &Inventory) -> (String, usize) {
    if left.raw_directory_image.cfb_version != right.raw_directory_image.cfb_version
        || left.raw_directory_image.sector_size != right.raw_directory_image.sector_size
    {
        return ("incomparable_cfb_geometry".into(), usize::MAX);
    }
    let (left_sids, right_sids) = allocation_sids(left, right);
    let left_image = normalized_directory_bytes(left, &left_sids);
    let right_image = normalized_directory_bytes(right, &right_sids);
    let difference_bytes = left_image
        .iter()
        .zip(&right_image)
        .filter(|(left, right)| left != right)
        .count()
        + left_image.len().abs_diff(right_image.len());
    if difference_bytes == 0 {
        ("match_after_allocation_normalization".into(), 0)
    } else {
        (
            "differ_after_allocation_normalization".into(),
            difference_bytes,
        )
    }
}

#[derive(Clone, Debug, Serialize)]
struct RawDirectoryOracle {
    mode: String,
    source_sha256: String,
    expected_sha256: String,
    output_sha256: String,
    source_expected_normalized: String,
    expected_output_normalized: String,
    source_output_normalized: String,
    source_expected_difference_bytes: usize,
    expected_output_difference_bytes: usize,
    source_output_difference_bytes: usize,
    source_expected_ok: bool,
    source_output_ok: bool,
    policy_gate: String,
}

fn raw_directory_oracle(
    operation: Operation,
    policy: Policy,
    source: &Inventory,
    expected: &Inventory,
    output: &Inventory,
    enforce_policy: bool,
) -> (RawDirectoryOracle, bool) {
    let (source_expected_normalized, source_expected_difference_bytes) =
        normalized_directory_comparison(source, expected);
    let (expected_output_normalized, expected_output_difference_bytes) =
        normalized_directory_comparison(expected, output);
    let (source_output_normalized, source_output_difference_bytes) =
        normalized_directory_comparison(source, output);
    let source_expected_ok = source_expected_normalized == "match_after_allocation_normalization";
    let source_output_ok = source_output_normalized == "match_after_allocation_normalization";
    let output_gate_required = enforce_policy
        && (operation == Operation::Format
            || (operation == Operation::Container && policy == Policy::Reuse));
    let policy_ok = source_expected_ok && (!output_gate_required || source_output_ok);
    let mode = match (operation, policy) {
        (Operation::Format, _) => "public_format_raw_source_model_gate".to_string(),
        (Operation::Container, Policy::Rewrite) => "rewrite_raw_normalization_report".to_string(),
        (Operation::Container, Policy::Reuse) => "reuse_raw_source_model_gate".to_string(),
    };
    let policy_gate = if !source_expected_ok {
        "failed_source_expected_model".to_string()
    } else if output_gate_required && !source_output_ok {
        match operation {
            Operation::Format => "failed_public_source_model".to_string(),
            Operation::Container => "failed_reuse_source_model".to_string(),
        }
    } else if operation == Operation::Container && policy == Policy::Rewrite {
        "rewrite_policy_report_only".to_string()
    } else {
        "passed_source_expected_model".to_string()
    };
    (
        RawDirectoryOracle {
            mode,
            source_sha256: source.summary.raw_directory_image_sha256.clone(),
            expected_sha256: expected.summary.raw_directory_image_sha256.clone(),
            output_sha256: output.summary.raw_directory_image_sha256.clone(),
            source_expected_normalized,
            expected_output_normalized,
            source_output_normalized,
            source_expected_difference_bytes,
            expected_output_difference_bytes,
            source_output_difference_bytes,
            source_expected_ok,
            source_output_ok,
            policy_gate,
        },
        policy_ok,
    )
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

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
struct PptSlideIdentity {
    slide_id: u32,
    persist_id: u32,
    flags: u32,
    text_count: u32,
    live_record_bytes: usize,
    live_record_sha256: String,
    #[serde(skip)]
    live_record: Vec<u8>,
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
    snapshot
        .slides()
        .iter()
        .map(|slide| {
            let offset = mapping.get(&slide.persist_id()).copied().ok_or_else(|| {
                "PPT live slide persist mapping is missing from the source projection".to_string()
            })?;
            let record = record_slice(
                &document,
                usize::try_from(offset)
                    .map_err(|_| "PPT live slide offset does not fit usize".to_string())?,
            )?;
            Ok(PptSlideIdentity {
                slide_id: slide.slide_id(),
                persist_id: slide.persist_id(),
                flags: slide.flags(),
                text_count: slide.text_count(),
                live_record_bytes: record.len(),
                live_record_sha256: sha256_hex(record),
                live_record: record.to_vec(),
            })
        })
        .collect()
}

fn doc_witness(value: DocSemanticWitness) -> SemanticWitness {
    SemanticWitness::Doc(Box::new(value))
}

fn ppt_witness(value: PptSemanticWitness) -> SemanticWitness {
    SemanticWitness::Ppt(Box::new(value))
}

#[cfg(test)]
fn ppt_slide_projection_matches(source: &[PptSlideIdentity], output: &[PptSlideIdentity]) -> bool {
    source.len() > 1
        && output.len() == source.len() - 1
        && output
            == source
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != 1)
                .map(|(_, slide)| slide.clone())
                .collect::<Vec<_>>()
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
struct PptPublicSlide {
    slide_id: u32,
    persist_id: u32,
    flags: u32,
    text_count: u32,
    list_text_bytes: usize,
    list_text_sha256: String,
    outline_refs_sha256: String,
    outline_interactions_sha256: String,
    slide_text_bytes: usize,
    slide_text_sha256: String,
    notes_status: String,
    notes_sha256: String,
    comments_status: String,
    comments_sha256: String,
    #[serde(skip)]
    list_text: String,
    #[serde(skip)]
    outline_refs: Vec<litchi_ppt::OutlineTextRef>,
    #[serde(skip)]
    outline_interactions: Vec<litchi_ppt::TextBodyInteractions>,
    #[serde(skip)]
    slide_text: String,
    #[serde(skip)]
    notes_identity: Option<(u32, u32, u32)>,
    #[serde(skip)]
    notes_text: Option<String>,
    #[serde(skip)]
    comments: Vec<litchi_ppt::slide::ParsedComment>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
struct PptSlideWitness {
    slide_id: u32,
    persist_id: u32,
    flags: u32,
    text_count: u32,
    live_record_bytes: usize,
    live_record_sha256: String,
    list_text_bytes: usize,
    list_text_sha256: String,
    outline_refs_sha256: String,
    outline_interactions_sha256: String,
    slide_text_bytes: usize,
    slide_text_sha256: String,
    notes_status: String,
    notes_sha256: String,
    comments_status: String,
    comments_sha256: String,
    #[serde(skip)]
    live_record: Vec<u8>,
    #[serde(skip)]
    list_text: String,
    #[serde(skip)]
    outline_refs: Vec<litchi_ppt::OutlineTextRef>,
    #[serde(skip)]
    outline_interactions: Vec<litchi_ppt::TextBodyInteractions>,
    #[serde(skip)]
    slide_text: String,
    #[serde(skip)]
    notes_identity: Option<(u32, u32, u32)>,
    #[serde(skip)]
    notes_text: Option<String>,
    #[serde(skip)]
    comments: Vec<litchi_ppt::slide::ParsedComment>,
}

fn debug_digest<T: std::fmt::Debug>(value: &T) -> String {
    sha256_hex(format!("{value:?}").as_bytes())
}

fn ppt_public_projection(bytes: &[u8]) -> Result<Vec<PptPublicSlide>, String> {
    let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.to_vec()))
        .map_err(|error| error.to_string())?;
    let presentation = package.presentation().map_err(|error| error.to_string())?;
    let entries = presentation.slide_directory().entries();
    let slides = presentation.slides().map_err(|error| error.to_string())?;
    if entries.len() != slides.len() {
        return Err("PPT public directory and slide views have different lengths".into());
    }
    entries
        .iter()
        .zip(slides.iter())
        .map(|(entry, slide)| {
            let text = slide.text().map_err(|error| error.to_string())?;
            // Notes and comments are optional dependencies.  Keep an explicit
            // unavailable state when a fixture exposes a malformed or newer
            // dependency that this public projection cannot decode; that is
            // evidence in the witness and does not make the required slide
            // identity/text projection disappear.
            let (notes_status, notes_sha256, notes_identity, notes_text) =
                match slide.speaker_notes() {
                    Err(error) => (
                        "unavailable".into(),
                        sha256_hex(error.to_string().as_bytes()),
                        None,
                        None,
                    ),
                    Ok(None) => ("absent".into(), String::new(), None, None),
                    Ok(Some(notes)) => match notes.text() {
                        Err(error) => (
                            "unavailable".into(),
                            sha256_hex(error.to_string().as_bytes()),
                            Some((notes.notes_id(), notes.persist_id(), notes.slide_id_ref())),
                            None,
                        ),
                        Ok(text) => {
                            let payload = format!(
                                "{}:{}:{}:{text}",
                                notes.notes_id(),
                                notes.persist_id(),
                                notes.slide_id_ref()
                            );
                            (
                                "present".into(),
                                sha256_hex(payload.as_bytes()),
                                Some((notes.notes_id(), notes.persist_id(), notes.slide_id_ref())),
                                Some(text.to_string()),
                            )
                        },
                    },
                };
            let (comments_status, comments_sha256, comments) = match slide.comments() {
                Err(error) => (
                    "unavailable".into(),
                    sha256_hex(error.to_string().as_bytes()),
                    Vec::new(),
                ),
                Ok(comments) if comments.is_empty() => ("absent".into(), String::new(), comments),
                Ok(comments) => ("present".into(), debug_digest(&comments), comments),
            };
            let list_text = entry.list_text().to_string();
            let outline_refs = entry.outline_text_refs().to_vec();
            let outline_interactions = entry.outline_text_interactions().to_vec();
            Ok(PptPublicSlide {
                slide_id: entry.slide_id(),
                persist_id: entry.persist_id(),
                flags: entry.flags(),
                text_count: entry.text_placeholder_count(),
                list_text_bytes: list_text.len(),
                list_text_sha256: sha256_hex(list_text.as_bytes()),
                outline_refs_sha256: debug_digest(&outline_refs),
                outline_interactions_sha256: debug_digest(&outline_interactions),
                slide_text_bytes: text.len(),
                slide_text_sha256: sha256_hex(text.as_bytes()),
                notes_status,
                notes_sha256,
                comments_status,
                comments_sha256,
                list_text,
                outline_refs,
                outline_interactions,
                slide_text: text.to_string(),
                notes_identity,
                notes_text,
                comments,
            })
        })
        .collect()
}

fn ppt_combined_projection(bytes: &[u8]) -> Result<Vec<PptSlideWitness>, String> {
    let raw = ppt_slide_projection(bytes)?;
    let public = ppt_public_projection(bytes)?;
    if raw.len() != public.len() {
        return Err("PPT raw and public slide projections have different lengths".into());
    }
    raw.into_iter()
        .zip(public)
        .map(|(raw, public)| {
            if raw.slide_id != public.slide_id
                || raw.persist_id != public.persist_id
                || raw.flags != public.flags
                || raw.text_count != public.text_count
            {
                return Err("PPT raw and public slide identities disagree".into());
            }
            Ok(PptSlideWitness {
                slide_id: raw.slide_id,
                persist_id: raw.persist_id,
                flags: raw.flags,
                text_count: raw.text_count,
                live_record_bytes: raw.live_record_bytes,
                live_record_sha256: raw.live_record_sha256,
                live_record: raw.live_record,
                list_text_bytes: public.list_text_bytes,
                list_text_sha256: public.list_text_sha256,
                outline_refs_sha256: public.outline_refs_sha256,
                outline_interactions_sha256: public.outline_interactions_sha256,
                slide_text_bytes: public.slide_text_bytes,
                slide_text_sha256: public.slide_text_sha256,
                notes_status: public.notes_status,
                notes_sha256: public.notes_sha256,
                comments_status: public.comments_status,
                comments_sha256: public.comments_sha256,
                list_text: public.list_text,
                outline_refs: public.outline_refs,
                outline_interactions: public.outline_interactions,
                slide_text: public.slide_text,
                notes_identity: public.notes_identity,
                notes_text: public.notes_text,
                comments: public.comments,
            })
        })
        .collect()
}

fn ppt_combined_projection_matches(source: &[PptSlideWitness], output: &[PptSlideWitness]) -> bool {
    source.len() > 1
        && output.len() == source.len() - 1
        && output
            == source
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != 1)
                .map(|(_, slide)| slide.clone())
                .collect::<Vec<_>>()
}

#[derive(Clone, Debug, Serialize, Default)]
struct PptSemanticWitness {
    selected_index: usize,
    source_slide_count: usize,
    output_slide_count: usize,
    source_order: Vec<PptSlideWitness>,
    output_order: Vec<PptSlideWitness>,
    removed_slide: Option<PptSlideWitness>,
    order_and_survivor_result: String,
    comparison_basis: String,
    dependency_scope: Vec<String>,
    error: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
enum SemanticWitness {
    Doc(Box<DocSemanticWitness>),
    Ppt(Box<PptSemanticWitness>),
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

fn unchanged_stream_bytes_match(
    source: &Inventory,
    expected: &Inventory,
    replacement_paths: &BTreeSet<Vec<String>>,
) -> bool {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    if source_paths != expected_paths {
        return false;
    }
    for (path, before) in &source.stream_bytes {
        if !replacement_paths.contains(path) && expected.stream_bytes.get(path) != Some(before) {
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

#[derive(Clone, Debug, Serialize, Default)]
struct CollectionWitness {
    name: String,
    source_status: String,
    output_status: String,
    source_count: usize,
    output_count: usize,
    source_sha256: String,
    output_sha256: String,
    result: String,
    direct_match: bool,
    direct_result: String,
}

#[derive(Clone, Debug, Serialize, Default)]
struct DocSemanticWitness {
    target_position: usize,
    source_paragraph_count: usize,
    output_paragraph_count: usize,
    source_target_text_sha256: String,
    output_target_text_sha256: String,
    requested_text_sha256: String,
    source_target_utf16_units: usize,
    output_target_utf16_units: usize,
    requested_text_utf16_units: usize,
    target_text_difference: String,
    target_utf16_length_difference: String,
    all_projection: String,
    comparison_basis: String,
    accepted_projection: CollectionWitness,
    rejected_projection: CollectionWitness,
    stories: Vec<CollectionWitness>,
    table_cells: CollectionWitness,
    field_results: CollectionWitness,
    revisions: CollectionWitness,
    revision_authors_source: Vec<String>,
    revision_authors_output: Vec<String>,
    embedded_objects: CollectionWitness,
    error: Option<String>,
}

#[derive(Clone, Debug)]
struct CollectionState {
    status: String,
    count: usize,
    sha256: String,
}

fn collection_state<T: std::fmt::Debug>(result: Result<(usize, T), String>) -> CollectionState {
    match result {
        Ok((count, value)) => CollectionState {
            status: if count == 0 { "empty" } else { "ok" }.into(),
            count,
            sha256: debug_digest(&value),
        },
        Err(error) => CollectionState {
            status: format!("unavailable:{error}"),
            count: 0,
            sha256: String::new(),
        },
    }
}

fn collection_witness(
    name: &str,
    source: CollectionState,
    output: CollectionState,
) -> CollectionWitness {
    let result = if source.status == "empty" && output.status == "empty" {
        "absent_equal"
    } else if source.status.starts_with("unavailable:") && source.status == output.status {
        "unavailable_equal"
    } else if (source.status == "ok" || source.status == "empty")
        && (output.status == "ok" || output.status == "empty")
        && source.count == output.count
        && source.sha256 == output.sha256
    {
        "equal"
    } else {
        "changed_or_unavailable"
    };
    CollectionWitness {
        name: name.into(),
        source_status: source.status,
        output_status: output.status,
        source_count: source.count,
        output_count: output.count,
        source_sha256: source.sha256,
        output_sha256: output.sha256,
        result: result.into(),
        direct_match: false,
        direct_result: "not_compared".into(),
    }
}

fn set_direct_result(witness: &mut CollectionWitness, comparison: Option<bool>) {
    match comparison {
        Some(true) => {
            witness.direct_match = true;
            witness.direct_result = "equal".into();
        },
        Some(false) => {
            witness.direct_match = false;
            witness.direct_result = "different".into();
        },
        None if witness.result == "unavailable_equal" => {
            witness.direct_match = true;
            witness.direct_result = "unavailable_equal".into();
        },
        None => {
            witness.direct_match = false;
            witness.direct_result = "unavailable_or_not_compared".into();
        },
    }
}

fn direct_slice_match<T: PartialEq>(
    witness: &mut CollectionWitness,
    source: Option<&[T]>,
    output: Option<&[T]>,
) {
    set_direct_result(
        witness,
        source.zip(output).map(|(source, output)| source == output),
    );
}

fn direct_value_match<T: PartialEq>(
    witness: &mut CollectionWitness,
    source: Option<&T>,
    output: Option<&T>,
) {
    set_direct_result(
        witness,
        source.zip(output).map(|(source, output)| source == output),
    );
}

fn paragraph_projection_witness(
    name: &str,
    source: Result<Vec<String>, String>,
    output: Result<Vec<String>, String>,
    replacement: &str,
) -> CollectionWitness {
    let source_values = source.as_ref().ok();
    let output_values = output.as_ref().ok();
    let source_state = match source_values {
        Some(values) => collection_state(Ok((values.len(), values.clone()))),
        None => collection_state::<Vec<String>>(Err(source
            .as_ref()
            .err()
            .cloned()
            .unwrap_or_else(|| "projection unavailable".into()))),
    };
    let output_state = match output_values {
        Some(values) => collection_state(Ok((values.len(), values.clone()))),
        None => collection_state::<Vec<String>>(Err(output
            .as_ref()
            .err()
            .cloned()
            .unwrap_or_else(|| "projection unavailable".into()))),
    };
    let mut witness = collection_witness(name, source_state, output_state);
    if let (Some(source_values), Some(output_values)) = (source_values, output_values) {
        let expected = source_values
            .iter()
            .enumerate()
            .map(|(index, value)| {
                if index == 0 {
                    replacement.to_string()
                } else {
                    value.clone()
                }
            })
            .collect::<Vec<_>>();
        if output_values == &expected {
            witness.result = "target_replaced".into();
            set_direct_result(&mut witness, Some(true));
        } else if output_values == source_values {
            witness.result = "equal".into();
            set_direct_result(&mut witness, Some(true));
        } else {
            set_direct_result(&mut witness, Some(false));
        }
    } else {
        set_direct_result(&mut witness, None);
    }
    witness
}

fn story_collection_witness(
    name: &str,
    story: litchi_doc::body_text::Story,
    source: CollectionState,
    output: CollectionState,
    source_items: Option<&[litchi_doc::body_text::TextItem]>,
    output_items: Option<&[litchi_doc::body_text::TextItem]>,
    replacement: &str,
) -> CollectionWitness {
    let mut witness = collection_witness(name, source, output);
    let Some((source_items, output_items)) = source_items.zip(output_items) else {
        set_direct_result(&mut witness, None);
        return witness;
    };
    if story == litchi_doc::body_text::Story::Main {
        let matches_target = source_items.len() == output_items.len()
            && source_items
                .iter()
                .zip(output_items)
                .all(|(before, after)| {
                    let expected = if before.target()
                        == litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0))
                    {
                        replacement
                    } else {
                        before.text()
                    };
                    after.target() == before.target() && after.text() == expected
                });
        if matches_target {
            witness.result = "target_replaced".into();
            set_direct_result(&mut witness, Some(true));
        } else {
            set_direct_result(&mut witness, Some(false));
        }
    } else if source_items == output_items {
        set_direct_result(&mut witness, Some(true));
    } else {
        set_direct_result(&mut witness, Some(false));
    }
    witness
}

fn doc_collection_state<T: std::fmt::Debug + Clone>(
    result: Result<Vec<T>, String>,
) -> (CollectionState, Option<Vec<T>>) {
    match result {
        Ok(values) => {
            let state = collection_state(Ok((values.len(), values.clone())));
            (state, Some(values))
        },
        Err(error) => (collection_state::<Vec<T>>(Err(error)), None),
    }
}

fn value_collection_state<T: std::fmt::Debug + Clone>(
    result: Result<(usize, T), String>,
) -> (CollectionState, Option<T>) {
    match result {
        Ok((count, value)) => {
            let state = collection_state(Ok((count, value.clone())));
            (state, Some(value))
        },
        Err(error) => (collection_state::<T>(Err(error)), None),
    }
}

fn doc_semantic_witness(
    source: &litchi_doc::body_text::Snapshot,
    output: &litchi_doc::body_text::Snapshot,
    replacement: &str,
) -> (bool, String, DocSemanticWitness) {
    let mut witness = DocSemanticWitness {
        target_position: 0,
        requested_text_sha256: sha256_hex(replacement.as_bytes()),
        requested_text_utf16_units: replacement.encode_utf16().count(),
        comparison_basis: "direct_text_and_collection_values_with_sha256_report_fields".into(),
        ..DocSemanticWitness::default()
    };
    let source_paragraphs = match source.paragraphs(litchi_doc::body_text::Projection::All) {
        Ok(value) => value,
        Err(error) => {
            let message = format!("DOC source paragraph read failed: {error}");
            witness.error = Some(message.clone());
            return (false, message, witness);
        },
    };
    let output_paragraphs = match output.paragraphs(litchi_doc::body_text::Projection::All) {
        Ok(value) => value,
        Err(error) => {
            let message = format!("DOC output paragraph read failed: {error}");
            witness.error = Some(message.clone());
            return (false, message, witness);
        },
    };
    witness.source_paragraph_count = source_paragraphs.len();
    witness.output_paragraph_count = output_paragraphs.len();
    let Some(source_target) = source_paragraphs.first().map(|value| value.text()) else {
        let message = "DOC source has no ordinary main-story paragraph 0".to_string();
        witness.error = Some(message.clone());
        return (false, message, witness);
    };
    let Some(output_target) = output_paragraphs.first().map(|value| value.text()) else {
        let message = "DOC output has no ordinary main-story paragraph 0".to_string();
        witness.error = Some(message.clone());
        return (false, message, witness);
    };
    witness.source_target_text_sha256 = sha256_hex(source_target.as_bytes());
    witness.output_target_text_sha256 = sha256_hex(output_target.as_bytes());
    witness.source_target_utf16_units = source_target.encode_utf16().count();
    witness.output_target_utf16_units = output_target.encode_utf16().count();
    witness.target_text_difference = if source_target != replacement {
        "changed".into()
    } else {
        "same_as_requested".into()
    };
    witness.target_utf16_length_difference =
        if witness.source_target_utf16_units != witness.requested_text_utf16_units {
            "changed".into()
        } else {
            "same_length".into()
        };
    let projection_ok = doc_text_projection_matches(
        &source_paragraphs
            .iter()
            .map(|paragraph| paragraph.text().to_string())
            .collect::<Vec<_>>(),
        &output_paragraphs
            .iter()
            .map(|paragraph| paragraph.text().to_string())
            .collect::<Vec<_>>(),
        replacement,
    );
    witness.all_projection = if projection_ok {
        "target_replaced_and_survivors_equal".into()
    } else {
        "mismatch".into()
    };

    witness.accepted_projection = paragraph_projection_witness(
        "accepted_projection",
        source
            .paragraphs(litchi_doc::body_text::Projection::Accepted)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        output
            .paragraphs(litchi_doc::body_text::Projection::Accepted)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        replacement,
    );
    witness.rejected_projection = paragraph_projection_witness(
        "rejected_projection",
        source
            .paragraphs(litchi_doc::body_text::Projection::Rejected)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        output
            .paragraphs(litchi_doc::body_text::Projection::Rejected)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        replacement,
    );

    let story_specs = [
        (litchi_doc::body_text::Story::Main, "main"),
        (litchi_doc::body_text::Story::Footnote, "footnote"),
        (litchi_doc::body_text::Story::Header, "header"),
        (litchi_doc::body_text::Story::Comment, "comment"),
        (litchi_doc::body_text::Story::Endnote, "endnote"),
        (litchi_doc::body_text::Story::Textbox, "textbox"),
        (
            litchi_doc::body_text::Story::HeaderTextbox,
            "header_textbox",
        ),
    ];
    let mut collections_ok = true;
    for (story, name) in story_specs {
        let source_result = source
            .story_paragraphs(story)
            .map_err(|error| error.to_string());
        let output_result = output
            .story_paragraphs(story)
            .map_err(|error| error.to_string());
        let (source_state, source_items) = doc_collection_state(source_result);
        let (output_state, output_items) = doc_collection_state(output_result);
        let item_witness = story_collection_witness(
            name,
            story,
            source_state,
            output_state,
            source_items.as_deref(),
            output_items.as_deref(),
            replacement,
        );
        collections_ok &= item_witness.direct_match;
        witness.stories.push(item_witness);
    }

    let (source_state, source_items) =
        doc_collection_state(source.table_cells().map_err(|error| error.to_string()));
    let (output_state, output_items) =
        doc_collection_state(output.table_cells().map_err(|error| error.to_string()));
    witness.table_cells = collection_witness("table_cells", source_state, output_state);
    direct_slice_match(
        &mut witness.table_cells,
        source_items.as_deref(),
        output_items.as_deref(),
    );
    collections_ok &= witness.table_cells.direct_match;
    drop((source_items, output_items));

    let (source_state, source_items) =
        doc_collection_state(source.field_results().map_err(|error| error.to_string()));
    let (output_state, output_items) =
        doc_collection_state(output.field_results().map_err(|error| error.to_string()));
    witness.field_results = collection_witness("field_results", source_state, output_state);
    direct_slice_match(
        &mut witness.field_results,
        source_items.as_deref(),
        output_items.as_deref(),
    );
    collections_ok &= witness.field_results.direct_match;
    drop((source_items, output_items));

    let (source_state, source_items) =
        doc_collection_state(source.revisions().map_err(|error| error.to_string()));
    let (output_state, output_items) =
        doc_collection_state(output.revisions().map_err(|error| error.to_string()));
    witness.revisions = collection_witness("revisions", source_state, output_state);
    direct_slice_match(
        &mut witness.revisions,
        source_items.as_deref(),
        output_items.as_deref(),
    );
    collections_ok &= witness.revisions.direct_match;
    witness.revision_authors_source = source_items
        .as_ref()
        .map(|items| {
            items
                .iter()
                .map(|revision| format!("{}:{}", revision.author_index, revision.author))
                .collect()
        })
        .unwrap_or_default();
    witness.revision_authors_output = output_items
        .as_ref()
        .map(|items| {
            items
                .iter()
                .map(|revision| format!("{}:{}", revision.author_index, revision.author))
                .collect()
        })
        .unwrap_or_default();
    drop((source_items, output_items));

    let source_objects = source
        .embedded_objects()
        .map(|value| (value.len(), value))
        .map_err(|error| error.to_string());
    let output_objects = output
        .embedded_objects()
        .map(|value| (value.len(), value))
        .map_err(|error| error.to_string());
    let (source_state, source_objects) = value_collection_state(source_objects);
    let (output_state, output_objects) = value_collection_state(output_objects);
    witness.embedded_objects = collection_witness("embedded_objects", source_state, output_state);
    direct_value_match(
        &mut witness.embedded_objects,
        source_objects.as_ref(),
        output_objects.as_ref(),
    );
    collections_ok &= witness.embedded_objects.direct_match;
    collections_ok &= witness.accepted_projection.direct_match;
    collections_ok &= witness.rejected_projection.direct_match;

    let semantic_ok = projection_ok
        && source_target != replacement
        && witness.source_target_utf16_units != witness.requested_text_utf16_units
        && collections_ok;
    if !semantic_ok {
        let message = "DOC semantic witness did not prove a changed paragraph-0 replacement with unchanged applicable collections".to_string();
        witness.error = Some(message.clone());
        return (false, message, witness);
    }
    (true, String::new(), witness)
}

#[derive(Clone, Debug, Serialize)]
struct Oracle {
    source_structurally_valid: bool,
    expected_structurally_valid: bool,
    output_structurally_valid: bool,
    semantic_reopen_ok: bool,
    semantic_witness: SemanticWitness,
    semantic_directory_metadata_matches_expected: bool,
    source_directory_shape_matches_expected: bool,
    source_directory_metadata_differences: Vec<DirectoryMetadataDifference>,
    directory_metadata_differences: Vec<DirectoryMetadataDifference>,
    raw_directory: RawDirectoryOracle,
    raw_directory_policy_ok: bool,
    source_stream_paths_match_expected: bool,
    output_stream_paths_match_expected: bool,
    output_stream_bytes_match_expected: bool,
    unchanged_stream_bytes_match_expected: bool,
    unchanged_source_output_stream_bytes_match: bool,
    changed_stream_paths_allowed: bool,
    actual_changed_stream_paths_allowed: bool,
    root_clsid_preserved: bool,
    storage_clsids_preserved: bool,
    no_deleted_or_new_streams: bool,
    oracle_ok: bool,
    failure_reasons: Vec<String>,
}

fn finalize_length_proof(case: Case, oracle: &Oracle, proof: &mut ChangedLengthProof) {
    proof.format_specific_semantic_length_proven = match (case, &oracle.semantic_witness) {
        (Case::DocFloat | Case::DocNoHf, SemanticWitness::Doc(witness)) => {
            proof.doc_source_target_utf16_units = Some(witness.source_target_utf16_units);
            proof.doc_requested_target_utf16_units = Some(witness.requested_text_utf16_units);
            witness.target_text_difference == "changed"
                && witness.target_utf16_length_difference == "changed"
                && witness.all_projection == "target_replaced_and_survivors_equal"
        },
        (Case::Ppt45543, SemanticWitness::Ppt(witness)) => {
            proof.ppt_source_live_slide_count = Some(witness.source_slide_count);
            proof.ppt_output_live_slide_count = Some(witness.output_slide_count);
            if let Some(removed) = &witness.removed_slide {
                proof.ppt_removed_slide_id = Some(removed.slide_id);
                proof.ppt_removed_slide_persist_id = Some(removed.persist_id);
            }
            witness.removed_slide.is_some()
                && witness.source_slide_count == witness.output_slide_count + 1
                && witness.order_and_survivor_result
                    == "source_order_with_selected_slide_removed_and_survivors_equal"
        },
        _ => false,
    };
    proof.logical_stream_length_change_proven =
        proof.any_stream_length_changed && proof.format_specific_semantic_length_proven;
}

fn semantic_reopen_check(
    case: Case,
    source_bytes: &[u8],
    output_bytes: &[u8],
    text: &str,
) -> (bool, String, SemanticWitness) {
    match case {
        Case::DocFloat | Case::DocNoHf => {
            let limits = litchi_doc::tracked_revision::Limits::default();
            let source = match litchi_doc::body_text::Snapshot::open(source_bytes.to_vec(), limits)
            {
                Ok(snapshot) => snapshot,
                Err(error) => {
                    return (
                        false,
                        format!("DOC source public reopen failed: {error}"),
                        doc_witness(DocSemanticWitness {
                            error: Some(format!("DOC source public reopen failed: {error}")),
                            ..DocSemanticWitness::default()
                        }),
                    );
                },
            };
            let output = match litchi_doc::body_text::Snapshot::open(output_bytes.to_vec(), limits)
            {
                Ok(snapshot) => snapshot,
                Err(error) => {
                    return (
                        false,
                        format!("DOC output public reopen failed: {error}"),
                        doc_witness(DocSemanticWitness {
                            error: Some(format!("DOC output public reopen failed: {error}")),
                            ..DocSemanticWitness::default()
                        }),
                    );
                },
            };
            let (ok, failure, witness) = doc_semantic_witness(&source, &output, text);
            (ok, failure, doc_witness(witness))
        },
        Case::Ppt45543 => {
            let source = match litchi_ppt::slide_order::Snapshot::from_bytes(source_bytes.to_vec())
            {
                Ok(snapshot) => snapshot,
                Err(error) => {
                    return (
                        false,
                        format!("PPT source public reopen failed: {error}"),
                        ppt_witness(PptSemanticWitness {
                            error: Some(format!("PPT source public reopen failed: {error}")),
                            ..PptSemanticWitness::default()
                        }),
                    );
                },
            };
            let output = match litchi_ppt::slide_order::Snapshot::from_bytes(output_bytes.to_vec())
            {
                Ok(snapshot) => snapshot,
                Err(error) => {
                    return (
                        false,
                        format!("PPT output public reopen failed: {error}"),
                        ppt_witness(PptSemanticWitness {
                            error: Some(format!("PPT output public reopen failed: {error}")),
                            ..PptSemanticWitness::default()
                        }),
                    );
                },
            };
            let expected_count = source.slide_count().saturating_sub(1);
            if output.slide_count() != expected_count {
                let message = format!(
                    "PPT slide count {} does not equal source count minus one {expected_count}",
                    output.slide_count()
                );
                return (
                    false,
                    message.clone(),
                    ppt_witness(PptSemanticWitness {
                        selected_index: 1,
                        source_slide_count: source.slide_count(),
                        output_slide_count: output.slide_count(),
                        error: Some(message),
                        ..PptSemanticWitness::default()
                    }),
                );
            }
            let source_projection = match ppt_combined_projection(source_bytes) {
                Ok(projection) => projection,
                Err(error) => {
                    return (
                        false,
                        format!("PPT source identity projection failed: {error}"),
                        ppt_witness(PptSemanticWitness {
                            selected_index: 1,
                            source_slide_count: source.slide_count(),
                            output_slide_count: output.slide_count(),
                            error: Some(format!("PPT source identity projection failed: {error}")),
                            ..PptSemanticWitness::default()
                        }),
                    );
                },
            };
            let output_projection = match ppt_combined_projection(output_bytes) {
                Ok(projection) => projection,
                Err(error) => {
                    return (
                        false,
                        format!("PPT output identity projection failed: {error}"),
                        ppt_witness(PptSemanticWitness {
                            selected_index: 1,
                            source_slide_count: source.slide_count(),
                            output_slide_count: output.slide_count(),
                            source_order: source_projection.clone(),
                            error: Some(format!("PPT output identity projection failed: {error}")),
                            ..PptSemanticWitness::default()
                        }),
                    );
                },
            };
            let order_ok = ppt_combined_projection_matches(&source_projection, &output_projection);
            let removed_slide = source_projection.get(1).cloned();
            let witness = PptSemanticWitness {
                selected_index: 1,
                source_slide_count: source_projection.len(),
                output_slide_count: output_projection.len(),
                source_order: source_projection.clone(),
                output_order: output_projection.clone(),
                removed_slide,
                order_and_survivor_result: if order_ok {
                    "source_order_with_selected_slide_removed_and_survivors_equal".into()
                } else {
                    "mismatch".into()
                },
                comparison_basis: "direct_public_values_and_live_record_bytes".into(),
                dependency_scope: vec![
                    "per-survivor slide text/list-text/outline/notes/comments values and live persisted record bytes are compared directly; digests are report fields".into(),
                    "unaffected top-level CFB streams and storages are compared by the complete stream/path/directory oracle".into(),
                    "dependencies embedded inside the allowed PowerPoint Document stream are not separately projected by this witness".into(),
                ],
                error: None,
            };
            if !order_ok {
                return (
                    false,
                    "PPT output live order or survivor payload differs from source with slide 1 removed".into(),
                    ppt_witness(witness),
                );
            }
            (true, String::new(), ppt_witness(witness))
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
    replacement_paths: &BTreeSet<Vec<String>>,
    enforce_policy: bool,
) -> Oracle {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    let source_stream_paths_match_expected = source_paths == expected_paths;
    let (output_stream_paths_match_expected, output_stream_bytes_match_expected) =
        exact_output_streams(expected, output);
    let unchanged_stream_bytes_match_expected =
        unchanged_stream_bytes_match(source, expected, replacement_paths);
    let unchanged_source_output_stream_bytes_match =
        unchanged_stream_bytes_match(source, output, replacement_paths);
    let actual_changed_stream_paths_allowed =
        changed_stream_paths_allowed(args.case, source, output);
    let changed_stream_paths_allowed = changed_stream_paths_allowed(args.case, source, expected);
    let (root_clsid_preserved, storage_clsids_preserved) =
        preserved_clsids(source, expected, output);
    let (semantic_reopen_ok, semantic_failure, semantic_witness) =
        semantic_reopen_check(args.case, source_bytes, output_bytes, &args.text);
    let source_directory_metadata_differences = directory_metadata_differences(source, expected);
    let directory_metadata_differences = directory_metadata_differences(expected, output);
    let semantic_directory_metadata_matches_expected =
        semantic_directory_metadata_matches(expected, output);
    let source_directory_shape_matches_expected = directory_shape_matches(source, expected);
    let (raw_directory, raw_directory_policy_ok) = raw_directory_oracle(
        args.operation,
        args.policy,
        source,
        expected,
        output,
        enforce_policy,
    );
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
    if !source_directory_shape_matches_expected {
        failure_reasons.push("public edit changed the CFB directory path/type/CLSID shape".into());
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
    if !unchanged_source_output_stream_bytes_match {
        failure_reasons.push(
            "one or more source streams changed outside the replacement set in the measured output"
                .into(),
        );
    }
    if !changed_stream_paths_allowed {
        failure_reasons
            .push("a changed stream is outside the allowed public-edit stream set".into());
    }
    if !actual_changed_stream_paths_allowed {
        failure_reasons
            .push("a measured output stream is outside the allowed public-edit stream set".into());
    }
    if !semantic_reopen_ok {
        failure_reasons.push(semantic_failure);
    }
    if !semantic_directory_metadata_matches_expected {
        failure_reasons.push("semantic CFB directory metadata differs from expected".into());
    }
    if !raw_directory_policy_ok {
        failure_reasons.push(
            "raw directory source/model policy gate failed outside planner-owned allocation fields"
                .into(),
        );
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
        && source_directory_shape_matches_expected
        && output_stream_bytes_match_expected
        && unchanged_stream_bytes_match_expected
        && unchanged_source_output_stream_bytes_match
        && changed_stream_paths_allowed
        && actual_changed_stream_paths_allowed
        && root_clsid_preserved
        && storage_clsids_preserved
        && semantic_reopen_ok
        && semantic_directory_metadata_matches_expected
        && raw_directory_policy_ok;
    Oracle {
        source_structurally_valid,
        expected_structurally_valid,
        output_structurally_valid,
        semantic_reopen_ok,
        semantic_witness,
        semantic_directory_metadata_matches_expected,
        source_directory_shape_matches_expected,
        source_directory_metadata_differences,
        directory_metadata_differences,
        raw_directory,
        raw_directory_policy_ok,
        source_stream_paths_match_expected,
        output_stream_paths_match_expected,
        output_stream_bytes_match_expected,
        unchanged_stream_bytes_match_expected,
        unchanged_source_output_stream_bytes_match,
        changed_stream_paths_allowed,
        actual_changed_stream_paths_allowed,
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

#[inline(never)]
fn measured_public_format(source: &[u8], args: &Args) -> Result<Vec<u8>, BoxError> {
    public_format_edit(source, args)
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
    let output = measured_public_format(source, args)?;
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
    phase_contract: String,
    input: String,
    policy: String,
    policy_applied: bool,
    policy_application_scope: String,
    policy_argument_effect: String,
    policy_contract: String,
    timing_claim: bool,
    allocator_instrumented: bool,
    allocation_ownership_contract: String,
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
    oracle_controls: Vec<OracleControl>,
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
    replacement_paths: &BTreeSet<Vec<String>>,
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
        replacement_paths,
        true,
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

#[derive(Clone, Debug, Serialize)]
struct OracleControl {
    name: String,
    status: String,
    rejected: bool,
    failure_reasons: Vec<String>,
}

fn control_args(args: &Args) -> Args {
    Args {
        case: args.case,
        input: args.input.clone(),
        operation: Operation::Container,
        policy: Policy::Reuse,
        warmups: 0,
        samples: 0,
        text: args.text.clone(),
    }
}

fn evaluate_oracle_control(
    name: &str,
    args: &Args,
    source: &Inventory,
    expected: &Inventory,
    source_bytes: &[u8],
    expected_bytes: &[u8],
    output_bytes: Result<Vec<u8>, BoxError>,
    replacement_paths: &BTreeSet<Vec<String>>,
) -> OracleControl {
    let output_bytes = match output_bytes {
        Ok(bytes) => bytes,
        Err(error) => {
            return OracleControl {
                name: name.into(),
                status: format!("not_run:{error}"),
                rejected: true,
                failure_reasons: vec![error.to_string()],
            };
        },
    };
    let output = match inventory(&output_bytes) {
        Ok(value) => value,
        Err(error) => {
            return OracleControl {
                name: name.into(),
                status: "rejected_invalid_cfb".into(),
                rejected: true,
                failure_reasons: vec![error.to_string()],
            };
        },
    };
    let oracle = oracle_for_output(
        args,
        source,
        expected,
        &output,
        source_bytes,
        expected_bytes,
        &output_bytes,
        replacement_paths,
        true,
    );
    OracleControl {
        name: name.into(),
        status: if oracle.oracle_ok {
            "unexpectedly_accepted".into()
        } else {
            "rejected".into()
        },
        rejected: !oracle.oracle_ok,
        failure_reasons: oracle.failure_reasons,
    }
}

fn oracle_controls(
    args: &Args,
    source: &Inventory,
    expected: &Inventory,
    source_bytes: &[u8],
    expected_bytes: &[u8],
    replacement_paths: &BTreeSet<Vec<String>>,
) -> Vec<OracleControl> {
    let control_args = control_args(args);
    let mut controls = Vec::new();
    let Some(missing_stream_path) = source
        .stream_bytes
        .iter()
        .find(|(_, bytes)| bytes.is_empty())
        .map(|(path, _)| path.clone())
        .or_else(|| source.stream_bytes.keys().next().cloned())
    else {
        return controls;
    };
    let missing_stream = (|| -> Result<Vec<u8>, BoxError> {
        let mut editor = Editor::open(
            expected_bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        editor.remove_stream(&missing_stream_path)?;
        Ok(editor.finish()?)
    })();
    controls.push(evaluate_oracle_control(
        "missing_stream",
        &control_args,
        source,
        expected,
        source_bytes,
        expected_bytes,
        missing_stream,
        replacement_paths,
    ));

    if let Some((path, bytes)) = source
        .stream_bytes
        .iter()
        .find(|(path, bytes)| !replacement_paths.contains(*path) && !bytes.is_empty())
    {
        let mut changed = bytes.clone();
        changed[0] ^= 1;
        let same_length = (|| -> Result<Vec<u8>, BoxError> {
            let mut editor = Editor::open(
                expected_bytes.to_vec(),
                Targets::default(),
                Limits::default(),
            )?;
            editor.put_stream_shared(path, Arc::from(changed.into_boxed_slice()))?;
            Ok(editor.finish()?)
        })();
        controls.push(evaluate_oracle_control(
            "untouched_same_length_stream_mutation",
            &control_args,
            source,
            expected,
            source_bytes,
            expected_bytes,
            same_length,
            replacement_paths,
        ));
    }

    let root_metadata = mutate_directory_byte(expected_bytes, expected, &[], 0x50);
    controls.push(evaluate_oracle_control(
        "root_clsid_mutation",
        &control_args,
        source,
        expected,
        source_bytes,
        expected_bytes,
        root_metadata,
        replacement_paths,
    ));

    if let Some(path) = expected.storage_clsids.keys().next() {
        let storage_metadata = mutate_directory_byte(expected_bytes, expected, path, 0x50);
        controls.push(evaluate_oracle_control(
            "storage_clsid_mutation",
            &control_args,
            source,
            expected,
            source_bytes,
            expected_bytes,
            storage_metadata,
            replacement_paths,
        ));
    }

    if let Some((path, _entry)) = expected.directory_entries.iter().find(|(path, entry)| {
        !path.is_empty() && entry.entry_type == litchi_cfb::consts::STGTY_STREAM
    }) {
        let state_mutation = mutate_directory_byte(expected_bytes, expected, path, 0x60);
        controls.push(evaluate_oracle_control(
            "stream_state_bits_mutation",
            &control_args,
            source,
            expected,
            source_bytes,
            expected_bytes,
            state_mutation,
            replacement_paths,
        ));
        let timestamp_mutation = mutate_directory_byte(expected_bytes, expected, path, 0x64);
        controls.push(evaluate_oracle_control(
            "stream_timestamp_mutation",
            &control_args,
            source,
            expected,
            source_bytes,
            expected_bytes,
            timestamp_mutation,
            replacement_paths,
        ));
    }

    controls.push(evaluate_oracle_control(
        "same_length_source_swap",
        &control_args,
        source,
        expected,
        source_bytes,
        expected_bytes,
        Ok(source_bytes.to_vec()),
        replacement_paths,
    ));

    match args.case {
        Case::DocFloat | Case::DocNoHf => {
            let wrong_args = Args {
                text: format!("{} wrong-target-control", args.text),
                ..control_args.clone()
            };
            controls.push(evaluate_oracle_control(
                "wrong_doc_target_text",
                &control_args,
                source,
                expected,
                source_bytes,
                expected_bytes,
                public_format_edit(source_bytes, &wrong_args),
                replacement_paths,
            ));
        },
        Case::Ppt45543 => {
            let wrong_slide = (|| -> Result<Vec<u8>, BoxError> {
                let snapshot =
                    litchi_ppt::slide_order::Snapshot::from_bytes(source_bytes.to_vec())?;
                let mut edit = snapshot.edit()?;
                edit.remove_slide(Position::new(0))?;
                Ok(edit.commit()?.snapshot().bytes().to_vec())
            })();
            controls.push(evaluate_oracle_control(
                "wrong_ppt_slide_identity",
                &control_args,
                source,
                expected,
                source_bytes,
                expected_bytes,
                wrong_slide,
                replacement_paths,
            ));

            if let Some((path, bytes)) = expected.stream_bytes.iter().find(|(path, bytes)| {
                path.last()
                    .is_some_and(|name| name.eq_ignore_ascii_case("PowerPoint Document"))
                    && !bytes.is_empty()
            }) {
                let mut changed = bytes.clone();
                changed[0] ^= 1;
                let survivor_payload = (|| -> Result<Vec<u8>, BoxError> {
                    let mut editor = Editor::open(
                        expected_bytes.to_vec(),
                        Targets::default(),
                        Limits::default(),
                    )?;
                    editor.put_stream_shared(path, Arc::from(changed.into_boxed_slice()))?;
                    Ok(editor.finish()?)
                })();
                controls.push(evaluate_oracle_control(
                    "ppt_survivor_payload_mutation",
                    &control_args,
                    source,
                    expected,
                    source_bytes,
                    expected_bytes,
                    survivor_payload,
                    replacement_paths,
                ));
            }
        },
    }
    controls
}

pub fn run(timing_claim: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    if args.operation == Operation::Format && args.policy == Policy::Rewrite {
        return failure(
            "--policy is only meaningful for --operation container; public format uses its default route",
        );
    }
    let source_bytes = std::fs::read(&args.input)?;
    let source = inventory(&source_bytes)?;

    // The expected stream bytes are produced by one real public format edit,
    // outside every timed or allocator-counted interval.  This also makes the
    // common-container operation a direct exact-replacement control.
    let expected_bytes = public_format_edit(&source_bytes, &args)?;
    let expected = inventory(&expected_bytes)?;
    let replacements = derive_replacements(&source, &expected, args.case)?;
    let replacement_paths = replacements
        .iter()
        .map(|replacement| replacement.path.clone())
        .collect::<BTreeSet<_>>();
    let mut length_proof = changed_length_proof(&source, &expected);
    let replacements_json = replacement_summaries(&source, &expected, &replacements);
    let expected_oracle = oracle_for_output(
        &args,
        &source,
        &expected,
        &expected,
        &source_bytes,
        &expected_bytes,
        &expected_bytes,
        &replacement_paths,
        false,
    );
    finalize_length_proof(args.case, &expected_oracle, &mut length_proof);
    let oracle_controls = oracle_controls(
        &args,
        &source,
        &expected,
        &source_bytes,
        &expected_bytes,
        &replacement_paths,
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
                        &replacement_paths,
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
                        &replacement_paths,
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
                        &replacement_paths,
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
                        &replacement_paths,
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
            "common_container_open_replace_and_validate_finish_control".into()
        },
        phase_contract: if args.operation == Operation::Format {
            "whole_ns is the public format owner open_edit_commit lifecycle; probe oracle validation is outside the timed interval".into()
        } else {
            "open_ns=Editor::open; stage_ns=put_streams_shared replace_and_validate; finish_ns=Editor::finish; whole_ns spans the same editor lifecycle".into()
        },
        input: args.input.display().to_string(),
        policy: args.policy.name().into(),
        policy_applied,
        policy_application_scope: if policy_applied {
            "common_container_editor".into()
        } else {
            "not_applied_public_format_route".into()
        },
        policy_argument_effect: if policy_applied {
            "applied_to_common_container_editor".into()
        } else {
            "ignored_public_format_default_route".into()
        },
        policy_contract: "Reuse preserves the normalized raw directory image outside planner-owned allocation fields; Rewrite may normalize physical directory layout and timestamps. Both policies must preserve logical streams, semantic directory metadata, root/storage CLSIDs, and public edit meaning; raw differences are reported in raw_directory and directory_metadata_differences".into(),
        timing_claim,
        allocator_instrumented: alloc_metrics::instrumented(),
        allocation_ownership_contract: "Each allocation region reports boundary-relative allocations. The opened Editor is retained from open through replace_and_validate; the staged Editor is retained until finish; the returned Vec is retained across finish. Format output is also retained outside its region for validation. retained_bytes is live ownership at the region boundary, not RSS.".into(),
        directory_metadata_fields: vec![
            "entry_type".into(),
            "name_utf16".into(),
            "clsid".into(),
            "bytes".into(),
            "start_sector".into(),
            "is_minifat".into(),
            "raw_left_sibling".into(),
            "raw_right_sibling".into(),
            "raw_child".into(),
            "raw_color".into(),
            "raw_state_bits".into(),
            "raw_creation_time".into(),
            "raw_modification_time".into(),
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
        oracle_controls,
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
                directory_entries: Vec::new(),
                raw_directory_image_bytes: 0,
                raw_directory_image_sha256: String::new(),
            },
            stream_bytes,
            storage_clsids: BTreeMap::new(),
            directory_entries: BTreeMap::new(),
            raw_directory_image: RawDirectoryImage {
                bytes: Vec::new(),
                sector_size: 512,
                cfb_version: 3,
            },
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
                live_record_bytes: 0,
                live_record_sha256: String::new(),
                live_record: Vec::new(),
            },
            PptSlideIdentity {
                slide_id: 20,
                persist_id: 200,
                flags: 0,
                text_count: 2,
                live_record_bytes: 0,
                live_record_sha256: String::new(),
                live_record: Vec::new(),
            },
            PptSlideIdentity {
                slide_id: 30,
                persist_id: 300,
                flags: 0,
                text_count: 3,
                live_record_bytes: 0,
                live_record_sha256: String::new(),
                live_record: Vec::new(),
            },
        ];
        let removed_wrong_slide = vec![source[0].clone(), source[1].clone()];
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
