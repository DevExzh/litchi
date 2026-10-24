//! DOC preservation oracle reused by the change-0917 retention probe.
//!
//! This module keeps the qualified 0730 DOC checks that are meaningful for the
//! current repeated-edit probe: complete CFB stream inventories, directory
//! metadata and normalized directory images, public DOC projections, and
//! survivor collections.  Timing, allocator instrumentation, PPT support, and
//! the old 0730 executable are intentionally outside this module.

use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{DirectoryEntry, OleFile, validate_source};
use litchi_core::{OwnedSource, Position};
use serde::Serialize;
use sha2::{Digest, Sha256};

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

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
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
    let mut ole = OleFile::open(
        Cursor::new(bytes.to_vec()),
        litchi_cfb::OleFileLimits::default(),
    )?;
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
    validate_source(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        litchi_cfb::SharedOleFileLimits::default(),
        litchi_core::ValidationLimits::default(),
    )
    .map(|report| report.is_complete() && !report.has_errors())
    .unwrap_or(false)
}

#[derive(Clone, Debug, Serialize)]
struct DirectoryMetadataDifference {
    path: Vec<String>,
    field: String,
    expected: String,
    actual: String,
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
    source: &Inventory,
    expected: &Inventory,
    output: &Inventory,
) -> (RawDirectoryOracle, bool) {
    let (source_expected_normalized, source_expected_difference_bytes) =
        normalized_directory_comparison(source, expected);
    let (expected_output_normalized, expected_output_difference_bytes) =
        normalized_directory_comparison(expected, output);
    let (source_output_normalized, source_output_difference_bytes) =
        normalized_directory_comparison(source, output);
    let source_expected_ok = source_expected_normalized == "match_after_allocation_normalization";
    let source_output_ok = source_output_normalized == "match_after_allocation_normalization";
    let policy_gate = if !source_expected_ok {
        "failed_source_expected_model"
    } else if !source_output_ok {
        "failed_public_source_model"
    } else {
        "passed_source_expected_model"
    };
    (
        RawDirectoryOracle {
            mode: "public_format_source_model_gate".into(),
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
            policy_gate: policy_gate.into(),
        },
        source_expected_ok && source_output_ok,
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

fn debug_digest<T: std::fmt::Debug>(value: &T) -> String {
    sha256_hex(format!("{value:?}").as_bytes())
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
    source: &litchi_doc::Package,
    output: &litchi_doc::Package,
    replacement: &str,
) -> (bool, String, DocSemanticWitness) {
    let mut witness = DocSemanticWitness {
        target_position: 0,
        requested_text_sha256: sha256_hex(replacement.as_bytes()),
        requested_text_utf16_units: replacement.encode_utf16().count(),
        comparison_basis: "direct_text_and_collection_values_with_sha256_report_fields".into(),
        ..DocSemanticWitness::default()
    };
    let source_paragraphs = match source.body_paragraphs(litchi_doc::body_text::Projection::All) {
        Ok(value) => value,
        Err(error) => {
            let message = format!("DOC source paragraph read failed: {error}");
            witness.error = Some(message.clone());
            return (false, message, witness);
        },
    };
    let output_paragraphs = match output.body_paragraphs(litchi_doc::body_text::Projection::All) {
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
            .body_paragraphs(litchi_doc::body_text::Projection::Accepted)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        output
            .body_paragraphs(litchi_doc::body_text::Projection::Accepted)
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
            .body_paragraphs(litchi_doc::body_text::Projection::Rejected)
            .map(|paragraphs| {
                paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text().to_string())
                    .collect()
            })
            .map_err(|error| error.to_string()),
        output
            .body_paragraphs(litchi_doc::body_text::Projection::Rejected)
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

    let (source_state, source_items) = doc_collection_state(
        source
            .tracked_revisions()
            .map_err(|error| error.to_string()),
    );
    let (output_state, output_items) = doc_collection_state(
        output
            .tracked_revisions()
            .map_err(|error| error.to_string()),
    );
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
        .map_err(|error| error.to_string())
        .and_then(|objects| objects.inventory().map_err(|error| error.to_string()))
        .map(|value| (value.len(), value))
        .map_err(|error| error.to_string());
    let output_objects = output
        .embedded_objects()
        .map_err(|error| error.to_string())
        .and_then(|objects| objects.inventory().map_err(|error| error.to_string()))
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
enum SemanticWitness {
    Doc(Box<DocSemanticWitness>),
}

fn doc_witness(value: DocSemanticWitness) -> SemanticWitness {
    SemanticWitness::Doc(Box::new(value))
}

fn allowed_changed_stream(path: &[String]) -> bool {
    let Some(name) = path.last() else {
        return false;
    };
    name.eq_ignore_ascii_case("WordDocument")
        || name.eq_ignore_ascii_case("0Table")
        || name.eq_ignore_ascii_case("1Table")
}

fn changed_stream_paths_allowed(source: &Inventory, expected: &Inventory) -> bool {
    source
        .stream_bytes
        .keys()
        .filter(|path| source.stream_bytes.get(*path) != expected.stream_bytes.get(*path))
        .all(|path| allowed_changed_stream(path))
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
    source.stream_bytes.iter().all(|(path, before)| {
        replacement_paths.contains(path) || expected.stream_bytes.get(path) == Some(before)
    })
}

fn derive_replacement_paths(
    source: &Inventory,
    expected: &Inventory,
) -> Result<BTreeSet<Vec<String>>, BoxError> {
    let source_paths = source.stream_bytes.keys().collect::<BTreeSet<_>>();
    let expected_paths = expected.stream_bytes.keys().collect::<BTreeSet<_>>();
    if source_paths != expected_paths {
        return failure(
            "public edit added or deleted CFB streams; no explicit stream mapping is available",
        );
    }
    if !changed_stream_paths_allowed(source, expected) {
        return failure(
            "public edit changed an unapproved stream; explicit stream mapping is required",
        );
    }
    Ok(source
        .stream_bytes
        .keys()
        .filter(|path| source.stream_bytes.get(*path) != expected.stream_bytes.get(*path))
        .cloned()
        .collect())
}

fn semantic_reopen_check(
    source_bytes: &[u8],
    output_bytes: &[u8],
    replacement: &str,
) -> (bool, String, SemanticWitness) {
    let source = match litchi_doc::Package::from_bytes(source_bytes) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            let message = format!("DOC source public reopen failed: {error}");
            return (
                false,
                message.clone(),
                doc_witness(DocSemanticWitness {
                    error: Some(message),
                    ..DocSemanticWitness::default()
                }),
            );
        },
    };
    let output = match litchi_doc::Package::from_bytes(output_bytes) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            let message = format!("DOC output public reopen failed: {error}");
            return (
                false,
                message.clone(),
                doc_witness(DocSemanticWitness {
                    error: Some(message),
                    ..DocSemanticWitness::default()
                }),
            );
        },
    };
    let (ok, failure, witness) = doc_semantic_witness(&source, &output, replacement);
    (ok, failure, doc_witness(witness))
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

fn oracle_for_output(
    source: &Inventory,
    expected: &Inventory,
    output: &Inventory,
    source_bytes: &[u8],
    expected_bytes: &[u8],
    output_bytes: &[u8],
    replacement_paths: &BTreeSet<Vec<String>>,
    replacement: &str,
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
    let actual_changed_stream_paths_allowed = changed_stream_paths_allowed(source, output);
    let changed_stream_paths_allowed = changed_stream_paths_allowed(source, expected);
    let (root_clsid_preserved, storage_clsids_preserved) =
        preserved_clsids(source, expected, output);
    let (semantic_reopen_ok, semantic_failure, semantic_witness) =
        semantic_reopen_check(source_bytes, output_bytes, replacement);
    let source_directory_metadata_differences = directory_metadata_differences(source, expected);
    let directory_metadata_differences = directory_metadata_differences(expected, output);
    let semantic_directory_metadata_matches_expected =
        semantic_directory_metadata_matches(expected, output);
    let source_directory_shape_matches_expected = directory_shape_matches(source, expected);
    let (raw_directory, raw_directory_policy_ok) = raw_directory_oracle(source, expected, output);
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

#[derive(Clone, Debug, Serialize)]
pub struct Check {
    pub oracle_ok: bool,
    pub source_structurally_valid: bool,
    pub expected_structurally_valid: bool,
    pub output_structurally_valid: bool,
    pub semantic_reopen_ok: bool,
    pub semantic_directory_metadata_matches_expected: bool,
    pub unchanged_stream_bytes_match_expected: bool,
    pub unchanged_source_output_stream_bytes_match: bool,
    pub output_stream_paths_match_expected: bool,
    pub output_stream_bytes_match_expected: bool,
    pub root_clsid_preserved: bool,
    pub storage_clsids_preserved: bool,
    pub failure_reasons: Vec<String>,
}

/// Applies the focused 0730 DOC oracle to one repeated-edit output.
///
/// `expected_bytes` is the zero-retention output for the same fixture and
/// edit count.  The oracle compares complete stream bytes, directory metadata,
/// normalized directory images, and public DOC survivor projections.
pub fn check_doc_output(
    source_bytes: &[u8],
    expected_bytes: &[u8],
    output_bytes: &[u8],
    replacement: &str,
) -> Result<Check, BoxError> {
    let source = inventory(source_bytes)?;
    let expected = inventory(expected_bytes)?;
    let output = inventory(output_bytes)?;
    let replacement_paths = derive_replacement_paths(&source, &expected)?;
    let oracle = oracle_for_output(
        &source,
        &expected,
        &output,
        source_bytes,
        expected_bytes,
        output_bytes,
        &replacement_paths,
        replacement,
    );
    Ok(Check {
        oracle_ok: oracle.oracle_ok,
        source_structurally_valid: oracle.source_structurally_valid,
        expected_structurally_valid: oracle.expected_structurally_valid,
        output_structurally_valid: oracle.output_structurally_valid,
        semantic_reopen_ok: oracle.semantic_reopen_ok,
        semantic_directory_metadata_matches_expected: oracle
            .semantic_directory_metadata_matches_expected,
        unchanged_stream_bytes_match_expected: oracle.unchanged_stream_bytes_match_expected,
        unchanged_source_output_stream_bytes_match: oracle
            .unchanged_source_output_stream_bytes_match,
        output_stream_paths_match_expected: oracle.output_stream_paths_match_expected,
        output_stream_bytes_match_expected: oracle.output_stream_bytes_match_expected,
        root_clsid_preserved: oracle.root_clsid_preserved,
        storage_clsids_preserved: oracle.storage_clsids_preserved,
        failure_reasons: oracle.failure_reasons,
    })
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
            &[("WordDocument", b"document"), ("0Table", b"table")],
            "root",
        );
        let actual = fake_inventory(&[("WordDocument", b"document")], "root");
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
    fn oracle_control_rejects_metadata_change() {
        let source = fake_inventory(&[("WordDocument", b"before")], "root-a");
        let expected = fake_inventory(&[("WordDocument", b"after")], "root-a");
        let output = fake_inventory(&[("WordDocument", b"after")], "root-b");
        let (root_preserved, storage_preserved) = preserved_clsids(&source, &expected, &output);
        assert!(!root_preserved);
        assert!(storage_preserved);
    }
}
