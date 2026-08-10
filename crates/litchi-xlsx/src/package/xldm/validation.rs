//! MS-XLDM structural and cross-file validation rules.

use super::codec::{checked_usize, crc32, invalid, limit};
use super::model::{
    BOM, BackupLog, CRC_SIZE, FileEntry, FileGroupClass, GeneratedNameKind, MAX_PATH_BYTES,
};
use crate::error::Result;
use std::collections::{HashMap, HashSet};

pub(super) fn validate_backup_log(
    log: &BackupLog,
    directory: &[FileEntry],
    partitions: usize,
    backup: usize,
) -> Result<()> {
    let mut expected = HashMap::<String, &FileEntry>::new();
    for (index, entry) in directory.iter().enumerate() {
        if index != partitions && index != backup {
            expected.insert(entry.path.clone(), entry);
        }
    }
    let mut seen = HashSet::new();
    for group in &log.file_groups {
        let persist = if group.persist_location_path.is_empty() {
            None
        } else {
            Some(normalize_generated_folder(&group.persist_location_path)?)
        };
        for file in &group.files {
            if !seen.insert(file.storage_path.clone()) {
                return Err(invalid(format!(
                    "duplicate backup-log StoragePath '{}'",
                    file.storage_path
                )));
            }
            let entry = expected.get(&file.storage_path).ok_or_else(|| {
                invalid(format!(
                    "backup-log StoragePath '{}' is absent from the virtual directory",
                    file.storage_path
                ))
            })?;
            if entry.stored_size.0.checked_sub(CRC_SIZE as u64) != Some(u64::from(file.size)) {
                return Err(invalid(format!(
                    "backup-log size mismatch for '{}'",
                    file.storage_path
                )));
            }
            if entry.last_write_timestamp != file.last_write_timestamp {
                return Err(invalid(format!(
                    "backup-log timestamp mismatch for '{}'",
                    file.storage_path
                )));
            }
            if let Some(persist) = &persist {
                let parent = file
                    .generated
                    .normalized_path
                    .rsplit_once('/')
                    .map_or("", |value| value.0);
                if parent != persist && !parent.starts_with(&format!("{persist}/")) {
                    return Err(invalid(format!(
                        "StoragePath '{}' is outside PersistLocationPath",
                        file.storage_path
                    )));
                }
            }
        }
    }
    if seen.len() != expected.len() {
        return Err(invalid(
            "backup log does not enumerate every non-marker virtual-directory file",
        ));
    }
    Ok(())
}

pub(super) fn normalize_generated_path(path: &str) -> Result<String> {
    if path.is_empty()
        || path.len() > MAX_PATH_BYTES
        || path.starts_with('/')
        || path.starts_with('\\')
        || path.as_bytes().get(1) == Some(&b':')
    {
        return Err(invalid(
            "generated storage path is empty, oversized, or rooted",
        ));
    }
    let segments: Vec<_> = path.split(['/', '\\']).collect();
    if segments.iter().any(|segment| {
        segment.is_empty() || *segment == "." || *segment == ".." || !valid_generated_token(segment)
    }) {
        return Err(invalid(format!(
            "generated storage path '{path}' violates section 2.2 characters"
        )));
    }
    Ok(segments.join("/"))
}
fn normalize_generated_folder(path: &str) -> Result<String> {
    let normalized = normalize_generated_path(&format!("{path}/placeholder.0.db.xml"))?;
    Ok(normalized
        .rsplit_once('/')
        .map_or("", |value| value.0)
        .to_owned())
}
fn valid_generated_token(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || (b'#'..=b'.').contains(&byte)
                || matches!(
                    byte,
                    b'!' | b'=' | b'@' | b'[' | b']' | b'^' | b'{' | b'}' | b'~' | b'_'
                )
        })
}
fn valid_id(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || (b'#'..=b'.').contains(&byte)
                || matches!(
                    byte,
                    b'!' | b'=' | b'@' | b'[' | b']' | b'^' | b'{' | b'}' | b'~'
                )
        })
}
fn digits(value: &str) -> bool {
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
}
fn folder(value: &str, suffix: &str, zero: bool) -> bool {
    let Some(prefix) = value.strip_suffix(suffix) else {
        return false;
    };
    let Some((id, version)) = prefix.rsplit_once('.') else {
        return false;
    };
    valid_id(id) && digits(version) && (!zero || version == "0")
}
pub(super) fn validate_generated_hierarchy(parents: &[&str]) -> Result<()> {
    if parents.is_empty() {
        return Ok(());
    }
    if parents.len() > 4 || !folder(parents[0], ".db", false) {
        return Err(invalid(
            "generated path must begin in a section 2.2 database folder",
        ));
    }
    if parents.len() == 1 {
        return Ok(());
    }
    if folder(parents[1], ".cub", true) {
        if parents.len() >= 3 && !folder(parents[2], ".det", false) {
            return Err(invalid("cube child folder must be a measure-group folder"));
        }
        if parents.len() == 4 && !folder(parents[3], ".prt", false) {
            return Err(invalid(
                "measure-group child folder must be a partition folder",
            ));
        }
    } else if folder(parents[1], ".dim", true) || folder(parents[1], ".ds", true) {
        if parents.len() != 2 {
            return Err(invalid(
                "dimension and data-source folders cannot contain generated subfolders",
            ));
        }
    } else {
        return Err(invalid("unknown generated database child folder"));
    }
    Ok(())
}

pub(super) fn classify_generated_name(
    name: &str,
    parent: Option<&str>,
) -> Result<GeneratedNameKind> {
    if name == "MdxScript.0.scr.xml" {
        return Ok(GeneratedNameKind::MdxScriptMetadata);
    }
    if let Some(prefix) = name.strip_suffix(".db.xml")
        && versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DatabaseDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".dsv.xml")
        && versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceViewDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".cub.xml")
        && versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::CubeDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".ds.xml")
        && versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceOrDimensionDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".dim.xml")
        && versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceOrDimensionDefinition);
    }
    if let Some(prefix) = name
        .strip_prefix("info.")
        .and_then(|value| value.strip_suffix(".xml"))
        && digits(prefix)
    {
        return match parent {
            Some(value) if folder(value, ".cub", true) => Ok(GeneratedNameKind::CubeInformation),
            Some(value) if folder(value, ".prt", false) => {
                Ok(GeneratedNameKind::PartitionInformation)
            },
            Some(value) if folder(value, ".dim", true) => Ok(GeneratedNameKind::TableInformation),
            _ => Err(invalid(
                "info file is outside a cube, partition, or dimension folder",
            )),
        };
    }
    for (suffix, kind) in [
        (".det.xml", GeneratedNameKind::MeasureGroupMetadata),
        (".prt.xml", GeneratedNameKind::PartitionMetadata),
    ] {
        if let Some(prefix) = name.strip_suffix(suffix)
            && versioned_id(prefix)
        {
            return Ok(kind);
        }
    }
    if let Some(prefix) = name.strip_suffix(".tbl.xml") {
        let Some((stem, version)) = prefix.rsplit_once('.') else {
            return Err(invalid("invalid tbl.xml generated name"));
        };
        if !digits(version) {
            return Err(invalid("invalid tbl.xml version"));
        }
        if let Some(value) = stem.strip_prefix("R$") {
            if dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::TableRelationshipMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("H$") {
            if dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::ColumnHierarchyMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("U$") {
            if dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::UserHierarchyMetadata);
            }
        } else if valid_id(stem) {
            return Ok(GeneratedNameKind::TableMetadata);
        }
    }
    if let Some(prefix) = name.strip_suffix(".dictionary") {
        let values: Vec<_> = prefix.split('.').collect();
        if values.len() >= 3 && digits(values[0]) && values[1..].iter().all(|value| valid_id(value))
        {
            return Ok(GeneratedNameKind::ColumnDictionary);
        }
    }
    if let Some(prefix) = name.strip_suffix(".hidx") {
        let Some((version, rest)) = prefix.split_once('.') else {
            return Err(invalid("invalid hidx name"));
        };
        if digits(version) && rest.starts_with("H$") && dollar_ids(&rest[2..], 2) {
            return Ok(GeneratedNameKind::ColumnHashIndex);
        }
    }
    if let Some(prefix) = name.strip_suffix(".idf") {
        return classify_idf(prefix);
    }
    Err(invalid(format!(
        "unrecognized section 2.2 generated file name '{name}'"
    )))
}
fn versioned_id(value: &str) -> bool {
    value
        .rsplit_once('.')
        .is_some_and(|(id, version)| valid_id(id) && digits(version))
}
fn dollar_ids(value: &str, minimum: usize) -> bool {
    let values: Vec<_> = value.split('$').collect();
    values.len() >= minimum && values.iter().all(|value| valid_id(value))
}
fn classify_idf(value: &str) -> Result<GeneratedNameKind> {
    let Some((version, rest)) = value.split_once('.') else {
        return Err(invalid("invalid idf generated name"));
    };
    if !digits(version) {
        return Err(invalid("invalid idf version"));
    }
    if let Some(rest) = rest.strip_prefix("R$")
        && rest.ends_with(".INDEX.0")
        && dollar_ids(rest.trim_end_matches(".INDEX.0"), 2)
    {
        return Ok(GeneratedNameKind::TableRelationshipIndex);
    }
    if let Some(rest) = rest.strip_prefix("H$") {
        for (suffix, kind) in [
            (".POS_TO_ID.0", GeneratedNameKind::ColumnPositionToId),
            (".ID_TO_POS.0", GeneratedNameKind::ColumnIdToPosition),
        ] {
            if let Some(ids) = rest.strip_suffix(suffix)
                && dollar_ids(ids, 2)
            {
                return Ok(kind);
            }
        }
    }
    if let Some(rest) = rest.strip_prefix("U$") {
        for (suffix, kind) in [
            (".CHILD_COUNT.0", GeneratedNameKind::UserHierarchyChildCount),
            (
                ".FIRST_CHILD_POS.0",
                GeneratedNameKind::UserHierarchyFirstChildPosition,
            ),
            (
                ".PARENT_POS.0",
                GeneratedNameKind::UserHierarchyParentPosition,
            ),
            (
                ".MULTI_LEVEL_ID.0",
                GeneratedNameKind::UserHierarchyMultilevelId,
            ),
        ] {
            if let Some(ids) = rest.strip_suffix(suffix)
                && dollar_ids(ids, 2)
            {
                return Ok(kind);
            }
        }
    }
    let values: Vec<_> = rest.split('.').collect();
    if values.len() >= 3
        && values.last() == Some(&"0")
        && values[..values.len() - 1]
            .iter()
            .all(|value| valid_id(value))
    {
        return Ok(GeneratedNameKind::ColumnData);
    }
    Err(invalid("unrecognized section 2.2 idf generated name"))
}
pub(super) fn validate_kind_location(kind: GeneratedNameKind, parents: &[&str]) -> Result<()> {
    let valid = match kind {
        GeneratedNameKind::DatabaseDefinition => parents.is_empty(),
        GeneratedNameKind::DataSourceViewDefinition
        | GeneratedNameKind::CubeDefinition
        | GeneratedNameKind::DataSourceOrDimensionDefinition => parents.len() == 1,
        GeneratedNameKind::CubeInformation
        | GeneratedNameKind::MdxScriptMetadata
        | GeneratedNameKind::MeasureGroupMetadata => {
            parents.len() == 2 && folder(parents[1], ".cub", true)
        },
        GeneratedNameKind::PartitionMetadata => {
            parents.len() == 3 && folder(parents[2], ".det", false)
        },
        GeneratedNameKind::PartitionInformation => {
            parents.len() == 4 && folder(parents[3], ".prt", false)
        },
        GeneratedNameKind::TableInformation
        | GeneratedNameKind::TableMetadata
        | GeneratedNameKind::TableRelationshipMetadata
        | GeneratedNameKind::ColumnHierarchyMetadata
        | GeneratedNameKind::UserHierarchyMetadata
        | GeneratedNameKind::ColumnData
        | GeneratedNameKind::TableRelationshipIndex
        | GeneratedNameKind::ColumnPositionToId
        | GeneratedNameKind::ColumnIdToPosition
        | GeneratedNameKind::ColumnHashIndex
        | GeneratedNameKind::ColumnDictionary
        | GeneratedNameKind::UserHierarchyChildCount
        | GeneratedNameKind::UserHierarchyFirstChildPosition
        | GeneratedNameKind::UserHierarchyParentPosition
        | GeneratedNameKind::UserHierarchyMultilevelId => {
            parents.len() == 2 && folder(parents[1], ".dim", true)
        },
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "generated file appears outside its section 2.2 folder",
        ))
    }
}
pub(super) fn kind_allowed_for_group(kind: GeneratedNameKind, class: FileGroupClass) -> bool {
    match class {
        FileGroupClass::Database => kind == GeneratedNameKind::DatabaseDefinition,
        FileGroupClass::DataSource => kind == GeneratedNameKind::DataSourceOrDimensionDefinition,
        FileGroupClass::DataSourceView => kind == GeneratedNameKind::DataSourceViewDefinition,
        FileGroupClass::Cube => matches!(
            kind,
            GeneratedNameKind::CubeDefinition | GeneratedNameKind::CubeInformation
        ),
        FileGroupClass::MdxScript => kind == GeneratedNameKind::MdxScriptMetadata,
        FileGroupClass::MeasureGroup => kind == GeneratedNameKind::MeasureGroupMetadata,
        FileGroupClass::Partition => matches!(
            kind,
            GeneratedNameKind::PartitionMetadata | GeneratedNameKind::PartitionInformation
        ),
        FileGroupClass::Dimension => !matches!(
            kind,
            GeneratedNameKind::DatabaseDefinition
                | GeneratedNameKind::DataSourceViewDefinition
                | GeneratedNameKind::CubeDefinition
                | GeneratedNameKind::CubeInformation
                | GeneratedNameKind::MdxScriptMetadata
                | GeneratedNameKind::MeasureGroupMetadata
                | GeneratedNameKind::PartitionMetadata
                | GeneratedNameKind::PartitionInformation
        ),
    }
}

pub(super) fn validate_paths(files: &[FileEntry]) -> Result<()> {
    let mut normalized = HashSet::new();
    for file in files {
        if file.path.is_empty()
            || file.path.len() > MAX_PATH_BYTES
            || file
                .path
                .bytes()
                .any(|byte| byte == 0 || byte.is_ascii_control())
        {
            return Err(invalid(
                "virtual-directory path is empty, oversized, or contains control bytes",
            ));
        }
        if file.path.starts_with('/')
            || file.path.starts_with('\\')
            || file.path.as_bytes().get(1) == Some(&b':')
        {
            return Err(invalid(format!(
                "rooted virtual-directory path '{}' is rejected",
                file.path
            )));
        }
        let segments: Vec<_> = file.path.split(['/', '\\']).collect();
        if segments
            .iter()
            .any(|segment| segment.is_empty() || *segment == "." || *segment == "..")
        {
            return Err(invalid(format!(
                "virtual-directory path '{}' contains traversal or empty segments",
                file.path
            )));
        }
        let path = segments.join("/").to_lowercase();
        if !normalized.insert(path) {
            return Err(invalid(format!(
                "duplicate case-insensitive virtual-directory path '{}'",
                file.path
            )));
        }
    }
    for path in &normalized {
        let segments: Vec<_> = path.split('/').collect();
        let mut prefix = String::new();
        for segment in segments.iter().take(segments.len().saturating_sub(1)) {
            if !prefix.is_empty() {
                prefix.push('/');
            }
            prefix.push_str(segment);
            if normalized.contains(&prefix) {
                return Err(invalid(format!(
                    "virtual-directory file '{prefix}' is also used as an ancestor directory"
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn validate_allocations(
    bytes: &[u8],
    data_offset: usize,
    directory_offset: usize,
    files: &mut [FileEntry],
    order: &[usize],
) -> Result<()> {
    let first_start = checked_usize(files[order[0]].offset.0, "file offset")?;
    if first_start != data_offset + BOM.len() {
        return Err(invalid(
            "partition allocation must immediately follow the files-section byte-order mark",
        ));
    }
    let mut previous_end = first_start;
    let last_position = order.len() - 1;
    let mut total = 0usize;
    for (position, index) in order.iter().enumerate() {
        let entry = &mut files[*index];
        let start = checked_usize(entry.offset.0, "file offset")?;
        let size = checked_usize(entry.stored_size.0, "file size")?;
        if size < CRC_SIZE {
            return Err(invalid(format!(
                "allocation '{}' is smaller than its CRC marker",
                entry.path
            )));
        }
        let end = start
            .checked_add(size)
            .ok_or_else(|| limit("allocation range"))?;
        if start < data_offset + BOM.len() || end > directory_offset {
            return Err(invalid(format!(
                "allocation '{}' crosses the files-section boundary",
                entry.path
            )));
        }
        if position > 0 {
            let gap = &bytes[previous_end..start];
            if position == last_position {
                if gap != BOM {
                    return Err(invalid(
                        "backup log must be preceded by exactly one byte-order mark",
                    ));
                }
            } else if !gap.is_empty() {
                return Err(invalid(
                    "unexpected gap or overlap between serial MS-XLDM allocations",
                ));
            }
        }
        let payload_end = end - CRC_SIZE;
        let expected = crc32(&bytes[start..payload_end]);
        let marker =
            u32::from_le_bytes(bytes[payload_end..end].try_into().unwrap_or_else(|error| {
                crate::error::panic_error_invariant(
                    "operation was checked before extraction",
                    error,
                )
            }));
        if marker != expected {
            return Err(invalid(format!(
                "CRC mismatch for allocation '{}'",
                entry.path
            )));
        }
        entry.crc32 = marker;
        total = total
            .checked_add(size)
            .ok_or_else(|| limit("allocated bytes"))?;
        previous_end = end;
    }
    if bytes[previous_end..directory_offset]
        .iter()
        .any(|byte| *byte != 0)
    {
        return Err(invalid("files-section page padding must be zero"));
    }
    if total > directory_offset - data_offset {
        return Err(limit("allocated bytes"));
    }
    Ok(())
}
