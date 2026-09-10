//! Native version-150 tabular path and backup-log rules.
//!
//! The canonical XLDM profile stores generated names directly in
//! `StoragePath`.  Native tabular stores use an opaque hexadecimal storage key
//! there and keep the logical Windows path in `Path`.  This module owns that
//! profile-specific projection so the canonical section-2.2 classifier remains
//! strict and unchanged.

use super::codec::{invalid, limit};
use super::compression::{CodecError, CodecLimits, XpressFraming, xpress_stream_size};
use super::model::{
    BackupLog, CRC_SIZE, FileEntry, FileGroupClass, GeneratedNameKind, GeneratedPath,
    MAX_PATH_BYTES, MAX_STORAGE_BYTES,
};
use crate::error::Result;
use std::collections::{HashMap, HashSet};

/// Classify a native `Path` after anchoring it to the backup-log `ServerRoot`.
///
/// The relative path is retained in its normalized slash form.  Native object
/// and column identifiers may contain spaces and Unicode, so this classifier
/// deliberately uses path-safe component checks instead of the narrower ASCII
/// identifier grammar used by the canonical profile.
pub(super) fn classify_tabular_source_path(
    source_path: &str,
    server_root: &str,
) -> Result<GeneratedPath> {
    let normalized_path = relative_native_path(source_path, server_root, "source path")?;
    let segments: Vec<_> = normalized_path.split('/').collect();
    let file = *segments
        .last()
        .ok_or_else(|| invalid("native source path has no file name"))?;
    let parents = &segments[..segments.len() - 1];
    validate_native_hierarchy(parents)?;
    let kind = classify_native_name(file, parents.last().copied())?;
    validate_native_kind_location(kind, parents)?;
    Ok(GeneratedPath {
        normalized_path,
        kind,
    })
}

/// Validate the native tabular backup log against its opaque directory keys.
///
/// The native `Size` value is the decoded member length.  We preflight the
/// explicit 16-bit Xpress framing to prove that length without materializing
/// any member.  Native timestamps use a different domain from the virtual
/// directory and are therefore retained independently.
pub(super) fn validate_backup_log(
    log: &BackupLog,
    directory: &[FileEntry],
    partitions: usize,
    backup: usize,
    bytes: &[u8],
) -> Result<()> {
    validate_native_root(&log.server_root)?;
    let mut expected = HashMap::<&str, &FileEntry>::new();
    for (index, entry) in directory.iter().enumerate() {
        if index != partitions && index != backup {
            expected.insert(entry.path.as_str(), entry);
        }
    }

    let mut seen = HashSet::new();
    let mut seen_generated = HashSet::new();
    let mut decoded_total = 0usize;
    for group in &log.file_groups {
        let persist = relative_native_path(
            &group.persist_location_path,
            &log.server_root,
            "persist location path",
        )?;
        validate_native_persist_location(&persist, group.class)?;
        for file in &group.files {
            validate_storage_key(&file.storage_path)?;
            if !seen.insert(file.storage_path.as_str()) {
                return Err(invalid(format!(
                    "duplicate native backup-log StoragePath '{}'",
                    file.storage_path
                )));
            }
            if !seen_generated.insert(file.generated.normalized_path.as_str()) {
                return Err(invalid(format!(
                    "duplicate native logical source path '{}'",
                    file.generated.normalized_path
                )));
            }
            let entry = expected.get(file.storage_path.as_str()).ok_or_else(|| {
                invalid(format!(
                    "native backup-log StoragePath '{}' is absent from the virtual directory",
                    file.storage_path
                ))
            })?;
            let payload = payload_slice(bytes, entry)?;
            let decoded = xpress_stream_size(
                payload,
                XpressFraming::Tabular16,
                CodecLimits {
                    max_input_bytes: MAX_STORAGE_BYTES,
                    max_output_bytes: MAX_STORAGE_BYTES,
                    ..CodecLimits::default()
                },
            )
            .map_err(|error| map_codec_error(error, &file.storage_path))?;
            if decoded != usize::try_from(file.size).unwrap_or(usize::MAX) {
                return Err(invalid(format!(
                    "native backup-log decoded size mismatch for '{}'",
                    file.storage_path
                )));
            }
            decoded_total = decoded_total
                .checked_add(decoded)
                .ok_or_else(|| limit("tabular decoded bytes"))?;
            if decoded_total > MAX_STORAGE_BYTES {
                return Err(limit("tabular decoded bytes"));
            }
            if !source_belongs_to_group(&file.generated.normalized_path, &persist) {
                return Err(invalid(format!(
                    "native source path '{}' is outside PersistLocationPath",
                    file.source_path
                )));
            }
            if !kind_allowed_for_native_group(file.generated.kind, group.class) {
                return Err(invalid(format!(
                    "native source path '{}' is incompatible with file-group class {}",
                    file.source_path,
                    group.class.code()
                )));
            }
        }
    }
    if seen.len() != expected.len() {
        return Err(invalid(
            "native backup log does not enumerate every non-marker virtual-directory file",
        ));
    }
    Ok(())
}

fn map_codec_error(error: CodecError, key: &str) -> crate::error::Error {
    match error {
        CodecError::Invalid(message) => {
            invalid(format!("invalid native Xpress member '{key}': {message}"))
        },
        CodecError::LimitExceeded(_) | CodecError::IntegerOverflow => limit("tabular member size"),
    }
}

fn payload_slice<'a>(bytes: &'a [u8], entry: &FileEntry) -> Result<&'a [u8]> {
    let start = usize::try_from(entry.offset.0).map_err(|_source| limit("file offset"))?;
    let size = usize::try_from(entry.stored_size.0).map_err(|_source| limit("file size"))?;
    if size < CRC_SIZE {
        return Err(invalid("native allocation is smaller than its CRC marker"));
    }
    let end = start
        .checked_add(size)
        .ok_or_else(|| limit("native payload range"))?;
    bytes
        .get(start..end - CRC_SIZE)
        .ok_or_else(|| invalid("native payload range is outside storage"))
}

fn validate_native_root(root: &str) -> Result<()> {
    if root.is_empty()
        || root.len() > MAX_PATH_BYTES
        || root.chars().any(|character| character.is_control())
    {
        return Err(invalid(
            "native ServerRoot is empty, oversized, or contains control bytes",
        ));
    }
    Ok(())
}

fn relative_native_path(path: &str, root: &str, label: &str) -> Result<String> {
    validate_native_root(root)?;
    let root = root.trim_end_matches(['\\', '/']);
    let tail = path
        .strip_prefix(root)
        .ok_or_else(|| invalid(format!("native {label} is outside ServerRoot")))?;
    let tail = tail
        .strip_prefix('\\')
        .or_else(|| tail.strip_prefix('/'))
        .ok_or_else(|| invalid(format!("native {label} has a ServerRoot prefix collision")))?;
    normalize_native_relative(tail, label)
}

fn normalize_native_relative(path: &str, label: &str) -> Result<String> {
    if path.is_empty() || path.len() > MAX_PATH_BYTES {
        return Err(invalid(format!("native {label} is empty or oversized")));
    }
    let segments: Vec<_> = path.split(['\\', '/']).collect();
    if segments.iter().any(|segment| {
        segment.is_empty() || *segment == "." || *segment == ".." || !native_component(segment)
    }) {
        return Err(invalid(format!(
            "native {label} contains an invalid path component"
        )));
    }
    Ok(segments.join("/"))
}

fn native_component(value: &str) -> bool {
    !value.is_empty()
        && value != "."
        && value != ".."
        && value
            .chars()
            .all(|character| !character.is_control() && character != ':')
}

fn validate_storage_key(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_PATH_BYTES
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'A'..=b'F').contains(&byte))
    {
        return Err(invalid(format!(
            "native StoragePath '{value}' must be an uppercase hexadecimal key"
        )));
    }
    Ok(())
}

fn native_digits(value: &str) -> bool {
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
}

fn native_versioned_id(value: &str) -> bool {
    value
        .rsplit_once('.')
        .is_some_and(|(id, version)| native_component(id) && native_digits(version))
}

fn native_folder(value: &str, suffix: &str) -> bool {
    value.strip_suffix(suffix).is_some_and(native_versioned_id)
}

fn native_dollar_ids(value: &str, minimum: usize) -> bool {
    let values: Vec<_> = value.split('$').collect();
    values.len() >= minimum && values.iter().all(|value| native_component(value))
}

fn validate_native_hierarchy(parents: &[&str]) -> Result<()> {
    if parents.is_empty() {
        return Ok(());
    }
    if parents.len() > 4 || !native_folder(parents[0], ".db") {
        return Err(invalid(
            "native generated path must begin in a database folder",
        ));
    }
    if parents.len() == 1 {
        return Ok(());
    }
    if native_folder(parents[1], ".cub") {
        if parents.len() >= 3 && !native_folder(parents[2], ".det") {
            return Err(invalid(
                "native cube child folder must be a measure-group folder",
            ));
        }
        if parents.len() == 4 && !native_folder(parents[3], ".prt") {
            return Err(invalid(
                "native measure-group child folder must be a partition folder",
            ));
        }
    } else if native_folder(parents[1], ".dim") || native_folder(parents[1], ".ds") {
        if parents.len() != 2 {
            return Err(invalid(
                "native dimension and data-source folders cannot contain generated subfolders",
            ));
        }
    } else {
        return Err(invalid("unknown native generated database child folder"));
    }
    Ok(())
}

fn classify_native_name(name: &str, parent: Option<&str>) -> Result<GeneratedNameKind> {
    if name == "0.CryptKey.bin" {
        return Ok(GeneratedNameKind::CryptographicKey);
    }
    if let Some(prefix) = name.strip_suffix(".db.xml")
        && native_versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DatabaseDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".dsv.xml")
        && native_versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceViewDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".cub.xml")
        && native_versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::CubeDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".ds.xml")
        && native_versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceOrDimensionDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".dim.xml")
        && native_versioned_id(prefix)
    {
        return Ok(GeneratedNameKind::DataSourceOrDimensionDefinition);
    }
    if let Some(prefix) = name.strip_suffix(".scr.xml")
        && prefix
            .rsplit_once('.')
            .is_some_and(|(id, version)| id == "MdxScript" && native_digits(version))
    {
        return Ok(GeneratedNameKind::MdxScriptMetadata);
    }
    if let Some(prefix) = name
        .strip_prefix("info.")
        .and_then(|value| value.strip_suffix(".xml"))
        && native_digits(prefix)
    {
        return match parent {
            Some(value) if native_folder(value, ".cub") => Ok(GeneratedNameKind::CubeInformation),
            Some(value) if native_folder(value, ".prt") => {
                Ok(GeneratedNameKind::PartitionInformation)
            },
            Some(value) if native_folder(value, ".dim") => Ok(GeneratedNameKind::TableInformation),
            _ => Err(invalid("native info file is outside a recognized folder")),
        };
    }
    for (suffix, kind) in [
        (".det.xml", GeneratedNameKind::MeasureGroupMetadata),
        (".prt.xml", GeneratedNameKind::PartitionMetadata),
    ] {
        if let Some(prefix) = name.strip_suffix(suffix)
            && native_versioned_id(prefix)
        {
            return Ok(kind);
        }
    }
    if let Some(prefix) = name.strip_suffix(".tbl.xml") {
        let Some((stem, version)) = prefix.rsplit_once('.') else {
            return Err(invalid("invalid native tbl.xml generated name"));
        };
        if !native_digits(version) {
            return Err(invalid("invalid native tbl.xml version"));
        }
        if let Some(value) = stem.strip_prefix("R$") {
            if native_dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::TableRelationshipMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("H$") {
            if native_dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::ColumnHierarchyMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("U$") {
            if native_dollar_ids(value, 2) {
                return Ok(GeneratedNameKind::UserHierarchyMetadata);
            }
        } else if native_component(stem) {
            return Ok(GeneratedNameKind::TableMetadata);
        }
    }
    if let Some(prefix) = name.strip_suffix(".dictionary") {
        let values: Vec<_> = prefix.split('.').collect();
        if values.len() >= 3
            && native_digits(values[0])
            && values[1..].iter().all(|value| native_component(value))
        {
            return Ok(GeneratedNameKind::ColumnDictionary);
        }
    }
    if let Some(prefix) = name.strip_suffix(".hidx") {
        let Some((version, rest)) = prefix.split_once('.') else {
            return Err(invalid("invalid native hidx name"));
        };
        if native_digits(version)
            && rest
                .strip_prefix("H$")
                .is_some_and(|value| native_dollar_ids(value, 2))
        {
            return Ok(GeneratedNameKind::ColumnHashIndex);
        }
    }
    if let Some(prefix) = name.strip_suffix(".idf") {
        return classify_native_idf(prefix);
    }
    Err(invalid(format!(
        "unrecognized native generated file name '{name}'"
    )))
}

fn classify_native_idf(value: &str) -> Result<GeneratedNameKind> {
    let Some((version, rest)) = value.split_once('.') else {
        return Err(invalid("invalid native idf generated name"));
    };
    if !native_digits(version) {
        return Err(invalid("invalid native idf version"));
    }
    if let Some(rest) = rest.strip_prefix("R$")
        && rest.ends_with(".INDEX.0")
        && native_dollar_ids(rest.trim_end_matches(".INDEX.0"), 2)
    {
        return Ok(GeneratedNameKind::TableRelationshipIndex);
    }
    if let Some(rest) = rest.strip_prefix("H$") {
        for (suffix, kind) in [
            (".POS_TO_ID.0", GeneratedNameKind::ColumnPositionToId),
            (".ID_TO_POS.0", GeneratedNameKind::ColumnIdToPosition),
        ] {
            if let Some(ids) = rest.strip_suffix(suffix)
                && native_dollar_ids(ids, 2)
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
                && native_dollar_ids(ids, 2)
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
            .all(|value| native_component(value))
    {
        return Ok(GeneratedNameKind::ColumnData);
    }
    Err(invalid("unrecognized native idf generated name"))
}

fn validate_native_kind_location(kind: GeneratedNameKind, parents: &[&str]) -> Result<()> {
    let valid = match kind {
        GeneratedNameKind::CryptographicKey => {
            parents.len() == 1 && native_folder(parents[0], ".db")
        },
        GeneratedNameKind::DatabaseDefinition => parents.is_empty(),
        GeneratedNameKind::DataSourceViewDefinition
        | GeneratedNameKind::CubeDefinition
        | GeneratedNameKind::DataSourceOrDimensionDefinition => {
            parents.len() == 1 && native_folder(parents[0], ".db")
        },
        GeneratedNameKind::CubeInformation
        | GeneratedNameKind::MdxScriptMetadata
        | GeneratedNameKind::MeasureGroupMetadata => {
            parents.len() == 2 && native_folder(parents[1], ".cub")
        },
        GeneratedNameKind::PartitionMetadata => {
            parents.len() == 3 && native_folder(parents[2], ".det")
        },
        GeneratedNameKind::PartitionInformation => {
            parents.len() == 4 && native_folder(parents[3], ".prt")
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
            parents.len() == 2 && native_folder(parents[1], ".dim")
        },
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "native generated file appears outside its section folder",
        ))
    }
}

pub(super) fn kind_allowed_for_native_group(
    kind: GeneratedNameKind,
    class: FileGroupClass,
) -> bool {
    match class {
        FileGroupClass::Database => {
            kind == GeneratedNameKind::DatabaseDefinition
                || kind == GeneratedNameKind::CryptographicKey
        },
        FileGroupClass::DataSource => kind == GeneratedNameKind::DataSourceOrDimensionDefinition,
        FileGroupClass::DataSourceView => kind == GeneratedNameKind::DataSourceViewDefinition,
        FileGroupClass::Cube => {
            matches!(
                kind,
                GeneratedNameKind::CubeDefinition | GeneratedNameKind::CubeInformation
            )
        },
        FileGroupClass::MdxScript => kind == GeneratedNameKind::MdxScriptMetadata,
        FileGroupClass::MeasureGroup => kind == GeneratedNameKind::MeasureGroupMetadata,
        FileGroupClass::Partition => {
            matches!(
                kind,
                GeneratedNameKind::PartitionMetadata | GeneratedNameKind::PartitionInformation
            )
        },
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
                | GeneratedNameKind::CryptographicKey
        ),
    }
}

fn source_belongs_to_group(source: &str, persist: &str) -> bool {
    if source
        .strip_prefix(persist)
        .is_some_and(|tail| tail.starts_with('/'))
    {
        return true;
    }
    let (source_parent, source_name) = source.rsplit_once('/').unwrap_or(("", source));
    let (persist_parent, persist_name) = persist.rsplit_once('/').unwrap_or(("", persist));
    if source_parent != persist_parent {
        return false;
    }
    let Some(source_base) = source_name.strip_suffix(".xml") else {
        return false;
    };
    matches!(
        (
            native_folder_identity(source_base),
            native_folder_identity(persist_name)
        ),
        (Some(source), Some(persist)) if source == persist
    )
}

fn validate_native_persist_location(path: &str, class: FileGroupClass) -> Result<()> {
    let segments: Vec<_> = path.split('/').collect();
    validate_native_hierarchy(&segments)?;
    let valid = match class {
        FileGroupClass::Database => segments.len() == 1 && native_folder(segments[0], ".db"),
        FileGroupClass::DataSource => segments.len() == 2 && native_folder(segments[1], ".ds"),
        FileGroupClass::DataSourceView => segments.len() == 1 && native_folder(segments[0], ".db"),
        FileGroupClass::Dimension => segments.len() == 2 && native_folder(segments[1], ".dim"),
        FileGroupClass::Cube | FileGroupClass::MdxScript => {
            segments.len() == 2 && native_folder(segments[1], ".cub")
        },
        FileGroupClass::MeasureGroup => segments.len() == 3 && native_folder(segments[2], ".det"),
        FileGroupClass::Partition => segments.len() == 4 && native_folder(segments[3], ".prt"),
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "native PersistLocationPath '{path}' is incompatible with file-group class {}",
            class.code()
        )))
    }
}

fn native_folder_identity(value: &str) -> Option<(&str, &str)> {
    [".db", ".cub", ".det", ".prt", ".dim", ".ds"]
        .iter()
        .find_map(|suffix| {
            value.strip_suffix(suffix).and_then(|prefix| {
                prefix
                    .rsplit_once('.')
                    .filter(|(_, version)| native_digits(version))
                    .map(|(id, _)| (id, *suffix))
            })
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROOT: &str = r"\\?\C:\native";

    #[test]
    fn classifies_native_versions_and_unicode_identifiers() {
        let script = classify_tabular_source_path(
            r"\\?\C:\native\DB.0.db\Model.28.cub\MdxScript.162.scr.xml",
            ROOT,
        )
        .unwrap();
        assert_eq!(script.kind, GeneratedNameKind::MdxScriptMetadata);
        assert_eq!(
            script.normalized_path,
            "DB.0.db/Model.28.cub/MdxScript.162.scr.xml"
        );

        let column = classify_tabular_source_path(
            r"\\?\C:\native\DB.0.db\Tabelle1.0.dim\35.Tabelle1.RowNumber 1.0.idf",
            ROOT,
        )
        .unwrap();
        assert_eq!(column.kind, GeneratedNameKind::ColumnData);

        let key =
            classify_tabular_source_path(r"\\?\C:\native\DB.0.db\0.CryptKey.bin", ROOT).unwrap();
        assert_eq!(key.kind, GeneratedNameKind::CryptographicKey);
    }

    #[test]
    fn rejects_root_prefix_collisions_traversal_and_unknown_persist_folders() {
        assert!(
            classify_tabular_source_path(
                r"\\?\C:\native-copy\DB.0.db\Model.1.cub\MdxScript.1.scr.xml",
                ROOT,
            )
            .is_err()
        );
        assert!(
            classify_tabular_source_path(r"\\?\C:\native\DB.0.db\..\MdxScript.1.scr.xml", ROOT,)
                .is_err()
        );
        assert!(
            validate_native_persist_location("DB.0.db/not-a-folder", FileGroupClass::Cube,)
                .is_err()
        );
        assert!(!source_belongs_to_group(
            "DB.0.db/MdxScript.1.scr.xml",
            "DB.0.db/not-a-folder"
        ));
    }
}
