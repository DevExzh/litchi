//! MS-XLDM semantic projection for backup-log metadata and generated paths.

use super::codec::{
    bool_value, bounded_leaf, exact_children, i32_value, i64_value, invalid, leaf_text, limit,
    valid_upper_guid,
};
use super::model::{
    BackupLog, FileGroup, FileGroupClass, GeneratedPath, LoggedFile, MAX_FILES, MAX_PARTITIONS,
    Node, WriteAccess, XmlEncoding,
};
use super::validation::{
    classify_generated_name, kind_allowed_for_group, normalize_generated_path,
    validate_generated_hierarchy, validate_kind_location,
};
use crate::error::Result;
use std::collections::HashSet;

/// Classify and validate one section 2.2 generated storage path.
pub fn classify_generated_path(path: &str) -> Result<GeneratedPath> {
    let normalized_path = normalize_generated_path(path)?;
    let segments: Vec<_> = normalized_path.split('/').collect();
    let file = *segments.last().unwrap_or_else(|| {
        crate::error::panic_missing_invariant("required value was checked before extraction")
    });
    let parents = &segments[..segments.len() - 1];
    validate_generated_hierarchy(parents)?;
    let kind = classify_generated_name(file, parents.last().copied())?;
    validate_kind_location(kind, parents)?;
    Ok(GeneratedPath {
        normalized_path,
        kind,
    })
}

pub(super) fn parse_backup_log(root: &Node, encoding: XmlEncoding) -> Result<BackupLog> {
    let names = [
        "BackupRestoreSyncVersion",
        "ServerRoot",
        "SvrEncryptPwdFlag",
        "ServerEnableBinaryXML",
        "ServerEnableCompression",
        "CompressionFlag",
        "EncryptionFlag",
        "ObjectName",
        "ObjectId",
        "Write",
        "OlapInfo",
        "Collations",
        "Languages",
        "FileGroups",
    ];
    let values = exact_children(root, "BackupLog", &names)?;
    let backup_restore_sync_version = i32_value(values[0])?;
    if backup_restore_sync_version != 1153 {
        return Err(invalid(
            "backup-log BackupRestoreSyncVersion must equal 1153",
        ));
    }
    let server_root = bounded_leaf(values[1], "server root")?;
    if !bool_value(values[2])? {
        return Err(invalid("backup-log SvrEncryptPwdFlag must be true"));
    }
    if bool_value(values[3])?
        || bool_value(values[4])?
        || bool_value(values[5])?
        || bool_value(values[6])?
    {
        return Err(invalid(
            "backup-log binary XML, compression, and encryption flags must be false",
        ));
    }
    let object_name = bounded_leaf(values[7], "object name")?;
    let object_id = bounded_leaf(values[8], "object id")?;
    let write_access = match leaf_text(values[9])? {
        "ReadWrite" => WriteAccess::ReadWrite,
        "ReadOnly" => WriteAccess::ReadOnly,
        "ReadOnlyExclusive" => WriteAccess::ReadOnlyExclusive,
        _ => return Err(invalid("unknown backup-log Write value")),
    };
    let is_olap = bool_value(values[10])?;
    let collations = parse_repeated_strings(values[11], "Collations", "Collation", "collation")?;
    let languages = parse_languages(values[12])?;
    let file_groups = parse_file_groups(values[13])?;
    Ok(BackupLog {
        backup_restore_sync_version,
        server_root,
        object_name,
        object_id,
        write_access,
        is_olap,
        collations,
        languages,
        file_groups,
        encoding,
    })
}

fn parse_repeated_strings(
    node: &Node,
    root: &str,
    child: &str,
    label: &str,
) -> Result<Vec<String>> {
    if node.name != root
        || node.attributes != 0
        || !node.text.trim().is_empty()
        || node.children.is_empty()
    {
        return Err(invalid(format!("{root} must contain at least one {child}")));
    }
    if node.children.len() > MAX_FILES {
        return Err(limit(label));
    }
    let mut seen = HashSet::new();
    let mut output = Vec::with_capacity(node.children.len());
    for value in &node.children {
        if value.name != child {
            return Err(invalid(format!("unexpected child in {root}")));
        }
        let value = bounded_leaf(value, label)?;
        if !seen.insert(value.to_lowercase()) {
            return Err(invalid(format!("duplicate case-insensitive {label}")));
        }
        output.push(value);
    }
    Ok(output)
}
fn parse_languages(node: &Node) -> Result<Vec<i32>> {
    if node.name != "Languages"
        || node.attributes != 0
        || !node.text.trim().is_empty()
        || node.children.is_empty()
    {
        return Err(invalid("Languages must contain at least one Language"));
    }
    if node.children.len() > MAX_FILES {
        return Err(limit("language count"));
    }
    let mut seen = HashSet::new();
    let mut output = Vec::with_capacity(node.children.len());
    for value in &node.children {
        if value.name != "Language" {
            return Err(invalid("unexpected child in Languages"));
        }
        let value = i32_value(value)?;
        if !seen.insert(value) {
            return Err(invalid("duplicate backup-log language"));
        }
        output.push(value);
    }
    Ok(output)
}

fn parse_file_groups(node: &Node) -> Result<Vec<FileGroup>> {
    if node.name != "FileGroups"
        || node.attributes != 0
        || !node.text.trim().is_empty()
        || node.children.is_empty()
    {
        return Err(invalid("FileGroups must contain at least one FileGroup"));
    }
    if node.children.len() > MAX_PARTITIONS {
        return Err(limit("file-group count"));
    }
    let names = [
        "Class",
        "ID",
        "Name",
        "ObjectVersion",
        "PersistLocation",
        "PersistLocationPath",
        "StorageLocationPath",
        "ObjectID",
        "FileList",
    ];
    let mut output = Vec::with_capacity(node.children.len());
    let mut identities = HashSet::new();
    let mut total_files = 0usize;
    for group in &node.children {
        let values = exact_children(group, "FileGroup", &names)?;
        let class = FileGroupClass::parse(i32_value(values[0])?)?;
        let id = bounded_leaf(values[1], "file-group ID")?;
        let name = bounded_leaf(values[2], "file-group name")?;
        let object_version = i32_value(values[3])?;
        let persist_location = i32_value(values[4])?;
        if object_version < 0 || persist_location < 0 {
            return Err(invalid("file-group versions cannot be negative"));
        }
        let persist_location_path = bounded_leaf(values[5], "persist location path")?;
        let storage_location_path = bounded_leaf(values[6], "storage location path")?;
        let object_id = bounded_leaf(values[7], "file-group ObjectID")?;
        if !valid_upper_guid(&object_id) {
            return Err(invalid("file-group ObjectID must be an uppercase UUID"));
        }
        if !identities.insert((class, object_id.clone())) {
            return Err(invalid("duplicate file-group class and ObjectID"));
        }
        let files = parse_logged_files(values[8], class)?;
        total_files = total_files
            .checked_add(files.len())
            .ok_or_else(|| limit("logged file count"))?;
        if total_files > MAX_FILES {
            return Err(limit("logged file count"));
        }
        output.push(FileGroup {
            class,
            id,
            name,
            object_version,
            persist_location,
            persist_location_path,
            storage_location_path,
            object_id,
            files,
        });
    }
    Ok(output)
}

fn parse_logged_files(node: &Node, class: FileGroupClass) -> Result<Vec<LoggedFile>> {
    if node.name != "FileList"
        || node.attributes != 0
        || !node.text.trim().is_empty()
        || node.children.is_empty()
    {
        return Err(invalid("FileList must contain at least one BackupFile"));
    }
    if node.children.len() > MAX_FILES {
        return Err(limit("logged file count"));
    }
    let names = ["Path", "StoragePath", "LastWriteTime", "Size"];
    node.children
        .iter()
        .map(|file| {
            let values = exact_children(file, "BackupFile", &names)?;
            let source_path = bounded_leaf(values[0], "source path")?;
            let storage_path = bounded_leaf(values[1], "storage path")?;
            let last_write_timestamp = i64_value(values[2])?;
            let signed_size = i32_value(values[3])?;
            if signed_size < 0 {
                return Err(invalid("logged file size cannot be negative"));
            }
            let generated = classify_generated_path(&storage_path)?;
            if !kind_allowed_for_group(generated.kind, class) {
                return Err(invalid(format!(
                    "generated path '{}' is incompatible with file-group class {}",
                    storage_path,
                    class.code()
                )));
            }
            Ok(LoggedFile {
                source_path,
                storage_path,
                last_write_timestamp,
                size: u32::try_from(signed_size)
                    .map_err(|_source| invalid("file-log size exceeds u32"))?,
                generated,
            })
        })
        .collect()
}
