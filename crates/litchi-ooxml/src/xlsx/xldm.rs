//! Safe structural inspection of MS-XLDM storage streams.
//!
//! This module implements only the outer storage described by MS-XLDM 2.1:
//! header, partition marker, serial file allocations, CRC markers, page
//! padding, and virtual directory. Member payloads are never decompressed,
//! decrypted, parsed as model metadata, evaluated, or used for I/O.

use crate::error::{OoxmlError, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;
use std::collections::HashSet;

pub const XLDM_PAGE_SIZE: usize = 4096;
pub const XLDM_STREAM_SIGNATURE: &str = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";

const MAX_STORAGE_BYTES: usize = 512 * 1024 * 1024;
const MAX_DIRECTORY_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_NODES: usize = 500_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_TEXT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FILES: usize = 100_000;
const MAX_PARTITIONS: usize = 65_536;
const MAX_PATH_BYTES: usize = 32 * 1024;
const CRC_SIZE: usize = 4;
const BOM: [u8; 2] = [0xFF, 0xFE];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmXmlEncoding {
    Utf16Le,
    Utf8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmCompression {
    Xpress,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmFileKind {
    Partitions,
    BackupLog,
    CryptographicKey,
    XmlMetadata,
    OpaqueBinary,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XldmWriteAccess {
    ReadWrite,
    ReadOnly,
    ReadOnlyExclusive,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XldmFileGroupClass {
    Database,
    DataSource,
    Dimension,
    Cube,
    MeasureGroup,
    Partition,
    DataSourceView,
    MdxScript,
}

impl XldmFileGroupClass {
    pub const fn code(self) -> i32 {
        match self {
            Self::Database => 100002,
            Self::DataSource => 100003,
            Self::Dimension => 100006,
            Self::Cube => 100010,
            Self::MeasureGroup => 100016,
            Self::Partition => 100021,
            Self::DataSourceView => 100053,
            Self::MdxScript => 100060,
        }
    }
    fn parse(value: i32) -> Result<Self> {
        match value {
            100002 => Ok(Self::Database),
            100003 => Ok(Self::DataSource),
            100006 => Ok(Self::Dimension),
            100010 => Ok(Self::Cube),
            100016 => Ok(Self::MeasureGroup),
            100021 => Ok(Self::Partition),
            100053 => Ok(Self::DataSourceView),
            100060 => Ok(Self::MdxScript),
            _ => Err(invalid(format!("unknown MS-XLDM file-group class {value}"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XldmGeneratedNameKind {
    DatabaseDefinition,
    DataSourceViewDefinition,
    CubeDefinition,
    DataSourceOrDimensionDefinition,
    CubeInformation,
    PartitionInformation,
    TableInformation,
    MdxScriptMetadata,
    MeasureGroupMetadata,
    PartitionMetadata,
    TableMetadata,
    TableRelationshipMetadata,
    ColumnHierarchyMetadata,
    UserHierarchyMetadata,
    ColumnData,
    TableRelationshipIndex,
    ColumnPositionToId,
    ColumnIdToPosition,
    ColumnHashIndex,
    ColumnDictionary,
    UserHierarchyChildCount,
    UserHierarchyFirstChildPosition,
    UserHierarchyParentPosition,
    UserHierarchyMultilevelId,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmGeneratedPath {
    pub normalized_path: String,
    pub kind: XldmGeneratedNameKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmLoggedFile {
    pub source_path: String,
    pub storage_path: String,
    pub last_write_timestamp: i64,
    pub size: u32,
    pub generated: XldmGeneratedPath,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmFileGroup {
    pub class: XldmFileGroupClass,
    pub id: String,
    pub name: String,
    pub object_version: i32,
    pub persist_location: i32,
    pub persist_location_path: String,
    /// Unused by MS-XLDM and retained without interpretation.
    pub storage_location_path: String,
    pub object_id: String,
    pub files: Vec<XldmLoggedFile>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmBackupLog {
    pub backup_restore_sync_version: i32,
    /// Originating filesystem root; retained as text and never accessed.
    pub server_root: String,
    pub object_name: String,
    pub object_id: String,
    pub write_access: XldmWriteAccess,
    pub is_olap: bool,
    pub collations: Vec<String>,
    pub languages: Vec<i32>,
    pub file_groups: Vec<XldmFileGroup>,
    pub encoding: XldmXmlEncoding,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct XldmOffset(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct XldmSize(pub u64);

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmHeader {
    pub backup_restore_sync_version: i32,
    pub fault_code: u32,
    pub encryption_key_version: i32,
    pub compression: XldmCompression,
    pub directory_offset: XldmOffset,
    pub directory_size: XldmSize,
    pub file_count: u32,
    pub object_id: String,
    pub data_offset: XldmOffset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmPartitionMarker<'a> {
    pub partition_count: usize,
    pub encoding: XldmXmlEncoding,
    /// Encoded XML is retained exactly; unused fields, including connection
    /// strings, are not projected or acted upon.
    pub encoded_xml: &'a [u8],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmFileEntry {
    pub path: String,
    pub kind: XldmFileKind,
    pub offset: XldmOffset,
    /// Stored size includes the four-byte CRC marker.
    pub stored_size: XldmSize,
    pub crc32: u32,
    pub delete: bool,
    pub created_timestamp: i64,
    pub access_timestamp: i64,
    pub last_write_timestamp: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmStorage<'a> {
    pub header: XldmHeader,
    pub header_encoding: XldmXmlEncoding,
    pub directory_encoding: XldmXmlEncoding,
    pub partition_marker: XldmPartitionMarker<'a>,
    pub backup_log: XldmBackupLog,
    pub files: Vec<XldmFileEntry>,
    bytes: &'a [u8],
}

impl XldmStorage<'_> {
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return one member's bytes without its CRC marker. The returned bytes
    /// remain compressed, encrypted, or otherwise encoded exactly as stored.
    pub fn file_payload(&self, index: usize) -> Option<&[u8]> {
        let entry = self.files.get(index)?;
        let start = usize::try_from(entry.offset.0).ok()?;
        let size = usize::try_from(entry.stored_size.0).ok()?;
        self.bytes
            .get(start..start.checked_add(size)?.checked_sub(CRC_SIZE)?)
    }

    pub fn file_stored_bytes(&self, index: usize) -> Option<&[u8]> {
        let entry = self.files.get(index)?;
        let start = usize::try_from(entry.offset.0).ok()?;
        let size = usize::try_from(entry.stored_size.0).ok()?;
        self.bytes.get(start..start.checked_add(size)?)
    }
}

#[derive(Clone)]
struct Node {
    name: String,
    attributes: usize,
    children: Vec<Node>,
    text: String,
}

/// Validate and inspect the outer MS-XLDM virtual storage.
pub fn inspect_xldm(bytes: &[u8]) -> Result<XldmStorage<'_>> {
    if bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if bytes.len() < XLDM_PAGE_SIZE * 3 || bytes.len() % XLDM_PAGE_SIZE != 0 {
        return Err(invalid(
            "MS-XLDM storage must contain at least three complete 4096-byte pages",
        ));
    }
    if bytes[..2] != BOM {
        return Err(invalid("MS-XLDM header byte-order mark is missing"));
    }
    let signature = utf16le(XLDM_STREAM_SIGNATURE);
    if bytes.get(2..2 + signature.len()) != Some(signature.as_slice()) {
        return Err(invalid("MS-XLDM stream storage signature is invalid"));
    }
    let header_xml_start = 2 + signature.len();
    let close = utf16le("</BackupLog>");
    let relative_end = memchr::memmem::rfind(&bytes[header_xml_start..XLDM_PAGE_SIZE], &close)
        .ok_or_else(|| invalid("MS-XLDM header BackupLog closing element is missing"))?;
    let header_xml_end = header_xml_start
        .checked_add(relative_end)
        .and_then(|value| value.checked_add(close.len()))
        .ok_or_else(|| limit("header XML range"))?;
    if bytes[header_xml_end..XLDM_PAGE_SIZE]
        .iter()
        .any(|byte| *byte != 0)
    {
        return Err(invalid("MS-XLDM header padding must be zero"));
    }
    let (header_xml, header_encoding) = decode_xml(&bytes[header_xml_start..header_xml_end], true)?;
    let header = parse_header(&parse_xml(&header_xml)?)?;
    let data_offset = checked_usize(header.data_offset.0, "data offset")?;
    let directory_offset = checked_usize(header.directory_offset.0, "directory offset")?;
    let directory_size = checked_usize(header.directory_size.0, "directory size")?;
    if data_offset != XLDM_PAGE_SIZE {
        return Err(invalid(
            "MS-XLDM data offset must equal the 4096-byte header size",
        ));
    }
    if directory_offset < XLDM_PAGE_SIZE * 2 || directory_offset % XLDM_PAGE_SIZE != 0 {
        return Err(invalid(
            "MS-XLDM directory offset must be page aligned after the files section",
        ));
    }
    if directory_size == 0 || directory_size > MAX_DIRECTORY_BYTES {
        return Err(limit("directory bytes"));
    }
    let directory_end = directory_offset
        .checked_add(directory_size)
        .ok_or_else(|| limit("directory range"))?;
    if directory_end > bytes.len() {
        return Err(invalid("MS-XLDM virtual directory extends beyond storage"));
    }
    if bytes[directory_end..].iter().any(|byte| *byte != 0) {
        return Err(invalid("MS-XLDM directory padding must be zero"));
    }
    if bytes.get(data_offset..data_offset + 2) != Some(&BOM) {
        return Err(invalid("MS-XLDM files-section byte-order mark is missing"));
    }
    if bytes.get(directory_offset..directory_offset + 2) == Some(&BOM) {
        return Err(invalid(
            "MS-XLDM virtual directory must not have a byte-order mark",
        ));
    }
    let (directory_xml, directory_encoding) =
        decode_xml(&bytes[directory_offset..directory_end], false)?;
    let mut files = parse_directory(&parse_xml(&directory_xml)?)?;
    if files.len() != header.file_count as usize {
        return Err(invalid(format!(
            "header file count {} does not match directory count {}",
            header.file_count,
            files.len()
        )));
    }
    if files.len() < 2 {
        return Err(invalid(
            "MS-XLDM storage requires partition and backup-log allocations",
        ));
    }
    validate_paths(&files)?;
    let mut order: Vec<usize> = (0..files.len()).collect();
    order.sort_by_key(|index| files[*index].offset);
    validate_allocations(bytes, data_offset, directory_offset, &mut files, &order)?;
    let first = order[0];
    let partition_bytes = payload_slice(bytes, &files[first])?;
    let (partitions_xml, partition_encoding) = decode_xml(partition_bytes, false)?;
    let partition_count = parse_partitions(&parse_xml(&partitions_xml)?)?;
    files[first].kind = XldmFileKind::Partitions;
    let last = *order.last().unwrap();
    files[last].kind = XldmFileKind::BackupLog;
    let (backup_xml, backup_encoding) = decode_xml(payload_slice(bytes, &files[last])?, false)?;
    let backup_log = parse_backup_log(&parse_xml(&backup_xml)?, backup_encoding)?;
    validate_backup_log(&backup_log, &files, first, last)?;
    Ok(XldmStorage {
        header,
        header_encoding,
        directory_encoding,
        partition_marker: XldmPartitionMarker {
            partition_count,
            encoding: partition_encoding,
            encoded_xml: partition_bytes,
        },
        backup_log,
        files,
        bytes,
    })
}

/// Revalidate and return the original byte stream exactly.
pub fn write_xldm(storage: &XldmStorage<'_>) -> Result<Vec<u8>> {
    inspect_xldm(storage.bytes)?;
    Ok(storage.bytes.to_vec())
}

/// Classify and validate one section 2.2 generated storage path.
pub fn classify_xldm_generated_path(path: &str) -> Result<XldmGeneratedPath> {
    let normalized_path = normalize_generated_path(path)?;
    let segments: Vec<_> = normalized_path.split('/').collect();
    let file = *segments.last().unwrap();
    let parents = &segments[..segments.len() - 1];
    validate_generated_hierarchy(parents)?;
    let kind = classify_generated_name(file, parents.last().copied())?;
    validate_kind_location(kind, parents)?;
    Ok(XldmGeneratedPath {
        normalized_path,
        kind,
    })
}

fn parse_backup_log(root: &Node, encoding: XldmXmlEncoding) -> Result<XldmBackupLog> {
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
        "ReadWrite" => XldmWriteAccess::ReadWrite,
        "ReadOnly" => XldmWriteAccess::ReadOnly,
        "ReadOnlyExclusive" => XldmWriteAccess::ReadOnlyExclusive,
        _ => return Err(invalid("unknown backup-log Write value")),
    };
    let is_olap = bool_value(values[10])?;
    let collations = parse_repeated_strings(values[11], "Collations", "Collation", "collation")?;
    let languages = parse_languages(values[12])?;
    let file_groups = parse_file_groups(values[13])?;
    Ok(XldmBackupLog {
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

fn parse_file_groups(node: &Node) -> Result<Vec<XldmFileGroup>> {
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
        let class = XldmFileGroupClass::parse(i32_value(values[0])?)?;
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
        output.push(XldmFileGroup {
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

fn parse_logged_files(node: &Node, class: XldmFileGroupClass) -> Result<Vec<XldmLoggedFile>> {
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
            let generated = classify_xldm_generated_path(&storage_path)?;
            if !kind_allowed_for_group(generated.kind, class) {
                return Err(invalid(format!(
                    "generated path '{}' is incompatible with file-group class {}",
                    storage_path,
                    class.code()
                )));
            }
            Ok(XldmLoggedFile {
                source_path,
                storage_path,
                last_write_timestamp,
                size: signed_size as u32,
                generated,
            })
        })
        .collect()
}

fn validate_backup_log(
    log: &XldmBackupLog,
    directory: &[XldmFileEntry],
    partitions: usize,
    backup: usize,
) -> Result<()> {
    let mut expected = HashMap::<String, &XldmFileEntry>::new();
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

fn normalize_generated_path(path: &str) -> Result<String> {
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
fn validate_generated_hierarchy(parents: &[&str]) -> Result<()> {
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

fn classify_generated_name(name: &str, parent: Option<&str>) -> Result<XldmGeneratedNameKind> {
    if name == "MdxScript.0.scr.xml" {
        return Ok(XldmGeneratedNameKind::MdxScriptMetadata);
    }
    if let Some(prefix) = name.strip_suffix(".db.xml") {
        if versioned_id(prefix) {
            return Ok(XldmGeneratedNameKind::DatabaseDefinition);
        }
    }
    if let Some(prefix) = name.strip_suffix(".dsv.xml") {
        if versioned_id(prefix) {
            return Ok(XldmGeneratedNameKind::DataSourceViewDefinition);
        }
    }
    if let Some(prefix) = name.strip_suffix(".cub.xml") {
        if versioned_id(prefix) {
            return Ok(XldmGeneratedNameKind::CubeDefinition);
        }
    }
    if let Some(prefix) = name.strip_suffix(".ds.xml") {
        if versioned_id(prefix) {
            return Ok(XldmGeneratedNameKind::DataSourceOrDimensionDefinition);
        }
    }
    if let Some(prefix) = name.strip_suffix(".dim.xml") {
        if versioned_id(prefix) {
            return Ok(XldmGeneratedNameKind::DataSourceOrDimensionDefinition);
        }
    }
    if let Some(prefix) = name
        .strip_prefix("info.")
        .and_then(|value| value.strip_suffix(".xml"))
    {
        if digits(prefix) {
            return match parent {
                Some(value) if folder(value, ".cub", true) => {
                    Ok(XldmGeneratedNameKind::CubeInformation)
                },
                Some(value) if folder(value, ".prt", false) => {
                    Ok(XldmGeneratedNameKind::PartitionInformation)
                },
                Some(value) if folder(value, ".dim", true) => {
                    Ok(XldmGeneratedNameKind::TableInformation)
                },
                _ => Err(invalid(
                    "info file is outside a cube, partition, or dimension folder",
                )),
            };
        }
    }
    for (suffix, kind) in [
        (".det.xml", XldmGeneratedNameKind::MeasureGroupMetadata),
        (".prt.xml", XldmGeneratedNameKind::PartitionMetadata),
    ] {
        if let Some(prefix) = name.strip_suffix(suffix) {
            if versioned_id(prefix) {
                return Ok(kind);
            }
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
                return Ok(XldmGeneratedNameKind::TableRelationshipMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("H$") {
            if dollar_ids(value, 2) {
                return Ok(XldmGeneratedNameKind::ColumnHierarchyMetadata);
            }
        } else if let Some(value) = stem.strip_prefix("U$") {
            if dollar_ids(value, 2) {
                return Ok(XldmGeneratedNameKind::UserHierarchyMetadata);
            }
        } else if valid_id(stem) {
            return Ok(XldmGeneratedNameKind::TableMetadata);
        }
    }
    if let Some(prefix) = name.strip_suffix(".dictionary") {
        let values: Vec<_> = prefix.split('.').collect();
        if values.len() >= 3 && digits(values[0]) && values[1..].iter().all(|value| valid_id(value))
        {
            return Ok(XldmGeneratedNameKind::ColumnDictionary);
        }
    }
    if let Some(prefix) = name.strip_suffix(".hidx") {
        let Some((version, rest)) = prefix.split_once('.') else {
            return Err(invalid("invalid hidx name"));
        };
        if digits(version) && rest.starts_with("H$") && dollar_ids(&rest[2..], 2) {
            return Ok(XldmGeneratedNameKind::ColumnHashIndex);
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
fn classify_idf(value: &str) -> Result<XldmGeneratedNameKind> {
    let Some((version, rest)) = value.split_once('.') else {
        return Err(invalid("invalid idf generated name"));
    };
    if !digits(version) {
        return Err(invalid("invalid idf version"));
    }
    if let Some(rest) = rest.strip_prefix("R$") {
        if rest.ends_with(".INDEX.0") && dollar_ids(rest.trim_end_matches(".INDEX.0"), 2) {
            return Ok(XldmGeneratedNameKind::TableRelationshipIndex);
        }
    }
    if let Some(rest) = rest.strip_prefix("H$") {
        for (suffix, kind) in [
            (".POS_TO_ID.0", XldmGeneratedNameKind::ColumnPositionToId),
            (".ID_TO_POS.0", XldmGeneratedNameKind::ColumnIdToPosition),
        ] {
            if let Some(ids) = rest.strip_suffix(suffix) {
                if dollar_ids(ids, 2) {
                    return Ok(kind);
                }
            }
        }
    }
    if let Some(rest) = rest.strip_prefix("U$") {
        for (suffix, kind) in [
            (
                ".CHILD_COUNT.0",
                XldmGeneratedNameKind::UserHierarchyChildCount,
            ),
            (
                ".FIRST_CHILD_POS.0",
                XldmGeneratedNameKind::UserHierarchyFirstChildPosition,
            ),
            (
                ".PARENT_POS.0",
                XldmGeneratedNameKind::UserHierarchyParentPosition,
            ),
            (
                ".MULTI_LEVEL_ID.0",
                XldmGeneratedNameKind::UserHierarchyMultilevelId,
            ),
        ] {
            if let Some(ids) = rest.strip_suffix(suffix) {
                if dollar_ids(ids, 2) {
                    return Ok(kind);
                }
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
        return Ok(XldmGeneratedNameKind::ColumnData);
    }
    Err(invalid("unrecognized section 2.2 idf generated name"))
}
fn validate_kind_location(kind: XldmGeneratedNameKind, parents: &[&str]) -> Result<()> {
    let valid = match kind {
        XldmGeneratedNameKind::DatabaseDefinition => parents.is_empty(),
        XldmGeneratedNameKind::DataSourceViewDefinition
        | XldmGeneratedNameKind::CubeDefinition
        | XldmGeneratedNameKind::DataSourceOrDimensionDefinition => parents.len() == 1,
        XldmGeneratedNameKind::CubeInformation
        | XldmGeneratedNameKind::MdxScriptMetadata
        | XldmGeneratedNameKind::MeasureGroupMetadata => {
            parents.len() == 2 && folder(parents[1], ".cub", true)
        },
        XldmGeneratedNameKind::PartitionMetadata => {
            parents.len() == 3 && folder(parents[2], ".det", false)
        },
        XldmGeneratedNameKind::PartitionInformation => {
            parents.len() == 4 && folder(parents[3], ".prt", false)
        },
        _ => parents.len() == 2 && folder(parents[1], ".dim", true),
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "generated file appears outside its section 2.2 folder",
        ))
    }
}
fn kind_allowed_for_group(kind: XldmGeneratedNameKind, class: XldmFileGroupClass) -> bool {
    match class {
        XldmFileGroupClass::Database => kind == XldmGeneratedNameKind::DatabaseDefinition,
        XldmFileGroupClass::DataSource => {
            kind == XldmGeneratedNameKind::DataSourceOrDimensionDefinition
        },
        XldmFileGroupClass::DataSourceView => {
            kind == XldmGeneratedNameKind::DataSourceViewDefinition
        },
        XldmFileGroupClass::Cube => matches!(
            kind,
            XldmGeneratedNameKind::CubeDefinition | XldmGeneratedNameKind::CubeInformation
        ),
        XldmFileGroupClass::MdxScript => kind == XldmGeneratedNameKind::MdxScriptMetadata,
        XldmFileGroupClass::MeasureGroup => kind == XldmGeneratedNameKind::MeasureGroupMetadata,
        XldmFileGroupClass::Partition => matches!(
            kind,
            XldmGeneratedNameKind::PartitionMetadata | XldmGeneratedNameKind::PartitionInformation
        ),
        XldmFileGroupClass::Dimension => !matches!(
            kind,
            XldmGeneratedNameKind::DatabaseDefinition
                | XldmGeneratedNameKind::DataSourceViewDefinition
                | XldmGeneratedNameKind::CubeDefinition
                | XldmGeneratedNameKind::CubeInformation
                | XldmGeneratedNameKind::MdxScriptMetadata
                | XldmGeneratedNameKind::MeasureGroupMetadata
                | XldmGeneratedNameKind::PartitionMetadata
                | XldmGeneratedNameKind::PartitionInformation
        ),
    }
}
fn bounded_leaf(node: &Node, label: &str) -> Result<String> {
    let value = leaf_text(node)?;
    if value.len() > MAX_PATH_BYTES {
        return Err(limit(label));
    }
    Ok(value.to_owned())
}

fn parse_header(root: &Node) -> Result<XldmHeader> {
    let names = [
        "BackupRestoreSyncVersion",
        "Fault",
        "faultcode",
        "ErrorCode",
        "EncryptionFlag",
        "EncryptionKey",
        "ApplyCompression",
        "m_cbOffsetHeader",
        "DataSize",
        "Files",
        "ObjectID",
        "m_cbOffsetData",
    ];
    let values = exact_children(root, "BackupLog", &names)?;
    let backup_restore_sync_version = i32_value(values[0])?;
    if backup_restore_sync_version != 140 {
        return Err(invalid("BackupRestoreSyncVersion must equal 140"));
    }
    if bool_value(values[1])? {
        return Err(invalid("header Fault must be false"));
    }
    let fault_code = u32_value(values[2])?;
    if !bool_value(values[3])? {
        return Err(invalid("header ErrorCode must be true"));
    }
    if bool_value(values[4])? {
        return Err(invalid("header EncryptionFlag must be false"));
    }
    let encryption_key_version = i32_value(values[5])?;
    if !bool_value(values[6])? {
        return Err(invalid("header ApplyCompression must be true"));
    }
    let directory_offset = XldmOffset(u64_value(values[7])?);
    let directory_size = XldmSize(u64_value(values[8])?);
    let file_count = u32_value(values[9])?;
    if file_count as usize > MAX_FILES {
        return Err(limit("file count"));
    }
    let object_id = leaf_text(values[10])?.to_owned();
    if !valid_upper_guid(&object_id) {
        return Err(invalid("header ObjectID must be an uppercase UUID"));
    }
    let data_offset = XldmOffset(u64_value(values[11])?);
    Ok(XldmHeader {
        backup_restore_sync_version,
        fault_code,
        encryption_key_version,
        compression: XldmCompression::Xpress,
        directory_offset,
        directory_size,
        file_count,
        object_id,
        data_offset,
    })
}

fn parse_directory(root: &Node) -> Result<Vec<XldmFileEntry>> {
    if root.name != "VirtualDirectory" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free VirtualDirectory root"));
    }
    if root.children.len() > MAX_FILES {
        return Err(limit("file count"));
    }
    let names = [
        "Path",
        "Size",
        "m_cbOffsetHeader",
        "Delete",
        "CreatedTimestamp",
        "Access",
        "LastWriteTime",
    ];
    root.children
        .iter()
        .map(|child| {
            let values = exact_children(child, "BackupFile", &names)?;
            let path = leaf_text(values[0])?.to_owned();
            let stored_size = XldmSize(u64_value(values[1])?);
            let offset = XldmOffset(u64_value(values[2])?);
            let delete = bool_value(values[3])?;
            let created_timestamp = i64_value(values[4])?;
            let access_timestamp = i64_value(values[5])?;
            let last_write_timestamp = i64_value(values[6])?;
            let lower = path.to_ascii_lowercase();
            let kind = if lower.ends_with("cryptkey.bin") {
                XldmFileKind::CryptographicKey
            } else if lower.ends_with(".xml") {
                XldmFileKind::XmlMetadata
            } else {
                XldmFileKind::OpaqueBinary
            };
            Ok(XldmFileEntry {
                path,
                kind,
                offset,
                stored_size,
                crc32: 0,
                delete,
                created_timestamp,
                access_timestamp,
                last_write_timestamp,
            })
        })
        .collect()
}

fn parse_partitions(root: &Node) -> Result<usize> {
    if root.name != "Partitions" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free Partitions root"));
    }
    if root.children.len() > MAX_PARTITIONS {
        return Err(limit("partition count"));
    }
    let names = [
        "ObjectPath",
        "Name",
        "DataSize",
        "Location",
        "DataSourceID",
        "ConnectionString",
    ];
    for partition in &root.children {
        let values = exact_children(partition, "Partition", &names)?;
        let _ = i64_value(values[2])?;
        for index in [0usize, 1, 3, 4, 5] {
            if leaf_text(values[index])?.len() > MAX_XML_TEXT_BYTES {
                return Err(limit("partition field bytes"));
            }
        }
    }
    Ok(root.children.len())
}

fn validate_paths(files: &[XldmFileEntry]) -> Result<()> {
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

fn validate_allocations(
    bytes: &[u8],
    data_offset: usize,
    directory_offset: usize,
    files: &mut [XldmFileEntry],
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
        let marker = u32::from_le_bytes(bytes[payload_end..end].try_into().unwrap());
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

fn payload_slice<'a>(bytes: &'a [u8], entry: &XldmFileEntry) -> Result<&'a [u8]> {
    let start = checked_usize(entry.offset.0, "file offset")?;
    let size = checked_usize(entry.stored_size.0, "file size")?;
    let end = start
        .checked_add(size)
        .and_then(|value| value.checked_sub(CRC_SIZE))
        .ok_or_else(|| limit("payload range"))?;
    bytes
        .get(start..end)
        .ok_or_else(|| invalid("payload range is outside storage"))
}

fn parse_xml(xml: &str) -> Result<Node> {
    if xml.len() > MAX_DIRECTORY_BYTES {
        return Err(limit("XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(element)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(decoded.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::CData(_) => {
                return Err(invalid(
                    "DTDs, processing instructions, and CDATA are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(element: &BytesStart<'_>) -> Result<Node> {
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            attributes += 1;
        }
    }
    Ok(Node {
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}
fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
fn exact_children<'a>(root: &'a Node, root_name: &str, names: &[&str]) -> Result<Vec<&'a Node>> {
    if root.name != root_name || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid(format!(
            "expected attribute-free {root_name} element"
        )));
    }
    if root.children.len() != names.len() {
        return Err(invalid(format!("{root_name} has an invalid child count")));
    }
    for (child, expected) in root.children.iter().zip(names) {
        if child.name != *expected {
            return Err(invalid(format!("expected {expected} in {root_name}")));
        }
    }
    Ok(root.children.iter().collect())
}
fn leaf_text(node: &Node) -> Result<&str> {
    if node.attributes != 0 || !node.children.is_empty() {
        return Err(invalid(format!(
            "{} must be an attribute-free leaf",
            node.name
        )));
    }
    Ok(&node.text)
}
fn bool_value(node: &Node) -> Result<bool> {
    match leaf_text(node)?.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{} is not an XML boolean", node.name))),
    }
}
fn u64_value(node: &Node) -> Result<u64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_| invalid(format!("{} is not an unsigned 64-bit integer", node.name)))
}
fn u32_value(node: &Node) -> Result<u32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_| invalid(format!("{} is not an unsigned 32-bit integer", node.name)))
}
fn i64_value(node: &Node) -> Result<i64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_| invalid(format!("{} is not a signed 64-bit integer", node.name)))
}
fn i32_value(node: &Node) -> Result<i32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_| invalid(format!("{} is not a signed 32-bit integer", node.name)))
}

fn decode_xml(bytes: &[u8], require_utf16: bool) -> Result<(String, XldmXmlEncoding)> {
    if bytes.is_empty() {
        return Err(invalid("empty XML allocation"));
    }
    if bytes.starts_with(&BOM) {
        return Err(invalid("unexpected XML byte-order mark"));
    }
    if require_utf16 || bytes.get(1) == Some(&0) {
        if bytes.len() % 2 != 0 {
            return Err(invalid("odd-length UTF-16LE XML"));
        }
        let words: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        Ok((
            String::from_utf16(&words).map_err(xml_error)?,
            XldmXmlEncoding::Utf16Le,
        ))
    } else {
        Ok((
            std::str::from_utf8(bytes).map_err(xml_error)?.to_owned(),
            XldmXmlEncoding::Utf8,
        ))
    }
}
fn utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}
fn valid_upper_guid(value: &str) -> bool {
    if value.len() != 36 {
        return false;
    }
    value.bytes().enumerate().all(|(index, byte)| {
        if matches!(index, 8 | 13 | 18 | 23) {
            byte == b'-'
        } else {
            byte.is_ascii_digit() || (b'A'..=b'F').contains(&byte)
        }
    })
}
fn checked_usize(value: u64, name: &str) -> Result<usize> {
    usize::try_from(value).map_err(|_| limit(name))
}
fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for byte in bytes {
        let mut index = ((crc >> 24) ^ u32::from(*byte)) & 0xFF;
        let mut table = index << 24;
        for _ in 0..8 {
            table = if table & 0x8000_0000 != 0 {
                (table << 1) ^ 0x04C1_1DB7
            } else {
                table << 1
            };
        }
        index = table;
        crc = (crc << 8) ^ index;
    }
    crc
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("MS-XLDM {name} limit exceeded"))
}

#[cfg(test)]
pub(crate) fn test_xldm_bytes() -> Vec<u8> {
    let payload = b"compressed-looking model metadata";
    let log = test_backup_log("Model.1.db.xml", payload.len() as i32, 100002);
    build_test_storage(&[
        ("Partitions", partitions_xml().as_bytes()),
        ("Model.1.db.xml", payload),
        ("BackupLog", log.as_bytes()),
    ])
}

#[cfg(test)]
fn partitions_xml() -> String {
    "<Partitions><Partition><ObjectPath></ObjectPath><Name></Name><DataSize>0</DataSize><Location></Location><DataSourceID></DataSourceID><ConnectionString></ConnectionString></Partition></Partitions>".into()
}

#[cfg(test)]
fn test_backup_log(path: &str, size: i32, class: i32) -> String {
    format!(
        "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>Model</ObjectId><Write>ReadWrite</Write><OlapInfo>false</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups><FileGroup><Class>{class}</Class><ID>Model</ID><Name>Model</Name><ObjectVersion>1</ObjectVersion><PersistLocation>1</PersistLocation><PersistLocationPath></PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>11111111-2222-3333-4444-555555555555</ObjectID><FileList><BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile></FileList></FileGroup></FileGroups></BackupLog>"
    )
}

#[cfg(test)]
fn build_test_storage(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut bytes = vec![0; XLDM_PAGE_SIZE];
    bytes.extend_from_slice(&BOM);
    let mut allocations = Vec::new();
    for (index, (_, payload)) in entries.iter().enumerate() {
        if index + 1 == entries.len() {
            bytes.extend_from_slice(&BOM);
        }
        let offset = bytes.len();
        bytes.extend_from_slice(payload);
        bytes.extend_from_slice(&crc32(payload).to_le_bytes());
        allocations.push((offset, payload.len() + CRC_SIZE));
    }
    let directory_offset = bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE;
    bytes.resize(directory_offset, 0);
    let mut directory = String::from("<VirtualDirectory>");
    for ((path, _), (offset, size)) in entries.iter().zip(&allocations) {
        directory.push_str(&format!("<BackupFile><Path>{path}</Path><Size>{size}</Size><m_cbOffsetHeader>{offset}</m_cbOffsetHeader><Delete>false</Delete><CreatedTimestamp>0</CreatedTimestamp><Access>0</Access><LastWriteTime>0</LastWriteTime></BackupFile>"));
    }
    directory.push_str("</VirtualDirectory>");
    let directory_bytes = utf16le(&directory);
    bytes.extend_from_slice(&directory_bytes);
    bytes.resize(bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE, 0);
    let header = format!(
        "<BackupLog><BackupRestoreSyncVersion>140</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>0</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>0</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{}</DataSize><Files>{}</Files><ObjectID>01234567-89AB-CDEF-0123-456789ABCDEF</ObjectID><m_cbOffsetData>4096</m_cbOffsetData></BackupLog>",
        directory_bytes.len(),
        entries.len()
    );
    let mut page = Vec::new();
    page.extend_from_slice(&BOM);
    page.extend_from_slice(&utf16le(XLDM_STREAM_SIGNATURE));
    page.extend_from_slice(&utf16le(&header));
    page.resize(XLDM_PAGE_SIZE, 0);
    bytes[..XLDM_PAGE_SIZE].copy_from_slice(&page);
    bytes
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inspects_typed_storage_and_writes_byte_exactly() {
        let bytes = test_xldm_bytes();
        let storage = inspect_xldm(&bytes).unwrap();
        assert_eq!(storage.header.backup_restore_sync_version, 140);
        assert_eq!(storage.partition_marker.partition_count, 1);
        assert_eq!(storage.files.len(), 3);
        assert_eq!(storage.files[0].kind, XldmFileKind::Partitions);
        assert_eq!(storage.files[2].kind, XldmFileKind::BackupLog);
        assert_eq!(
            storage.backup_log.file_groups[0].class,
            XldmFileGroupClass::Database
        );
        assert_eq!(
            storage.backup_log.file_groups[0].files[0].generated.kind,
            XldmGeneratedNameKind::DatabaseDefinition
        );
        assert_eq!(
            storage.file_payload(1).unwrap(),
            b"compressed-looking model metadata"
        );
        assert_eq!(write_xldm(&storage).unwrap(), bytes);
    }

    #[test]
    fn rejects_header_signature_version_padding_and_offsets() {
        let base = test_xldm_bytes();
        for mutation in [0usize, 4] {
            let mut bytes = base.clone();
            bytes[mutation] ^= 1;
            assert!(inspect_xldm(&bytes).is_err());
        }
        let mut bytes = base;
        bytes[XLDM_PAGE_SIZE - 1] = 1;
        assert!(inspect_xldm(&bytes).is_err());
    }

    #[test]
    fn rejects_crc_corruption_without_interpreting_payload() {
        let mut bytes = test_xldm_bytes();
        let storage = inspect_xldm(&bytes).unwrap();
        let offset = storage.files[1].offset.0 as usize;
        bytes[offset] ^= 1;
        assert!(inspect_xldm(&bytes).is_err());
    }

    #[test]
    fn rejects_overlap_gap_and_nonzero_page_padding() {
        let mut bytes = test_xldm_bytes();
        let storage = inspect_xldm(&bytes).unwrap();
        let padding = storage.files.last().unwrap().offset.0 as usize
            + storage.files.last().unwrap().stored_size.0 as usize;
        bytes[padding] = 1;
        assert!(inspect_xldm(&bytes).is_err());
    }

    #[test]
    fn rejects_traversal_duplicate_and_ancestor_cycles() {
        for paths in [
            ["Partitions", "../escape", "BackupLog"],
            ["Partitions", "A/B", "a/b"],
            ["Partitions", "A", "A/B"],
        ] {
            let entries: Vec<_> = paths
                .iter()
                .enumerate()
                .map(|(index, path)| {
                    (
                        *path,
                        if index == 0 {
                            partitions_xml().as_bytes().to_vec()
                        } else {
                            vec![index as u8]
                        },
                    )
                })
                .collect();
            let refs: Vec<_> = entries
                .iter()
                .map(|(path, data)| (*path, data.as_slice()))
                .collect();
            assert!(inspect_xldm(&build_test_storage(&refs)).is_err());
        }
    }

    #[test]
    fn rejects_directory_count_and_partition_shape_mismatches() {
        let bad_partition = b"<Partitions><Wrong/></Partitions>";
        assert!(
            inspect_xldm(&build_test_storage(&[
                ("Partitions", bad_partition),
                ("Data", b"x"),
                ("BackupLog", b"log")
            ]))
            .is_err()
        );
    }

    #[test]
    fn classifies_every_generated_filename_family() {
        let cases = [
            ("Model.1.db.xml", XldmGeneratedNameKind::DatabaseDefinition),
            (
                "Model.1.db/View.2.dsv.xml",
                XldmGeneratedNameKind::DataSourceViewDefinition,
            ),
            (
                "Model.1.db/Cube.2.cub.xml",
                XldmGeneratedNameKind::CubeDefinition,
            ),
            (
                "Model.1.db/Cube.0.cub/info.3.xml",
                XldmGeneratedNameKind::CubeInformation,
            ),
            (
                "Model.1.db/Cube.0.cub/MdxScript.0.scr.xml",
                XldmGeneratedNameKind::MdxScriptMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det.xml",
                XldmGeneratedNameKind::MeasureGroupMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det/Table.3.prt.xml",
                XldmGeneratedNameKind::PartitionMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det/Table.4.prt/info.5.xml",
                XldmGeneratedNameKind::PartitionInformation,
            ),
            (
                "Model.1.db/Table.0.dim/Table.1.tbl.xml",
                XldmGeneratedNameKind::TableMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/R$Table$Rel.1.tbl.xml",
                XldmGeneratedNameKind::TableRelationshipMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/H$Table$Col.1.tbl.xml",
                XldmGeneratedNameKind::ColumnHierarchyMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/U$Table$Hier.1.tbl.xml",
                XldmGeneratedNameKind::UserHierarchyMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/1.Table.Col.0.idf",
                XldmGeneratedNameKind::ColumnData,
            ),
            (
                "Model.1.db/Table.0.dim/1.R$Table$Rel.INDEX.0.idf",
                XldmGeneratedNameKind::TableRelationshipIndex,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.POS_TO_ID.0.idf",
                XldmGeneratedNameKind::ColumnPositionToId,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.ID_TO_POS.0.idf",
                XldmGeneratedNameKind::ColumnIdToPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.hidx",
                XldmGeneratedNameKind::ColumnHashIndex,
            ),
            (
                "Model.1.db/Table.0.dim/1.Table.Col.dictionary",
                XldmGeneratedNameKind::ColumnDictionary,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.CHILD_COUNT.0.idf",
                XldmGeneratedNameKind::UserHierarchyChildCount,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.FIRST_CHILD_POS.0.idf",
                XldmGeneratedNameKind::UserHierarchyFirstChildPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.PARENT_POS.0.idf",
                XldmGeneratedNameKind::UserHierarchyParentPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.MULTI_LEVEL_ID.0.idf",
                XldmGeneratedNameKind::UserHierarchyMultilevelId,
            ),
        ];
        for (path, expected) in cases {
            assert_eq!(
                classify_xldm_generated_path(path).unwrap().kind,
                expected,
                "{path}"
            );
        }
    }

    #[test]
    fn rejects_backup_log_size_class_and_name_mismatches() {
        let payload = b"data";
        for (path, size, class) in [
            ("Model.1.db.xml", 3, 100002),
            ("Model.1.db.xml", 4, 100006),
            ("../escape", 4, 100002),
        ] {
            let log = test_backup_log(path, size, class);
            let bytes = build_test_storage(&[
                ("Partitions", partitions_xml().as_bytes()),
                ("Model.1.db.xml", payload),
                ("BackupLog", log.as_bytes()),
            ]);
            assert!(inspect_xldm(&bytes).is_err());
        }
    }

    #[test]
    fn rejects_malformed_backup_log_flags_and_enums() {
        for replacement in [
            ("<SvrEncryptPwdFlag>true", "<SvrEncryptPwdFlag>false"),
            ("<Write>ReadWrite", "<Write>Execute"),
            ("<Class>100002", "<Class>100004"),
        ] {
            let payload = b"data";
            let log =
                test_backup_log("Model.1.db.xml", 4, 100002).replace(replacement.0, replacement.1);
            let bytes = build_test_storage(&[
                ("Partitions", partitions_xml().as_bytes()),
                ("Model.1.db.xml", payload),
                ("BackupLog", log.as_bytes()),
            ]);
            assert!(inspect_xldm(&bytes).is_err());
        }
    }
}
