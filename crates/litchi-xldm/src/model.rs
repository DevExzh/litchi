//! Public MS-XLDM storage model and bounded parsing primitives.

use crate::error::{Error, Result};

pub const XLDM_PAGE_SIZE: usize = 4096;
pub const XLDM_STREAM_SIGNATURE: &str = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";

pub(super) const MAX_STORAGE_BYTES: usize = 512 * 1024 * 1024;
pub(super) const MAX_DIRECTORY_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_XML_NODES: usize = 500_000;
pub(super) const MAX_XML_DEPTH: usize = 128;
pub(super) const MAX_XML_TEXT_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_FILES: usize = 100_000;
pub(super) const MAX_PARTITIONS: usize = 65_536;
pub(super) const MAX_PATH_BYTES: usize = 32 * 1024;
pub(super) const CRC_SIZE: usize = 4;
pub(super) const BOM: [u8; 2] = [0xFF, 0xFE];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XmlEncoding {
    Utf16Le,
    Utf8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Compression {
    Xpress,
}

/// Storage rules selected from the header, retained through exact writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StorageProfile {
    /// Canonical MS-XLDM version-140 storage and metadata.
    Xldm140,
    /// Version-150 tabular compatibility rules backed by the native fixture.
    Tabular150,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileKind {
    Partitions,
    BackupLog,
    CryptographicKey,
    XmlMetadata,
    OpaqueBinary,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum WriteAccess {
    ReadWrite,
    ReadOnly,
    ReadOnlyExclusive,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FileGroupClass {
    Database,
    DataSource,
    Dimension,
    Cube,
    MeasureGroup,
    Partition,
    DataSourceView,
    MdxScript,
}

impl FileGroupClass {
    #[must_use]
    pub const fn code(self) -> i32 {
        match self {
            Self::Database => 100_002,
            Self::DataSource => 100_003,
            Self::Dimension => 100_006,
            Self::Cube => 100_010,
            Self::MeasureGroup => 100_016,
            Self::Partition => 100_021,
            Self::DataSourceView => 100_053,
            Self::MdxScript => 100_060,
        }
    }
    pub(super) fn parse(value: i32) -> Result<Self> {
        match value {
            100_002 => Ok(Self::Database),
            100_003 => Ok(Self::DataSource),
            100_006 => Ok(Self::Dimension),
            100_010 => Ok(Self::Cube),
            100_016 => Ok(Self::MeasureGroup),
            100_021 => Ok(Self::Partition),
            100_053 => Ok(Self::DataSourceView),
            100_060 => Ok(Self::MdxScript),
            _ => Err(Error::Invalid(format!(
                "unknown MS-XLDM file-group class {value}"
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum GeneratedNameKind {
    CryptographicKey,
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
pub struct GeneratedPath {
    /// Logical path relative to the model's storage root. In the tabular profile
    /// this comes from the source Path, separately from the opaque StoragePath key.
    pub normalized_path: String,
    pub kind: GeneratedNameKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoggedFile {
    pub source_path: String,
    pub storage_path: String,
    pub last_write_timestamp: i64,
    /// Stored payload bytes for Xldm140; decoded member bytes for Tabular150.
    pub size: u32,
    pub generated: GeneratedPath,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileGroup {
    pub class: FileGroupClass,
    pub id: String,
    pub name: String,
    pub object_version: i32,
    pub persist_location: i32,
    pub persist_location_path: String,
    /// Unused by MS-XLDM and retained without interpretation.
    pub storage_location_path: String,
    pub object_id: String,
    pub files: Vec<LoggedFile>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackupLog {
    /// Semantic version 1153. Tabular150's lexical spelling is 11.53; the
    /// original spelling remains in the retained storage bytes.
    pub backup_restore_sync_version: i32,
    /// Originating filesystem root; retained as text and never accessed.
    pub server_root: String,
    pub object_name: String,
    pub object_id: String,
    pub write_access: WriteAccess,
    pub is_olap: bool,
    pub collations: Vec<String>,
    pub languages: Vec<i32>,
    pub file_groups: Vec<FileGroup>,
    pub encoding: XmlEncoding,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct Offset(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct Size(pub u64);

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header {
    pub backup_restore_sync_version: i32,
    pub fault_code: u32,
    pub encryption_key_version: i32,
    pub compression: Compression,
    pub directory_offset: Offset,
    pub directory_size: Size,
    pub file_count: u32,
    pub object_id: String,
    pub data_offset: Offset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PartitionMarker<'a> {
    pub partition_count: usize,
    pub encoding: XmlEncoding,
    /// Encoded XML is retained exactly; unused fields, including connection
    /// strings, are not projected or acted upon.
    pub encoded_xml: &'a [u8],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileEntry {
    pub path: String,
    pub kind: FileKind,
    pub offset: Offset,
    /// Stored size includes the four-byte CRC marker.
    pub stored_size: Size,
    pub crc32: u32,
    pub delete: bool,
    pub created_timestamp: i64,
    pub access_timestamp: i64,
    pub last_write_timestamp: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Storage<'a> {
    pub header: Header,
    pub header_encoding: XmlEncoding,
    pub directory_encoding: XmlEncoding,
    pub partition_marker: PartitionMarker<'a>,
    pub backup_log: BackupLog,
    pub files: Vec<FileEntry>,
    pub(super) bytes: &'a [u8],
    pub(super) profile: StorageProfile,
}

impl Storage<'_> {
    #[must_use]
    pub fn profile(&self) -> StorageProfile {
        self.profile
    }

    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes
    }

    /// Return the exact source buffer retained by this borrowed inspection.
    ///
    /// The slice points into the caller-owned input; this owner never takes
    /// ownership of or copies the source during inspection.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.bytes
    }

    /// Return one member's bytes without its CRC marker. The returned bytes
    /// remain compressed, encrypted, or otherwise encoded exactly as stored.
    #[must_use]
    pub fn file_payload(&self, index: usize) -> Option<&[u8]> {
        let entry = self.files.get(index)?;
        let start = usize::try_from(entry.offset.0).ok()?;
        let size = usize::try_from(entry.stored_size.0).ok()?;
        self.bytes
            .get(start..start.checked_add(size)?.checked_sub(CRC_SIZE)?)
    }

    #[must_use]
    pub fn file_stored_bytes(&self, index: usize) -> Option<&[u8]> {
        let entry = self.files.get(index)?;
        let start = usize::try_from(entry.offset.0).ok()?;
        let size = usize::try_from(entry.stored_size.0).ok()?;
        self.bytes.get(start..start.checked_add(size)?)
    }
}

#[cfg(test)]
pub(crate) fn test_xldm140_storage<'a>(bytes: &'a [u8], paths: &[&str]) -> Storage<'a> {
    Storage {
        header: Header {
            backup_restore_sync_version: 1153,
            fault_code: 0,
            encryption_key_version: 0,
            compression: Compression::Xpress,
            directory_offset: Offset(0),
            directory_size: Size(0),
            file_count: u32::try_from(paths.len()).expect("test directory fits u32"),
            object_id: String::new(),
            data_offset: Offset(0),
        },
        header_encoding: XmlEncoding::Utf16Le,
        directory_encoding: XmlEncoding::Utf16Le,
        partition_marker: PartitionMarker {
            partition_count: 0,
            encoding: XmlEncoding::Utf16Le,
            encoded_xml: &[],
        },
        backup_log: BackupLog {
            backup_restore_sync_version: 1153,
            server_root: String::new(),
            object_name: String::new(),
            object_id: String::new(),
            write_access: WriteAccess::ReadWrite,
            is_olap: false,
            collations: Vec::new(),
            languages: Vec::new(),
            file_groups: Vec::new(),
            encoding: XmlEncoding::Utf16Le,
        },
        files: paths
            .iter()
            .map(|path| FileEntry {
                path: (*path).to_owned(),
                kind: FileKind::OpaqueBinary,
                offset: Offset(0),
                stored_size: Size(0),
                crc32: 0,
                delete: false,
                created_timestamp: 0,
                access_timestamp: 0,
                last_write_timestamp: 0,
            })
            .collect(),
        bytes,
        profile: StorageProfile::Xldm140,
    }
}

#[derive(Clone)]
pub(super) struct Node {
    pub(super) name: String,
    pub(super) attributes: usize,
    pub(super) children: Vec<Node>,
    pub(super) text: String,
}
