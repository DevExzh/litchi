//! Source-backed models for non-simple OLE Property Set storage.

use super::codec;
use crate::object::directory::Metadata;
use litchi_cfb::{OleError, SharedOleFile, directory_names_equal};
use std::sync::Arc;

/// Bounds applied by the non-simple Property Set owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum complete CFB source size accepted by the shared reader.
    pub max_input_bytes: u64,
    /// Maximum CFB directory stream size.
    pub max_directory_bytes: u64,
    /// Maximum decoded FAT/DIFAT/MiniFAT table size.
    pub max_allocation_table_bytes: u64,
    /// Maximum number of CFB directory entries retained in the catalog.
    pub max_directory_entries: usize,
    /// Maximum UTF-8 bytes in one directory name or CLSID spelling.
    pub max_name_bytes: usize,
    /// Maximum aggregate UTF-8 metadata bytes retained by the catalog.
    pub max_total_name_bytes: usize,
    /// Maximum aggregate UTF-8 bytes retained by copied directory paths and
    /// bounded path indexes.
    pub max_total_path_bytes: usize,
    /// Maximum storage nesting depth below the CFB root.
    pub max_storage_depth: usize,
    /// Maximum number of stream elements.
    pub max_streams: usize,
    /// Maximum size of one stream element.
    pub max_stream_bytes: u64,
    /// Maximum aggregate logical stream bytes in one source.
    pub max_total_stream_bytes: u64,
    /// Maximum size of the required `CONTENTS` stream.
    pub max_contents_bytes: u64,
    /// Maximum property descriptors admitted by the owner.
    pub max_properties: usize,
    /// Maximum indirect-property references admitted by one contents stream.
    pub max_indirect_properties: usize,
    /// Maximum rendered candidate size.
    pub max_output_bytes: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_directory_bytes: 64 * 1024 * 1024,
            max_allocation_table_bytes: 64 * 1024 * 1024,
            max_directory_entries: 65_536,
            max_name_bytes: 256,
            max_total_name_bytes: 64 * 1024 * 1024,
            max_total_path_bytes: 64 * 1024 * 1024,
            max_storage_depth: 64,
            max_streams: 65_536,
            max_stream_bytes: 128 * 1024 * 1024,
            max_total_stream_bytes: 512 * 1024 * 1024,
            max_contents_bytes: 16 * 1024 * 1024,
            max_properties: 16_384,
            max_indirect_properties: 16_384,
            max_output_bytes: 512 * 1024 * 1024,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<(), OleError> {
        if self.max_input_bytes == 0
            || self.max_directory_bytes == 0
            || self.max_allocation_table_bytes == 0
            || self.max_directory_entries == 0
            || self.max_name_bytes == 0
            || self.max_total_name_bytes == 0
            || self.max_total_path_bytes == 0
            || self.max_storage_depth == 0
            || self.max_streams == 0
            || self.max_stream_bytes == 0
            || self.max_total_stream_bytes == 0
            || self.max_contents_bytes == 0
            || self.max_properties == 0
            || self.max_indirect_properties == 0
            || self.max_output_bytes == 0
        {
            return Err(OleError::InvalidFormat(
                "non-simple Property Set limits must be non-zero".into(),
            ));
        }
        Ok(())
    }

    pub(crate) fn cfb_limits(self) -> Result<litchi_cfb::SharedOleFileLimits, OleError> {
        litchi_cfb::SharedOleFileLimits::new(self.max_input_bytes)?
            .with_max_directory_bytes(self.max_directory_bytes)?
            .with_max_allocation_table_bytes(self.max_allocation_table_bytes)
    }
}

/// Physical element kind required by one indirect Property Set value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ElementKind {
    /// A CFB stream.
    Stream,
    /// A CFB storage, including an empty storage.
    Storage,
}

/// A bounded physical CFB element descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Element {
    path: Vec<String>,
    kind: Option<ElementKind>,
    metadata: Option<Metadata>,
    size: u64,
}

impl Element {
    pub(crate) fn new(
        path: Vec<String>,
        kind: Option<ElementKind>,
        metadata: Option<Metadata>,
        size: u64,
    ) -> Self {
        Self {
            path,
            kind,
            metadata,
            size,
        }
    }

    /// The physical CFB path relative to the non-simple root storage.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// The typed kind, or `None` for an unrecognized source directory kind.
    #[must_use]
    pub const fn kind(&self) -> Option<ElementKind> {
        self.kind
    }

    /// Source directory metadata for a known CFB element kind.
    #[must_use]
    pub const fn metadata(&self) -> Option<Metadata> {
        self.metadata
    }

    /// Declared logical stream size, or zero for a storage.
    #[must_use]
    pub const fn size(&self) -> u64 {
        self.size
    }
}

#[derive(Debug, Clone)]
pub(crate) struct EntryRecord {
    pub(crate) path: Vec<String>,
    pub(crate) kind: Option<ElementKind>,
    pub(crate) metadata: Option<Metadata>,
    pub(crate) raw_kind: u8,
    pub(crate) size: u64,
}

impl EntryRecord {
    pub(crate) fn try_element(&self) -> Result<Element, OleError> {
        let mut path = Vec::new();
        path.try_reserve_exact(self.path.len())
            .map_err(|source| OleError::Allocation {
                resource: "non-simple Property Set element view path",
                source,
            })?;
        for component in &self.path {
            let mut owned = String::new();
            owned
                .try_reserve_exact(component.len())
                .map_err(|source| OleError::Allocation {
                    resource: "non-simple Property Set element view name",
                    source,
                })?;
            owned.push_str(component);
            path.push(owned);
        }
        Ok(Element::new(path, self.kind, self.metadata, self.size))
    }

    pub(crate) fn is_root(&self) -> bool {
        self.path.is_empty()
    }
}

/// Immutable source state for one non-simple Property Set CFB payload.
#[derive(Debug)]
pub(crate) struct State {
    pub(crate) source: Arc<[u8]>,
    pub(crate) cfb: Arc<SharedOleFile>,
    pub(crate) records: Arc<[EntryRecord]>,
    pub(crate) root_metadata: Metadata,
    pub(crate) root_name: String,
    pub(crate) limits: Limits,
}

/// A lazy, source-preserving non-simple Property Set storage snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub(crate) state: Arc<State>,
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.state.source.as_ref() == other.state.source.as_ref()
    }
}

impl Eq for Snapshot {}

impl Snapshot {
    /// Opens an owned compound-file payload with explicit finite limits.
    pub fn open(source: Vec<u8>, limits: Limits) -> Result<Self, OleError> {
        limits.validate()?;
        let source_len = source.len() as u64;
        if source_len > limits.max_input_bytes {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set input bytes",
                observed: source_len,
                maximum: limits.max_input_bytes,
            });
        }
        Self::open_shared(source.into(), limits)
    }

    /// Opens an immutable shared compound-file payload without copying it.
    pub fn open_shared(source: Arc<[u8]>, limits: Limits) -> Result<Self, OleError> {
        limits.validate()?;
        Ok(Self {
            state: codec::capture(source, limits)?,
        })
    }

    /// Exact source bytes retained by this snapshot.
    #[must_use]
    pub fn source_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.state.source)
    }

    /// The configured owner limits.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.limits
    }

    /// Source root-storage directory metadata.
    #[must_use]
    pub fn root_metadata(&self) -> Metadata {
        self.state.root_metadata
    }

    /// Returns every source directory element, including empty storages.
    pub fn elements(&self) -> Result<Vec<Element>, OleError> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.state.records.len())
            .map_err(|source| OleError::Allocation {
                resource: "non-simple Property Set element views",
                source,
            })?;
        for record in self.state.records.iter().filter(|record| !record.is_root()) {
            output.push(record.try_element()?);
        }
        Ok(output)
    }

    /// Reads and validates the required `CONTENTS` PropertySetStream lazily.
    pub fn contents(&self) -> Result<crate::property_set::Stream, OleError> {
        codec::read_contents(&self.state)
    }

    /// Reads one source stream lazily and shares the resulting allocation.
    pub fn stream(&self, path: &[&str]) -> Result<Option<Arc<[u8]>>, OleError> {
        let Some(record) = codec::find_record(&self.state.records, path) else {
            return Ok(None);
        };
        if record.kind != Some(ElementKind::Stream) {
            return Err(OleError::InvalidFormat(format!(
                "non-simple element {path:?} is not a stream"
            )));
        }
        let bytes = codec::read_stream(&self.state, record)?;
        Ok(Some(bytes))
    }

    /// Returns the bounded source descriptor for one physical element.
    pub fn element(&self, path: &[&str]) -> Result<Option<Element>, OleError> {
        match codec::find_record(&self.state.records, path) {
            Some(record) => Ok(Some(record.try_element()?)),
            None => Ok(None),
        }
    }

    /// Starts an isolated source-bound edit.
    #[must_use]
    pub fn edit(&self) -> crate::property_set::non_simple::Editor {
        crate::property_set::non_simple::Editor::new(self.clone())
    }
}

pub(crate) fn same_path(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| directory_names_equal(left, right))
}
