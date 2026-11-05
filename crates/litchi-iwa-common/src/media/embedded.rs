//! Archive-free inventory of the embedded media of an iWork package: every
//! `TSP.DataInfo` record with its materialized member and the references
//! that keep it alive (record 0375).
//!
//! A format owner feeds an [`Inventory`] with the package metadata's data
//! records and component data references, the `MessageInfo.data_references`
//! of every object and the `DataMetadataMap` entries, then [`Inventory::finish`]
//! resolves the members and returns the assets ordered by identifier.

use std::collections::{HashMap, HashSet};
use std::num::NonZeroU64;
use std::path::{Component, Path};

use super::Type;

/// Inventory failures.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A data identifier is 0.
    #[error("embedded media identifiers must be non-zero")]
    ZeroIdentifier,
    /// Two data records share an identifier.
    #[error("duplicate DataInfo identifier {0}")]
    DuplicateIdentifier(u64),
    /// A data record names an unsafe file.
    #[error("unsafe DataInfo filename {0:?}")]
    UnsafeFileName(String),
    /// A reference count overflowed.
    #[error("embedded media reference count overflow")]
    Overflow,
}

/// Result type of the inventory.
pub type Result<T> = std::result::Result<T, Error>;

/// The non-zero identifier of one embedded asset (`TSP.DataInfo.identifier`).
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct AssetId(NonZeroU64);

impl AssetId {
    /// Construct an identifier.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`].
    pub const fn new(raw: u64) -> Result<Self> {
        match NonZeroU64::new(raw) {
            Some(value) => Ok(Self(value)),
            None => Err(Error::ZeroIdentifier),
        }
    }

    /// The native identifier.
    #[must_use]
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

/// One embedded asset.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Embedded {
    data: AssetId,
    preferred_file_name: String,
    package_path: Option<String>,
    media_type: Type,
    size: Option<u64>,
    declared_size: Option<u64>,
    digest: Vec<u8>,
    component_reference_count: u64,
    component_reference_record_count: usize,
    message_reference_count: usize,
    has_data_metadata: bool,
}

impl Embedded {
    /// The identifier.
    #[must_use]
    pub const fn data(&self) -> AssetId {
        self.data
    }

    /// The name iWork prefers when exporting the asset.
    #[must_use]
    pub fn preferred_file_name(&self) -> &str {
        &self.preferred_file_name
    }

    /// The materialized `Data/` member, when present.
    #[must_use]
    pub fn package_path(&self) -> Option<&str> {
        self.package_path.as_deref()
    }

    /// The type of the materialized (else the preferred) file name.
    #[must_use]
    pub const fn media_type(&self) -> Type {
        self.media_type
    }

    /// The materialized length.
    #[must_use]
    pub const fn size(&self) -> Option<u64> {
        self.size
    }

    /// The length the data record declares.
    #[must_use]
    pub const fn declared_size(&self) -> Option<u64> {
        self.declared_size
    }

    /// The declared digest (normally SHA-1).
    #[must_use]
    pub fn digest(&self) -> &[u8] {
        &self.digest
    }

    /// The sum of component-level object reference counts.
    #[must_use]
    pub const fn component_reference_count(&self) -> u64 {
        self.component_reference_count
    }

    /// The number of component data-reference records, zero-count ones
    /// included.
    #[must_use]
    pub const fn component_reference_record_count(&self) -> usize {
        self.component_reference_record_count
    }

    /// The number of `MessageInfo.data_references` occurrences.
    #[must_use]
    pub const fn message_reference_count(&self) -> usize {
        self.message_reference_count
    }

    /// Whether the `DataMetadataMap` lists the asset.
    #[must_use]
    pub const fn has_data_metadata(&self) -> bool {
        self.has_data_metadata
    }

    /// Whether the asset has a materialized member.
    #[must_use]
    pub const fn is_materialized(&self) -> bool {
        self.package_path.is_some() && self.size.is_some()
    }

    /// Whether anything references the asset.
    #[must_use]
    pub const fn is_referenced(&self) -> bool {
        self.component_reference_record_count != 0
            || self.message_reference_count != 0
            || self.has_data_metadata
    }
}

struct DataRecord {
    data: AssetId,
    preferred_file_name: String,
    file_name: Option<String>,
    declared_size: Option<u64>,
    digest: Vec<u8>,
}

/// Collects the embedded-media facts of one package.
#[derive(Default)]
pub struct Inventory {
    records: Vec<DataRecord>,
    identifiers: HashSet<u64>,
    component_counts: HashMap<u64, u64>,
    component_records: HashMap<u64, usize>,
    message_counts: HashMap<u64, usize>,
    data_metadata: HashSet<u64>,
}

/// The `Data/` member a data-record file name names.
///
/// # Errors
///
/// [`Error::UnsafeFileName`] for an empty, absolute or escaping name.
pub fn data_entry_name(file_name: &str) -> Result<String> {
    let path = Path::new(file_name);
    if file_name.is_empty()
        || file_name.contains(['\0', '\\'])
        || path.is_absolute()
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(Error::UnsafeFileName(file_name.to_owned()));
    }
    Ok(format!("Data/{file_name}"))
}

impl Inventory {
    /// An empty inventory.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// One `TSP.DataInfo` record.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`] or [`Error::DuplicateIdentifier`].
    pub fn data_info(
        &mut self,
        identifier: u64,
        preferred_file_name: &str,
        file_name: Option<&str>,
        declared_size: Option<u64>,
        digest: &[u8],
    ) -> Result<()> {
        let data = AssetId::new(identifier)?;
        if !self.identifiers.insert(identifier) {
            return Err(Error::DuplicateIdentifier(identifier));
        }
        self.records.push(DataRecord {
            data,
            preferred_file_name: preferred_file_name.to_owned(),
            file_name: file_name.map(str::to_owned),
            declared_size,
            digest: digest.to_vec(),
        });
        Ok(())
    }

    /// One component data-reference record.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`] or [`Error::Overflow`].
    pub fn component_reference(&mut self, identifier: u64) -> Result<()> {
        AssetId::new(identifier)?;
        let records = self.component_records.entry(identifier).or_insert(0);
        *records = records.checked_add(1).ok_or(Error::Overflow)?;
        self.component_counts.entry(identifier).or_insert(0);
        Ok(())
    }

    /// One owner (object reference count) of a component data reference.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`] or [`Error::Overflow`].
    pub fn component_owner(&mut self, identifier: u64, count: u32) -> Result<()> {
        AssetId::new(identifier)?;
        let total = self.component_counts.entry(identifier).or_insert(0);
        *total = total.checked_add(u64::from(count)).ok_or(Error::Overflow)?;
        Ok(())
    }

    /// One `MessageInfo.data_references` occurrence.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`] or [`Error::Overflow`].
    pub fn message_reference(&mut self, identifier: u64) -> Result<()> {
        AssetId::new(identifier)?;
        let count = self.message_counts.entry(identifier).or_insert(0);
        *count = count.checked_add(1).ok_or(Error::Overflow)?;
        Ok(())
    }

    /// One `DataMetadataMap` entry.
    ///
    /// # Errors
    ///
    /// [`Error::ZeroIdentifier`].
    pub fn data_metadata(&mut self, identifier: u64) -> Result<()> {
        AssetId::new(identifier)?;
        self.data_metadata.insert(identifier);
        Ok(())
    }

    /// Resolve the members (`member_size` gives a package member's length)
    /// and return the assets ordered by identifier.
    ///
    /// # Errors
    ///
    /// [`Error::UnsafeFileName`].
    pub fn finish(self, member_size: impl Fn(&str) -> Option<u64>) -> Result<Vec<Embedded>> {
        let mut assets = Vec::with_capacity(self.records.len());
        for record in self.records {
            let package_path = record
                .file_name
                .as_deref()
                .filter(|name| !name.is_empty())
                .map(data_entry_name)
                .transpose()?
                .filter(|path| member_size(path).is_some());
            let size = package_path.as_deref().and_then(&member_size);
            let type_name = package_path
                .as_deref()
                .and_then(|path| Path::new(path).file_name())
                .and_then(|name| name.to_str())
                .unwrap_or(&record.preferred_file_name);
            let media_type = Path::new(type_name)
                .extension()
                .and_then(|extension| extension.to_str())
                .map_or(Type::Unknown, Type::from_extension);
            let id = record.data.get();
            assets.push(Embedded {
                data: record.data,
                preferred_file_name: record.preferred_file_name,
                package_path,
                media_type,
                size,
                declared_size: record.declared_size,
                digest: record.digest,
                component_reference_count: self.component_counts.get(&id).copied().unwrap_or(0),
                component_reference_record_count: self
                    .component_records
                    .get(&id)
                    .copied()
                    .unwrap_or(0),
                message_reference_count: self.message_counts.get(&id).copied().unwrap_or(0),
                has_data_metadata: self.data_metadata.contains(&id),
            });
        }
        assets.sort_unstable_by_key(|asset| asset.data);
        Ok(assets)
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Inventory, data_entry_name};
    use crate::media::Type;

    #[test]
    fn inventories_resolve_members_and_references() {
        let mut inventory = Inventory::new();
        inventory
            .data_info(7, "photo.png", Some("photo-7.png"), Some(10), b"d")
            .unwrap();
        inventory.data_info(3, "clip.mov", None, None, b"").unwrap();
        assert_eq!(
            inventory.data_info(3, "x", None, None, b""),
            Err(Error::DuplicateIdentifier(3))
        );
        inventory.component_reference(7).unwrap();
        inventory.component_owner(7, 2).unwrap();
        inventory.message_reference(3).unwrap();
        let assets = inventory
            .finish(|path| (path == "Data/photo-7.png").then_some(10))
            .unwrap();
        assert_eq!(assets[0].data().get(), 3);
        assert!(!assets[0].is_materialized() && assets[0].is_referenced());
        assert_eq!(assets[0].media_type(), Type::Video);
        assert!(assets[1].is_materialized() && assets[1].is_referenced());
        assert_eq!(assets[1].component_reference_count(), 2);
        assert!(data_entry_name("../x").is_err());
    }
}
