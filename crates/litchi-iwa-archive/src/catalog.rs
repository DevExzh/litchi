use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_iwa_core::{Archive, SnappyStream};

use crate::package::{Catalog, SourceProvenance};
use crate::zip::{ZipArchive, is_iwa_name, parse_directory_index_components, parse_iwa_components};
use crate::{Limits, Result};

/// One parsed `.iwa` component in deterministic member-name order.
#[derive(Debug)]
pub struct Component {
    name: Box<str>,
    archive: Archive,
}

impl Component {
    pub(crate) fn new(name: &str, archive: Archive) -> Self {
        Self {
            name: name.into(),
            archive,
        }
    }

    /// Return the normalized ZIP member name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the neutral parsed IWA archive.
    #[must_use]
    pub const fn archive(&self) -> &Archive {
        &self.archive
    }

    /// Consume the component and return its owned name and archive.
    #[must_use]
    pub fn into_parts(self) -> (String, Archive) {
        (self.name.into(), self.archive)
    }
}

/// Deterministic parsed `.iwa` components from one physical iWork ZIP input.
///
/// This catalog owns only ZIP/IWA ingress. Metadata, media, package
/// transactions, object indexing, and application-specific message decoding
/// remain in their respective adapter crates.
#[derive(Debug)]
pub struct ComponentCatalog {
    components: Box<[Component]>,
}

impl ComponentCatalog {
    /// Parse a ZIP bundle from memory using the default physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP, nested `Index.zip`, Snappy stream, or IWA
    /// framing is malformed, encrypted, or exceeds a resource ceiling.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a ZIP bundle from memory under explicit physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the limits are invalid, input exceeds a ceiling,
    /// or the ZIP/IWA stream is malformed or encrypted.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let archive = ZipArchive::new_with_limits(bytes, validated_limits)?;
        let components = parse_iwa_components(&archive, validated_limits)?.into_boxed_slice();
        Ok(Self { components })
    }

    pub(crate) fn from_directory_index_zip(bytes: &[u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let input_size = u64::try_from(bytes.len()).map_err(|_error| {
            crate::Error::InvalidBundle(
                "directory bundle Index.zip length does not fit u64".to_owned(),
            )
        })?;
        validated_limits.check_input_size(input_size, "directory bundle Index.zip")?;
        let archive = ZipArchive::new_with_limits(bytes, validated_limits)?;
        let components =
            parse_directory_index_components(&archive, validated_limits)?.into_boxed_slice();
        Ok(Self { components })
    }

    pub(crate) fn from_logical_entries<'a>(
        entries: impl IntoIterator<Item = (&'a str, &'a [u8])>,
        limits: Limits,
    ) -> Result<Self> {
        Self::from_validated_logical_entries(entries, 0, limits)
    }

    pub(crate) fn from_validated_logical_entries<'a>(
        entries: impl IntoIterator<Item = (&'a str, &'a [u8])>,
        component_capacity: usize,
        limits: Limits,
    ) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let mut components = Vec::new();
        components
            .try_reserve_exact(component_capacity)
            .map_err(|_error| crate::Error::Allocation {
                resource: "logical IWA component catalog",
                amount: component_capacity,
            })?;
        let mut decompressed_iwa_bytes = 0;
        for (name, data) in entries {
            if !is_iwa_name(name) {
                continue;
            }
            if let Some((component, decompressed_bytes)) =
                parse_component(name, data, validated_limits)?
            {
                decompressed_iwa_bytes = validated_limits
                    .charge_iwa_total_bytes(decompressed_iwa_bytes, decompressed_bytes)?;
                if components.len() == components.capacity() {
                    components
                        .try_reserve(1)
                        .map_err(|_error| crate::Error::Allocation {
                            resource: "logical IWA component catalog",
                            amount: 1,
                        })?;
                }
                components.push(component);
            }
        }
        components.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        Ok(Self {
            components: components.into_boxed_slice(),
        })
    }

    /// Return the number of parsed components.
    #[must_use]
    pub fn len(&self) -> usize {
        self.components.len()
    }

    /// Return whether no IWA components were found.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.components.is_empty()
    }

    /// Iterate over components in deterministic normalized-name order.
    pub fn iter(&self) -> impl Iterator<Item = &Component> {
        self.components.iter()
    }

    /// Find one component by normalized ZIP member name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Component> {
        self.components
            .binary_search_by(|component| component.name().cmp(name))
            .ok()
            .map(|index| &self.components[index])
    }

    /// Return the component at a compact catalog ordinal, if it exists.
    #[must_use]
    pub fn get_index(&self, index: usize) -> Option<&Component> {
        self.components.get(index)
    }

    fn from_package_catalog(catalog: &Catalog, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let mut components = Vec::new();
        let mut decompressed_iwa_bytes = 0;
        for entry in catalog.iter() {
            if !is_iwa_name(entry.name()) {
                continue;
            }
            if entry.is_opaque() {
                return Err(crate::Error::InvalidBundle(format!(
                    "IWA component {} uses an unsupported ZIP compression method",
                    entry.name()
                )));
            }
            if let Some((component, decompressed_bytes)) =
                parse_component(entry.name(), entry.data(), validated_limits)?
            {
                decompressed_iwa_bytes = validated_limits
                    .charge_iwa_total_bytes(decompressed_iwa_bytes, decompressed_bytes)?;
                components
                    .try_reserve(1)
                    .map_err(|_error| crate::Error::Allocation {
                        resource: "IWA component catalog",
                        amount: 1,
                    })?;
                components.push(component);
            }
        }
        components.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        Ok(Self {
            components: components.into_boxed_slice(),
        })
    }
}

impl IntoIterator for ComponentCatalog {
    type Item = Component;
    type IntoIter = std::vec::IntoIter<Component>;

    fn into_iter(self) -> Self::IntoIter {
        self.components.into_vec().into_iter()
    }
}

/// One immutable physical package snapshot and its parsed IWA components.
///
/// This is the shared ingress boundary for format owners that need both raw
/// package members (metadata, editing, or exact preservation) and semantic IWA
/// components. The ZIP envelope is traversed once: components are decoded from
/// the [`Catalog`]'s already-materialized logical entries instead of reopening
/// the source bytes through a second ZIP reader.
#[derive(Debug)]
pub struct SourceCatalog {
    package: Catalog,
    components: ComponentCatalog,
    limits: Limits,
}

impl SourceCatalog {
    /// Parse borrowed package bytes using the default physical limits.
    ///
    /// The source is copied exactly once into immutable shared storage.
    ///
    /// # Errors
    ///
    /// Returns an error when ZIP ingress, legacy normalization, Snappy/IWA
    /// decoding, or a configured physical ceiling fails.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse borrowed package bytes under caller-selected physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source or any package component is malformed,
    /// encrypted, ambiguous, or over budget.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let package = Catalog::from_bytes_with_limits(bytes, limits)?;
        Self::from_package(package, limits)
    }

    /// Parse an already-owned immutable package source without copying it.
    ///
    /// # Errors
    ///
    /// Returns an error when ZIP ingress, legacy normalization, Snappy/IWA
    /// decoding, or a configured physical ceiling fails.
    pub fn from_shared_bytes(source: Arc<[u8]>) -> Result<Self> {
        Self::from_shared_bytes_with_limits(source, Limits::default())
    }

    /// Parse an already-owned immutable package source under explicit limits.
    ///
    /// The exact [`Arc`] allocation remains authoritative for the lifetime of
    /// this snapshot and can be reused by preserve-mode writes.
    ///
    /// # Errors
    ///
    /// Returns an error when the source or any package component is malformed,
    /// encrypted, ambiguous, or over budget.
    pub fn from_shared_bytes_with_limits(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        let package = Catalog::from_shared_bytes_with_limits(source, limits)?;
        Self::from_package(package, limits)
    }

    /// Snapshot an immutable positional source using the default limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source changes during the bounded read or the
    /// resulting package cannot be decoded.
    pub fn from_read_at(source: &dyn ReadAt) -> Result<Self> {
        Self::from_read_at_with_limits(source, Limits::default())
    }

    /// Snapshot an immutable positional source under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source changes during the bounded read or the
    /// resulting package cannot be decoded.
    pub fn from_read_at_with_limits(source: &dyn ReadAt, limits: Limits) -> Result<Self> {
        let package = Catalog::from_read_at_with_limits(source, limits)?;
        Self::from_package(package, limits)
    }

    fn from_package(package: Catalog, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let components = ComponentCatalog::from_package_catalog(&package, validated_limits)?;
        Ok(Self {
            package,
            components,
            limits: validated_limits,
        })
    }

    /// Return the checked physical profile that authorized this snapshot.
    ///
    /// A downstream format projection must derive its physical assumptions
    /// from this value rather than accepting a second, potentially weaker
    /// profile after ZIP, Snappy, and IWA validation has already completed.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Borrow the physical/logical package catalog retained by this snapshot.
    #[must_use]
    pub const fn package(&self) -> &Catalog {
        &self.package
    }

    /// Borrow parsed IWA components in deterministic normalized-name order.
    #[must_use]
    pub const fn components(&self) -> &ComponentCatalog {
        &self.components
    }

    /// Clone the authoritative immutable source handle without copying bytes.
    #[must_use]
    pub fn shared_source(&self) -> Arc<[u8]> {
        self.package.shared_source()
    }

    /// Borrow the authoritative physical source bytes.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.package.source_bytes()
    }

    /// Return whether logical members still describe the source ZIP exactly.
    #[must_use]
    pub const fn source_is_exact(&self) -> bool {
        self.package.source_is_exact()
    }

    /// Return whether this snapshot is exact or normalized from legacy ZIP.
    #[must_use]
    pub const fn source_provenance(&self) -> SourceProvenance {
        self.package.source_provenance()
    }

    /// Consume this snapshot without cloning either catalog.
    #[must_use]
    pub fn into_parts(self) -> (Catalog, ComponentCatalog) {
        (self.package, self.components)
    }

    /// Consume this snapshot and retain only its parsed IWA components.
    ///
    /// The physical package catalog, including its authoritative source bytes
    /// and decoded logical entries, is dropped before this method returns.
    /// This is useful for read-only semantic projections that do not provide
    /// preserve-mode writing or package metadata access.
    #[must_use]
    pub fn into_components(self) -> ComponentCatalog {
        self.components
    }
}

pub(crate) fn parse_component(
    name: &str,
    compressed_data: &[u8],
    limits: Limits,
) -> Result<Option<(Component, u64)>> {
    // OperationStorage is a separate persistence format despite its `.iwa`
    // suffix in legacy documents. It remains a raw package member but is not
    // part of the IWA object graph.
    if name.rsplit('/').next() == Some("OperationStorage.iwa")
        && compressed_data.starts_with(b"bvxn")
    {
        return Ok(None);
    }

    let decompressed =
        SnappyStream::decompress_with_limits(compressed_data, limits.snappy_limits()?)?;
    let decompressed_bytes = u64::try_from(decompressed.as_bytes().len()).map_err(|_error| {
        crate::Error::InvalidBundle("decompressed IWA stream length does not fit u64".to_owned())
    })?;
    let archive =
        Archive::parse_with_limits(decompressed.as_bytes(), limits.effective_archive_limits()?)?;
    Ok(Some((Component::new(name, archive), decompressed_bytes)))
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn iwa_bytes(identifier: u64, message_type: u32) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: message_type,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        Ok(SnappyStream::compress(&archive.to_bytes()?)?)
    }

    fn iwa_bytes_with_payload(
        identifier: u64,
        message_type: u32,
        payload_bytes: usize,
    ) -> Result<(Vec<u8>, u64)> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: message_type,
                    data: vec![0; payload_bytes],
                }],
            )?],
        };
        let decompressed = archive.to_bytes()?;
        let decompressed_bytes = u64::try_from(decompressed.len()).map_err(|_error| {
            crate::Error::InvalidBundle("test IWA stream length does not fit u64".to_owned())
        })?;
        Ok((SnappyStream::compress(&decompressed)?, decompressed_bytes))
    }

    fn indexed_catalog() -> Result<ComponentCatalog> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Alpha.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Index/Bravo.iwa", &iwa_bytes(2, 6001)?)?;
        writer.write_stored("Index/Charlie.iwa", &iwa_bytes(3, 6002)?)?;
        ComponentCatalog::from_bytes(&writer.finish_to_bytes()?)
    }

    #[test]
    fn gets_first_component_by_index() -> Result<()> {
        let catalog = indexed_catalog()?;

        assert_eq!(
            catalog.get_index(0).map(Component::name),
            Some("Index/Alpha.iwa")
        );
        Ok(())
    }

    #[test]
    fn gets_in_range_component_by_index() -> Result<()> {
        let catalog = indexed_catalog()?;

        assert_eq!(
            catalog.get_index(1).map(Component::name),
            Some("Index/Bravo.iwa")
        );
        Ok(())
    }

    #[test]
    fn returns_none_for_out_of_range_component_index() -> Result<()> {
        let catalog = indexed_catalog()?;

        assert!(catalog.get_index(catalog.len()).is_none());
        Ok(())
    }

    #[test]
    fn parses_direct_iwa_and_skips_operation_storage() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Index/OperationStorage.iwa", b"bvxn opaque data")?;
        let bytes = writer.finish_to_bytes()?;

        let catalog = ComponentCatalog::from_bytes(&bytes)?;
        assert_eq!(catalog.len(), 1);
        assert!(!catalog.is_empty());
        assert_eq!(
            catalog.get("Index/Document.iwa").map(Component::name),
            Some("Index/Document.iwa")
        );
        let component = catalog.iter().next().ok_or_else(|| {
            crate::Error::InvalidBundle("component catalog unexpectedly empty".to_owned())
        })?;
        assert_eq!(component.name(), "Index/Document.iwa");
        assert_eq!(component.archive().objects[0].messages[0].type_, 6000);
        Ok(())
    }

    #[test]
    fn aggregate_decompressed_iwa_budget_is_exact_for_direct_and_source_catalogs() -> Result<()> {
        let (alpha, alpha_bytes) = iwa_bytes_with_payload(1, 6_000, 8 * 1024)?;
        let (bravo, bravo_bytes) = iwa_bytes_with_payload(2, 6_001, 8 * 1024)?;
        let exact_total = alpha_bytes.checked_add(bravo_bytes).ok_or_else(|| {
            crate::Error::InvalidBundle("test IWA aggregate length overflowed u64".to_owned())
        })?;
        let max_component_bytes = alpha_bytes.max(bravo_bytes);
        let max_entry_bytes = u64::try_from(alpha.len().max(bravo.len())).map_err(|_error| {
            crate::Error::InvalidBundle("test ZIP member length does not fit u64".to_owned())
        })?;
        let max_iwa_stream_bytes = usize::try_from(max_component_bytes).map_err(|_error| {
            crate::Error::InvalidBundle("test IWA stream length does not fit usize".to_owned())
        })?;

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Alpha.iwa", &alpha)?;
        writer.write_stored("Index/Bravo.iwa", &bravo)?;
        let bytes = writer.finish_to_bytes()?;
        let input_bytes = u64::try_from(bytes.len()).map_err(|_error| {
            crate::Error::InvalidBundle("test ZIP length does not fit u64".to_owned())
        })?;

        let exact = Limits::new(
            input_bytes,
            2,
            max_entry_bytes,
            exact_total,
            max_iwa_stream_bytes,
        )?;
        assert_eq!(
            ComponentCatalog::from_bytes_with_limits(&bytes, exact)?.len(),
            2
        );
        assert_eq!(
            SourceCatalog::from_bytes_with_limits(&bytes, exact)?
                .components()
                .len(),
            2
        );

        let exceeded = Limits::new(
            input_bytes,
            2,
            max_entry_bytes,
            exact_total - 1,
            max_iwa_stream_bytes,
        )?;
        for result in [
            ComponentCatalog::from_bytes_with_limits(&bytes, exceeded).map(|_catalog| ()),
            SourceCatalog::from_bytes_with_limits(&bytes, exceeded).map(|_catalog| ()),
        ] {
            let Err(error) = result else {
                return Err(crate::Error::InvalidBundle(
                    "aggregate IWA budget should reject the test package".to_owned(),
                ));
            };
            assert!(matches!(
                error,
                crate::Error::Limit {
                    kind: crate::LimitKind::IwaTotalBytes,
                    observed,
                    maximum,
                } if observed == exact_total && maximum == exact_total - 1
            ));
        }
        Ok(())
    }

    #[test]
    fn consumes_component_name_and_archive() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let bytes = writer.finish_to_bytes()?;

        let component = ComponentCatalog::from_bytes(&bytes)?
            .into_iter()
            .next()
            .ok_or_else(|| {
                crate::Error::InvalidBundle("component catalog unexpectedly empty".to_owned())
            })?;
        let (name, archive) = component.into_parts();

        assert_eq!(name, "Index/Document.iwa");
        assert_eq!(archive.objects[0].messages[0].type_, 6000);
        Ok(())
    }

    #[test]
    fn parses_nested_index_and_rejects_encryption() -> Result<()> {
        let mut index = StreamingArchiveWriter::new();
        index.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let index_bytes = index.finish_to_bytes()?;

        let mut outer = StreamingArchiveWriter::new();
        outer.write_stored("legacy.pages/Index.zip", &index_bytes)?;
        let outer_bytes = outer.finish_to_bytes()?;
        assert_eq!(ComponentCatalog::from_bytes(&outer_bytes)?.len(), 1);

        let mut encrypted = StreamingArchiveWriter::new();
        encrypted.write_stored(".iwpv2", b"metadata")?;
        encrypted.write_stored("Index/Document.iwa", b"ciphertext")?;
        let encrypted_bytes = encrypted.finish_to_bytes()?;
        let result = ComponentCatalog::from_bytes(&encrypted_bytes);
        assert!(matches!(result, Err(crate::Error::Encrypted)));
        Ok(())
    }

    #[test]
    fn rejects_mixed_direct_and_legacy_representations() -> Result<()> {
        let mut index = StreamingArchiveWriter::new();
        index.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let index_bytes = index.finish_to_bytes()?;

        let mut outer = StreamingArchiveWriter::new();
        outer.write_stored("legacy.pages/Index.zip", &index_bytes)?;
        outer.write_stored("Index/CalculationEngine.iwa", &iwa_bytes(2, 7000)?)?;
        let outer_bytes = outer.finish_to_bytes()?;

        assert!(matches!(
            ComponentCatalog::from_bytes(&outer_bytes),
            Err(crate::Error::InvalidBundle(message)) if message.contains("mixes direct IWA")
        ));
        Ok(())
    }

    #[test]
    fn rejects_input_above_profile() -> Result<()> {
        let limits = Limits::new(1, 10, 100, 100, 100)?;
        let result = ComponentCatalog::from_bytes_with_limits(b"not a zip", limits);
        assert!(matches!(result, Err(crate::Error::Limit { .. })));
        Ok(())
    }

    #[test]
    fn source_catalog_reuses_shared_source_and_traverses_direct_zip_once() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Metadata/DocumentIdentifier", b"shared-source")?;
        let source: Arc<[u8]> = writer.finish_to_bytes()?.into();

        crate::zip::reset_test_parse_count();
        let catalog = SourceCatalog::from_shared_bytes(Arc::clone(&source))?;

        assert!(Arc::ptr_eq(&source, &catalog.shared_source()));
        assert_eq!(crate::zip::test_parse_count(), 1);
        assert!(catalog.source_is_exact());
        assert_eq!(catalog.source_provenance(), SourceProvenance::ExactZip);
        assert_eq!(catalog.package().len(), 2);
        assert_eq!(catalog.components().len(), 1);
        assert_eq!(
            catalog
                .components()
                .get("Index/Document.iwa")
                .map(Component::name),
            Some("Index/Document.iwa")
        );
        Ok(())
    }

    #[test]
    fn source_catalog_retains_the_validated_ingress_profile() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let source: Arc<[u8]> = writer.finish_to_bytes()?.into();
        let source_len = u64::try_from(source.len()).map_err(|_error| {
            crate::Error::InvalidBundle("test source length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(source_len, 7, source_len, source_len, 1_024 * 1_024)?;

        let catalog = SourceCatalog::from_shared_bytes_with_limits(source, limits)?;

        assert_eq!(catalog.limits(), limits);
        Ok(())
    }

    #[test]
    fn source_catalog_component_handoff_releases_physical_source() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Metadata/DocumentIdentifier", b"physical-source")?;
        let source: Arc<[u8]> = writer.finish_to_bytes()?.into();
        let source_lifetime = Arc::downgrade(&source);
        let catalog = SourceCatalog::from_shared_bytes(Arc::clone(&source))?;

        drop(source);
        assert!(source_lifetime.upgrade().is_some());

        let components = catalog.into_components();

        assert!(source_lifetime.upgrade().is_none());
        assert_eq!(components.len(), 1);
        assert_eq!(
            components.get("Index/Document.iwa").map(Component::name),
            Some("Index/Document.iwa")
        );
        Ok(())
    }

    #[test]
    fn source_catalog_normalizes_legacy_index_during_two_required_traversals() -> Result<()> {
        let mut index = StreamingArchiveWriter::new();
        index.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let index_bytes = index.finish_to_bytes()?;

        let mut outer = StreamingArchiveWriter::new();
        outer.write_stored("legacy.pages/Index.zip", &index_bytes)?;
        outer.write_stored("legacy.pages/Metadata/DocumentIdentifier", b"legacy-source")?;
        let source: Arc<[u8]> = outer.finish_to_bytes()?.into();

        crate::zip::reset_test_parse_count();
        let catalog = SourceCatalog::from_shared_bytes(Arc::clone(&source))?;

        assert!(Arc::ptr_eq(&source, &catalog.shared_source()));
        assert_eq!(crate::zip::test_parse_count(), 2);
        assert!(!catalog.source_is_exact());
        assert_eq!(catalog.source_provenance(), SourceProvenance::LegacyZip);
        assert!(
            catalog
                .package()
                .iter()
                .any(|entry| entry.name() == "Metadata/DocumentIdentifier")
        );
        assert_eq!(catalog.components().len(), 1);
        assert_eq!(
            catalog
                .components()
                .get("Index/Document.iwa")
                .map(Component::name),
            Some("Index/Document.iwa")
        );
        Ok(())
    }

    #[test]
    fn source_catalog_component_projection_matches_component_only_ingress() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Bravo.iwa", &iwa_bytes(2, 6001)?)?;
        writer.write_stored("Index/Alpha.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Index/OperationStorage.iwa", b"bvxn opaque data")?;
        let bytes = writer.finish_to_bytes()?;

        let source = SourceCatalog::from_bytes(&bytes)?;
        let components = ComponentCatalog::from_bytes(&bytes)?;
        let source_projection = source
            .components()
            .iter()
            .map(|component| {
                (
                    component.name(),
                    component.archive().objects[0].messages[0].type_,
                )
            })
            .collect::<Vec<_>>();
        let component_projection = components
            .iter()
            .map(|component| {
                (
                    component.name(),
                    component.archive().objects[0].messages[0].type_,
                )
            })
            .collect::<Vec<_>>();

        assert_eq!(source_projection, component_projection);
        Ok(())
    }
}
