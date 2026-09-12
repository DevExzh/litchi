//! Lazy physical ingress for Keynote slide-table merge queries.
//!
//! This reader intentionally keeps the package boundary separate from the
//! semantic [`crate::Package`] API. Physical ZIP/IWA validation produces one
//! checked component catalog and compact object index; slide/table selection
//! and the selected merge-owner payload are decoded only when a query is
//! made. The semantic show, slides, text, cells, and other package values are
//! never constructed by this type.

use std::{fmt, path::Path, sync::Arc};

#[cfg(feature = "internal-iwork-source")]
use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind, SourceCatalog};

use super::super::{
    Package, PayloadLimitKind, ReadError, ReadOptions, SemanticLimitKind, read_source,
};
use super::SlideTableMergesError;

/// A lazy, metadata-only Keynote merged-cell reader.
///
/// `MergeReader` retains one validated component catalog and the compact object
/// index required to resolve rooted slide/table ownership. It does not retain
/// a ZIP reassembly catalog, build a semantic [`crate::Document`], decode slide
/// values, or expose native object identifiers. Clones share the immutable
/// package state and its parsed archive allocations.
///
/// Queries accept exact visible slide names, checked zero-based slide
/// positions, and the existing checked [`crate::TableSelector`] vocabulary.
/// They return the same checked [`crate::slide::table::merge::Region`] values
/// as [`Package::slide_table_merges`].
///
/// Physical ingress is validated at construction. Query-time validation is
/// limited to the selected slide/table graph and merge-owner payload, so an
/// unrelated malformed cell or slide does not force a full semantic decode.
///
/// ```no_run
/// use litchi_keynote::{MergeReader, SlideSelector, TableSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let reader = MergeReader::open("deck.key")?;
/// let regions = reader.slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
/// println!("{regions:?}");
/// # Ok(())
/// # }
/// ```
pub struct MergeReader {
    package: Package,
}

impl fmt::Debug for MergeReader {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MergeReader")
            .field("options", &self.package.read_options())
            .finish_non_exhaustive()
    }
}

impl Clone for MergeReader {
    fn clone(&self) -> Self {
        Self {
            package: self.package.clone(),
        }
    }
}

impl MergeReader {
    /// Open a Keynote package for lazy merged-cell queries with default limits.
    ///
    /// The file is read through the same descriptor-safe bounded ingress used
    /// by [`Package::open`]. Only the checked component catalog and object
    /// index are retained after this function returns.
    ///
    /// # Errors
    ///
    /// Returns a typed merge error when the path cannot be read, physical
    /// ingress exceeds a limit, or the source is not a valid Keynote package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, SlideTableMergesError> {
        Self::open_with_options(path, ReadOptions::default())
    }

    /// Open a Keynote package for lazy merged-cell queries under explicit
    /// physical and semantic limits.
    ///
    /// Semantic limits govern the compact object index and the selected query
    /// traversal.  They do not trigger construction of the complete semantic
    /// show or any cell-bearing value.
    ///
    /// # Errors
    ///
    /// Returns a typed merge error when bounded filesystem or package ingress
    /// fails.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: ReadOptions,
    ) -> Result<Self, SlideTableMergesError> {
        let source = read_source(path.as_ref(), options.archive()).map_err(map_read_error)?;
        Self::from_shared_bytes_with_options(source, options)
    }

    /// Parse Keynote package bytes for lazy merged-cell queries with defaults.
    ///
    /// The input is parsed into one immutable component catalog; the ZIP
    /// allocation is released after component validation because this reader
    /// does not provide source-preserving writes. No semantic `Document` or
    /// slide snapshot is built.
    ///
    /// # Errors
    ///
    /// Returns a typed merge error when physical ingress or Keynote root
    /// validation fails.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, SlideTableMergesError> {
        Self::from_bytes_with_options(bytes, ReadOptions::default())
    }

    /// Parse Keynote package bytes under explicit physical and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns a typed merge error when the source is malformed, belongs to a
    /// different iWork format, or exceeds the configured profile.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: ReadOptions,
    ) -> Result<Self, SlideTableMergesError> {
        let length =
            u64::try_from(bytes.len()).map_err(|_| SlideTableMergesError::InvalidSource)?;
        if length > options.archive().max_input_bytes() {
            return Err(SlideTableMergesError::LimitExceeded {
                kind: super::SlideTableMergesLimitKind::InputBytes,
                observed: length,
                maximum: options.archive().max_input_bytes(),
            });
        }
        let source = SourceCatalog::from_bytes_with_limits(bytes, options.archive())
            .map_err(map_archive_error)?;
        let components = Arc::new(source.into_components());
        let package =
            Package::from_merge_components(components, options).map_err(map_read_error)?;
        Ok(Self { package })
    }

    /// Parse an already shared immutable Keynote source without copying its
    /// ZIP allocation.
    ///
    /// This is useful for a coordinator that already owns the exact source
    /// bytes. The source still goes through one bounded `SourceCatalog` parse,
    /// while clones share the resulting immutable component state.
    ///
    /// # Errors
    ///
    /// Returns the same typed merge failures as [`Self::from_bytes_with_options`].
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self, SlideTableMergesError> {
        Self::from_shared_bytes_with_options(bytes, ReadOptions::default())
    }

    /// Parse an already shared immutable source under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns a typed merge error when physical ingress, Keynote format
    /// classification, or bounded object indexing fails.
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: ReadOptions,
    ) -> Result<Self, SlideTableMergesError> {
        let source = SourceCatalog::from_shared_bytes_with_limits(bytes, options.archive())
            .map_err(map_archive_error)?;
        let components = Arc::new(source.into_components());
        let package =
            Package::from_merge_components(components, options).map_err(map_read_error)?;
        Ok(Self { package })
    }

    /// Build a metadata-only reader from an already checked component
    /// catalog.  The catalog and each shared archive allocation remain owned
    /// by the coordinator; this call only adds the Keynote object index and a
    /// shared reader handle.
    #[cfg(feature = "internal-iwork-source")]
    #[doc(hidden)]
    pub fn __from_shared_catalog(
        catalog: Arc<ComponentCatalog>,
        options: ReadOptions,
    ) -> Result<Self, SlideTableMergesError> {
        let package = Package::from_merge_components(catalog, options).map_err(map_read_error)?;
        Ok(Self { package })
    }

    /// Read all validated merged-cell regions for one selected slide table.
    ///
    /// Selection is performed through the rooted slide-owned drawable and
    /// z-order graph.  The selected table model's merge-owner storage is then
    /// decoded with the shared borrowed Buffa wire adapter.  The source remains
    /// immutable and no native identifiers cross this API.
    ///
    /// # Errors
    ///
    /// Returns selector, source, or bounded resource errors from the focused
    /// Keynote merge path.
    pub fn slide_table_merges<'slide>(
        &self,
        slide: impl Into<crate::SlideSelector<'slide>>,
        table: impl Into<crate::TableSelector>,
    ) -> Result<Vec<crate::slide::table::merge::Region>, SlideTableMergesError> {
        self.package
            .slide_table_merges_from_components(slide, table)
    }

    /// Return the physical and semantic profiles retained by this reader.
    #[must_use]
    pub fn read_options(&self) -> ReadOptions {
        self.package.read_options()
    }
}

fn map_read_error(error: ReadError) -> SlideTableMergesError {
    match error {
        ReadError::NotKeynote => SlideTableMergesError::UnsupportedSource,
        ReadError::Archive(error) => map_archive_error(error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: map_semantic_limit(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: map_payload_limit(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableMergesError::Allocation { amount },
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => SlideTableMergesError::InvalidSource,
    }
}

fn map_archive_error(error: ArchiveError) -> SlideTableMergesError {
    match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableMergesError::LimitExceeded {
            kind: map_archive_limit(kind),
            observed,
            maximum,
        },
        ArchiveError::Allocation { amount, .. } => SlideTableMergesError::Allocation { amount },
        ArchiveError::Io(_)
        | ArchiveError::Zip { .. }
        | ArchiveError::Iwa(_)
        | ArchiveError::InvalidLimits(_)
        | ArchiveError::Encrypted
        | ArchiveError::SourceChanged { .. }
        | ArchiveError::DirectoryChanged { .. }
        | ArchiveError::Reassembly(_)
        | ArchiveError::InvalidBundle(_) => SlideTableMergesError::InvalidSource,
    }
}

fn map_archive_limit(
    kind: ArchiveLimitKind,
) -> super::super::slide_table_merges::SlideTableMergesLimitKind {
    use super::super::slide_table_merges::SlideTableMergesLimitKind;

    match kind {
        ArchiveLimitKind::InputBytes => SlideTableMergesLimitKind::InputBytes,
        ArchiveLimitKind::OutputBytes => SlideTableMergesLimitKind::OutputBytes,
        ArchiveLimitKind::Entries => SlideTableMergesLimitKind::Entries,
        ArchiveLimitKind::MemberNameBytes => SlideTableMergesLimitKind::EntryBytes,
        ArchiveLimitKind::MetadataBytes => SlideTableMergesLimitKind::TotalBytes,
        ArchiveLimitKind::CompressedEntryBytes => SlideTableMergesLimitKind::EntryBytes,
        ArchiveLimitKind::EntryBytes => SlideTableMergesLimitKind::EntryBytes,
        ArchiveLimitKind::TotalBytes => SlideTableMergesLimitKind::TotalBytes,
        ArchiveLimitKind::IwaStreamBytes => SlideTableMergesLimitKind::PayloadObjects,
        ArchiveLimitKind::IwaTotalBytes => SlideTableMergesLimitKind::TotalBytes,
    }
}

fn map_semantic_limit(
    kind: SemanticLimitKind,
) -> super::super::slide_table_merges::SlideTableMergesLimitKind {
    use super::super::slide_table_merges::SlideTableMergesLimitKind;

    match kind {
        SemanticLimitKind::Objects => SlideTableMergesLimitKind::PayloadObjects,
        SemanticLimitKind::Slides => SlideTableMergesLimitKind::Components,
        SemanticLimitKind::References => SlideTableMergesLimitKind::References,
        SemanticLimitKind::TextStorages => SlideTableMergesLimitKind::PayloadMessages,
        SemanticLimitKind::TextFragments => SlideTableMergesLimitKind::PayloadItems,
        SemanticLimitKind::TextBytes => SlideTableMergesLimitKind::Retained,
    }
}

fn map_payload_limit(
    kind: PayloadLimitKind,
) -> super::super::slide_table_merges::SlideTableMergesLimitKind {
    use super::super::slide_table_merges::SlideTableMergesLimitKind;

    match kind {
        PayloadLimitKind::Bytes => SlideTableMergesLimitKind::InputBytes,
        PayloadLimitKind::Fields => SlideTableMergesLimitKind::WireFields,
        PayloadLimitKind::Nesting => SlideTableMergesLimitKind::WireNesting,
        PayloadLimitKind::Work => SlideTableMergesLimitKind::WireWork,
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::Ordering;

    use super::MergeReader;
    use crate::{Package, SlideSelector, TableSelector};
    use litchi_iwa_common::table::merge::Region;

    const NATIVE_SOURCE: &[u8] =
        include_bytes!("../../../../../test-data/iwork/keynote/slide-table-merges-native.key");

    #[test]
    fn reader_matches_package_geometry_without_semantic_snapshot() {
        let reader = MergeReader::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let expected = [Region::new(3, 1, 2, 2).expect("checked region")];

        assert_eq!(
            reader
                .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
                .expect("reader merge read"),
            expected
        );
        assert_eq!(
            package
                .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
                .expect("package merge read"),
            expected
        );
    }

    #[test]
    fn clones_share_the_reader_state_and_retain_options() {
        let reader = MergeReader::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let clone = reader.clone();
        assert_eq!(clone.read_options(), reader.read_options());
        assert_eq!(
            clone.slide_table_merges(0, 0).expect("clone merge read"),
            [Region::new(3, 1, 2, 2).expect("checked region")]
        );
    }

    #[test]
    fn name_selection_uses_metadata_without_semantic_decode() {
        let reader = MergeReader::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        assert_eq!(
            reader
                .package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
        assert!(matches!(
            reader.slide_table_merges(SlideSelector::name("missing"), TableSelector::index(0)),
            Err(super::super::SlideTableMergesError::SlideNameNotFound)
        ));
        assert_eq!(
            reader
                .package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
    }

    #[cfg(feature = "internal-iwork-source")]
    #[allow(
        deprecated,
        reason = "The synthetic Keynote nodes exercise the legacy required envelope fields."
    )]
    #[test]
    fn shared_catalog_keeps_original_archive_allocations() {
        use litchi_iwa_archive::{ComponentCatalog, Limits};

        let records = ComponentCatalog::from_bytes(NATIVE_SOURCE)
            .expect("native Keynote source")
            .into_iter()
            .map(|component| {
                let (name, archive) = component.into_parts();
                (name, std::sync::Arc::new(archive))
            })
            .collect::<Vec<_>>();
        let lifetimes = records
            .iter()
            .map(|(_, archive)| std::sync::Arc::downgrade(archive))
            .collect::<Vec<_>>();
        let catalog = std::sync::Arc::new(
            ComponentCatalog::__from_shared_archives(
                records
                    .iter()
                    .map(|(name, archive)| (name.as_str(), std::sync::Arc::clone(archive))),
                Limits::default(),
            )
            .expect("shared component catalog"),
        );
        let reader = MergeReader::__from_shared_catalog(
            std::sync::Arc::clone(&catalog),
            crate::ReadOptions::default(),
        )
        .expect("shared catalog reader");
        assert_eq!(
            reader
                .package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
        for (name, archive) in &records {
            assert!(std::ptr::eq(
                catalog.get(name).expect("shared component").archive(),
                std::sync::Arc::as_ptr(archive),
            ));
            assert_eq!(std::sync::Arc::strong_count(archive), 2);
        }
        assert_eq!(
            reader
                .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
                .expect("shared catalog merge read"),
            [Region::new(3, 1, 2, 2).expect("checked region")]
        );
        assert_eq!(
            reader
                .package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
        let clone = reader.clone();
        drop(reader);
        drop(catalog);
        drop(records);
        assert!(
            lifetimes
                .iter()
                .all(|lifetime| lifetime.upgrade().is_some())
        );
        drop(clone);
        assert!(
            lifetimes
                .iter()
                .all(|lifetime| lifetime.upgrade().is_none())
        );
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn missing_name_scan_accumulates_references_across_many_slides()
    -> Result<(), Box<dyn std::error::Error>> {
        use std::sync::Arc;

        use crate::{
            MAX_OBJECTS, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES,
            ReadOptions, SemanticLimits,
        };
        use litchi_iwa_archive::{ComponentCatalog, Limits};
        use litchi_iwa_core::{ArchiveObject, RawMessage};
        use litchi_iwa_protos::{kn, tsp};
        use prost::Message as _;

        fn reference(identifier: u64) -> tsp::Reference {
            tsp::Reference {
                identifier,
                deprecated_type: Some(7),
                deprecated_is_external: Some(false),
            }
        }

        let package = Package::from_bytes(NATIVE_SOURCE)?;
        let show_identifier = package.root_show_identifier()?;
        let components = ComponentCatalog::from_bytes(NATIVE_SOURCE)?;
        let mut parts = components
            .into_iter()
            .map(|component| component.into_parts())
            .collect::<Vec<_>>();
        let mut next_identifier = parts
            .iter()
            .flat_map(|(_, archive)| archive.objects.iter())
            .filter_map(|object| object.archive_info.identifier)
            .max()
            .ok_or("native source has no objects")?
            .checked_add(1)
            .ok_or("native object identifier overflow")?;
        let mut extra_objects = Vec::new();
        let mut show_rewritten = false;
        const EXTRA_SLIDES: usize = 8;

        for (_, archive) in &mut parts {
            for object in &mut archive.objects {
                if object.archive_info.identifier != Some(show_identifier) {
                    continue;
                }
                for message in &mut object.messages {
                    if message.type_ != 2 {
                        continue;
                    }
                    let mut show = kn::ShowArchive::decode(message.data.as_slice())?;
                    for _ in 0..EXTRA_SLIDES {
                        let node_identifier = next_identifier;
                        let slide_identifier = next_identifier
                            .checked_add(1)
                            .ok_or("synthetic slide identifier overflow")?;
                        next_identifier = slide_identifier
                            .checked_add(1)
                            .ok_or("synthetic slide identifier overflow")?;
                        show.slide_tree.slides.push(reference(node_identifier));
                        let node = kn::SlideNodeArchive {
                            slide: Some(reference(slide_identifier)),
                            is_skipped: false,
                            has_builds: false,
                            has_transition: false,
                            ..Default::default()
                        };
                        extra_objects.push(ArchiveObject::new(
                            node_identifier,
                            vec![RawMessage {
                                type_: 4,
                                data: node.encode_to_vec(),
                            }],
                        )?);
                        extra_objects.push(ArchiveObject::new(
                            slide_identifier,
                            vec![RawMessage {
                                type_: 5,
                                data: Vec::new(),
                            }],
                        )?);
                    }
                    message.data = show.encode_to_vec();
                    show_rewritten = true;
                }
            }
        }
        assert!(show_rewritten, "native source must contain the show object");
        parts
            .first_mut()
            .ok_or("native source has no components")?
            .1
            .objects
            .extend(extra_objects);
        let catalog = Arc::new(ComponentCatalog::__from_shared_archives(
            parts
                .iter()
                .map(|(name, archive)| (name.as_str(), Arc::new(archive.clone()))),
            Limits::default(),
        )?);
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            5,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )?;
        let reader = MergeReader::__from_shared_catalog(
            catalog,
            ReadOptions::new(Limits::default(), semantic),
        )?;
        assert!(matches!(
            reader.slide_table_merges(SlideSelector::name("missing"), TableSelector::index(0)),
            Err(super::super::SlideTableMergesError::LimitExceeded {
                kind: super::super::SlideTableMergesLimitKind::References,
                observed: 6,
                maximum: 5,
            })
        ));
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_only_package_ingress_remains_unsupported_for_package_reads() {
        let components = std::sync::Arc::new(
            litchi_iwa_archive::ComponentCatalog::from_bytes(NATIVE_SOURCE)
                .expect("native Keynote components"),
        );
        let package = Package::from_classified_components(
            components,
            crate::Limits::default(),
            crate::SemanticLimits::default(),
        )
        .expect("semantic-only Keynote package");
        assert!(matches!(
            package.slide_table_merges(SlideSelector::index(0), TableSelector::index(0)),
            Err(super::super::SlideTableMergesError::UnsupportedSource)
        ));
    }
}
