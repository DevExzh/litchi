//! Bounded, format-neutral Apple iWork reading.
//!
//! [`Document`] detects Pages, Keynote, or Numbers once, consumes the same
//! immutable package state into its concrete format owner, and eagerly
//! validates an archive-free semantic snapshot. Public values never expose
//! native object numbers, protobuf messages, package internals, or concrete
//! format types.
//!
//! ```no_run
//! use litchi::iwork::{Document, TextRole};
//!
//! # fn read(value: &[u8]) -> Result<(), litchi::iwork::Error> {
//! let document = Document::from_bytes(value)?;
//! let snapshot = document.snapshot();
//! println!("{}: {}", document.format(), snapshot.summary());
//! for text in snapshot.iter_text() {
//!     if text.role() != TextRole::TableName {
//!         println!("{}", text.value());
//!     }
//! }
//! # Ok(())
//! # }
//! ```

mod error;
mod limits;
mod model;

use std::{fmt, path::Path, sync::Arc};

pub use error::{Error, ErrorKind, Resource, Result, Stage};
pub use limits::{Options, SnapshotLimits, SourceLimits};
pub use model::{
    Cell, CellView, Document, Section, SectionKind, Slide, Snapshot, Summary, Table, Text,
    TextRole, Value, ValueKind,
};

/// Apple application family selected for one semantic document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    /// Apple Pages.
    Pages,
    /// Apple Keynote.
    Keynote,
    /// Apple Numbers.
    Numbers,
}

impl Format {
    /// Return the stable application name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Pages => "Pages",
            Self::Keynote => "Keynote",
            Self::Numbers => "Numbers",
        }
    }
}

impl fmt::Display for Format {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.name())
    }
}

impl Document {
    /// Open a packaged iWork file or app-authored package directory.
    ///
    /// The path is snapshotted once under the default finite resource
    /// profile. Later semantic access performs no ambient filesystem reads.
    ///
    /// # Errors
    ///
    /// Returns a facade-owned [`Error`] when the path is inaccessible,
    /// unsafe, unrecognized, malformed, changes during capture, or exceeds a
    /// configured resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, Options::default())
    }

    /// Open a packaged iWork file or app-authored package directory under
    /// explicit checked options.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`].
    pub fn open_with_options(path: impl AsRef<Path>, options: Options) -> Result<Self> {
        let physical = options.source().detector()?;
        let prepared =
            litchi_iwa_detect::PreparedSource::__from_path_with_semantic_metadata(path, physical)
                .map_err(map_detection)?
                .ok_or_else(Error::unrecognized)?;
        decode(prepared, options.snapshot())
    }

    /// Read borrowed package bytes under the default finite resource profile.
    ///
    /// The bytes are copied once into immutable shared ownership before any
    /// semantic value can escape. Use [`Self::from_shared_bytes`] when the
    /// caller already owns an `Arc<[u8]>`.
    ///
    /// # Errors
    ///
    /// Returns a facade-owned [`Error`] when the input is unrecognized,
    /// malformed, or over budget.
    pub fn from_bytes(value: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(value, Options::default())
    }

    /// Read borrowed package bytes under explicit checked options.
    ///
    /// # Errors
    ///
    /// Returns a facade-owned [`Error`] when the input is unrecognized,
    /// malformed, or over budget.
    pub fn from_bytes_with_options(value: &[u8], options: Options) -> Result<Self> {
        admit_input(value.len(), options.source())?;
        let physical = options.source().detector()?;
        let prepared =
            litchi_iwa_detect::PreparedSource::__from_bytes_with_semantic_metadata(value, physical)
                .map_err(map_detection)?
                .ok_or_else(Error::unrecognized)?;
        decode(prepared, options.snapshot())
    }

    /// Read already-owned immutable package bytes without copying them.
    ///
    /// # Errors
    ///
    /// Returns a facade-owned [`Error`] when the input is unrecognized,
    /// malformed, or over budget.
    pub fn from_shared_bytes(value: Arc<[u8]>) -> Result<Self> {
        Self::from_shared_bytes_with_options(value, Options::default())
    }

    /// Read already-owned immutable package bytes under explicit options.
    ///
    /// # Errors
    ///
    /// Returns a facade-owned [`Error`] when the input is unrecognized,
    /// malformed, or over budget.
    pub fn from_shared_bytes_with_options(value: Arc<[u8]>, options: Options) -> Result<Self> {
        admit_input(value.len(), options.source())?;
        let physical = options.source().detector()?;
        let prepared =
            litchi_iwa_detect::PreparedSource::__from_shared_bytes_with_semantic_metadata(
                value, physical,
            )
            .map_err(map_detection)?
            .ok_or_else(Error::unrecognized)?;
        decode(prepared, options.snapshot())
    }
}

fn admit_input(length: usize, limits: SourceLimits) -> Result<()> {
    let observed = u64::try_from(length).map_err(|_error| {
        Error::limit(
            None,
            Stage::Input,
            Resource::InputBytes,
            u64::MAX,
            limits.max_input_bytes(),
        )
    })?;
    if observed > limits.max_input_bytes() {
        return Err(Error::limit(
            None,
            Stage::Input,
            Resource::InputBytes,
            observed,
            limits.max_input_bytes(),
        ));
    }
    Ok(())
}

fn decode(prepared: litchi_iwa_detect::PreparedSource, limits: SnapshotLimits) -> Result<Document> {
    let aggregate = limits.aggregate()?;
    match prepared.format() {
        litchi_iwa_detect::Format::Pages => {
            let document = litchi_pages::__semantic_document_from_prepared_source(
                prepared,
                limits.max_sections(),
                limits.max_text_bytes(),
            )
            .map_err(map_pages)?;
            let data = litchi_iwa_structured::StructuredData::from_pages_document_with_limits(
                document, aggregate,
            )
            .map_err(|error| map_aggregate(error, Format::Pages))?;
            Ok(Document::from_data(Format::Pages, data))
        },
        litchi_iwa_detect::Format::Keynote => {
            let semantic = keynote_limits(limits)?;
            let document =
                litchi_keynote::__semantic_document_from_prepared_source(prepared, semantic)
                    .map_err(map_keynote)?;
            let data = litchi_iwa_structured::StructuredData::from_keynote_document_with_limits(
                document, aggregate,
            )
            .map_err(|error| map_aggregate(error, Format::Keynote))?;
            Ok(Document::from_data(Format::Keynote, data))
        },
        litchi_iwa_detect::Format::Numbers => {
            let semantic = numbers_limits(limits)?;
            let tables =
                litchi_numbers::__compatibility_tables_from_prepared_source(prepared, semantic)
                    .map_err(map_numbers)?;
            let data = litchi_iwa_structured::StructuredData::from_parts_with_limits(
                tables,
                Vec::new(),
                Vec::new(),
                aggregate,
            )
            .map_err(|error| map_aggregate(error, Format::Numbers))?;
            Ok(Document::from_data(Format::Numbers, data))
        },
    }
}

fn keynote_limits(limits: SnapshotLimits) -> Result<litchi_keynote::SemanticLimits> {
    let defaults = litchi_keynote::SemanticLimits::default();
    litchi_keynote::SemanticLimits::new(
        defaults.max_objects(),
        limits.max_slides(),
        defaults.max_references(),
        defaults.max_text_storages(),
        defaults.max_text_fragments(),
        limits.max_text_bytes(),
    )
    .map_err(|_error| Error::invariant(Some(Format::Keynote), Stage::Validation))
}

fn numbers_limits(limits: SnapshotLimits) -> Result<litchi_numbers::PackageSemanticLimits> {
    let defaults = litchi_numbers::PackageSemanticLimits::default();
    litchi_numbers::PackageSemanticLimits::new(
        defaults.max_objects(),
        defaults.max_sheets(),
        limits.max_tables(),
        defaults.max_references(),
    )
    .and_then(|profile| {
        profile.with_projection_limits(defaults.max_materialized_cells(), limits.max_text_bytes())
    })
    .map_err(|_error| Error::invariant(Some(Format::Numbers), Stage::Validation))
}

fn map_detection(error: litchi_iwa_detect::Error) -> Error {
    match error {
        litchi_iwa_detect::Error::Io(_error) => {
            Error::new(ErrorKind::Io, Stage::Input, None, None, None, None)
        },
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::limit(
            None,
            Stage::Detection,
            detection_resource(kind),
            observed,
            maximum,
        ),
        litchi_iwa_detect::Error::InvalidLimits => Error::new(
            ErrorKind::InvalidOptions,
            Stage::Validation,
            None,
            None,
            None,
            None,
        ),
        litchi_iwa_detect::Error::Allocation { amount } => {
            Error::allocation(None, Stage::Detection, usize_u64(amount))
        },
        litchi_iwa_detect::Error::Encrypted => Error::encrypted(),
        litchi_iwa_detect::Error::SourceChanged => Error::source_changed(),
        litchi_iwa_detect::Error::IwaCore(_)
        | litchi_iwa_detect::Error::IwaCommon(_)
        | litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_) => Error::invalid_data(None, Stage::Detection),
        _ => Error::invalid_data(None, Stage::Detection),
    }
}

fn detection_resource(kind: litchi_iwa_detect::LimitKind) -> Resource {
    match kind {
        litchi_iwa_detect::LimitKind::InputBytes => Resource::InputBytes,
        litchi_iwa_detect::LimitKind::OutputBytes => Resource::OutputBytes,
        litchi_iwa_detect::LimitKind::Entries => Resource::Entries,
        litchi_iwa_detect::LimitKind::MemberNameBytes => Resource::MemberNameBytes,
        litchi_iwa_detect::LimitKind::MetadataBytes => Resource::MetadataBytes,
        litchi_iwa_detect::LimitKind::CompressedEntryBytes => Resource::CompressedEntryBytes,
        litchi_iwa_detect::LimitKind::EntryBytes => Resource::EntryBytes,
        litchi_iwa_detect::LimitKind::TotalBytes => Resource::ExpandedBytes,
        litchi_iwa_detect::LimitKind::IwaStreamBytes => Resource::DecodedBytes,
        litchi_iwa_detect::LimitKind::IwaTotalBytes => Resource::AggregateDecodedBytes,
        _ => Resource::SemanticWork,
    }
}

fn map_pages(error: litchi_pages::PackageError) -> Error {
    match error {
        litchi_pages::PackageError::Io(_error) => Error::new(
            ErrorKind::Io,
            Stage::Input,
            Some(Format::Pages),
            None,
            None,
            None,
        ),
        litchi_pages::PackageError::SectionNamesTooLarge { observed, limit } => Error::limit(
            Some(Format::Pages),
            Stage::Semantic,
            Resource::TextBytes,
            usize_u64(observed),
            usize_u64(limit),
        ),
        litchi_pages::PackageError::Semantic(error) => match error {
            litchi_pages::Error::TooManySections { actual, limit } => Error::limit(
                Some(Format::Pages),
                Stage::Semantic,
                Resource::Sections,
                usize_u64(actual),
                usize_u64(limit),
            ),
            litchi_pages::Error::TooManyBodyStorages { actual, limit } => Error::limit(
                Some(Format::Pages),
                Stage::Semantic,
                Resource::TextStorages,
                usize_u64(actual),
                usize_u64(limit),
            ),
            litchi_pages::Error::TextTooLarge { observed, limit } => Error::limit(
                Some(Format::Pages),
                Stage::Semantic,
                Resource::TextBytes,
                usize_u64(observed),
                usize_u64(limit),
            ),
            litchi_pages::Error::InvalidSectionIndex { .. } => {
                Error::invalid_data(Some(Format::Pages), Stage::Semantic)
            },
            _ => Error::invalid_data(Some(Format::Pages), Stage::Semantic),
        },
        litchi_pages::PackageError::Archive(_) | litchi_pages::PackageError::InvalidFormat(_) => {
            Error::invalid_data(Some(Format::Pages), Stage::Semantic)
        },
        _ => Error::invalid_data(Some(Format::Pages), Stage::Semantic),
    }
}

fn map_keynote(error: litchi_keynote::ReadError) -> Error {
    match error {
        litchi_keynote::ReadError::Io(_error) => Error::new(
            ErrorKind::Io,
            Stage::Input,
            Some(Format::Keynote),
            None,
            None,
            None,
        ),
        litchi_keynote::ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::limit(
            Some(Format::Keynote),
            Stage::Semantic,
            keynote_resource(kind),
            usize_u64(observed),
            usize_u64(maximum),
        ),
        litchi_keynote::ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::limit(
            Some(Format::Keynote),
            Stage::Semantic,
            keynote_payload_resource(kind),
            usize_u64(observed),
            usize_u64(maximum),
        ),
        litchi_keynote::ReadError::Allocation { amount, .. } => {
            Error::allocation(Some(Format::Keynote), Stage::Semantic, usize_u64(amount))
        },
        _ => Error::invalid_data(Some(Format::Keynote), Stage::Semantic),
    }
}

fn keynote_resource(kind: litchi_keynote::SemanticLimitKind) -> Resource {
    match kind {
        litchi_keynote::SemanticLimitKind::Objects => Resource::Objects,
        litchi_keynote::SemanticLimitKind::Slides => Resource::Slides,
        litchi_keynote::SemanticLimitKind::References => Resource::References,
        litchi_keynote::SemanticLimitKind::TextStorages => Resource::TextStorages,
        litchi_keynote::SemanticLimitKind::TextFragments => Resource::TextFragments,
        litchi_keynote::SemanticLimitKind::TextBytes => Resource::TextBytes,
        _ => Resource::SemanticWork,
    }
}

fn keynote_payload_resource(kind: litchi_keynote::PayloadLimitKind) -> Resource {
    match kind {
        litchi_keynote::PayloadLimitKind::Bytes => Resource::PayloadBytes,
        litchi_keynote::PayloadLimitKind::Fields => Resource::Fields,
        litchi_keynote::PayloadLimitKind::Nesting => Resource::NestingDepth,
        litchi_keynote::PayloadLimitKind::Work => Resource::SemanticWork,
        _ => Resource::SemanticWork,
    }
}

fn map_numbers(error: litchi_numbers::PackageError) -> Error {
    if let Some(resource_error) = error.resource_error() {
        return map_numbers_resource_error(resource_error);
    }
    match error {
        litchi_numbers::PackageError::Io(_error) => Error::new(
            ErrorKind::Io,
            Stage::Input,
            Some(Format::Numbers),
            None,
            None,
            None,
        ),
        litchi_numbers::PackageError::InputTooLarge { observed, maximum } => Error::limit(
            Some(Format::Numbers),
            Stage::Input,
            Resource::InputBytes,
            observed,
            maximum,
        ),
        litchi_numbers::PackageError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::limit(
            Some(Format::Numbers),
            Stage::Semantic,
            numbers_resource(kind),
            usize_u64(observed),
            usize_u64(maximum),
        ),
        _ => Error::invalid_data(Some(Format::Numbers), Stage::Semantic),
    }
}

fn map_numbers_resource_error(error: litchi_numbers::PackageResourceError) -> Error {
    match error {
        litchi_numbers::PackageResourceError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::limit(
            Some(Format::Numbers),
            Stage::Semantic,
            numbers_payload_resource(kind),
            usize_u64(observed),
            usize_u64(maximum),
        ),
        litchi_numbers::PackageResourceError::Allocation { amount } => {
            Error::allocation(Some(Format::Numbers), Stage::Semantic, usize_u64(amount))
        },
        _ => Error::invalid_data(Some(Format::Numbers), Stage::Semantic),
    }
}

fn numbers_payload_resource(kind: litchi_numbers::PackagePayloadLimitKind) -> Resource {
    match kind {
        litchi_numbers::PackagePayloadLimitKind::InputBytes => Resource::PayloadBytes,
        litchi_numbers::PackagePayloadLimitKind::Fields => Resource::Fields,
        litchi_numbers::PackagePayloadLimitKind::OutputBytes => Resource::OutputBytes,
        litchi_numbers::PackagePayloadLimitKind::Nesting => Resource::NestingDepth,
        litchi_numbers::PackagePayloadLimitKind::RewriteWork => Resource::SemanticWork,
        litchi_numbers::PackagePayloadLimitKind::TableRows
        | litchi_numbers::PackagePayloadLimitKind::TableColumns
        | litchi_numbers::PackagePayloadLimitKind::TableCells
        | litchi_numbers::PackagePayloadLimitKind::MaterializedCells => Resource::Cells,
        _ => Resource::SemanticWork,
    }
}

fn numbers_resource(kind: litchi_numbers::SemanticLimitKind) -> Resource {
    match kind {
        litchi_numbers::SemanticLimitKind::Objects => Resource::Objects,
        litchi_numbers::SemanticLimitKind::Sheets => Resource::Sheets,
        litchi_numbers::SemanticLimitKind::Tables => Resource::Tables,
        litchi_numbers::SemanticLimitKind::References => Resource::References,
        litchi_numbers::SemanticLimitKind::MaterializedCells => Resource::Cells,
        litchi_numbers::SemanticLimitKind::OutputTextBytes
        | litchi_numbers::SemanticLimitKind::TextBytes => Resource::TextBytes,
        litchi_numbers::SemanticLimitKind::FormulaRenderDepth
        | litchi_numbers::SemanticLimitKind::FormulaDepth => Resource::NestingDepth,
        litchi_numbers::SemanticLimitKind::FormulaWireBytes => Resource::PayloadBytes,
        litchi_numbers::SemanticLimitKind::FormulaRenderWork
        | litchi_numbers::SemanticLimitKind::FormulaWork => Resource::SemanticWork,
        _ => Resource::SemanticWork,
    }
}

fn map_aggregate(error: litchi_iwa_structured::Error, format: Format) -> Error {
    match error {
        litchi_iwa_structured::Error::TooManyTables { actual, limit } => Error::limit(
            Some(format),
            Stage::Validation,
            Resource::Tables,
            usize_u64(actual),
            usize_u64(limit),
        ),
        litchi_iwa_structured::Error::TooManySlides { actual, limit } => Error::limit(
            Some(format),
            Stage::Validation,
            Resource::Slides,
            usize_u64(actual),
            usize_u64(limit),
        ),
        litchi_iwa_structured::Error::TooManySections { actual, limit } => Error::limit(
            Some(format),
            Stage::Validation,
            Resource::Sections,
            usize_u64(actual),
            usize_u64(limit),
        ),
        litchi_iwa_structured::Error::TextTooLarge { observed, maximum } => Error::limit(
            Some(format),
            Stage::Validation,
            Resource::TextBytes,
            usize_u64(observed),
            usize_u64(maximum),
        ),
        litchi_iwa_structured::Error::InvalidSlideIndex { .. }
        | litchi_iwa_structured::Error::InvalidSectionIndex { .. } => {
            Error::invariant(Some(format), Stage::Validation)
        },
        _ => Error::invalid_data(Some(format), Stage::Validation),
    }
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_names_are_stable() {
        assert_eq!(Format::Pages.name(), "Pages");
        assert_eq!(Format::Keynote.to_string(), "Keynote");
        assert_eq!(Format::Numbers.name(), "Numbers");
    }

    #[test]
    fn empty_and_non_zip_inputs_are_unrecognized() {
        for value in [&[][..], b"not a zip".as_slice()] {
            let error = Document::from_bytes(value)
                .err()
                .unwrap_or_else(|| panic!("input must be unrecognized"));
            assert_eq!(error.kind(), ErrorKind::Unrecognized);
            assert_eq!(error.stage(), Stage::Detection);
            assert_eq!(error.format(), None);
        }
    }

    #[test]
    fn checked_options_report_exact_resource() {
        let error = SnapshotLimits::new(
            0,
            SnapshotLimits::HARD_MAX_SLIDES,
            SnapshotLimits::HARD_MAX_SECTIONS,
            SnapshotLimits::HARD_MAX_TEXT_BYTES,
        )
        .err()
        .unwrap_or_else(|| panic!("zero table limit must fail"));
        assert_eq!(error.kind(), ErrorKind::InvalidOptions);
        assert_eq!(error.resource(), Some(Resource::Tables));
        assert_eq!(error.observed(), Some(0));
        assert_eq!(
            error.maximum(),
            Some(SnapshotLimits::HARD_MAX_TABLES as u64)
        );
    }

    #[test]
    fn input_limit_is_checked_before_copying() {
        let physical = SourceLimits::new(
            1,
            SourceLimits::HARD_MAX_ENTRIES,
            SourceLimits::HARD_MAX_ENTRY_BYTES,
            SourceLimits::HARD_MAX_EXPANDED_BYTES,
            SourceLimits::HARD_MAX_DECODED_BYTES_PER_ITEM,
        )
        .unwrap_or_else(|error| panic!("valid tightened profile: {error}"));
        let error =
            Document::from_bytes_with_options(&[0_u8; 2], Options::default().with_source(physical))
                .err()
                .unwrap_or_else(|| panic!("two bytes must exceed a one-byte profile"));
        assert_eq!(error.kind(), ErrorKind::LimitExceeded);
        assert_eq!(error.stage(), Stage::Input);
        assert_eq!(error.resource(), Some(Resource::InputBytes));
        assert_eq!(error.observed(), Some(2));
        assert_eq!(error.maximum(), Some(1));
    }

    #[test]
    fn detector_failures_map_to_exact_content_free_root_categories() {
        let resources = [
            (
                litchi_iwa_detect::LimitKind::InputBytes,
                Resource::InputBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::OutputBytes,
                Resource::OutputBytes,
            ),
            (litchi_iwa_detect::LimitKind::Entries, Resource::Entries),
            (
                litchi_iwa_detect::LimitKind::MemberNameBytes,
                Resource::MemberNameBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::MetadataBytes,
                Resource::MetadataBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::CompressedEntryBytes,
                Resource::CompressedEntryBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::EntryBytes,
                Resource::EntryBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::TotalBytes,
                Resource::ExpandedBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::IwaStreamBytes,
                Resource::DecodedBytes,
            ),
            (
                litchi_iwa_detect::LimitKind::IwaTotalBytes,
                Resource::AggregateDecodedBytes,
            ),
        ];
        for (kind, resource) in resources {
            let error = map_detection(litchi_iwa_detect::Error::LimitExceeded {
                kind,
                observed: 9,
                maximum: 8,
            });
            assert_eq!(error.kind(), ErrorKind::LimitExceeded);
            assert_eq!(error.stage(), Stage::Detection);
            assert_eq!(error.resource(), Some(resource));
            assert_eq!(error.observed(), Some(9));
            assert_eq!(error.maximum(), Some(8));
            assert!(std::error::Error::source(&error).is_none());
        }

        let encrypted = map_detection(litchi_iwa_detect::Error::Encrypted);
        assert_eq!(encrypted.kind(), ErrorKind::Encrypted);
        assert_eq!(encrypted.stage(), Stage::Detection);
        assert!(std::error::Error::source(&encrypted).is_none());

        let allocation = map_detection(litchi_iwa_detect::Error::Allocation { amount: 17 });
        assert_eq!(allocation.kind(), ErrorKind::Allocation);
        assert_eq!(allocation.stage(), Stage::Detection);
        assert_eq!(allocation.resource(), Some(Resource::Memory));
        assert_eq!(allocation.observed(), Some(17));

        let changed = map_detection(litchi_iwa_detect::Error::SourceChanged);
        assert_eq!(changed.kind(), ErrorKind::SourceChanged);
        assert_eq!(changed.stage(), Stage::Input);

        let invalid = map_detection(litchi_iwa_detect::Error::InvalidLimits);
        assert_eq!(invalid.kind(), ErrorKind::InvalidOptions);
        assert_eq!(invalid.stage(), Stage::Validation);
    }

    #[test]
    fn keynote_limits_map_every_known_resource_exactly() {
        let semantic_resources = [
            (
                litchi_keynote::SemanticLimitKind::Objects,
                Resource::Objects,
            ),
            (litchi_keynote::SemanticLimitKind::Slides, Resource::Slides),
            (
                litchi_keynote::SemanticLimitKind::References,
                Resource::References,
            ),
            (
                litchi_keynote::SemanticLimitKind::TextStorages,
                Resource::TextStorages,
            ),
            (
                litchi_keynote::SemanticLimitKind::TextFragments,
                Resource::TextFragments,
            ),
            (
                litchi_keynote::SemanticLimitKind::TextBytes,
                Resource::TextBytes,
            ),
        ];
        for (kind, resource) in semantic_resources {
            let error = map_keynote(litchi_keynote::ReadError::SemanticLimit {
                kind,
                observed: 9,
                maximum: 8,
                path: litchi_keynote::SemanticPath::Package,
            });
            assert_limit_shape(error, Format::Keynote, Stage::Semantic, resource);
        }

        let payload_resources = [
            (
                litchi_keynote::PayloadLimitKind::Bytes,
                Resource::PayloadBytes,
            ),
            (litchi_keynote::PayloadLimitKind::Fields, Resource::Fields),
            (
                litchi_keynote::PayloadLimitKind::Nesting,
                Resource::NestingDepth,
            ),
            (
                litchi_keynote::PayloadLimitKind::Work,
                Resource::SemanticWork,
            ),
        ];
        for (kind, resource) in payload_resources {
            let error = map_keynote(litchi_keynote::ReadError::PayloadLimit {
                kind,
                observed: 9,
                maximum: 8,
                path: litchi_keynote::SemanticPath::Package,
            });
            assert_limit_shape(error, Format::Keynote, Stage::Semantic, resource);
        }
    }

    #[test]
    fn pages_body_storage_limit_uses_the_shared_exact_resource() {
        let error = map_pages(litchi_pages::PackageError::Semantic(
            litchi_pages::Error::TooManyBodyStorages {
                actual: 9,
                limit: 8,
            },
        ));
        assert_limit_shape(
            error,
            Format::Pages,
            Stage::Semantic,
            Resource::TextStorages,
        );
    }

    #[test]
    fn pages_text_limit_preserves_the_actual_observed_bytes() {
        let error = map_pages(litchi_pages::PackageError::Semantic(
            litchi_pages::Error::TextTooLarge {
                observed: 20,
                limit: 4,
            },
        ));
        assert_eq!(error.kind(), ErrorKind::LimitExceeded);
        assert_eq!(error.stage(), Stage::Semantic);
        assert_eq!(error.format(), Some(Format::Pages));
        assert_eq!(error.resource(), Some(Resource::TextBytes));
        assert_eq!(error.observed(), Some(20));
        assert_eq!(error.maximum(), Some(4));
        assert!(std::error::Error::source(&error).is_none());
    }

    #[test]
    fn numbers_limits_map_every_known_resource_exactly() {
        let resources = [
            (
                litchi_numbers::SemanticLimitKind::Objects,
                Resource::Objects,
            ),
            (litchi_numbers::SemanticLimitKind::Sheets, Resource::Sheets),
            (litchi_numbers::SemanticLimitKind::Tables, Resource::Tables),
            (
                litchi_numbers::SemanticLimitKind::References,
                Resource::References,
            ),
            (
                litchi_numbers::SemanticLimitKind::MaterializedCells,
                Resource::Cells,
            ),
            (
                litchi_numbers::SemanticLimitKind::OutputTextBytes,
                Resource::TextBytes,
            ),
            (
                litchi_numbers::SemanticLimitKind::FormulaRenderWork,
                Resource::SemanticWork,
            ),
            (
                litchi_numbers::SemanticLimitKind::FormulaRenderDepth,
                Resource::NestingDepth,
            ),
            (
                litchi_numbers::SemanticLimitKind::FormulaWireBytes,
                Resource::PayloadBytes,
            ),
            (
                litchi_numbers::SemanticLimitKind::FormulaWork,
                Resource::SemanticWork,
            ),
            (
                litchi_numbers::SemanticLimitKind::TextBytes,
                Resource::TextBytes,
            ),
            (
                litchi_numbers::SemanticLimitKind::FormulaDepth,
                Resource::NestingDepth,
            ),
        ];
        for (kind, resource) in resources {
            let error = map_numbers(litchi_numbers::PackageError::SemanticLimit {
                kind,
                observed: 9,
                maximum: 8,
                path: litchi_numbers::PackageSemanticPath::Package,
            });
            assert_limit_shape(error, Format::Numbers, Stage::Semantic, resource);
        }
    }

    #[test]
    fn numbers_payload_limits_map_every_known_resource_exactly() {
        use litchi_numbers::{PackagePayloadLimitKind as LimitKind, PackageResourceError};

        let resources = [
            (LimitKind::InputBytes, Resource::PayloadBytes),
            (LimitKind::Fields, Resource::Fields),
            (LimitKind::OutputBytes, Resource::OutputBytes),
            (LimitKind::Nesting, Resource::NestingDepth),
            (LimitKind::RewriteWork, Resource::SemanticWork),
            (LimitKind::TableRows, Resource::Cells),
            (LimitKind::TableColumns, Resource::Cells),
            (LimitKind::TableCells, Resource::Cells),
            (LimitKind::MaterializedCells, Resource::Cells),
        ];
        for (kind, resource) in resources {
            let error = map_numbers_resource_error(PackageResourceError::LimitExceeded {
                kind,
                observed: 9,
                maximum: 8,
            });
            assert_limit_shape(error, Format::Numbers, Stage::Semantic, resource);
        }
    }

    #[test]
    fn numbers_allocation_preserves_amount_without_internal_resource_text() {
        let error = map_numbers_resource_error(litchi_numbers::PackageResourceError::Allocation {
            amount: 17,
        });
        assert_eq!(error.kind(), ErrorKind::Allocation);
        assert_eq!(error.stage(), Stage::Semantic);
        assert_eq!(error.format(), Some(Format::Numbers));
        assert_eq!(error.resource(), Some(Resource::Memory));
        assert_eq!(error.observed(), Some(17));
        assert_eq!(error.maximum(), None);
        assert_eq!(error.to_string(), "iWork allocation failed");
        assert!(std::error::Error::source(&error).is_none());
    }

    #[test]
    fn aggregate_index_failures_are_internal_validation_invariants() {
        for (format, error) in [
            (
                Format::Keynote,
                litchi_iwa_structured::Error::InvalidSlideIndex {
                    expected: 0,
                    actual: 1,
                },
            ),
            (
                Format::Pages,
                litchi_iwa_structured::Error::InvalidSectionIndex {
                    expected: 0,
                    actual: 1,
                },
            ),
        ] {
            let error = map_aggregate(error, format);
            assert_eq!(error.kind(), ErrorKind::Invariant);
            assert_eq!(error.stage(), Stage::Validation);
            assert_eq!(error.format(), Some(format));
            assert_eq!(error.resource(), None);
            assert_eq!(error.observed(), None);
            assert_eq!(error.maximum(), None);
        }
    }

    #[test]
    fn mapped_errors_are_content_free_and_stage_neutral_in_display() {
        const PRIVATE: &str = "private-path/member/authored-value";
        let invalid = [
            map_keynote(litchi_keynote::ReadError::InvalidFormat(PRIVATE.to_owned())),
            map_numbers(litchi_numbers::PackageError::InvalidFormat(
                PRIVATE.to_owned(),
            )),
            map_pages(litchi_pages::PackageError::InvalidFormat(
                PRIVATE.to_owned(),
            )),
        ];
        for error in invalid {
            assert_eq!(error.kind(), ErrorKind::InvalidData);
            assert!(!format!("{error:?}").contains(PRIVATE));
            assert!(!error.to_string().contains(PRIVATE));
            assert_eq!(error.resource(), None);
            assert_eq!(error.observed(), None);
            assert_eq!(error.maximum(), None);
        }

        let allocation = map_keynote(litchi_keynote::ReadError::Allocation {
            resource: PRIVATE,
            amount: 17,
        });
        assert_eq!(allocation.kind(), ErrorKind::Allocation);
        assert_eq!(allocation.resource(), Some(Resource::Memory));
        assert_eq!(allocation.observed(), Some(17));
        assert_eq!(allocation.maximum(), None);
        assert_eq!(allocation.to_string(), "iWork allocation failed");
        assert!(!format!("{allocation:?}").contains(PRIVATE));

        let invariant = Error::invariant(Some(Format::Numbers), Stage::Validation);
        assert_eq!(invariant.to_string(), "iWork invariant failed");
        assert_eq!(invariant.resource(), None);
        assert_eq!(invariant.observed(), None);
        assert_eq!(invariant.maximum(), None);
    }

    fn assert_limit_shape(error: Error, format: Format, stage: Stage, resource: Resource) {
        assert_eq!(error.kind(), ErrorKind::LimitExceeded);
        assert_eq!(error.stage(), stage);
        assert_eq!(error.format(), Some(format));
        assert_eq!(error.resource(), Some(resource));
        assert_eq!(error.observed(), Some(9));
        assert_eq!(error.maximum(), Some(8));
        assert!(std::error::Error::source(&error).is_none());
    }
}
