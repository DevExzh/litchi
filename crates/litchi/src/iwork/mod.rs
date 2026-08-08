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

use std::{fmt, sync::Arc};

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
        let prepared = litchi_iwa_detect::PreparedSource::from_bytes_with_limits(value, physical)
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
            litchi_iwa_detect::PreparedSource::from_shared_bytes_with_limits(value, physical)
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
            let document = {
                let package =
                    litchi_pages::Package::__from_prepared_source(prepared).map_err(map_pages)?;
                package.semantic_document().clone()
            };
            let data = litchi_iwa_structured::StructuredData::from_pages_document_with_limits(
                document, aggregate,
            )
            .map_err(|error| map_aggregate(error, Format::Pages))?;
            Ok(Document::from_data(Format::Pages, data))
        },
        litchi_iwa_detect::Format::Keynote => {
            let semantic = keynote_limits(limits)?;
            let document = {
                let package = litchi_keynote::Package::__from_prepared_source(prepared, semantic)
                    .map_err(map_keynote)?;
                package.semantic_snapshot().map_err(map_keynote)?
            };
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
        litchi_iwa_detect::Error::IwaCore(_)
        | litchi_iwa_detect::Error::IwaCommon(_)
        | litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_) => Error::invalid_data(None, Stage::Detection),
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
                Resource::SemanticWork,
                usize_u64(actual),
                usize_u64(limit),
            ),
            litchi_pages::Error::TextTooLarge { limit } => Error::limit(
                Some(Format::Pages),
                Stage::Semantic,
                Resource::TextBytes,
                usize_u64(limit).saturating_add(1),
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
            observed, maximum, ..
        } => Error::limit(
            Some(Format::Keynote),
            Stage::Semantic,
            Resource::SemanticWork,
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
        litchi_keynote::SemanticLimitKind::Slides => Resource::Slides,
        litchi_keynote::SemanticLimitKind::TextBytes => Resource::TextBytes,
        _ => Resource::SemanticWork,
    }
}

fn map_numbers(error: litchi_numbers::PackageError) -> Error {
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

fn numbers_resource(kind: litchi_numbers::SemanticLimitKind) -> Resource {
    match kind {
        litchi_numbers::SemanticLimitKind::Tables => Resource::Tables,
        litchi_numbers::SemanticLimitKind::MaterializedCells => Resource::Cells,
        litchi_numbers::SemanticLimitKind::OutputTextBytes
        | litchi_numbers::SemanticLimitKind::TextBytes => Resource::TextBytes,
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
            Error::invalid_data(Some(format), Stage::Validation)
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
}
