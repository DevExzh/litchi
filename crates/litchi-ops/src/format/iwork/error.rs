//! The typed leaf errors of the iWork owners, kept whole behind one
//! [`Source`](crate::error::Source) variant (record 0296).

use std::fmt;

use crate::error::Code;

/// A Pages, Keynote or Numbers leaf error, as the owner raised it.
///
/// [`Error::leaf`](crate::Error::leaf) returns it inside
/// [`Source::Iwork`](crate::error::Source::Iwork); the standard error source
/// chain reaches the leaf error itself.
#[derive(Debug)]
#[non_exhaustive]
pub enum LeafError {
    /// iWork package detection or preparation refused the input.
    Detection(litchi_iwa_detect::Error),
    /// The Pages package could not be read.
    #[cfg(feature = "pages")]
    Pages(litchi_pages::PackageError),
    /// A Pages section-text transaction failed.
    #[cfg(feature = "pages")]
    PagesText(litchi_pages::section::text::Error),
    /// Writing the Pages package to a sink failed.
    #[cfg(feature = "pages")]
    PagesWrite(litchi_pages::WriteError),
    /// A Pages paragraph transaction failed.
    #[cfg(feature = "pages")]
    PagesParagraphs(litchi_pages::section::paragraphs::Error),
    /// A Pages section-name transaction failed.
    #[cfg(feature = "pages")]
    PagesName(litchi_pages::section::name::Error),
    /// A Pages character-style transaction failed (record 0309).
    #[cfg(feature = "pages")]
    PagesTextStyle(litchi_pages::section::text_style::Error),
    /// A Pages body-table value transaction failed (record 0309).
    #[cfg(feature = "pages")]
    PagesTableValues(litchi_pages::table::values::Error),
    /// A Pages header or footer text transaction failed.
    #[cfg(feature = "pages")]
    PagesHeaderFooter(litchi_pages::header_footer::text::Error),
    /// A Pages footnote text transaction failed.
    #[cfg(feature = "pages")]
    PagesFootnote(litchi_pages::footnote::text::Error),
    /// A Pages body-footnote lifecycle transaction failed.
    #[cfg(feature = "pages")]
    PagesFootnotes(litchi_pages::footnote::body::transaction::Error),
    /// A Pages document-properties transaction failed.
    #[cfg(feature = "pages")]
    PagesMetadata(litchi_pages::metadata::Error),
    /// Saving a Pages package directory failed (record 0301).
    #[cfg(feature = "pages")]
    PagesSave(litchi_pages::SaveError),
    /// A Pages body-table deletion failed.
    #[cfg(feature = "pages")]
    PagesTableDeletion(litchi_pages::table::deletion::Error),
    /// The Keynote package could not be read.
    #[cfg(feature = "keynote")]
    Keynote(litchi_keynote::ReadError),
    /// A Keynote slide title or body transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteText(litchi_keynote::slide::text::Error),
    /// A Keynote slide skip transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteEdit(litchi_keynote::edit::Error),
    /// A Keynote slide order transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteOrder(litchi_keynote::slide::order::Error),
    /// A Keynote slide deletion failed.
    #[cfg(feature = "keynote")]
    KeynoteDeletion(litchi_keynote::slide::delete::Error),
    /// Writing the Keynote package to a sink failed.
    #[cfg(feature = "keynote")]
    KeynoteWrite(litchi_keynote::WriteError),
    /// A Keynote speaker-notes transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteNotes(litchi_keynote::slide::notes::Error),
    /// A Keynote slide-drawable transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteDrawables(litchi_keynote::slide::drawable::Error),
    /// A Keynote movie or audio removal failed (record 0309).
    #[cfg(feature = "keynote")]
    KeynoteMedia(litchi_keynote::slide::media::lifecycle::Error),
    /// A Keynote placeholder visibility change failed (record 0309).
    #[cfg(feature = "keynote")]
    KeynotePlaceholder(litchi_keynote::slide::placeholder::Error),
    /// A Keynote slide-creation transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteCreation(litchi_keynote::slide::creation::Error),
    /// Saving a Keynote package directory failed (record 0306).
    #[cfg(feature = "keynote")]
    KeynoteSave(litchi_keynote::SaveError),
    /// The Numbers package could not be read.
    #[cfg(feature = "numbers")]
    Numbers(litchi_numbers::Error),
    /// The Numbers workbook text could not be rendered.
    #[cfg(feature = "numbers")]
    NumbersText(litchi_numbers::workbook::ReadError),
    /// A Numbers table-cell transaction failed.
    #[cfg(feature = "numbers")]
    NumbersCells(litchi_numbers::table::cells::Error),
    /// Writing the Numbers package to a sink failed.
    #[cfg(feature = "numbers")]
    NumbersWrite(litchi_numbers::WriteError),
    /// A Numbers transaction other than the cell transaction failed: sheet
    /// and table names, sheet order, table relocation, merges, comments,
    /// display formats, formula text, or a fresh workbook build.
    #[cfg(feature = "numbers")]
    NumbersEdit(super::numbers::EditError),
    /// Saving a Numbers package directory failed (record 0306).
    #[cfg(feature = "numbers")]
    NumbersSave(litchi_numbers::SaveError),
}

impl fmt::Display for LeafError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Detection(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::Pages(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesWrite(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for LeafError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Detection(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::Pages(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesWrite(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.source_error(),
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => Some(error),
        }
    }
}

impl LeafError {
    /// The unified code of the leaf error.
    ///
    /// Read failures of a package detection accepted are corrupt files,
    /// except typed resource ceilings (`limit_exceeded`) and I/O. Edit
    /// failures are the request's fault: a selector that matches nothing is
    /// `path_not_found`, a refused value `invalid_value`, a native shape the
    /// leaf does not rewrite yet `unsupported_capability`, a protected table
    /// `document_protected`, and a failed readback `internal_error`.
    #[must_use]
    pub(crate) fn code(&self) -> Code {
        match self {
            Self::Detection(error) => detection_code(error),
            #[cfg(feature = "pages")]
            Self::Pages(error) => pages_code(error),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => pages_text_code(error),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => pages_paragraphs_code(error),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => pages_name_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => pages_text_style_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => pages_table_values_code(error),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => pages_region_code(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => pages_footnote_code(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => pages_footnotes_code(error),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => pages_metadata_code(error),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => match error.write_error() {
                Some(_) => Code::IoError,
                None if matches!(error, litchi_pages::SaveError::Bundle(_)) => Code::InvalidInput,
                None => Code::IoError,
            },
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => pages_deletion_code(error),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => keynote_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => keynote_text_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => keynote_edit_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => keynote_order_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => keynote_deletion_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => keynote_notes_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => keynote_drawables_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => keynote_media_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => keynote_placeholder_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => keynote_creation_code(error),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => numbers_code(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => match error {
                litchi_numbers::workbook::ReadError::Limit { .. } => Code::LimitExceeded,
                _ => Code::InternalError,
            },
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => numbers_cells_code(error),
            #[cfg(feature = "pages")]
            Self::PagesWrite(_) => Code::IoError,
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(_) => Code::IoError,
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(_) => Code::IoError,
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.code(),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => match error {
                litchi_keynote::SaveError::Bundle(_) => Code::InvalidInput,
                _ => Code::IoError,
            },
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => match error {
                litchi_numbers::SaveError::Bundle(_) => Code::InvalidInput,
                _ => Code::IoError,
            },
        }
    }
}

/// The code of a physical iWork archive failure: a resource ceiling is
/// `limit_exceeded` (record 0306: source limits and package-directory
/// captures report through it), I/O is `io_error`, the rest is a malformed
/// package.
fn archive_code(error: &litchi_iwa_archive::Error) -> Code {
    use litchi_iwa_archive::Error as E;
    match error {
        E::Limit { .. } => Code::LimitExceeded,
        E::Io(_) => Code::IoError,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

fn detection_code(error: &litchi_iwa_detect::Error) -> Code {
    match error {
        litchi_iwa_detect::Error::LimitExceeded { .. } => Code::LimitExceeded,
        litchi_iwa_detect::Error::Io(_) => Code::IoError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "pages")]
fn pages_code(error: &litchi_pages::PackageError) -> Code {
    use litchi_pages::PackageError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::Limit { .. }
        | E::Semantic(
            litchi_pages::Error::TooManySections { .. }
            | litchi_pages::Error::TextTooLarge { .. }
            | litchi_pages::Error::TooManyBodyStorages { .. },
        )
        | E::SectionNamesTooLarge { .. }
        | E::PayloadLimit { .. }
        | E::ObjectLimit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "pages")]
fn pages_text_code(error: &litchi_pages::section::text::Error) -> Code {
    use litchi_pages::section::text::Error as E;
    match error {
        E::NameNotFound | E::PositionNotFound { .. } => Code::PathNotFound,
        E::AmbiguousSelector { .. } | E::BodySectionCount { .. } => Code::InvalidPath,
        E::SpanOutOfBounds { .. } | E::SurrogateBoundary { .. } => Code::OutOfRange,
        E::SectionBreakReplacement
        | E::FootnoteAnchorReplacement
        | E::ObjectMarkerReplacement
        | E::DependentContent
        | E::UnsupportedSource => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::Allocation { .. } | E::OperationAlreadyStaged | E::Verification => Code::InternalError,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InvalidValue,
    }
}

#[cfg(feature = "pages")]
fn pages_paragraphs_code(error: &litchi_pages::section::paragraphs::Error) -> Code {
    use litchi_pages::section::paragraphs::Error as E;
    match error {
        E::Text(error) => pages_text_code(error),
        E::PositionNotFound { .. } => Code::PathNotFound,
        E::ParagraphBreak => Code::InvalidValue,
        E::InterleavedAnchors | E::AnchoredParagraph => Code::UnsupportedCapability,
        E::LastParagraph => Code::InvalidArgument,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_name_code(error: &litchi_pages::section::name::Error) -> Code {
    use litchi_pages::section::name::Error as E;
    match error {
        E::NameNotFound | E::PositionNotFound { .. } => Code::PathNotFound,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::UnsupportedSource => Code::UnsupportedCapability,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::InvalidValue,
    }
}

#[cfg(feature = "pages")]
fn pages_text_style_code(error: &litchi_pages::section::text_style::Error) -> Code {
    use litchi_pages::section::text_style::Error as E;
    match error {
        E::SectionNotFound => Code::PathNotFound,
        E::InvalidSpan | E::EmptyStyle => Code::InvalidArgument,
        E::UnsupportedDependency | E::UnsupportedSource => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_placeholder_code(error: &litchi_keynote::slide::placeholder::Error) -> Code {
    use litchi_keynote::slide::placeholder::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } | E::PlaceholderNotFound { .. } => {
            Code::PathNotFound
        },
        E::UnsupportedSource => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::InvalidSource | E::AmbiguousSelector => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_media_code(error: &litchi_keynote::slide::media::lifecycle::Error) -> Code {
    use litchi_keynote::slide::media::lifecycle::Error as E;
    match error {
        E::SlideNameNotFound
        | E::SlidePositionNotFound { .. }
        | E::MoviePositionNotFound { .. } => Code::PathNotFound,
        E::UnsupportedSource | E::UnsupportedComment | E::AudioPoster | E::KindMismatch { .. } => {
            Code::UnsupportedCapability
        },
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict | E::Read => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_table_values_code(error: &litchi_pages::table::values::Error) -> Code {
    use litchi_pages::table::values::Error as E;
    match error {
        E::TableNotFound => Code::PathNotFound,
        E::InvalidCell => Code::InvalidValue,
        E::Formulas | E::UnsupportedStorage | E::UnsupportedSource => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_region_code(error: &litchi_pages::header_footer::text::Error) -> Code {
    use litchi_pages::header_footer::text::Error as E;
    match error {
        E::NotFound | E::PositionOutOfBounds => Code::PathNotFound,
        E::TextTooLarge | E::StructuralMarker => Code::InvalidValue,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::UnsupportedCapability,
    }
}

#[cfg(feature = "pages")]
fn pages_footnote_code(error: &litchi_pages::footnote::text::Error) -> Code {
    use litchi_pages::footnote::text::Error as E;
    match error {
        E::NotFound => Code::PathNotFound,
        E::TextTooLarge | E::CustomMarkTooLarge | E::StructuralMarker => Code::InvalidValue,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::UnsupportedCapability,
    }
}

#[cfg(feature = "pages")]
fn pages_footnotes_code(error: &litchi_pages::footnote::body::transaction::Error) -> Code {
    use litchi_pages::footnote::body::transaction::Error as E;
    match error {
        E::NotFound => Code::PathNotFound,
        E::PositionOutOfBounds => Code::OutOfRange,
        E::PositionOccupied | E::StructuralMarker | E::TextTooLarge | E::CustomMarkTooLarge => {
            Code::InvalidValue
        },
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::UnsupportedCapability,
    }
}

#[cfg(feature = "pages")]
fn pages_metadata_code(error: &litchi_pages::metadata::Error) -> Code {
    use litchi_pages::metadata::Error as E;
    match error {
        E::InvalidValue => Code::InvalidValue,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict | E::InvalidSource => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_deletion_code(error: &litchi_pages::table::deletion::Error) -> Code {
    use litchi_pages::table::deletion::Error as E;
    match error {
        E::TableNotFound => Code::PathNotFound,
        E::AmbiguousTableName | E::AmbiguousSelector => Code::InvalidPath,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::PatchConflict => Code::InvalidInput,
        E::Verification => Code::InternalError,
        _ => Code::UnsupportedCapability,
    }
}

#[cfg(feature = "keynote")]
fn keynote_code(error: &litchi_keynote::ReadError) -> Code {
    use litchi_keynote::ReadError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::SemanticLimit { .. } | E::PayloadLimit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "keynote")]
fn keynote_text_code(error: &litchi_keynote::slide::text::Error) -> Code {
    use litchi_keynote::slide::text::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } | E::TextStorageNotFound { .. } => {
            Code::PathNotFound
        },
        E::AmbiguousSelector | E::EmptySlideName => Code::InvalidPath,
        E::SpanOutOfBounds { .. } | E::SurrogateBoundary { .. } => Code::OutOfRange,
        E::UnsupportedSource
        | E::UnsupportedKind { .. }
        | E::ObjectMarkerReplacement
        | E::DependentContent => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::OperationAlreadyStaged => Code::InternalError,
        E::InvalidSource => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_edit_code(error: &litchi_keynote::edit::Error) -> Code {
    use litchi_keynote::edit::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } => Code::PathNotFound,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_order_code(error: &litchi_keynote::slide::order::Error) -> Code {
    use litchi_keynote::slide::order::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } => Code::PathNotFound,
        E::DestinationOutOfRange { .. } => Code::OutOfRange,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_deletion_code(error: &litchi_keynote::slide::delete::Error) -> Code {
    use litchi_keynote::slide::delete::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } => Code::PathNotFound,
        E::CannotDeleteFinalSlide => Code::InvalidArgument,
        _ => Code::UnsupportedCapability,
    }
}

#[cfg(feature = "keynote")]
fn keynote_notes_code(error: &litchi_keynote::slide::notes::Error) -> Code {
    use litchi_keynote::slide::notes::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } | E::NotesStorageNotFound => {
            Code::PathNotFound
        },
        E::AmbiguousSelector | E::EmptySlideName => Code::InvalidPath,
        E::SpanOutOfBounds { .. } | E::SurrogateBoundary { .. } => Code::OutOfRange,
        E::UnsupportedSource | E::ObjectMarkerReplacement | E::DependentContent => {
            Code::UnsupportedCapability
        },
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_drawables_code(error: &litchi_keynote::slide::drawable::Error) -> Code {
    use litchi_keynote::slide::drawable::Error as E;
    match error {
        E::SlideNameNotFound
        | E::SlidePositionNotFound { .. }
        | E::DrawablePositionNotFound { .. } => Code::PathNotFound,
        E::AmbiguousSelector | E::EmptySlideName => Code::InvalidPath,
        E::ZPositionOutOfRange { .. } => Code::OutOfRange,
        E::InvalidGeometry | E::InvalidImage | E::InvalidSpan => Code::InvalidValue,
        E::DrawableRemoved => Code::InvalidArgument,
        E::UnsupportedSource | E::Unsupported { .. } => Code::UnsupportedCapability,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_creation_code(error: &litchi_keynote::slide::creation::Error) -> Code {
    use litchi_keynote::slide::creation::Error as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } => Code::PathNotFound,
        E::AmbiguousSelector | E::EmptySlideName => Code::InvalidPath,
        E::LayoutNameNotFound | E::AmbiguousLayout | E::LayoutPositionNotFound { .. } => {
            Code::InvalidValue
        },
        E::PositionOutOfRange { .. } => Code::OutOfRange,
        E::UnsupportedSource | E::UnsupportedTopology | E::UnsupportedContent { .. } => {
            Code::UnsupportedCapability
        },
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "numbers")]
fn numbers_code(error: &litchi_numbers::Error) -> Code {
    use litchi_numbers::Error as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::SemanticLimit { .. } | E::InputTooLarge { .. } => Code::LimitExceeded,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "numbers")]
fn numbers_cells_code(error: &litchi_numbers::table::cells::Error) -> Code {
    use litchi_numbers::table::cells::Error as E;
    match error {
        E::SheetNotFound | E::TableNotFound => Code::PathNotFound,
        E::InvalidAddress | E::OutOfBounds { .. } => Code::OutOfRange,
        E::AmbiguousSource { .. } | E::DuplicatePosition { .. } => Code::InvalidPath,
        E::LimitExceeded { .. } => Code::LimitExceeded,
        E::TableLocked { .. } => Code::DocumentProtected,
        E::UnsupportedSource { .. } | E::UnsupportedDependency { .. } => {
            Code::UnsupportedCapability
        },
        E::InvalidSource { .. } | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}
