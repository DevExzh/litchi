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
    /// The Keynote package could not be read.
    #[cfg(feature = "keynote")]
    Keynote(litchi_keynote::ReadError),
    /// A Keynote slide title or body transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteText(litchi_keynote::SlideTextError),
    /// A Keynote slide skip transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteEdit(litchi_keynote::EditError),
    /// A Keynote slide order transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteOrder(litchi_keynote::SlideOrderError),
    /// A Keynote slide deletion failed.
    #[cfg(feature = "keynote")]
    KeynoteDeletion(litchi_keynote::slide::delete::Error),
    /// Writing the Keynote package to a sink failed.
    #[cfg(feature = "keynote")]
    KeynoteWrite(litchi_keynote::WriteError),
    /// The Numbers package could not be read.
    #[cfg(feature = "numbers")]
    Numbers(litchi_numbers::PackageError),
    /// The Numbers workbook text could not be rendered.
    #[cfg(feature = "numbers")]
    NumbersText(litchi_numbers::DocumentReadError),
    /// A Numbers table-cell transaction failed.
    #[cfg(feature = "numbers")]
    NumbersCells(litchi_numbers::table::cells::Error),
    /// Writing the Numbers package to a sink failed.
    #[cfg(feature = "numbers")]
    NumbersWrite(litchi_numbers::WriteError),
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
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => error.fmt(formatter),
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
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => Some(error),
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
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => numbers_code(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => match error {
                litchi_numbers::DocumentReadError::Limit { .. } => Code::LimitExceeded,
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
        }
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
        E::SectionNamesTooLarge { .. } | E::PayloadLimit { .. } | E::ObjectLimit { .. } => {
            Code::LimitExceeded
        },
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

#[cfg(feature = "keynote")]
fn keynote_code(error: &litchi_keynote::ReadError) -> Code {
    use litchi_keynote::ReadError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::SemanticLimit { .. } | E::PayloadLimit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "keynote")]
fn keynote_text_code(error: &litchi_keynote::SlideTextError) -> Code {
    use litchi_keynote::SlideTextError as E;
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
fn keynote_edit_code(error: &litchi_keynote::EditError) -> Code {
    use litchi_keynote::EditError as E;
    match error {
        E::SlideNameNotFound | E::SlidePositionNotFound { .. } => Code::PathNotFound,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_order_code(error: &litchi_keynote::SlideOrderError) -> Code {
    use litchi_keynote::SlideOrderError as E;
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

#[cfg(feature = "numbers")]
fn numbers_code(error: &litchi_numbers::PackageError) -> Code {
    use litchi_numbers::PackageError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
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
