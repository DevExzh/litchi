//! Apple iWork inputs (Pages, Keynote, Numbers) for the read classes.
//!
//! `litchi::open` opens iWork packages as `File::Pages`, `File::Keynote` and
//! `File::Numbers` since record 0296, and `litchi_py.open` returns them as
//! unified files. The family views (`File::document`, `presentation`,
//! `workbook`) do not cover iWork, so the `Document`, `Presentation` and
//! `Workbook` read classes read an iWork input through the format's
//! archive-free semantic document, which the facade re-exports as
//! `litchi::{pages, keynote, numbers}::semantic`. An iWork package
//! directory, which `litchi::open` still refuses, takes the same route.
//!
//! These are the documents the facade's `Document`, `Presentation` and
//! `sheet::Workbook` used to project, so the Python reads keep their output
//! (risk R6 in record 0003): Pages text and one paragraph per text block,
//! Keynote text with the show title and each slide's text, Numbers sheet
//! names and the sheet/table/cell text layout.

use std::path::Path;

use litchi::{Code, FileFormat};

use crate::errors::Failure;

/// Whether `error` is `litchi::open`'s refusal of an iWork input of
/// `format` (a package directory).
pub(crate) fn is_refused(error: &litchi::Error, format: FileFormat) -> bool {
    error.code() == Code::UnsupportedFormat && error.detected() == Some(format)
}

/// Classify one of the three formats' semantic read errors, which share
/// their shape.
macro_rules! failure {
    ($error:expr, $kind:path, $other_application:pat) => {{
        let error = $error;
        let message = error.to_string();
        let code = match &error {
            ReadError::Io { kind } if matches!(kind, $kind) => Code::FileNotFound,
            ReadError::Io { .. } | ReadError::InvalidSource => Code::IoError,
            $other_application => Code::UnsupportedFormat,
            ReadError::Limit { .. } => Code::LimitExceeded,
            ReadError::Allocation { .. } => Code::InternalError,
            // `InvalidFormat`, and any variant added later: a malformed
            // input, as the unified classifiers treat unnamed leaf variants.
            _ => Code::CorruptFile,
        };
        Failure::new(code, message)
    }};
}

/// Open a Pages document: the read-only projection of its package (a
/// single-file package or a package directory).
pub(crate) fn open_pages(path: &Path) -> Result<litchi::pages::semantic::Document, Failure> {
    litchi::pages::semantic::open(path).map_err(pages_failure)
}

/// Classify a Pages package error.
fn pages_failure(error: litchi::pages::semantic::PackageError) -> Failure {
    use litchi::pages::semantic::{Error as SemanticError, PackageError};
    let message = error.to_string();
    let code = match &error {
        PackageError::Io(error) if error.kind() == std::io::ErrorKind::NotFound => {
            Code::FileNotFound
        },
        PackageError::Io(_) => Code::IoError,
        PackageError::UnexpectedFormat { .. } | PackageError::NotOfficeFile => {
            Code::UnsupportedFormat
        },
        PackageError::Limit { .. }
        | PackageError::PayloadLimit { .. }
        | PackageError::ObjectLimit { .. }
        | PackageError::SectionNamesTooLarge { .. }
        | PackageError::Semantic(
            SemanticError::TooManySections { .. }
            | SemanticError::TextTooLarge { .. }
            | SemanticError::TooManyBodyStorages { .. },
        ) => Code::LimitExceeded,
        PackageError::Allocation { .. } => Code::InternalError,
        // `InvalidFormat`, and any variant added later: a malformed input,
        // as the unified classifiers treat unnamed leaf variants.
        _ => Code::CorruptFile,
    };
    Failure::new(code, message)
}

/// Open a Keynote document.
pub(crate) fn open_keynote(path: &Path) -> Result<litchi::keynote::semantic::Document, Failure> {
    litchi::keynote::semantic::Document::open(path).map_err(keynote_failure)
}

/// Classify a Keynote read error.
pub(crate) fn keynote_failure(error: litchi::keynote::semantic::DocumentReadError) -> Failure {
    use litchi::keynote::semantic::{DocumentIoKind, DocumentReadError as ReadError};
    failure!(error, DocumentIoKind::NotFound, ReadError::NotKeynote)
}

/// Open a Numbers document.
pub(crate) fn open_numbers(path: &Path) -> Result<litchi::numbers::semantic::Workbook, Failure> {
    litchi::numbers::semantic::Workbook::open(path).map_err(numbers_failure)
}

/// Classify a Numbers read error.
pub(crate) fn numbers_failure(error: litchi::numbers::semantic::ReadError) -> Failure {
    use litchi::numbers::semantic::ReadError;
    use litchi::numbers::workbook::IoKind;
    failure!(error, IoKind::NotFound, ReadError::UnexpectedFormat { .. })
}

/// The paragraph-like text values of a Pages document, in section order:
/// each section's heading, paragraphs and non-empty text storages. This is
/// the projection the facade's `Document::paragraphs` made, one value per
/// text block (a block can hold `\n`).
pub(crate) fn pages_paragraphs(document: &litchi::pages::semantic::Document) -> Vec<String> {
    document
        .sections()
        .iter()
        .flat_map(|section| {
            section
                .heading()
                .into_iter()
                .chain(section.paragraphs().iter().map(String::as_str))
                .chain(
                    section
                        .text_storages()
                        .iter()
                        .filter(|storage| !storage.is_empty())
                        .map(|storage| storage.text()),
                )
        })
        .map(str::to_owned)
        .collect()
}
