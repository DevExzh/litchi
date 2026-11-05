//! Apple iWork inputs (Pages, Keynote, Numbers) for the read classes.
//!
//! `litchi::open` opens iWork packages and package directories as
//! `File::Pages`, `File::Keynote` and `File::Numbers` (records 0296 and
//! 0306), and `litchi_py.open` returns them as unified files. The family
//! views (`File::document`, `presentation`, `workbook`) do not cover iWork,
//! so the `Document`, `Presentation` and `Workbook` read classes hold the
//! opened `File` and project the leaf package it retains: Pages text and one
//! paragraph per text block, Keynote text and each slide's text, Numbers
//! sheet names and the sheet/table/cell text layout. These are the
//! projections the classes made before (risk R6 in record 0003), now read
//! from the one `File` open instead of a second, semantic-only open.

use litchi::{Code, File, FileFormat};

use crate::errors::Failure;

/// The `File` of an iWork variant was expected but another one arrived.
fn not_iwork(format: FileFormat) -> Failure {
    Failure::new(
        Code::InternalError,
        format!("expected a {format:?} file from the unified open"),
    )
}

/// The Pages package an opened `File::Pages` retains.
pub(crate) fn pages(file: &File) -> Result<&litchi::pages::Package, Failure> {
    match file {
        File::Pages(handle) => handle.pages(),
        _ => None,
    }
    .ok_or_else(|| not_iwork(FileFormat::Pages))
}

/// The Keynote package an opened `File::Keynote` retains.
pub(crate) fn keynote(file: &File) -> Result<&litchi::keynote::Package, Failure> {
    match file {
        File::Keynote(handle) => handle.keynote(),
        _ => None,
    }
    .ok_or_else(|| not_iwork(FileFormat::Keynote))
}

/// The Numbers package an opened `File::Numbers` retains.
pub(crate) fn numbers(file: &File) -> Result<&litchi::numbers::Package, Failure> {
    match file {
        File::Numbers(handle) => handle.numbers(),
        _ => None,
    }
    .ok_or_else(|| not_iwork(FileFormat::Numbers))
}

/// Classify a Keynote slide read error.
pub(crate) fn keynote_failure(error: &litchi::keynote::ReadError) -> Failure {
    use litchi::keynote::ReadError;
    let code = match error {
        ReadError::Io(_) => Code::IoError,
        ReadError::SemanticLimit { .. } | ReadError::PayloadLimit { .. } => Code::LimitExceeded,
        ReadError::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    };
    Failure::new(code, error.to_string())
}

/// The paragraph-like text values of a Pages document, in section order:
/// each section's heading, paragraphs and non-empty text storages. This is
/// the projection the facade's `Document::paragraphs` made, one value per
/// text block (a block can hold `\n`).
pub(crate) fn pages_paragraphs(package: &litchi::pages::Package) -> Vec<String> {
    package
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
