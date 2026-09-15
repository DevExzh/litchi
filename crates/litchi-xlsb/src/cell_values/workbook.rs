#![allow(
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Workbook-level publication for lossless cell-value edits.

use super::{Commit, Limits, Snapshot, Value};
use crate::external_link::ExternalLinkLimits;
use crate::package::error::{Error, Result};
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI, Part};

const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Read editable cell values from one binary worksheet part with safe limits.
///
/// # Errors
///
/// Returns an error when the part is missing, not a worksheet, or has an
/// invalid BIFF12 record stream.
pub fn read(package: &OpcPackage, worksheet: &PackURI) -> Result<Snapshot> {
    read_with_limits(package, worksheet, Limits::DEFAULT)
}

/// Read editable cell values with an explicit finite policy.
///
/// # Errors
///
/// Returns an error when the part is missing, not a worksheet, exceeds a
/// selected limit, or has an invalid BIFF12 record stream.
pub fn read_with_limits(
    package: &OpcPackage,
    worksheet: &PackURI,
    limits: Limits,
) -> Result<Snapshot> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    super::worksheet::read_with_limits(part.blob(), limits)
}

/// Apply one source-checked cell-value commit atomically to an OPC package.
///
/// Only the selected worksheet part is changed. The complete XLSB workbook is
/// reparsed before publication, so invalid candidates never escape this
/// boundary and every untouched part remains owned by the existing lossless
/// OPC package machinery.
///
/// # Errors
///
/// Returns an error for a stale patch, a non-worksheet target, or a candidate
/// that fails complete XLSB workbook validation.
pub fn apply(package: &mut OpcPackage, worksheet: &PackURI, commit: &Commit) -> Result<Snapshot> {
    apply_with_external_link_limits(package, worksheet, commit, ExternalLinkLimits::default())
}

/// Apply one source-checked cell-value commit while validating candidate
/// workbook reparses with an explicit external-link policy.
pub fn apply_with_external_link_limits(
    package: &mut OpcPackage,
    worksheet: &PackURI,
    commit: &Commit,
    external_link_limits: ExternalLinkLimits,
) -> Result<Snapshot> {
    let applied = apply_retaining_parse(package, worksheet, commit, external_link_limits)?;
    match applied {
        Applied::Unchanged(snapshot) => Ok(snapshot),
        Applied::Published { snapshot, workbook } => {
            *package = workbook.into_opc_package();
            Ok(snapshot)
        },
    }
}

/// The outcome of one validated cell-value publication.
///
/// Validating a candidate requires parsing it as a complete XLSB workbook.
/// This type hands that parse back instead of dropping it, so a caller that
/// owns a [`crate::Workbook`] publishes the parse it already paid for rather
/// than reparsing the same candidate bytes a second time. A caller that owns
/// only an [`OpcPackage`] takes the package back out of it.
pub(crate) enum Applied {
    /// The patch reproduced the stored worksheet bytes exactly, so there is no
    /// candidate: nothing was cloned, nothing was parsed, and the caller's
    /// package is already the published package.
    Unchanged(Snapshot),
    /// The patch changed the worksheet and the candidate passed every check.
    Published {
        /// The published snapshot, identical to the committed one.
        snapshot: Snapshot,
        /// The validated candidate workbook. Its package carries the updated
        /// worksheet part and no signatures. Boxed because a parsed workbook
        /// is far larger than a snapshot.
        workbook: Box<crate::Workbook>,
    },
}

/// Apply one source-checked cell-value commit and hand the validated candidate
/// parse back to the caller without publishing it.
///
/// Nothing the caller owns is touched: every refusal below leaves the caller's
/// package, and any workbook parsed from it, exactly as they were. An exact
/// no-op is answered from the stored bytes, without cloning the package or
/// parsing a candidate at all.
///
/// # Errors
///
/// Returns an error for a stale patch, a non-worksheet target, or a candidate
/// that fails complete XLSB workbook validation.
pub(crate) fn apply_retaining_parse(
    package: &OpcPackage,
    worksheet: &PackURI,
    commit: &Commit,
    external_link_limits: ExternalLinkLimits,
) -> Result<Applied> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    let updated = commit.patch().apply(part.blob())?;
    if updated.as_slice() == part.blob() {
        return Ok(Applied::Unchanged(commit.snapshot().clone()));
    }

    let mut candidate = package.clone();
    candidate.get_part_mut(worksheet)?.set_blob(updated);
    candidate.unsign();
    let parsed = crate::Workbook::from_opc_package_with_external_link_limits(
        candidate,
        external_link_limits,
    )?;
    let worksheet_index = (0..parsed.worksheet_count())
        .find_map(|index| {
            parsed
                .worksheet_uri(index)
                .ok()
                .filter(|candidate_uri| candidate_uri == worksheet)
                .map(|_| index)
        })
        .ok_or_else(|| {
            Error::WorksheetNotFound(format!("candidate worksheet part {}", worksheet.as_str()))
        })?;
    let _ = parsed.worksheet(worksheet_index)?;
    validate_dependencies(commit.snapshot(), &parsed)?;
    Ok(Applied::Published {
        snapshot: commit.snapshot().clone(),
        workbook: Box::new(parsed),
    })
}

fn validate_dependencies(snapshot: &Snapshot, workbook: &crate::Workbook) -> Result<()> {
    for cell in snapshot.cells() {
        let reference = cell.reference();
        let style = usize::try_from(cell.style().get()).map_err(|_| {
            Error::InvalidCellReference(format!(
                "cell ({}, {}) style index does not fit this platform",
                reference.row(),
                reference.column()
            ))
        })?;
        let format = workbook.styles().get_cell_format(style).ok_or_else(|| {
            Error::InvalidCellReference(format!(
                "cell ({}, {}) references missing style index {style}",
                reference.row(),
                reference.column()
            ))
        })?;
        for (resource, index, present) in [
            (
                "font",
                format.font_id,
                workbook
                    .styles()
                    .get_font(usize::try_from(format.font_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "fill",
                format.fill_id,
                workbook
                    .styles()
                    .get_fill(usize::try_from(format.fill_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "border",
                format.border_id,
                workbook
                    .styles()
                    .get_border(usize::try_from(format.border_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "number format",
                format.num_fmt_id,
                format.num_fmt_id < 164
                    || workbook.styles().get_num_fmt(format.num_fmt_id).is_some(),
            ),
        ] {
            if !present {
                return Err(Error::InvalidCellReference(format!(
                    "cell ({}, {}) style {style} references missing {resource} {index}",
                    reference.row(),
                    reference.column()
                )));
            }
        }
        match cell.value() {
            Value::SharedStringIndex(index) => {
                let index = usize::try_from(*index).map_err(|_| {
                    Error::InvalidCellReference(format!(
                        "cell ({}, {}) shared-string index does not fit this platform",
                        reference.row(),
                        reference.column()
                    ))
                })?;
                let string = workbook.shared_strings().get(index).ok_or_else(|| {
                    Error::InvalidCellReference(format!(
                        "cell ({}, {}) references missing shared-string index {index}",
                        reference.row(),
                        reference.column()
                    ))
                })?;
                validate_string_fonts(string, workbook, reference)?;
            },
            Value::RichString(string) => validate_string_fonts(string, workbook, reference)?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_string_fonts(
    string: &crate::package::shared_strings::SharedString,
    workbook: &crate::Workbook,
    reference: super::Reference,
) -> Result<()> {
    for font_id in string
        .runs
        .iter()
        .map(|run| run.font_id)
        .chain(string.phonetic.iter().map(|phonetic| phonetic.font_id))
    {
        if workbook.styles().get_font(usize::from(font_id)).is_none() {
            return Err(Error::InvalidCellReference(format!(
                "cell ({}, {}) rich string references missing font {font_id}",
                reference.row(),
                reference.column()
            )));
        }
    }
    Ok(())
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == WORKSHEET_CONTENT_TYPE {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: WORKSHEET_CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        })
    }
}
