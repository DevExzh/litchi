//! Workbook-level publication for lossless cell-value edits.

use super::{Commit, Limits, Snapshot, Value};
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
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    let updated = commit.patch().apply(part.blob())?;
    if updated.as_slice() == part.blob() {
        return Ok(commit.snapshot().clone());
    }

    let mut candidate = package.clone();
    candidate.get_part_mut(worksheet)?.set_blob(updated.clone());
    candidate.unsign();
    let parsed = crate::Workbook::from_opc_package(candidate.clone())?;
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
    *package = candidate;
    Ok(commit.snapshot().clone())
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
        if workbook.styles().get_cell_format(style).is_none() {
            return Err(Error::InvalidCellReference(format!(
                "cell ({}, {}) references missing style index {style}",
                reference.row(),
                reference.column()
            )));
        }
        if let Value::SharedStringIndex(index) = cell.value() {
            let index = usize::try_from(*index).map_err(|_| {
                Error::InvalidCellReference(format!(
                    "cell ({}, {}) shared-string index does not fit this platform",
                    reference.row(),
                    reference.column()
                ))
            })?;
            if workbook.shared_strings().get(index).is_none() {
                return Err(Error::InvalidCellReference(format!(
                    "cell ({}, {}) references missing shared-string index {index}",
                    reference.row(),
                    reference.column()
                )));
            }
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
