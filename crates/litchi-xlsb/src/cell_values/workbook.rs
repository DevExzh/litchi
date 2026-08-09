//! Workbook-level publication for lossless cell-value edits.

use super::{Commit, Limits, Snapshot};
use crate::package::error::{Error, Result};
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
    crate::Workbook::from_opc_package(candidate.clone())?;
    *package = candidate;
    super::worksheet::read(&updated)
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
