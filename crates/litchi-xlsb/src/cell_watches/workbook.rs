//! Workbook/package traversal for worksheet cell-watch snapshots.

use super::worksheet::{self, Commit, Snapshot};
use crate::package::error::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part};

const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// A worksheet part paired with its typed cell-watch snapshot.
#[derive(Debug, Clone)]
pub struct Sheet {
    /// Absolute OPC part name, retained for package-level diagnostics.
    pub part_name: String,
    /// Immutable worksheet stream snapshot.
    pub snapshot: Snapshot,
}

/// Read one worksheet's cell-watch and phonetic owner.
pub fn read(package: &OpcPackage, worksheet: &PackURI) -> Result<Snapshot> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    worksheet::read(part.blob())
}

/// Read every worksheet in deterministic part-name order.
pub fn load(package: &OpcPackage) -> Result<Vec<Sheet>> {
    let mut parts: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| part.content_type() == WORKSHEET_CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .collect();
    parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    parts
        .into_iter()
        .map(|part_name| {
            let snapshot = read(package, &part_name)?;
            Ok(Sheet {
                part_name: part_name.to_string(),
                snapshot,
            })
        })
        .collect()
}

/// Atomically apply a source-checked worksheet commit to an OPC package.
///
/// The candidate package is cloned and reparsed through the XLSB workbook
/// owner before publication. A failed parse or source guard leaves the input
/// package unchanged.
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
    worksheet::read(&updated)
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
