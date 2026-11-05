//! ODS document properties (record 0277): initial metadata at creation and
//! the `/metadata` set path.
//!
//! `set("/metadata", props)` commits through the leaf's ADR 0034 §7
//! metadata family (`Package::edit_metadata`, record 0288; the shared
//! `format::odf_metadata` path): `title`, `subject`, `description`,
//! `author`, `lastModifiedBy`, `language` and `keywords`; `null` removes a
//! field. Every other root-metadata key the OLE and OOXML formats accept is
//! a typed `unsupported_capability` naming the property, and an unknown key
//! is `unsupported_property`. A set never stamps `modified` (determinism,
//! ADR 0034 §9).

use litchi_core::{FileFormat, Metadata};
use litchi_ods::Package;

use super::write::{Commit, Patch};
use super::{Ods, core_error, odf_error};
use crate::error::Result;
use crate::format::odf_metadata;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Build the minimal spreadsheet (one empty `Sheet1`) with creation-time
/// metadata, through the leaf builder's fresh-authoring metadata.
pub(super) fn create(metadata: &Metadata) -> Result<Package> {
    let mut builder = litchi_ods::Builder::new();
    builder.set_metadata(metadata.clone()).map_err(core_error)?;
    builder
        .add_sheet(litchi_ods::Sheet::new("Sheet1").map_err(core_error)?)
        .map_err(core_error)?;
    builder.build().map_err(odf_error)
}

/// Whether `target` is exactly the `/metadata` node.
pub(crate) fn is_target(target: &TargetArg<'_>) -> bool {
    odf_metadata::is_target(target)
}

/// Commit `/metadata` assignments through the leaf metadata family.
pub(super) fn set(file: &Ods, props: &Props<'_>, limits: &SemanticLimits) -> Result<Commit> {
    let package = file.materialized()?;
    let mut edit = package.edit_metadata();
    odf_metadata::stage(&mut edit, props, FileFormat::Ods, limits)?;
    let (next, patch, changed) = odf_metadata::commit(edit, FileFormat::Ods)?;
    let snapshot_bytes = if changed {
        package
            .as_bytes()
            .len()
            .saturating_add(next.as_bytes().len())
    } else {
        0
    };
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(TransactionError::Limit {
            kind: LimitKind::SnapshotBytes,
            observed: snapshot_bytes,
            maximum: limits.max_snapshot_bytes(),
        }
        .into());
    }
    let after = if changed {
        file.with_package(next)?
    } else {
        file.clone()
    };
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Metadata(patch),
        operations: props.len(),
        snapshot_bytes,
    })
}
