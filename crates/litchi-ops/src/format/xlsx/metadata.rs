//! XLSX document metadata on the workbook root (record 0253).
//!
//! A set on `/metadata`, or on `/` as OfficeCLI spells it (record 0338),
//! with OfficeCLI's document-property keys commits one leaf
//! [`litchi_xlsx::metadata`] transaction. The handle's open-time metadata is
//! replaced by the committed value, so reads of `/` observe it.

use litchi_core::{FileFormat, Metadata};
use litchi_xlsx as xlsx;

use super::write::Commit;
use super::{Xlsx, XlsxOwner};
use crate::error::Result;
use crate::format::metadata::{self, Assignments};
use crate::options::OpenOptions;
use crate::props::Props;
use crate::target::TargetArg;

fn native(error: xlsx::metadata::Error) -> crate::Error {
    metadata::native_error(error, FileFormat::Xlsx)
}

/// Commit `update` to `workbook`'s metadata.
pub(super) fn commit(
    workbook: &xlsx::Package,
    update: impl FnOnce(&mut Metadata),
    custom: Option<&Assignments>,
) -> Result<(xlsx::Package, Metadata, bool)> {
    let mut edit = workbook.edit_metadata().map_err(native)?;
    edit.update(update).map_err(native)?;
    if let Some(assignments) = custom.filter(|assignments| assignments.has_custom()) {
        let props = assignments.apply_custom(edit.custom_props(), FileFormat::Xlsx)?;
        edit.set_custom_props(props);
    }
    // The staged value is the committed one; no second read of the parts.
    let metadata = edit.metadata().map_err(native)?;
    let commit = edit.commit().map_err(native)?;
    let changed = commit.diagnostics().changed();
    let after = commit.into_package();
    Ok((after, metadata, changed))
}

impl Xlsx {
    /// Set document properties on `/metadata`, or on the root `/`
    /// (OfficeCLI's spelling, record 0338): the root sets nothing else, so a
    /// key that is not a document property is `unsupported_property`.
    /// `None` for other targets and for an empty root set.
    pub(super) fn try_set_metadata(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        let root = metadata::is_root(target) && !props.is_empty();
        if !(metadata::is_target(target) || root) {
            return Ok(None);
        }
        let assignments = metadata::admit(props, FileFormat::Xlsx)?;
        self.set_metadata(&assignments, props.len()).map(Some)
    }

    fn set_metadata(
        &self,
        assignments: &Assignments,
        authored_operations: usize,
    ) -> Result<Commit> {
        let source = self.editor_for_write()?;
        let (target, metadata, changed) =
            commit(&source, |value| assignments.apply(value), Some(assignments))?;
        let after = if changed {
            self.adopt_metadata(target.clone(), metadata)
        } else {
            self.clone()
        };
        self.check_source_for_replay()?;
        Commit::from_metadata(self, after, source, target, changed, authored_operations)
    }

    /// Adopt a metadata revision: [`Self::adopt_workbook`] with the new
    /// metadata.
    fn adopt_metadata(&self, workbook: xlsx::Package, metadata: Metadata) -> Self {
        let options = OpenOptions::new().retain_source(self.0.retain_source);
        let signature = match self.signature() {
            super::Signature::RemovalAuthorized => super::Signature::Unsigned,
            signature => signature,
        };
        Self::with_owner(
            XlsxOwner::Eager(workbook),
            metadata,
            self.0.metadata_repairs.clone(),
            &options,
        )
        .with_origin(self.origin())
        .with_signature(signature)
    }
}

/// Create-time metadata: commit `metadata` to a fresh workbook.
pub(super) fn create(workbook: xlsx::Package, metadata: &Metadata) -> Result<Xlsx> {
    let (workbook, metadata, _) = commit(&workbook, |value| *value = metadata.clone(), None)?;
    Ok(Xlsx::from_workbook(workbook, &OpenOptions::new())?.with_metadata(metadata))
}

impl Xlsx {
    fn with_metadata(&self, metadata: Metadata) -> Self {
        let XlsxOwner::Eager(workbook) = &self.0.owner else {
            return self.clone();
        };
        self.adopt_metadata(workbook.clone(), metadata)
    }
}
