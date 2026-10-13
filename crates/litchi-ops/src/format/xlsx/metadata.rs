//! XLSX document metadata on the workbook root (record 0253).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
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
    workbook: &xlsx::Workbook,
    update: impl FnOnce(&mut Metadata),
) -> Result<(xlsx::Workbook, Metadata, bool)> {
    let mut edit = workbook.edit_metadata().map_err(native)?;
    edit.update(update).map_err(native)?;
    let commit = edit.commit().map_err(native)?;
    let changed = commit.diagnostics().changed();
    let after = commit.into_package();
    let metadata = after.metadata().map_err(native)?;
    Ok((after, metadata, changed))
}

impl Xlsx {
    /// Set document properties on the root path; `None` for other targets.
    pub(super) fn try_set_metadata(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        if !metadata::is_root(target) {
            return Ok(None);
        }
        let assignments = metadata::admit(props, FileFormat::Xlsx)?;
        self.set_metadata(&assignments).map(Some)
    }

    fn set_metadata(&self, assignments: &Assignments) -> Result<Commit> {
        let source = self.editor()?;
        let (target, metadata, changed) = commit(&source, |value| assignments.apply(value))?;
        let after = if changed {
            self.adopt_metadata(target.clone(), metadata)
        } else {
            self.clone()
        };
        self.check_source()?;
        Ok(Commit::from_metadata(after, source, target))
    }

    /// Adopt a metadata revision: [`Self::adopt_workbook`] with the new
    /// metadata.
    fn adopt_metadata(&self, workbook: xlsx::Workbook, metadata: Metadata) -> Self {
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
pub(super) fn create(workbook: xlsx::Workbook, metadata: &Metadata) -> Result<Xlsx> {
    let (workbook, metadata, _) = commit(&workbook, |value| *value = metadata.clone())?;
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
