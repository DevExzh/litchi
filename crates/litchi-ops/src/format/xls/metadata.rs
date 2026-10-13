//! XLS metadata: initial metadata at creation and the `/metadata` set path
//! (record 0254).

use litchi_core::{FileFormat, Metadata};
use litchi_xls::Package;

use super::Xls;
use super::write::{StepPatch, WriteCommit};
use crate::Props;
use crate::error::{Error, Result};
use crate::format::ole_metadata;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Build the minimal workbook with creation-time metadata.
pub(super) fn create(metadata: &Metadata) -> Result<Package> {
    let mut builder = litchi_xls::Builder::new();
    builder
        .set_metadata(metadata.clone())
        .map_err(|error| match &error {
            litchi_xls::Error::Metadata(inner) => {
                ole_metadata::metadata_error(inner, FileFormat::Xls)
            },
            _ => Error::from(error),
        })?;
    Ok(builder.build()?)
}

impl Xls {
    /// Apply `props` to the workbook properties through one native metadata
    /// transaction.
    pub(super) fn set_metadata(
        &self,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        if props.len() > limits.max_operations() {
            return Err(TransactionError::Limit {
                kind: LimitKind::Operations,
                observed: props.len(),
                maximum: limits.max_operations(),
            }
            .into());
        }
        let package = self.editable_package()?;
        let mut edit = package.edit_metadata()?;
        let operations = ole_metadata::apply(edit.properties_mut(), props, FileFormat::Xls)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        let snapshot_bytes =
            ole_metadata::snapshot_bytes(patch.before_bytes(), patch.after_bytes());
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(TransactionError::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: snapshot_bytes,
                maximum: limits.max_snapshot_bytes(),
            }
            .into());
        }
        self.check_source()?;
        Ok(WriteCommit {
            before: self.clone(),
            after: if changed {
                self.with_package(package)?
            } else {
                self.clone()
            },
            patch: StepPatch::Metadata(patch),
            operations,
            snapshot_bytes,
        })
    }
}
