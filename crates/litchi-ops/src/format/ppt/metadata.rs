//! PPT metadata: initial metadata at creation and the `/metadata` set path
//! (record 0254).

use litchi_core::{FileFormat, Metadata};
use litchi_ppt::Package;
use std::mem::size_of;

use super::Ppt;
use super::write::{StepPatch, WriteCommit};
use crate::Props;
use crate::error::{Error, Result};
use crate::format::ole_metadata;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Build the minimal presentation with creation-time metadata.
pub(super) fn create(metadata: &Metadata) -> Result<Package> {
    let mut builder = litchi_ppt::Builder::new();
    builder
        .set_metadata(metadata.clone())
        .map_err(|error| match &error {
            litchi_ppt::WriteError::Metadata(inner) => {
                ole_metadata::metadata_error(inner, FileFormat::Ppt)
            },
            _ => Error::from(litchi_ppt::Error::Authoring(error)),
        })?;
    Ok(builder.build()?)
}

impl Ppt {
    /// Apply `props` to the presentation properties through one native metadata
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
        self.preflight_snapshot_bytes(
            limits.max_snapshot_bytes(),
            size_of::<litchi_ppt::package::metadata::Edit>(),
        )?;
        let package = self.editable_package_for_write()?;
        let mut edit = package.edit_metadata()?;
        let operations = ole_metadata::apply(edit.properties_mut()?, props, FileFormat::Ppt)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        if changed && self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Ppt));
        }
        let staged_snapshot_bytes =
            ole_metadata::snapshot_bytes(patch.before_bytes(), patch.after_bytes());
        let (package, patch, snapshot_bytes) =
            if !changed && self.signature() == super::Signature::Signed {
                // Re-root a signed no-op on the original signed package; the
                // writable probe above intentionally edits an unsigned copy.
                let original = self.original_package_for_noop()?;
                self.preflight_snapshot_bytes(
                    limits.max_snapshot_bytes(),
                    size_of::<litchi_ppt::package::metadata::Edit>(),
                )?;
                let (package, patch) = original.edit_metadata()?.commit()?.into_parts();
                let canonical_snapshot_bytes =
                    ole_metadata::snapshot_bytes(patch.before_bytes(), patch.after_bytes());
                (
                    package,
                    patch,
                    staged_snapshot_bytes.max(canonical_snapshot_bytes),
                )
            } else {
                (package, patch, staged_snapshot_bytes)
            };
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(TransactionError::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: snapshot_bytes,
                maximum: limits.max_snapshot_bytes(),
            }
            .into());
        }
        self.check_source_for_write()?;
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
