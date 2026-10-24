//! XLS metadata: initial metadata at creation and the `/metadata` set path
//! (record 0254).

use litchi_core::{FileFormat, Metadata};
use litchi_xls::Package;
use std::mem::size_of;

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
        self.preflight_snapshot_bytes(
            limits.max_snapshot_bytes(),
            size_of::<litchi_xls::package::metadata::Edit>(),
        )?;
        let package =
            self.editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = package.edit_metadata()?;
        let operations = ole_metadata::apply(edit.properties_mut()?, props, FileFormat::Xls)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        if changed && self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Xls));
        }
        let staged_snapshot_bytes =
            ole_metadata::snapshot_bytes(patch.before_bytes(), patch.after_bytes());
        let (package, patch, snapshot_bytes) =
            if !changed && self.signature() == super::Signature::Signed {
                // Re-root the identity patch on the retained signed package;
                // the writable probe used an unsigned private package.
                let original = self.original_package_for_noop()?;
                let minimum = self
                    .retained_source_len()
                    .saturating_mul(2)
                    .saturating_add(size_of::<litchi_xls::Patch>())
                    .saturating_add(size_of::<litchi_xls::package::metadata::Edit>());
                if minimum > limits.max_snapshot_bytes() {
                    return Err(TransactionError::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: minimum,
                        maximum: limits.max_snapshot_bytes(),
                    }
                    .into());
                }
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
