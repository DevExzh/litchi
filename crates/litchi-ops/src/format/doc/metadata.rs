//! DOC metadata: initial metadata at creation and the `/metadata` set path
//! (record 0254).

use litchi_core::{FileFormat, Metadata};
use std::mem::size_of;

use super::{Doc, StepPatch, WriteCommit};
use crate::Props;
use crate::error::{Code, Error, Result};
use crate::format::ole_metadata;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Author creation-time metadata on the leaf builder.
pub(super) fn author(builder: &mut litchi_doc::Builder, metadata: &Metadata) -> Result<()> {
    builder
        .set_metadata(metadata.clone())
        .map_err(|error| match &error {
            litchi_doc::writer::WriteError::Metadata(inner) => {
                ole_metadata::metadata_error(inner, FileFormat::Doc)
            },
            _ => Error::from(error),
        })
}

impl Doc {
    /// Apply `props` to the document properties through one native metadata
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
            size_of::<litchi_doc::package::metadata::Edit>(),
        )?;
        let Some(package) = self.editable_package_for_write()? else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the DOC document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Doc)
            .with_suggestion("reopen the document without `OpenOptions::retain_source(false)`"));
        };
        let mut edit = package.edit_metadata()?;
        let operations = ole_metadata::apply(edit.properties_mut()?, props, FileFormat::Doc)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        if changed && self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Doc));
        }
        let staged_snapshot_bytes =
            ole_metadata::snapshot_bytes(patch.before_bytes(), patch.after_bytes());
        let (package, patch, snapshot_bytes) =
            if !changed && self.signature() == super::Signature::Signed {
                // The writable probe is an unsigned private package. Rebuild
                // an identity metadata patch over the retained signed owner
                // so replay and the facade have the same physical base.
                let original = self.original_package_for_noop()?;
                self.preflight_snapshot_bytes(
                    limits.max_snapshot_bytes(),
                    size_of::<litchi_doc::package::metadata::Edit>(),
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
                self.with_package(package)
            } else {
                self.clone()
            },
            patch: StepPatch::Metadata(patch),
            operations,
            snapshot_bytes,
        })
    }
}
