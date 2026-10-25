//! DOCX document properties on the root and `/metadata` (records 0338 and
//! 0342).
//!
//! A set on `/` (OfficeCLI's `set / title=...`) or on `/metadata` with
//! root-metadata keys commits one leaf [`litchi_docx::metadata`]
//! transaction: the core properties (record 0251), the application keys
//! (`company`, `manager`, `template`, `application` and their `extended.*`
//! spellings) and `custom.<name>` (record 0342). `contentType`, which no
//! OOXML part stores, is refused with `unsupported_property`. As for XLSX
//! and PPTX, a set never stamps `modified`.

use litchi_core::FileFormat;

use super::write::{Commit, WorkEdit};
use super::{Docx, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::format::metadata;
use crate::props::Props;
use crate::target::TargetArg;

impl Docx {
    /// Whether a set on `target` with `props` is a metadata set: `/metadata`,
    /// or the root with at least one property.
    pub(super) fn is_metadata_set(target: &TargetArg<'_>, props: &Props<'_>) -> bool {
        metadata::is_target(target) || (metadata::is_root(target) && !props.is_empty())
    }

    /// Commit the admitted document-property assignments of `props`; the
    /// handle's own edit policy applies (a signed document, one opened
    /// without its source, an authorized signature removal).
    fn commit_metadata(&self, props: &Props<'_>) -> Result<Docx> {
        self.edit_policy()?;
        let assignments = metadata::admit(props, FileFormat::Docx)?;
        if let Some(key) = assignments.unsupported_property() {
            return Err(Error::new(
                Code::UnsupportedProperty,
                "the format cannot store this metadata field",
            )
            .with_property(key)
            .with_format(FileFormat::Docx));
        }
        let package = if self.signature() == crate::format::signature::Signature::RemovalAuthorized
        {
            self.unsigned_package()?
        } else {
            match self.owner() {
                DocxOwner::Eager(package) => package.clone(),
                DocxOwner::Source(package) => self.promoted(package)?,
            }
        };
        let mut edit = package.edit_metadata();
        edit.update(|value| assignments.apply_core(value));
        if assignments.has_app() {
            let base = edit.app_props().map_err(native)?;
            edit.set_app_props(Some(assignments.apply_app(base)));
        }
        if assignments.has_custom() {
            let custom = assignments.apply_custom(edit.custom_props(), FileFormat::Docx)?;
            edit.set_custom_props(custom);
        }
        let commit = edit.commit().map_err(native)?;
        let changed = commit.diagnostics().changed();
        self.owner().check_source()?;
        Ok(if changed {
            self.adopt_eager_package(commit.into_package())
        } else {
            // An equal value publishes nothing (record 0253's no-op rule).
            self.clone()
        })
    }

    /// [`Self::commit_metadata`] for single-operation `File::set`, which
    /// keeps only the handle.
    pub(super) fn set_metadata_file(&self, props: &Props<'_>) -> Result<Docx> {
        self.commit_metadata(props)
    }

    /// [`Self::commit_metadata`] for a transaction, which keeps a
    /// main-document patch per step: a metadata step's is the handle's own
    /// editor committed empty (the main document is unchanged), and the
    /// coordinator compares the endpoints' core properties to see the change
    /// (record 0338).
    pub(super) fn set_metadata(&self, props: &Props<'_>, snapshot_budget: usize) -> Result<Commit> {
        let file = self.commit_metadata(props)?;
        let editor = self.edit_with_snapshot_budget(snapshot_budget)?;
        let (patch, diagnostics) = WorkEdit::from(editor)?.commit_patch()?;
        Ok(Commit::from_parts(file, patch, diagnostics))
    }
}

/// A leaf metadata error, with the application- and custom-properties
/// failures mapped as for the other OOXML formats.
fn native(error: litchi_docx::metadata::Error) -> Error {
    match error {
        litchi_docx::metadata::Error::Properties(error) => {
            metadata::native_error(error, FileFormat::Docx)
        },
        error => Error::from(error),
    }
}
