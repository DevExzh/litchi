//! DOCX core document properties on the root and `/metadata` (record 0338).
//!
//! A set on `/` (OfficeCLI's `set / title=...`) or on `/metadata` with
//! root-metadata keys commits one leaf [`litchi_docx::metadata`]
//! transaction, the core-properties family of record 0251. The leaf family
//! wraps the core-properties part only, so application keys
//! (`extended.company`, ...) and `custom.<name>` are refused with
//! `unsupported_property` (record 0253 F3). As for XLSX and PPTX, a set never
//! stamps `modified`.

use litchi_core::FileFormat;

use super::write::{Commit, WorkEdit};
use super::{Docx, DocxOwner};
use crate::error::Result;
use crate::format::metadata::{self, Scope};
use crate::props::Props;
use crate::target::TargetArg;

impl Docx {
    /// Whether a set on `target` with `props` is a metadata set: `/metadata`,
    /// or the root with at least one property.
    pub(super) fn is_metadata_set(target: &TargetArg<'_>, props: &Props<'_>) -> bool {
        metadata::is_target(target) || (metadata::is_root(target) && !props.is_empty())
    }

    /// Commit the admitted core-property assignments of `props`; the
    /// handle's own edit policy applies (a signed document, one opened
    /// without its source, an authorized signature removal).
    fn commit_metadata(&self, props: &Props<'_>) -> Result<Docx> {
        self.edit_policy()?;
        let assignments = metadata::admit_with(props, FileFormat::Docx, Scope::Core)?;
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
        edit.update(|value| assignments.apply(value));
        let commit = edit.commit()?;
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
    pub(super) fn set_metadata(&self, props: &Props<'_>) -> Result<Commit> {
        let file = self.commit_metadata(props)?;
        let (patch, diagnostics) = WorkEdit::from(self.edit()?)?.commit_patch()?;
        Ok(Commit::from_parts(file, patch, diagnostics))
    }
}
