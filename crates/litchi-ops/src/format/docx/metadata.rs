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
use super::{Docx, DocxEditor, DocxOwner};
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

    /// Commit the admitted core-property assignments of `props`.
    ///
    /// `editor` is the handle's editor, built first so the signed,
    /// unsigned-copy and retained-source policies of every DOCX write apply.
    pub(super) fn set_metadata(&self, editor: DocxEditor, props: &Props<'_>) -> Result<Commit> {
        let assignments = metadata::admit_with(props, FileFormat::Docx, Scope::Core)?;
        let package = match (editor, self.owner()) {
            (DocxEditor::Unsigned(package), _) => *package,
            (_, DocxOwner::Eager(package)) => package.clone(),
            (_, DocxOwner::Source(package)) => self.promoted(package)?,
        };
        let mut edit = package.edit_metadata();
        edit.update(|value| assignments.apply(value));
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        self.owner().check_source()?;
        // The transaction history keeps a main-document patch per step; a
        // metadata step's is empty, and the coordinator compares the core
        // properties of the endpoints to see the change (record 0338).
        let file = if changed {
            self.adopt_eager_package(commit.into_package())
        } else {
            // An equal value publishes nothing (record 0253's no-op rule).
            self.clone()
        };
        WorkEdit::from(file.edit()?)?.commit(&file)
    }
}
