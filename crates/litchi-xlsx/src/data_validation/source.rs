//! Guarded data-validation publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_data_validations`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_data_validations`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{
    Collection, Snapshot, replace_data_validation_collections_with_readback,
    validate_data_validation_collections,
};
use crate::error::{Result, invalid};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("data-validation", super::Patch);

impl Patch {
    /// The exact worksheet state the patch requires.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact worksheet state the patch produces.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        self.0.after()
    }
}

impl crate::workbook::source::deferred::Change for super::Patch {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        super::Patch::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        // A replay compares content: the selected worksheet, its owner and
        // relationships must hold the before state (record 0360).
        let current = Snapshot::load_source_backed(package, self.before().sheet_position())?;
        if !current.same_source(self.before()) {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        Ok(())
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        package.check_execution()?;
        crate::workbook::source::deferred::Change::check(self, package)?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().worksheet_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }
}

impl SourceBackedWorkbook {
    /// Begin an isolated data-validation edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector does not name a worksheet, its XML
    /// is invalid, or the source cannot be read.
    pub fn edit_data_validations<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        let before = Snapshot::load_source_backed(self.edit_package()?, selector)?;
        Ok(Edit::new(self.clone(), before))
    }

    /// Replay a data-validation patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_data_validations(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated data-validation edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    staged: Vec<Collection>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot) -> Self {
        Self {
            workbook,
            staged: before.collections().to_vec(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Complete staged data-validation collections.
    #[must_use]
    pub fn collections(&self) -> &[Collection] {
        &self.staged
    }

    /// Replace all staged data-validation collections after validation.
    pub fn set_collections(&mut self, value: Vec<Collection>) -> Result<bool> {
        validate_data_validation_collections(&value)?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Edit a clone atomically, retaining the previous staged state on error.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Vec<Collection>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.staged.clone();
        operation(&mut candidate)?;
        self.set_collections(candidate)
    }

    /// Remove every data-validation collection from the selected worksheet.
    pub fn clear(&mut self) -> bool {
        if self.staged.is_empty() {
            return false;
        }
        self.staged.clear();
        true
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = self.before.collections().to_vec();
    }

    fn changed(&self) -> bool {
        self.before.collections() != self.staged
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        validate_data_validation_collections(&self.staged)?;
        if !self.changed() {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }
        let (output, readback) = replace_data_validation_collections_with_readback(
            self.before.source_xml(),
            &self.staged,
        )?;
        if readback != self.staged {
            return Err(invalid(
                "data-validation publication changed the staged semantic state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&self.before, output, readback);
        self.before.check_execution()?;
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }
}
