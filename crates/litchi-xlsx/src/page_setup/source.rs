//! Guarded page-setup publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_page_setup`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_page_setup`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{Setup, Snapshot};
use crate::error::{Result, invalid};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("page-setup", super::Patch);

impl Patch {
    /// The exact worksheet state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact worksheet state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
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

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().worksheet_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }
}

impl SourceBackedWorkbook {
    /// Begin an isolated page-setup edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector does not name a worksheet, its XML
    /// is invalid, or the source cannot be read.
    pub fn edit_page_setup<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        let before = Snapshot::load_source_backed(self.edit_package()?, selector)?;
        Ok(Edit::new(self.clone(), before))
    }

    /// Replay a page-setup patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_page_setup(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated page-setup edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    staged: Option<Setup>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot) -> Self {
        Self {
            workbook,
            staged: before.page_setup().cloned(),
            before,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-setup state.
    #[must_use]
    pub const fn page_setup(&self) -> Option<&Setup> {
        self.staged.as_ref()
    }

    /// Create or replace the complete staged relationship-free settings.
    ///
    /// An existing printer-settings relationship, if present, is retained
    /// exactly by the worksheet codec and is never exposed for mutation here.
    pub fn set(&mut self, value: Setup) -> bool {
        if self.staged.as_ref() == Some(&value) {
            return false;
        }
        self.staged = Some(value);
        true
    }

    /// Remove direct page setup.
    ///
    /// Commit retains the typed refusal when printer settings are attached.
    pub fn remove(&mut self) -> bool {
        self.staged.take().is_some()
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = self.before.page_setup().cloned();
    }

    fn changed(&self) -> bool {
        self.before.page_setup() != self.staged.as_ref()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if !self.changed() {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }
        let output =
            super::replace_worksheet_page_setup(self.before.source_xml(), self.staged.as_ref())?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        if snapshot.page_setup() != self.staged.as_ref() {
            return Err(invalid(
                "page-setup publication changed the staged semantic state",
            ));
        }
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }
}
