//! Source-backed page-break edits of a [`SourceBackedWorkbook`] (record
//! 0360, the 0355a §1 shape).
//!
//! [`SourceBackedWorkbook::edit_page_breaks`] stages one worksheet's row and
//! column breaks without materializing any unselected part body;
//! [`Edit::commit`] validates the edit and returns a [`Commit`] holding the
//! new workbook, whose output is the byte-preserving splice the former
//! stream publisher wrote: the selected worksheet XML is the only replaced
//! member and every other ZIP member is raw-copied. An exact no-op returns
//! the source workbook, which writes the complete source bytes. A changed
//! edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_page_breaks`] replays a [`Patch`] onto a
//! workbook holding its before state.

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{Collection, PageBreaks, Snapshot};
use crate::error::{Result, invalid};
use crate::{Selector, SourceBackedWorkbook};

pub use super::LimitKind;

crate::workbook::source::deferred::source_family!("page-break", super::Patch, own_limits);

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

impl From<Patch> for super::edit::Patch {
    /// Replay a source-backed page-break change onto an owning
    /// [`crate::Package`] with [`crate::Package::apply_page_breaks`].
    fn from(patch: Patch) -> Self {
        Self(std::sync::Arc::unwrap_or_clone(patch.0))
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
        self.check(package)?;
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
    /// Begin an isolated page-break edit of one worksheet without
    /// materializing any unselected part body.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector does not name a worksheet, its
    /// page-break XML is invalid, or the source cannot be read.
    pub fn edit_page_breaks<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        let before = Snapshot::load_source_backed(self.edit_package()?, selector)?;
        Ok(Edit {
            workbook: self.clone(),
            staged: before.page_breaks().clone(),
            before,
        })
    }

    /// Replay a page-break patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when the selected worksheet
    /// does not hold the patch's before state, and the signed-source refusal
    /// for a changed patch on a signed source.
    pub fn apply_page_breaks(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated page-break edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    staged: PageBreaks,
}

impl Edit {
    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-break state.
    #[must_use]
    pub const fn page_breaks(&self) -> &PageBreaks {
        &self.staged
    }

    /// Replace the complete staged page-break state.
    pub fn set(&mut self, value: PageBreaks) -> Result<bool> {
        value.validate()?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Clone-edit the complete staged value; a failed closure changes nothing.
    pub fn edit(&mut self, edit: impl FnOnce(&mut PageBreaks) -> Result<()>) -> Result<bool> {
        let mut draft = self.staged.clone();
        edit(&mut draft)?;
        self.set(draft)
    }

    /// Replace horizontal row breaks.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_horizontal(collection))
    }

    /// Replace vertical column breaks.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_vertical(collection))
    }

    /// Remove horizontal row breaks.
    pub fn remove_horizontal(&mut self) -> bool {
        self.staged.remove_horizontal()
    }

    /// Remove vertical column breaks.
    pub fn remove_vertical(&mut self) -> bool {
        self.staged.remove_vertical()
    }

    /// Remove both authored collections.
    pub fn clear(&mut self) -> bool {
        self.staged.clear()
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = self.before.page_breaks().clone();
    }

    /// Validate the edit and return the committed workbook, the exact patch
    /// and the diagnostics.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid staged state, a rewrite that changed the
    /// staged semantics, or a changed edit of a signed or encrypted source.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        if self.before.page_breaks() == &self.staged {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }
        let output = super::replace(self.before.source_xml(), &self.staged)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        if snapshot.page_breaks() != &self.staged {
            return Err(invalid(
                "page-break publication changed the staged semantic state",
            ));
        }
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }
}
