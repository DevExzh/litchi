//! Guarded page-margin publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_page_margins`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_page_margins`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{Margins, Snapshot};
use crate::error::{Result, invalid};
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("page-margin", super::Patch);

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
    /// Begin an isolated page-margin edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Source-resolution errors are returned by the first accessor, staging
    /// operation, or commit that needs the selected worksheet.
    pub fn edit_page_margins<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone(), own_selector(selector.into())?))
    }

    /// Replay a page-margin patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_page_margins(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated page-margin edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    /// `None` means that the edit still inherits the source state. `Some`
    /// contains the complete replacement, including `None` for removal.
    staged: Option<Option<Margins>>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, selector: Selector<'static>) -> Self {
        Self {
            source: LazySource::new(workbook),
            selector,
            staged: None,
        }
    }

    fn resolve_before(&self) -> Result<&Snapshot> {
        let selector = self.selector.clone();
        self.source
            .resolve(|package| Snapshot::load_source_backed(package, selector))
    }

    fn effective(&self) -> Result<Option<&Margins>> {
        match &self.staged {
            Some(staged) => Ok(staged.as_ref()),
            None => Ok(self.resolve_before()?.page_margins()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Currently staged page-margin state.
    pub fn page_margins(&self) -> Result<Option<&Margins>> {
        self.effective()
    }

    /// Create or replace the complete staged page-margin state.
    pub fn set(&mut self, value: Margins) -> Result<bool> {
        if self.effective()?.is_some_and(|current| current == &value) {
            return Ok(false);
        }
        self.staged = Some(Some(value));
        Ok(true)
    }

    /// Remove direct page margins.
    pub fn remove(&mut self) -> Result<bool> {
        if self.effective()?.is_none() {
            return Ok(false);
        }
        self.staged = Some(None);
        Ok(true)
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = None;
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        let staged = self
            .staged
            .unwrap_or_else(|| before.page_margins().copied());
        before.check_execution()?;
        if before.page_margins() == staged.as_ref() {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        let output = super::replace_page_margins(before.source_xml(), staged.as_ref())?;
        let snapshot = Snapshot::from_rewritten_source(&before, output)?;
        snapshot.check_execution()?;
        if snapshot.page_margins() != staged.as_ref() {
            return Err(invalid(
                "page-margin publication changed the staged semantic state",
            ));
        }
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
