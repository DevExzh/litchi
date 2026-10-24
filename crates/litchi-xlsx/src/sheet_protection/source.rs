//! Guarded sheet-protection publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_sheet_protection`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_sheet_protection`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{Metadata, Snapshot, replace_protection, validate_metadata};
use crate::error::{Result, invalid};
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("sheet-protection", super::Patch);

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
    /// Begin an isolated sheet-protection edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Source-resolution errors are returned by the first accessor, staging
    /// operation, or commit that needs the selected worksheet.
    pub fn edit_sheet_protection<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone(), own_selector(selector.into())?))
    }

    /// Replay a sheet-protection patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_sheet_protection(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated sheet-protection edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    /// `None` inherits the source metadata; `Some` stores a complete
    /// replacement.
    staged: Option<Metadata>,
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

    fn effective(&self) -> Result<&Metadata> {
        match self.staged.as_ref() {
            Some(staged) => Ok(staged),
            None => Ok(self.resolve_before()?.metadata()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Complete staged protection metadata.
    pub fn metadata(&self) -> Result<&Metadata> {
        self.effective()
    }

    /// Replace the complete staged protection state after validation.
    pub fn set(&mut self, value: Metadata) -> Result<bool> {
        validate_metadata(&value)?;
        if self.effective()? == &value {
            return Ok(false);
        }
        self.staged = Some(value);
        Ok(true)
    }

    /// Edit a clone atomically, retaining the previous staged state on error.
    pub fn update(&mut self, operation: impl FnOnce(&mut Metadata) -> Result<()>) -> Result<bool> {
        let mut candidate = self.effective()?.clone();
        operation(&mut candidate)?;
        self.set(candidate)
    }

    /// Remove all sheet protection and protected ranges.
    pub fn clear(&mut self) -> Result<bool> {
        if self.effective()? == &Metadata::default() {
            return Ok(false);
        }
        self.staged = Some(Metadata::default());
        Ok(true)
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = None;
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        let staged = self.staged.unwrap_or_else(|| before.metadata().clone());
        before.check_execution()?;
        validate_metadata(&staged)?;
        if before.metadata() == &staged {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        let output = replace_protection(before.source_xml(), &staged)?;
        let snapshot = Snapshot::from_rewritten_source(&before, output)?;
        if snapshot.metadata() != &staged {
            return Err(invalid(
                "sheet-protection publication changed the staged semantic state",
            ));
        }
        before.check_execution()?;
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
