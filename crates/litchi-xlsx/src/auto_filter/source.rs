//! Guarded auto-filter publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_auto_filter`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_auto_filter`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::package::validate_definition;
use super::{Definition, Snapshot};
use crate::error::{Result, invalid};
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("auto-filter", super::Patch);

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
    /// Begin an isolated auto-filter edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Source-resolution errors are returned by the first accessor, staging
    /// operation, or commit that needs the selected worksheet.
    pub fn edit_auto_filter<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone(), own_selector(selector.into())?))
    }

    /// Replay a auto-filter patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_auto_filter(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated auto-filter edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    /// `None` inherits the source value; `Some` stores the complete
    /// replacement, including `None` for a cleared filter.
    staged: Option<Option<Definition>>,
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

    fn effective(&self) -> Result<Option<&Definition>> {
        match &self.staged {
            Some(staged) => Ok(staged.as_ref()),
            None => Ok(self.resolve_before()?.auto_filter()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Complete currently staged filter and sort state.
    pub fn auto_filter(&self) -> Result<Option<&Definition>> {
        self.effective()
    }

    /// Create or replace the complete staged filter and sort state.
    pub fn set(&mut self, value: Definition) -> Result<bool> {
        validate_definition(&value)?;
        if self.effective()?.is_some_and(|current| current == &value) {
            return Ok(false);
        }
        self.staged = Some(Some(value));
        Ok(true)
    }

    /// Edit a clone atomically, retaining the staged state on error.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Option<Definition>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.effective()?.cloned();
        operation(&mut candidate)?;
        if let Some(value) = candidate.as_ref() {
            validate_definition(value)?;
        }
        if self.effective()? == candidate.as_ref() {
            return Ok(false);
        }
        self.staged = Some(candidate);
        Ok(true)
    }

    /// Remove the direct worksheet auto-filter and its sort state.
    pub fn clear(&mut self) -> Result<bool> {
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
        let staged = self.staged.unwrap_or_else(|| before.auto_filter().cloned());
        before.check_execution()?;
        if before.auto_filter() == staged.as_ref() {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        if before.mutation_locked(staged.as_ref()) {
            return Err(invalid(
                "worksheet protection forbids the staged auto-filter or sort mutation",
            ));
        }
        if let Some(value) = staged.as_ref() {
            validate_definition(value)?;
        }
        let output = super::replace_auto_filter(before.source_xml(), staged.as_ref())?;
        let readback = super::parse_auto_filter(&output)?;
        if readback != staged {
            return Err(invalid(
                "auto-filter publication changed the staged semantic state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&before, output, readback)?;
        before.check_execution()?;
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
