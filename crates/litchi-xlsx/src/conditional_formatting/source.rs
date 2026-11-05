//! Guarded conditional-formatting publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_conditional_formats`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_conditional_formats`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::{Formatting, Snapshot};
use crate::error::{Result, invalid};
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("conditional-formatting", super::Patch);

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
    /// Begin an isolated conditional-formatting edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Source-resolution errors are returned by the first accessor, staging
    /// operation, or commit that needs the selected worksheet.
    pub fn edit_conditional_formats<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone(), own_selector(selector.into())?))
    }

    /// Replay a conditional-formatting patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_conditional_formats(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated conditional-formatting edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    /// `None` inherits the source collection; `Some` stores a complete
    /// replacement, including an empty collection.
    staged: Option<Vec<Formatting>>,
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

    fn effective(&self) -> Result<&[Formatting]> {
        match self.staged.as_deref() {
            Some(staged) => Ok(staged),
            None => Ok(self.resolve_before()?.collections()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Complete ordered collection currently staged for publication.
    pub fn collections(&self) -> Result<&[Formatting]> {
        self.effective()
    }

    /// Replace the complete ordered owner collection atomically.
    pub fn set_collections(&mut self, value: Vec<Formatting>) -> Result<bool> {
        super::package::validate_authored(
            &value,
            self.resolve_before()?.differential_format_count(),
        )?;
        if self.effective()? == value.as_slice() {
            return Ok(false);
        }
        self.staged = Some(value);
        Ok(true)
    }

    /// Mutate a clone and publish it to the edit only after complete validation.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Vec<Formatting>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.effective()?.to_vec();
        operation(&mut candidate)?;
        self.set_collections(candidate)
    }

    /// Remove all direct core conditional-formatting owners.
    pub fn clear(&mut self) -> Result<bool> {
        if self.effective()?.is_empty() {
            return Ok(false);
        }
        self.staged = Some(Vec::new());
        Ok(true)
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged = None;
    }

    /// Validate, rewrite, reopen, and freeze the isolated edit.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        let staged = self.staged.unwrap_or_else(|| before.collections().to_vec());
        before.check_execution()?;
        if before.collections() == staged.as_slice() {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        if before.mutation_locked() {
            return Err(invalid(
                "worksheet protection forbids conditional-formatting mutation",
            ));
        }
        let dxf_count = before.differential_format_count();
        super::package::validate_authored(&staged, dxf_count)?;
        let output =
            super::replace_conditional_formattings(before.source_xml(), &staged, dxf_count)?;
        let readback = super::package::parse_editable_conditional_formattings(&output, dxf_count)?;
        if readback != staged {
            return Err(invalid(
                "conditional-formatting publication changed the staged state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&before, output, readback)?;
        before.check_execution()?;
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
