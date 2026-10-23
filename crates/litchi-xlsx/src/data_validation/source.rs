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
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!("data-validation", super::Patch);

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
    /// Begin an isolated data-validation edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Source-resolution errors are returned by the first accessor, staging
    /// operation, or commit that needs the selected worksheet.
    pub fn edit_data_validations<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.check_execution()?;
        Ok(Edit::new(self.clone(), own_selector(selector.into())?))
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
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    /// `None` inherits the source collection; `Some` stores a complete
    /// replacement, including an empty collection.
    staged: Option<Vec<Collection>>,
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

    fn effective(&self) -> Result<&[Collection]> {
        match self.staged.as_deref() {
            Some(staged) => Ok(staged),
            None => Ok(self.resolve_before()?.collections()),
        }
    }

    /// Resolve the exact source state observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Complete staged data-validation collections.
    pub fn collections(&self) -> Result<&[Collection]> {
        self.effective()
    }

    /// Replace all staged data-validation collections after validation.
    pub fn set_collections(&mut self, value: Vec<Collection>) -> Result<bool> {
        validate_data_validation_collections(&value)?;
        if self.effective()? == value.as_slice() {
            return Ok(false);
        }
        self.staged = Some(value);
        Ok(true)
    }

    /// Edit a clone atomically, retaining the previous staged state on error.
    pub fn update(
        &mut self,
        operation: impl FnOnce(&mut Vec<Collection>) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.effective()?.to_vec();
        operation(&mut candidate)?;
        self.set_collections(candidate)
    }

    /// Remove every data-validation collection from the selected worksheet.
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

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        let staged = self.staged.unwrap_or_else(|| before.collections().to_vec());
        before.check_execution()?;
        validate_data_validation_collections(&staged)?;
        if before.collections() == staged.as_slice() {
            let patch = super::Patch::new(before.clone(), before);
            return commit_change(self.source.workbook(), patch);
        }
        let (output, readback) =
            replace_data_validation_collections_with_readback(before.source_xml(), &staged)?;
        if readback != staged {
            return Err(invalid(
                "data-validation publication changed the staged semantic state",
            ));
        }
        let snapshot = Snapshot::from_rewritten_source(&before, output, readback);
        before.check_execution()?;
        commit_change(self.source.workbook(), super::Patch::new(before, snapshot))
    }
}
