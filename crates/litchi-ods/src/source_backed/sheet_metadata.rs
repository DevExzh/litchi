//! Worksheet metadata edits of a [`SourceBackedSpreadsheet`] (record
//! 0358a): consolidation, label ranges, cell-range sources and detective
//! metadata, published by replacing only their spans in `content.xml`.
//!
//! [`SourceBackedSpreadsheet::edit_sheet_metadata`] (or
//! [`SourceBackedSpreadsheet::edit_sheet_metadata_with`] under explicit
//! limits and an execution context) starts an [`Edit`]; [`Edit::commit`]
//! returns a [`Commit`] holding the new spreadsheet, the reversible
//! [`Patch`] and the [`Diagnostics`];
//! [`SourceBackedSpreadsheet::apply_sheet_metadata`] replays a patch.
//! [`SourceBackedSpreadsheet::sheet_metadata`] reads the metadata as a
//! [`Snapshot`]. A changed edit of a signed source is refused at commit.

use std::fmt;

use litchi_core::{ExecutionContext, SourceVersion};

use crate::SourceBackedSpreadsheet;
use crate::sheet_metadata::{
    CellMetadataView, CellRange, CellSelector, Detective, Edit as Model, LabelRange,
    LabelRangesView, Limits, Options, Snapshot as Inner,
};

pub use litchi_core::edit::Diagnostics;
pub use litchi_odf_common::{Error, LimitKind};

super::content_patch!("a sheet metadata edit");

/// The result of a sheet metadata edit's `commit`: the new source-backed
/// spreadsheet, the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<SourceBackedSpreadsheet, Patch, Diagnostics>;

type Result<T> = std::result::Result<T, Error>;

/// The sheet metadata of one [`SourceBackedSpreadsheet`], read once.
pub struct Snapshot<'source> {
    owner: &'source SourceBackedSpreadsheet,
    inner: Inner,
    source_version: SourceVersion,
}

impl fmt::Debug for Snapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_version", &self.source_version)
            .field("source_bytes", &self.inner.source_xml().len())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for Snapshot<'source> {
    fn clone(&self) -> Self {
        Self {
            owner: self.owner,
            inner: self.inner.clone(),
            source_version: self.source_version,
        }
    }
}

impl<'source> Snapshot<'source> {
    pub(crate) fn from_owner(
        owner: &'source SourceBackedSpreadsheet,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        owner.check_source()?;
        let source_version = owner.source_version()?;
        let source = owner.content_xml_arc()?;
        let inner =
            Inner::parse_arc_with_context(source, limits, context, enforce_context_lineage)?;
        owner.check_source()?;
        Ok(Self {
            owner,
            inner,
            source_version,
        })
    }
    /// Exact source version captured by this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }
    /// Borrow exact content.xml bytes.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        self.inner.source_xml()
    }
    /// Read consolidation.
    pub fn consolidation(&self) -> Result<Option<&Options>> {
        self.check_source()?;
        Ok(self.inner.consolidation()?)
    }
    /// Read labels.
    pub fn label_ranges(&self) -> Result<LabelRangesView<'_>> {
        self.check_source()?;
        Ok(self.inner.label_ranges())
    }
    /// Read cell metadata.
    pub fn cell_metadata(
        &self,
        selector: CellSelector<'_>,
    ) -> Result<Option<CellMetadataView<'_>>> {
        self.check_source()?;
        Ok(self.inner.cell_metadata(selector)?)
    }
    /// An isolated edit of this metadata, committed under `context`.
    fn edit(&self, context: ExecutionContext) -> Result<Edit<'source>> {
        self.check_source()?;
        Ok(Edit {
            before: self.clone(),
            inner: self.inner.edit(),
            context,
        })
    }
    fn check_source(&self) -> Result<()> {
        let observed = self.owner.source_version()?;
        if observed == self.source_version {
            Ok(())
        } else {
            Err(litchi_core::Error::SourceChanged {
                expected: self.source_version,
                observed,
            }
            .into())
        }
    }
    fn is_signed(&self) -> Result<bool> {
        Ok(self.owner.sheet_metadata_signed()?)
    }
}

/// A staged sheet metadata edit of one [`SourceBackedSpreadsheet`]. Dropping
/// it discards it.
pub struct Edit<'source> {
    before: Snapshot<'source>,
    inner: Model,
    context: ExecutionContext,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source_version", &self.before.source_version)
            .finish_non_exhaustive()
    }
}

impl<'source> Edit<'source> {
    /// Base source snapshot.
    #[must_use]
    pub const fn before(&self) -> &Snapshot<'source> {
        &self.before
    }
    /// Stage consolidation replacement/removal.
    pub fn set_consolidation(&mut self, value: Option<Options>) -> Result<()> {
        Ok(self.inner.set_consolidation(value)?)
    }
    /// Clear consolidation.
    pub fn clear_consolidation(&mut self) -> Result<()> {
        Ok(self.inner.clear_consolidation()?)
    }
    /// Stage label insertion.
    pub fn add_label_range(&mut self, value: LabelRange) -> Result<()> {
        Ok(self.inner.add_label_range(value)?)
    }
    /// Stage label insertion at a source-order index.
    pub fn insert_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        Ok(self.inner.insert_label_range(index, value)?)
    }
    /// Stage label replacement.
    pub fn replace_label_range(&mut self, index: usize, value: LabelRange) -> Result<()> {
        Ok(self.inner.replace_label_range(index, value)?)
    }
    /// Stage label removal.
    pub fn remove_label_range(&mut self, index: usize) -> Result<LabelRange> {
        Ok(self.inner.remove_label_range(index)?)
    }
    /// Clear labels and their container.
    pub fn clear_label_ranges(&mut self) -> Result<()> {
        Ok(self.inner.clear_label_ranges()?)
    }
    /// Ensure labels container presence.
    pub fn ensure_label_ranges(&mut self) -> Result<()> {
        Ok(self.inner.ensure_label_ranges()?)
    }
    /// Set an existing cell source owner.
    pub fn set_cell_range_source(
        &mut self,
        selector: CellSelector<'_>,
        value: CellRange,
    ) -> Result<()> {
        Ok(self.inner.set_cell_range_source(selector, value)?)
    }
    /// Clear an existing cell source owner.
    pub fn clear_cell_range_source(&mut self, selector: CellSelector<'_>) -> Result<()> {
        Ok(self.inner.clear_cell_range_source(selector)?)
    }
    /// Set an existing cell detective owner.
    pub fn set_detective(&mut self, selector: CellSelector<'_>, value: Detective) -> Result<()> {
        Ok(self.inner.set_detective(selector, value)?)
    }
    /// Clear an existing cell detective owner.
    pub fn clear_detective(&mut self, selector: CellSelector<'_>) -> Result<()> {
        Ok(self.inner.clear_detective(selector)?)
    }
    /// Edit an existing cell detective owner.
    pub fn edit_detective<F>(&mut self, selector: CellSelector<'_>, update: F) -> Result<()>
    where
        F: FnOnce(&mut Detective) -> litchi_core::Result<()>,
    {
        Ok(self.inner.edit_detective(selector, update)?)
    }
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.inner = self.before.inner.edit();
    }

    /// Validate the staged changes under the edit's execution context and
    /// commit them to a new spreadsheet. Nothing is written.
    ///
    /// # Errors
    ///
    /// Refuses a changed edit of a signed source
    /// ([`RefusalKind::Protected`](litchi_core::edit::RefusalKind::Protected));
    /// returns the validation, limit or execution failure, or a source
    /// change.
    pub fn commit(mut self) -> Result<Commit> {
        self.before.check_source()?;
        let commit = self.inner.commit(&self.context)?;
        if commit.changed() && self.before.is_signed()? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "ods.source_backed.sheet_metadata",
                "signed-source refusal: changed ODS sheet metadata requires explicit unsign/resign policy",
            ));
        }
        self.before.check_source()?;
        let owner = self.before.owner;
        let after = std::sync::Arc::clone(commit.snapshot().source_arc());
        let change = super::Change::new(owner, std::sync::Arc::clone(&after));
        let changed = !change.is_noop();
        let package = if changed {
            owner.with_content(after)
        } else {
            owner.share()
        };
        Ok(Commit::new(
            package,
            Patch { change },
            Diagnostics::new(changed),
        ))
    }
}

impl SourceBackedSpreadsheet {
    /// Read the sheet metadata under finite default limits.
    ///
    /// # Errors
    ///
    /// Returns the metadata's parse, limit or execution failure, or a source
    /// change.
    pub fn sheet_metadata(&self) -> Result<Snapshot<'_>> {
        Snapshot::from_owner(
            self.effective()?,
            Limits::default(),
            &crate::sheet_metadata::default_context(),
            false,
        )
    }

    /// Read the sheet metadata under explicit limits and execution context.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_metadata`].
    pub fn sheet_metadata_with(
        &self,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Snapshot<'_>> {
        Snapshot::from_owner(self.effective()?, limits, context, true)
    }

    /// Start a sheet metadata edit under finite default limits.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_metadata`].
    pub fn edit_sheet_metadata(&self) -> Result<Edit<'_>> {
        self.sheet_metadata()?
            .edit(crate::sheet_metadata::default_context())
    }

    /// Start a sheet metadata edit under explicit limits; the metadata is
    /// read, and the edit committed, under `context`.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_metadata`].
    pub fn edit_sheet_metadata_with(
        &self,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Edit<'_>> {
        self.sheet_metadata_with(limits, context)?
            .edit(context.clone())
    }

    /// [`Self::apply_sheet_metadata`] admitting the patch's target `content.xml` under
    /// explicit destination limits and execution context: its output
    /// ceiling and a full parse under that policy.
    ///
    /// # Errors
    ///
    /// As [`Self::apply_sheet_metadata`], and the destination policy's limit or execution
    /// failure.
    pub fn apply_sheet_metadata_with(
        &self,
        patch: &Patch,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Commit> {
        if !patch.is_noop() {
            crate::sheet_metadata::admit_target(patch.change.after_arc(), limits, context)?;
        }
        self.apply_sheet_metadata(patch)
    }

    /// Apply a sheet metadata patch to a new spreadsheet; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this spreadsheet is over another
    /// source or does not hold the patch's before `content.xml`, and refuses
    /// a changed patch of a signed source.
    pub fn apply_sheet_metadata(&self, patch: &Patch) -> Result<Commit> {
        if !patch.is_noop() && self.sheet_metadata_signed()? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "ods.source_backed.sheet_metadata",
                "signed-source refusal: changed ODS sheet metadata patch cannot be applied",
            ));
        }
        let package = patch
            .change
            .publish(self, "ods.source_backed.sheet_metadata.apply")?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}
