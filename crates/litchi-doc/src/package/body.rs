//! Package-level body-text routes (record 0359a F1).
//!
//! The body-text engine (`body_text::{Snapshot, Edit, Commit, Patch}`) is
//! crate-private: [`Package::edit`] is its only staging route. This module
//! publishes, on [`Package`], every capability the engine had beyond the
//! primary edit's verbs — the body reads, prepared sub-edits and their
//! disjoint composition ([`JoinedEdits`]), the three-way merge of two
//! patches ([`ThreeWayPlan`]), durable-patch replay, transfer planning and an
//! undo/redo [`History`] of packages. Each commit is the primary
//! [`Commit`]: a new [`Package`] with its exact [`Patch`] and
//! [`super::Diagnostics`].

use litchi_core::Position;
use litchi_core::patch::{CompositionLimits, HistoryLimits};

use super::Package;
use super::edit::{Commit, Edit, Error, Patch, Result};
use crate::body_text::{
    self, CharacterProperty, EmbeddedTransferPlan, JoinError, Paragraph, PictureTransferPlan,
    Projection, Story, SubEdit, TextItem, TextTarget, ThreeWayConflict, TransferPlan,
};
use crate::parts::dofr::DofrArray;
use crate::parts::saved_selection::SavedSelection;

/// Bounded undo/redo history over immutable DOC packages: record each
/// [`Commit::package`] with its weight (for example its byte length).
pub type History = litchi_core::patch::History<Package>;

impl Package {
    /// The body-text engine snapshot of this package under its retained edit
    /// options.
    fn body(&self) -> Result<body_text::Snapshot> {
        self.body_snapshot(&self.retained_edit_options())
    }

    /// Main-story paragraphs under one revision `projection`.
    ///
    /// # Errors
    ///
    /// Returns an error when the body cannot be read.
    pub fn body_paragraphs(&self, projection: Projection) -> Result<Vec<Paragraph>> {
        Ok(self.body()?.paragraphs(projection)?)
    }

    /// The paragraphs of one story.
    ///
    /// # Errors
    ///
    /// Returns an error when the story cannot be read.
    pub fn story_paragraphs(&self, story: Story) -> Result<Vec<TextItem>> {
        Ok(self.body()?.story_paragraphs(story)?)
    }

    /// Every table cell's text item.
    ///
    /// # Errors
    ///
    /// Returns an error when the tables cannot be read.
    pub fn table_cells(&self) -> Result<Vec<TextItem>> {
        Ok(self.body()?.table_cells()?)
    }

    /// Every field result's text item.
    ///
    /// # Errors
    ///
    /// Returns an error when the fields cannot be read.
    pub fn field_results(&self) -> Result<Vec<TextItem>> {
        Ok(self.body()?.field_results()?)
    }

    /// The saved selection (`Selsf`), when present, bound to this package.
    ///
    /// # Errors
    ///
    /// Returns an error when the record is malformed.
    pub fn saved_selection(&self) -> Result<Option<SavedSelection>> {
        Ok(self.body()?.saved_selection()?)
    }

    /// The `RgDofr` records, when present, bound to this package.
    ///
    /// # Errors
    ///
    /// Returns an error when the records are malformed.
    pub fn dofr_records(&self) -> Result<Option<DofrArray>> {
        Ok(self.body()?.dofr_records()?)
    }

    /// Start a bounded composition of independently prepared, provably
    /// disjoint body edits.
    ///
    /// # Errors
    ///
    /// Returns an error when the body cannot be read for editing.
    pub fn compose(&self, limits: CompositionLimits) -> Result<JoinedEdits> {
        Ok(JoinedEdits {
            package: self.clone(),
            inner: self.body()?.compose(limits),
        })
    }

    /// Prepare one composable body-paragraph replacement.
    ///
    /// # Errors
    ///
    /// Returns the replacement's validation refusal or a composition bound.
    pub fn prepare_replace(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        replacement: impl Into<String>,
    ) -> Result<SubEdit> {
        Ok(self
            .body()?
            .prepare_replace(limits, identifier, position, replacement)?)
    }

    /// Prepare one composable body-paragraph bold toggle.
    ///
    /// # Errors
    ///
    /// Returns the toggle's validation refusal or a composition bound.
    pub fn prepare_bold(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        enabled: bool,
    ) -> Result<SubEdit> {
        Ok(self
            .body()?
            .prepare_bold(limits, identifier, position, enabled)?)
    }

    /// Prepare one composable text replacement of any text target.
    ///
    /// # Errors
    ///
    /// Returns the replacement's validation refusal or a composition bound.
    pub fn prepare_text(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        replacement: impl Into<String>,
    ) -> Result<SubEdit> {
        Ok(self
            .body()?
            .prepare_text(limits, identifier, target, replacement)?)
    }

    /// Prepare one composable character-property toggle of any text target.
    ///
    /// # Errors
    ///
    /// Returns the toggle's validation refusal or a composition bound.
    pub fn prepare_format(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        target: TextTarget,
        property: CharacterProperty,
        enabled: bool,
    ) -> Result<SubEdit> {
        Ok(self
            .body()?
            .prepare_format(limits, identifier, target, property, enabled)?)
    }

    /// Start a primary edit under `options`, opening its body snapshot
    /// afresh while reporting the open's deterministic, content-free phase
    /// boundaries to `observer` and its in-memory CFB parse to
    /// `cfb_observer` (the opt-in profiling equivalent of
    /// [`Package::edit_with_options`]).
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Package::edit_with_options`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn edit_profiled<'a>(
        &self,
        options: super::EditOptions,
        observer: impl FnMut(body_text::DiagnosticEvent) + 'a,
        cfb_observer: impl FnMut(body_text::CfbParseEvent) + 'a,
    ) -> Result<Edit<'a, super::edit::ProfiledOpen<'a>>> {
        let options = self.constrained_edit_options(options);
        Ok(Edit::from_profiled(
            self.clone(),
            options,
            observer,
            cfb_observer,
        ))
    }

    /// A bounded undo/redo history starting at this package.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Replay a durable body patch (`Patch::to_durable`) onto the exact
    /// artifact it was made from.
    ///
    /// # Errors
    ///
    /// Returns an error for a patch of another format or artifact, or one
    /// whose replay fails validation.
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Commit> {
        let inner = self.body()?.stage_durable(patch)?;
        Edit::from_body(self.clone(), inner).commit()
    }

    /// Plan an inert text transfer from `donor`'s `source` target into this
    /// package's `destination` target.
    ///
    /// # Errors
    ///
    /// Returns a refusal for a target with a drawing or field dependency.
    pub fn plan_text_transfer_from(
        &self,
        donor: &Self,
        source: TextTarget,
        destination: TextTarget,
    ) -> Result<TransferPlan> {
        Ok(self
            .body()?
            .plan_text_transfer_from(&donor.body()?, source, destination)?)
    }

    /// Plan a dependency-closed embedded-object transfer from `donor`.
    ///
    /// # Errors
    ///
    /// Returns a refusal for a missing source or an occupied destination.
    pub fn plan_embedded_transfer_from(
        &self,
        donor: &Self,
        source_storage_id: u32,
        destination_storage_id: u32,
    ) -> Result<EmbeddedTransferPlan> {
        Ok(self.body()?.plan_embedded_transfer_from(
            &donor.body()?,
            source_storage_id,
            destination_storage_id,
        )?)
    }

    /// Plan a picture transfer from `donor`'s `source` target into this
    /// package's `destination` target.
    ///
    /// # Errors
    ///
    /// Returns a refusal for an unsupported picture or target.
    pub fn plan_picture_transfer_from(
        &self,
        donor: &Self,
        source: TextTarget,
        destination: TextTarget,
    ) -> Result<PictureTransferPlan> {
        Ok(self
            .body()?
            .plan_picture_transfer_from(&donor.body()?, source, destination)?)
    }

    /// Plan the three-way merge of two patches made from this package.
    ///
    /// # Errors
    ///
    /// Returns [`body_text::Error::PatchConflict`] when either patch was made
    /// from another package, and a refusal for auxiliary-table changes.
    pub fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
        Ok(ThreeWayPlan {
            package: self.clone(),
            inner: self.body()?.plan_three_way(&left.inner, &right.inner)?,
        })
    }
}

/// A bounded composition of independently prepared, provably disjoint body
/// edits over one package. Dropping it discards it.
pub struct JoinedEdits {
    package: Package,
    inner: body_text::Composition,
}

impl std::fmt::Debug for JoinedEdits {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("JoinedEdits")
            .field("edits", &self.inner.len())
            .finish_non_exhaustive()
    }
}

impl JoinedEdits {
    /// Join one prepared edit when its lineage and semantic facet are disjoint
    /// from every accepted edit.
    ///
    /// # Errors
    ///
    /// Returns the structured refusal with the rejected edit.
    #[expect(
        clippy::result_large_err,
        reason = "composition refusals return the rejected prepared edit to the caller"
    )]
    pub fn join(&mut self, incoming: SubEdit) -> std::result::Result<&mut Self, JoinError> {
        self.inner.join(incoming)?;
        Ok(self)
    }

    /// Number of accepted prepared edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    /// Whether no edit has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Commit every accepted edit atomically in stable identifier order.
    ///
    /// # Errors
    ///
    /// Returns the staging, publication and reopen errors of [`Edit::commit`].
    pub fn commit(self) -> Result<Commit> {
        let inner = self.inner.into_edit()?;
        Edit::from_body(self.package, inner).commit()
    }
}

/// The three-way merge plan of two patches made from one package.
pub struct ThreeWayPlan {
    package: Package,
    inner: body_text::ThreeWayPlan,
}

impl std::fmt::Debug for ThreeWayPlan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ThreeWayPlan")
            .field("conflicts", &self.inner.conflicts().len())
            .finish_non_exhaustive()
    }
}

impl ThreeWayPlan {
    /// The semantic effects both patches change differently.
    #[must_use]
    pub fn conflicts(&self) -> &[ThreeWayConflict] {
        self.inner.conflicts()
    }

    /// Whether the plan can be committed.
    #[must_use]
    pub fn is_conflict_free(&self) -> bool {
        self.inner.is_conflict_free()
    }

    /// Publish the conflict-free combined plan as a new package.
    ///
    /// # Errors
    ///
    /// Returns [`body_text::Error::Conflict`] for a three-way plan whose
    /// branches write the same facet differently, and the staging, publication
    /// and reopen errors of [`Edit::commit`].
    pub fn commit(self) -> Result<Commit> {
        let inner = self.inner.into_edit().map_err(Error::from)?;
        Edit::from_body(self.package, inner).commit()
    }
}
