//! Package-bound main-document transaction orchestration.

use super::{
    CompactionPolicy, Diagnostics, HyperlinkTextReplacement, ParagraphTextReplacement,
    ParagraphTransfer, Patch, Result, RevisionAction, RevisionKind, RevisionSelector, RunAnchor,
    Snapshot, TableCellAddress, transaction,
};
use crate::Package;
use litchi_core::Position;
use litchi_core::patch::CompositionLimits;
use std::sync::{Mutex, TryLockError};

mod chart;
mod comments;
mod fields;
mod graph;
mod notes;
mod numbering;
mod picture;
mod smartart;
mod stories;
mod story_edit;
mod styles;
mod substitute;
mod threads;
pub use chart::{NewChart, chart_workbook};
pub use comments::{CommentChange, NewComment};
pub use graph::HyperlinkTarget;
pub use notes::{NewNote, NoteKind};
pub use picture::{NewPicture, PictureChange};
pub use smartart::{NewSmartArt, SmartArtParts};
pub use stories::NewStory;
pub use story_edit::StoryTarget;
pub(crate) use substitute::story_parts_of;

/// An isolated main-document edit retaining its complete source package.
///
/// `P` defaults to the eager [`crate::Package`]; source-backed owners use
/// [`crate::source_backed::Package`].
///
/// Dropping an edit rolls it back. Committing validates and returns a new
/// package; the source is unchanged on success and failure.
pub struct Edit<P = Package> {
    package: P,
    /// The native transaction is populated on the first operation (or
    /// projection/output query).  Keeping the source materializer here makes
    /// `Package::edit*` O(1) without introducing a fake empty snapshot.
    inner: Mutex<Option<transaction::Edit>>,
    materializer: fn(&P, super::EditOptions) -> Result<transaction::Edit>,
    options: super::EditOptions,
    compaction: CompactionPolicy,
}

fn deferred_error(operation: &'static str) -> transaction::Error {
    transaction::Error::Document(crate::Error::internal(
        operation,
        "document edit materializer is unavailable",
    ))
}

fn unavailable_materializer<P>(_: &P, _: super::EditOptions) -> Result<transaction::Edit> {
    Err(deferred_error("document.edit.materialize"))
}

impl<P> std::fmt::Debug for Edit<P> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl<P: Clone> Edit<P> {
    pub(crate) fn new(package: P, inner: transaction::Edit) -> Self {
        let options = match inner.snapshot_budget() {
            Some(maximum) => super::EditOptions::new().max_snapshot_bytes(maximum),
            None => super::EditOptions::new(),
        };
        let compaction = inner.compaction_policy();
        Self {
            package,
            inner: Mutex::new(Some(inner)),
            materializer: unavailable_materializer,
            options,
            compaction,
        }
    }

    pub(crate) fn deferred(
        package: P,
        options: super::EditOptions,
        materializer: fn(&P, super::EditOptions) -> Result<transaction::Edit>,
    ) -> Self {
        Self {
            package,
            inner: Mutex::new(None),
            materializer,
            options,
            compaction: CompactionPolicy::default(),
        }
    }

    fn build_inner(&self) -> Result<transaction::Edit> {
        let inner = (self.materializer)(&self.package, self.options)?;
        Ok(inner
            .with_options(self.options)
            .with_compaction_policy(self.compaction))
    }

    fn inner_slot_mut(&mut self) -> &mut Option<transaction::Edit> {
        match self.inner.get_mut() {
            Ok(slot) => slot,
            Err(poisoned) => poisoned.into_inner(),
        }
    }

    pub(crate) fn inner_mut(&mut self) -> Result<&mut transaction::Edit> {
        if self.inner_slot_mut().is_none() {
            let inner = self.build_inner()?;
            *self.inner_slot_mut() = Some(inner);
        }
        self.inner_slot_mut()
            .as_mut()
            .ok_or_else(|| deferred_error("document.edit"))
    }

    fn into_inner(mut self) -> Result<transaction::Edit> {
        if self.inner_slot_mut().is_none() {
            let inner = self.build_inner()?;
            *self.inner_slot_mut() = Some(inner);
        }
        let state = match self.inner.into_inner() {
            Ok(state) => state,
            Err(poisoned) => poisoned.into_inner(),
        };
        state.ok_or_else(|| deferred_error("document.edit"))
    }

    fn inner_ref<R>(&self, read: impl FnOnce(&transaction::Edit) -> R) -> Result<R> {
        let mut state = match self.inner.try_lock() {
            Ok(state) => state,
            Err(TryLockError::Poisoned(poisoned)) => poisoned.into_inner(),
            Err(TryLockError::WouldBlock) => return Err(deferred_error("document.edit.borrow")),
        };
        if state.is_none() {
            *state = Some(self.build_inner()?);
        }
        let inner = state
            .as_ref()
            .ok_or_else(|| deferred_error("document.edit"))?;
        Ok(read(inner))
    }

    pub(crate) fn is_strict(&self) -> Result<bool> {
        self.inner_ref(transaction::Edit::is_strict)
    }

    /// Configure the retained main-document XML-pair policy.
    ///
    /// The package owner remains unchanged. The bound is checked before each
    /// native projection is retained.
    pub fn with_options(mut self, options: super::EditOptions) -> Self {
        self.options = options;
        if let Some(inner) = self.inner_slot_mut().take() {
            *self.inner_slot_mut() = Some(inner.with_options(options));
        }
        self
    }

    /// Return the retained main-document XML-pair bound, if any.
    #[must_use]
    pub const fn max_snapshot_bytes_value(&self) -> Option<usize> {
        self.options.max_snapshot_bytes_value()
    }

    /// Apply one already-planned composition batch without opening a second
    /// document snapshot. Package owners use this to retain one validation and
    /// publication pass for all joined operations.
    pub(crate) fn apply_composition_operations(
        &mut self,
        operations: &[transaction::Operation],
    ) -> Result<&mut Self> {
        for operation in operations {
            self.inner_mut()?.apply_operation(operation)?;
        }
        Ok(self)
    }

    /// Change one direct run's local properties while preserving unrelated markup.
    ///
    /// See [`super::raw::Tree::format_run`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn format_run(
        &mut self,
        paragraph: Position,
        run: Position,
        change: impl FnOnce(&mut crate::properties::Run),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_run(paragraph, run, change)?;
        Ok(self)
    }

    /// Change one direct-body paragraph's local properties.
    ///
    /// See [`super::raw::Tree::format_paragraph`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn format_paragraph(
        &mut self,
        paragraph: Position,
        change: impl FnOnce(&mut crate::properties::Paragraph),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_paragraph(paragraph, change)?;
        Ok(self)
    }

    /// Insert a paragraph with typed direct properties at a body anchor.
    ///
    /// See [`super::raw::Tree::insert_formatted_paragraph`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn insert_formatted_paragraph(
        &mut self,
        at: super::BodyAnchor,
        paragraph: &super::NewParagraph,
    ) -> Result<Position> {
        self.inner_mut()?.insert_formatted_paragraph(at, paragraph)
    }

    /// Remove a paragraph after checking ranges, section breaks and story references.
    ///
    /// See [`super::raw::Tree::remove_paragraph`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn remove_paragraph(&mut self, position: Position) -> Result<&mut Self> {
        self.inner_mut()?.remove_paragraph(position)?;
        Ok(self)
    }

    /// Move a paragraph atomically, preserving its exact XML.
    ///
    /// See [`super::raw::Tree::move_paragraph`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn move_paragraph(
        &mut self,
        position: Position,
        to: super::BodyAnchor,
    ) -> Result<Position> {
        self.inner_mut()?.move_paragraph(position, to)
    }

    /// Swap two direct-body paragraphs while retaining their exact XML.
    ///
    /// See [`super::raw::Tree::swap_paragraphs`] for selection and
    /// preservation rules.
    pub fn swap_paragraphs(&mut self, first: Position, second: Position) -> Result<&mut Self> {
        self.inner_mut()?.swap_paragraphs(first, second)?;
        Ok(self)
    }

    /// Insert a typed run before a direct run position.
    ///
    /// See [`super::raw::Tree::insert_run`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn insert_run(
        &mut self,
        paragraph: Position,
        run: Position,
        new_run: &super::NewRun,
    ) -> Result<Position> {
        self.inner_mut()?.insert_run(paragraph, run, new_run)
    }

    /// Append a typed run after a paragraph's last child.
    ///
    /// See [`super::raw::Tree::append_run`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn append_run(&mut self, paragraph: Position, new_run: &super::NewRun) -> Result<Position> {
        self.inner_mut()?.append_run(paragraph, new_run)
    }

    /// Remove one direct run after checking field and story dependencies.
    ///
    /// See [`super::raw::Tree::remove_run`] for selection and preservation rules.
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document. Managed snapshots currently refuse this operation.
    pub fn remove_run(&mut self, paragraph: Position, run: Position) -> Result<&mut Self> {
        self.inner_mut()?.remove_run(paragraph, run)?;
        Ok(self)
    }

    /// Move one direct run within a paragraph while retaining its exact XML.
    ///
    /// See [`super::raw::Tree::move_run`] for selection and preservation
    /// rules.
    pub fn move_run(
        &mut self,
        paragraph: Position,
        run: Position,
        to: RunAnchor,
    ) -> Result<Position> {
        self.inner_mut()?.move_run(paragraph, run, to)
    }

    /// Move one direct run to another direct-body paragraph when the source
    /// and destination namespace scopes are proven equivalent.
    ///
    /// See [`super::raw::Tree::move_run_to`] for selection and
    /// preservation rules.
    pub fn move_run_to(
        &mut self,
        from_paragraph: Position,
        run: Position,
        to_paragraph: Position,
        to: RunAnchor,
    ) -> Result<Position> {
        self.inner_mut()?
            .move_run_to(from_paragraph, run, to_paragraph, to)
    }

    /// Swap two direct runs in one paragraph while retaining their exact XML.
    ///
    /// See [`super::raw::Tree::swap_runs`] for selection and preservation
    /// rules.
    pub fn swap_runs(
        &mut self,
        paragraph: Position,
        first: Position,
        second: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?.swap_runs(paragraph, first, second)?;
        Ok(self)
    }

    /// Swap direct runs in two direct-body paragraphs while retaining their
    /// exact XML fragments.
    pub fn swap_runs_to(
        &mut self,
        first_paragraph: Position,
        first: Position,
        second_paragraph: Position,
        second: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .swap_runs_to(first_paragraph, first, second_paragraph, second)?;
        Ok(self)
    }

    /// Return the source main-document projection.
    pub fn source(&self) -> Result<Snapshot> {
        self.inner_ref(|inner| inner.source().clone())
    }

    /// Return the staged main-document projection.
    pub fn projected(&self) -> Result<Snapshot> {
        self.inner_ref(|inner| inner.projected().clone())
    }

    /// Convert this staged edit into an immutable sub-edit for a package
    /// composition. The package owner remains responsible for publishing the
    /// resulting [`super::composition::Composition`] in one commit.
    ///
    /// # Errors
    /// Returns a common composition bound or identifier error.
    pub fn prepare(
        self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
    ) -> Result<super::composition::Prepared> {
        self.into_inner()?.prepare(limits, identifier)
    }

    /// Change the direct properties of the paragraph at `path`.
    ///
    /// See [`super::raw::Tree::format_paragraph_at`].
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document.
    pub fn format_paragraph_at(
        &mut self,
        path: &super::ElementPath,
        change: impl FnOnce(&mut crate::properties::Paragraph),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_paragraph_at(path, change)?;
        Ok(self)
    }

    /// Change the direct properties of the table at `path` (record 0342).
    ///
    /// See [`super::raw::Tree::format_table_at`].
    ///
    /// # Errors
    /// Returns a typed selection, property or resource refusal without
    /// changing the staged document.
    pub fn format_table_at(
        &mut self,
        path: &super::ElementPath,
        change: impl FnOnce(&mut crate::properties::Table),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_table_at(path, change)?;
        Ok(self)
    }

    /// Change the direct properties of the table row at `path` (record
    /// 0342).
    ///
    /// See [`super::raw::Tree::format_row_at`].
    ///
    /// # Errors
    /// Returns a typed selection, property or resource refusal without
    /// changing the staged document.
    pub fn format_row_at(
        &mut self,
        path: &super::ElementPath,
        change: impl FnOnce(&mut crate::properties::Row),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_row_at(path, change)?;
        Ok(self)
    }

    /// Change the direct properties of the run at `path`.
    ///
    /// See [`super::raw::Tree::format_run_at`].
    ///
    /// # Errors
    /// Returns a typed selection, property, dependency or resource refusal without
    /// changing the staged document.
    pub fn format_run_at(
        &mut self,
        path: &super::ElementPath,
        change: impl FnOnce(&mut crate::properties::Run),
    ) -> Result<&mut Self> {
        self.inner_mut()?.format_run_at(path, change)?;
        Ok(self)
    }

    /// Replace the text of the paragraph, run or hyperlink at `path`.
    ///
    /// See [`super::raw::Tree::replace_text_at`].
    ///
    /// # Errors
    /// Returns a typed selection, text or resource refusal without changing the
    /// staged document.
    pub fn replace_text_at(
        &mut self,
        path: &super::ElementPath,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?.replace_text_at(path, authored_text)?;
        Ok(self)
    }

    /// Replace a table cell's content with one paragraph of text.
    ///
    /// See [`super::raw::Tree::replace_cell_text_at`].
    ///
    /// # Errors
    /// Returns a typed selection, dependency, text or resource refusal without
    /// changing the staged document.
    pub fn replace_cell_text_at(
        &mut self,
        path: &super::ElementPath,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_cell_text_at(path, authored_text)?;
        Ok(self)
    }

    /// Set or remove the screen tip of the hyperlink at `hyperlink` (record
    /// 0345).
    ///
    /// See [`super::raw::Tree::set_hyperlink_tooltip_at`].
    ///
    /// # Errors
    /// Returns a typed selection, value or resource refusal without changing
    /// the staged document.
    pub fn set_hyperlink_tooltip_at(
        &mut self,
        hyperlink: &super::ElementPath,
        tooltip: Option<&str>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .set_hyperlink_tooltip_at(hyperlink, tooltip)?;
        Ok(self)
    }

    /// Insert a run into the paragraph or hyperlink at `owner`.
    ///
    /// See [`super::raw::Tree::insert_run_at`].
    ///
    /// # Errors
    /// Returns a typed selection, field, property or resource refusal without
    /// changing the staged document.
    pub fn insert_run_at(
        &mut self,
        owner: &super::ElementPath,
        at: Option<Position>,
        new_run: &super::NewRun,
    ) -> Result<Position> {
        self.inner_mut()?.insert_run_at(owner, at, new_run)
    }

    /// Insert a run holding one break of `kind` into the paragraph at
    /// `paragraph`, before its direct run `at` (after its last child for
    /// `None`), and return the run's position among the paragraph's direct
    /// runs (record 0319).
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::RunBreakType;
    /// use litchi_docx::document::{ElementPath, ElementStep};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Before");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.insert_break_at(&paragraph, None, RunBreakType::Page)?;
    /// let next = edit.commit()?.into_package();
    /// let runs = next.document()?.paragraphs()?[0].runs()?;
    /// assert_eq!(runs.len(), 2);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns a typed selection, field or resource refusal without changing
    /// the staged document.
    pub fn insert_break_at(
        &mut self,
        paragraph: &super::ElementPath,
        at: Option<Position>,
        kind: crate::RunBreakType,
    ) -> Result<Position> {
        let value = match kind {
            crate::RunBreakType::Page => Some("page"),
            crate::RunBreakType::Column => Some("column"),
            crate::RunBreakType::TextWrapping => None,
        };
        self.inner_mut()?
            .insert_run_markup_at(paragraph, at, |prefix| {
                let prefix = String::from_utf8_lossy(prefix);
                let q = |local: &str| {
                    if prefix.is_empty() {
                        local.to_owned()
                    } else {
                        format!("{prefix}:{local}")
                    }
                };
                Ok(match value {
                    Some(value) => format!(
                        r#"<{r}><{br} {ty}="{value}"/></{r}>"#,
                        r = q("r"),
                        br = q("br"),
                        ty = q("type"),
                    ),
                    None => format!("<{r}><{br}/></{r}>", r = q("r"), br = q("br")),
                }
                .into_bytes())
            })
    }

    /// Insert a paragraph into the body, a cell or a block content control.
    ///
    /// See [`super::raw::Tree::insert_paragraph_at`].
    ///
    /// # Errors
    /// Returns a typed selection, property or resource refusal without changing
    /// the staged document.
    pub fn insert_paragraph_at(
        &mut self,
        container: &super::ElementPath,
        anchor: super::ElementAnchor,
        paragraph: &super::NewParagraph,
    ) -> Result<Position> {
        self.inner_mut()?
            .insert_paragraph_at(container, anchor, paragraph)
    }

    /// Insert a table into the body, a cell or a block content control.
    ///
    /// See [`super::raw::Tree::insert_table_at`].
    ///
    /// # Errors
    /// Returns a typed selection, text or resource refusal without changing the
    /// staged document.
    pub fn insert_table_at(
        &mut self,
        container: &super::ElementPath,
        anchor: super::ElementAnchor,
        table: &super::NewTable,
    ) -> Result<Position> {
        self.inner_mut()?.insert_table_at(container, anchor, table)
    }

    /// Insert a row into the table at `table`.
    ///
    /// See [`super::raw::Tree::insert_row_at`].
    ///
    /// # Errors
    /// Returns a typed selection, text or resource refusal without changing the
    /// staged document.
    pub fn insert_row_at(
        &mut self,
        table: &super::ElementPath,
        at: Option<Position>,
        cells: Option<usize>,
        texts: &[String],
    ) -> Result<Position> {
        self.inner_mut()?.insert_row_at(table, at, cells, texts)
    }

    /// Insert a cell into the row at `row`, keeping the table grid in step
    /// (record 0345).
    ///
    /// See [`super::raw::Tree::insert_cell_at`].
    ///
    /// # Errors
    /// Returns a typed selection, text or resource refusal without changing the
    /// staged document.
    pub fn insert_cell_at(
        &mut self,
        row: &super::ElementPath,
        at: Option<Position>,
        cell: &super::NewCell,
    ) -> Result<Position> {
        self.inner_mut()?.insert_cell_at(row, at, cell)
    }

    /// Insert a grid column, with a cell in every row, into the table at
    /// `table` (record 0345).
    ///
    /// See [`super::raw::Tree::insert_column_at`].
    ///
    /// # Errors
    /// Returns a typed selection, merge, text or resource refusal without
    /// changing the staged document.
    pub fn insert_column_at(
        &mut self,
        table: &super::ElementPath,
        at: Option<Position>,
        width: Option<u32>,
        text: &str,
    ) -> Result<Position> {
        self.inner_mut()?.insert_column_at(table, at, width, text)
    }

    /// Remove a grid column and its cells from the table at `table`
    /// (record 0345).
    ///
    /// See [`super::raw::Tree::remove_column_at`].
    ///
    /// # Errors
    /// Returns a typed selection, merge, dependency or resource refusal
    /// without changing the staged document.
    pub fn remove_column_at(
        &mut self,
        table: &super::ElementPath,
        column: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?.remove_column_at(table, column)?;
        Ok(self)
    }

    /// Set the width of a grid column of the table at `table` (record
    /// 0345).
    ///
    /// See [`super::raw::Tree::set_column_width_at`].
    ///
    /// # Errors
    /// Returns a typed selection, value or resource refusal without changing
    /// the staged document.
    pub fn set_column_width_at(
        &mut self,
        table: &super::ElementPath,
        column: Position,
        width: u32,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .set_column_width_at(table, column, width)?;
        Ok(self)
    }

    /// Remove the element at `path`.
    ///
    /// See [`super::raw::Tree::remove_element_at`].
    ///
    /// # Errors
    /// Returns a typed selection, dependency or resource refusal without
    /// changing the staged document.
    pub fn remove_element_at(&mut self, path: &super::ElementPath) -> Result<&mut Self> {
        self.inner_mut()?.remove_element_at(path)?;
        Ok(self)
    }

    /// Move the element at `path` into `container`.
    ///
    /// See [`super::raw::Tree::move_element_at`].
    ///
    /// # Errors
    /// Returns a typed selection, dependency or resource refusal without
    /// changing the staged document.
    pub fn move_element_at(
        &mut self,
        path: &super::ElementPath,
        container: &super::ElementPath,
        anchor: super::ElementAnchor,
    ) -> Result<super::ElementPath> {
        self.inner_mut()?.move_element_at(path, container, anchor)
    }

    /// The exact bytes of the element at `path` in the staged document.
    ///
    /// # Errors
    /// Returns a typed selection refusal for a path that names no element.
    pub fn element_xml(&self, path: &super::ElementPath) -> Result<Vec<u8>> {
        self.inner_ref(|inner| inner.element_xml(path).map(ToOwned::to_owned))?
    }

    /// Remaining semantic-operation slots in the native document ledger.
    ///
    /// This is the native cap, not a limit on the number of caller commands.
    /// One caller command may stage several operations or only exact no-ops.
    #[must_use]
    pub fn remaining_operations(&self) -> usize {
        self.inner
            .try_lock()
            .ok()
            .and_then(|state| state.as_ref().map(transaction::Edit::remaining_operations))
            .unwrap_or(transaction::MAX_OPERATIONS)
    }

    /// Choose how changed main-document XML is compacted at commit.
    #[must_use]
    pub fn with_compaction_policy(mut self, policy: CompactionPolicy) -> Self {
        self.compaction = policy;
        if let Some(inner) = self.inner_slot_mut().take() {
            *self.inner_slot_mut() = Some(inner.with_compaction_policy(policy));
        }
        self
    }

    /// Return the selected compaction policy.
    #[must_use]
    pub const fn compaction_policy(&self) -> CompactionPolicy {
        self.compaction
    }

    /// Clone the staged operations while retaining the same source package.
    ///
    /// # Errors
    /// Returns an allocation or operation admission error.
    pub fn try_clone(&self) -> Result<Self> {
        let inner = {
            let state = match self.inner.try_lock() {
                Ok(state) => state,
                Err(TryLockError::Poisoned(poisoned)) => poisoned.into_inner(),
                Err(TryLockError::WouldBlock) => {
                    return Err(deferred_error("document.edit.clone"));
                },
            };
            match state.as_ref() {
                Some(inner) => Some(inner.try_clone()?),
                None => None,
            }
        };
        Ok(Self {
            package: self.package.clone(),
            inner: Mutex::new(inner),
            materializer: self.materializer,
            options: self.options,
            compaction: self.compaction,
        })
    }

    /// Atomically replace complete text across direct-body paragraphs.
    ///
    /// Positions must be non-empty, unique, and strictly increasing. Every
    /// paragraph retains its run boundaries, formatting, drawings, and unknown
    /// run XML under the same refusal rules as [`Self::replace_paragraph_text`].
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a checked leaf/resource failure without
    /// changing the edit.
    pub fn replace_body_paragraph_texts(
        &mut self,
        replacements: &[ParagraphTextReplacement],
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_body_paragraph_texts(replacements)?;
        Ok(self)
    }

    /// Replace all text in a direct-body paragraph while retaining run
    /// boundaries, formatting, drawings, and unknown run XML.
    ///
    /// Replacement characters are assigned to the existing text slots in
    /// order: each slot keeps up to its original character count and the final
    /// slot receives any remainder. Direct hyperlinks and other paragraph
    /// wrappers use their focused operations and are refused here.
    /// The input must expose its borrowed `str` view as well as a `String`
    /// conversion so managed callers can be admitted before conversion.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl Into<String> + AsRef<str>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_paragraph_text(position, authored_text)?;
        Ok(self)
    }

    /// Replace all text in one direct paragraph hyperlink while leaving its
    /// anchor, tooltip, relationship, target frame, and unknown XML untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_hyperlink_text(
        &mut self,
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_hyperlink_text(paragraph, hyperlink, authored_text)?;
        Ok(self)
    }

    /// Atomically replace direct hyperlinks across one or more direct-body
    /// paragraphs.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Each leaf remains inside one direct
    /// paragraph, so this bounded composite selector cannot cross ownership or
    /// relationship scopes.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked leaf failure also leaves the entire edit unchanged.
    pub fn replace_body_hyperlink_texts(
        &mut self,
        replacements: &[HyperlinkTextReplacement],
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_body_hyperlink_texts(replacements)?;
        Ok(self)
    }

    /// Replace text and structural characters in one direct run while
    /// retaining its `w:rPr`, drawings, and opaque run children. Tabs, line
    /// breaks, carriage returns, non-breaking hyphens, and soft hyphens map to
    /// their native `WordprocessingML` run elements.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_run_text(
        &mut self,
        paragraph: Position,
        run: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_run_text(paragraph, run, authored_text)?;
        Ok(self)
    }

    /// Replace only the first direct `w:t` child of one run while retaining
    /// its attributes and every other run child byte-for-byte. Tabs, line
    /// feeds, and carriage returns remain literal text-node characters.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot. Budget-managed edits
    /// retain the same explicit refusal as native source-preserving run edits.
    pub fn replace_run_text_node(
        &mut self,
        paragraph: Position,
        run: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_run_text_node(paragraph, run, authored_text)?;
        Ok(self)
    }

    /// Mark the owning complex field dirty when a direct run is a cached
    /// result. Runs outside a cached result, empty runs, and already-dirty
    /// fields are no-ops.
    pub fn mark_cached_run_field_dirty(
        &mut self,
        paragraph: Position,
        run: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .mark_cached_run_field_dirty(paragraph, run)?;
        Ok(self)
    }

    /// Replace the displayed text of one direct simple field while preserving
    /// its instruction, dirty/lock state, formatting, and opaque XML.
    ///
    /// Field instructions remain inert and are never evaluated or refreshed.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_simple_field_text(
        &mut self,
        paragraph: Position,
        field: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_simple_field_text(paragraph, field, authored_text)?;
        Ok(self)
    }

    /// Replace the result text between one complex field's direct-run
    /// `begin`/`separate`/`end` markers. Instruction and marker runs remain
    /// exact; nested complex fields inside the selected result are refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_complex_field_result_text(
        &mut self,
        paragraph: Position,
        field: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_complex_field_result_text(paragraph, field, authored_text)?;
        Ok(self)
    }

    /// Replace inert text inside one direct tracked insertion or deletion.
    /// Revision metadata and the wrapper itself remain byte-preserved.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_revision_text(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_revision_text(paragraph, kind, revision, authored_text)?;
        Ok(self)
    }

    /// Apply one direct inline tracked revision using Word redline semantics.
    ///
    /// Insertions are accepted by unwrapping their runs and rejected by
    /// removing them. Deletions are accepted by removing them and rejected by
    /// unwrapping their runs while converting `w:delText` to ordinary
    /// `w:t`. The wrapper's revision metadata is intentionally consumed by
    /// either disposition; bytes belonging to retained children remain exact,
    /// including opaque extension markup and namespace declarations hoisted
    /// from the removed wrapper when needed.
    ///
    /// Only a direct paragraph child is selectable. Moves, range markers,
    /// property/table revisions, nested revision owners, fields, controls,
    /// hyperlinks, and other structures whose ownership or dependency closure
    /// cannot be proven are refused before the edit changes state.
    ///
    /// Like every operation other than direct paragraph text replacement, a
    /// disposition is refused on a budget-managed transaction.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn apply_revision(
        &mut self,
        selector: RevisionSelector,
        action: RevisionAction,
    ) -> Result<&mut Self> {
        self.inner_mut()?.apply_revision(selector, action)?;
        Ok(self)
    }

    /// Accept one direct inline tracked insertion or deletion.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::apply_revision`].
    pub fn accept_revision(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .accept_revision(paragraph, kind, revision)?;
        Ok(self)
    }

    /// Reject one direct inline tracked insertion or deletion.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::apply_revision`].
    pub fn reject_revision(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .reject_revision(paragraph, kind, revision)?;
        Ok(self)
    }

    /// Replace text inside one direct inline content control while retaining
    /// `w:sdtPr`, data binding, lock state, wrapper metadata, and opaque XML.
    ///
    /// This operation supports an inline `w:sdtContent` whose direct children
    /// are runs; block controls and nested control structures are refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_content_control_text(
        &mut self,
        paragraph: Position,
        control: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_content_control_text(paragraph, control, authored_text)?;
        Ok(self)
    }

    /// Replace text in the deepest inline content control reached by a
    /// non-empty direct-child path. The first position selects a direct
    /// paragraph `w:sdt`; later positions select a direct `w:sdt` in the
    /// preceding `w:sdtContent`.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_content_control_text(
        &mut self,
        paragraph: Position,
        controls: &[Position],
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?.replace_nested_content_control_text(
            paragraph,
            controls,
            authored_text,
        )?;
        Ok(self)
    }

    /// Replace one direct hyperlink inside the deepest nested inline control.
    /// Every selector step is a direct child, so the path cannot cross owner
    /// boundaries or relationship scopes.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_content_control_hyperlink_text(
        &mut self,
        paragraph: Position,
        controls: &[Position],
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_nested_content_control_hyperlink_text(
                paragraph,
                controls,
                hyperlink,
                authored_text,
            )?;
        Ok(self)
    }

    /// Replace one direct paragraph inside the deepest block content control
    /// reached by a non-empty path. The first position selects a direct-body
    /// `w:sdt`; later positions select a direct nested `w:sdt`.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_block_content_control_paragraph_text(
        &mut self,
        controls: &[Position],
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_block_content_control_paragraph_text(controls, paragraph, authored_text)?;
        Ok(self)
    }

    /// Replace one direct hyperlink inside a path-addressed block-control
    /// paragraph without changing its relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_block_content_control_paragraph_hyperlink_text(
        &mut self,
        controls: &[Position],
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_block_content_control_paragraph_hyperlink_text(
                controls,
                paragraph,
                hyperlink,
                authored_text,
            )?;
        Ok(self)
    }

    /// Atomically replace direct hyperlinks across one or more direct
    /// paragraphs in the deepest block content control reached by `controls`.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Direct-child selection keeps every
    /// leaf within the same path-addressed control owner.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked path or leaf failure also leaves the entire edit
    /// unchanged.
    pub fn replace_block_content_control_paragraph_hyperlink_texts(
        &mut self,
        controls: &[Position],
        replacements: &[HyperlinkTextReplacement],
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_block_content_control_paragraph_hyperlink_texts(controls, replacements)?;
        Ok(self)
    }

    /// Replace text in a basic direct-body table cell.
    ///
    /// The supported cell contains one direct paragraph. Its existing runs,
    /// formatting, cell properties, drawings, and unknown run XML remain in
    /// place; nested tables, controls, and multiple cell paragraphs are
    /// refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_table_cell_text(table, row, cell, authored_text)?;
        Ok(self)
    }

    /// Replace one direct paragraph in a rich or multi-paragraph table cell.
    /// Other paragraphs, nested tables, cell properties, and opaque cell XML
    /// remain untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_paragraph_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?.replace_table_cell_paragraph_text(
            table,
            row,
            cell,
            paragraph,
            authored_text,
        )?;
        Ok(self)
    }

    /// Replace one direct paragraph in a cell reached through a non-empty
    /// body-to-nested-table path. Each later table must be a direct child of
    /// the preceding selected cell.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_table_cell_paragraph_text(
        &mut self,
        path: &[TableCellAddress],
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?.replace_nested_table_cell_paragraph_text(
            path,
            paragraph,
            authored_text,
        )?;
        Ok(self)
    }

    /// Replace one direct hyperlink inside a path-addressed nested-table cell
    /// paragraph without changing its relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_table_cell_paragraph_hyperlink_text(
        &mut self,
        path: &[TableCellAddress],
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_nested_table_cell_paragraph_hyperlink_text(
                path,
                paragraph,
                hyperlink,
                authored_text,
            )?;
        Ok(self)
    }

    /// Atomically replace direct hyperlinks across one or more direct
    /// paragraphs in the deepest cell reached by `path`.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Direct-child selection keeps every
    /// leaf within the same path-addressed cell owner.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked path or leaf failure also leaves the entire edit
    /// unchanged.
    pub fn replace_nested_table_cell_paragraph_hyperlink_texts(
        &mut self,
        path: &[TableCellAddress],
        replacements: &[HyperlinkTextReplacement],
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .replace_nested_table_cell_paragraph_hyperlink_texts(path, replacements)?;
        Ok(self)
    }

    /// Insert a compact plain-text paragraph at a projected zero-based
    /// position. `position == paragraph_count()` appends before the body-final
    /// section properties.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn insert_paragraph(
        &mut self,
        position: Position,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .insert_paragraph(position, authored_text)?;
        Ok(self)
    }

    /// Insert one non-mutating dependency-checked paragraph transfer plan.
    ///
    /// The plan is receiver-specific: every relationship reference was mapped
    /// to an equivalent relationship already owned by the exact target
    /// package. This method never copies or guesses package dependencies.
    ///
    /// # Errors
    ///
    /// Returns a stale-plan, position, operation-bound, or XML validation
    /// error without changing the projected snapshot.
    pub fn insert_paragraph_transfer(
        &mut self,
        position: Position,
        plan: &ParagraphTransfer,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .insert_paragraph_transfer(position, plan)?;
        Ok(self)
    }
}

impl Edit<Package> {
    /// Validate the staged document and publish a complete immutable package.
    ///
    /// ```
    /// use litchi_docx::Package;
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let source = Package::new()?;
    /// let edit = source.edit()?;
    /// drop(source);
    /// let commit = edit.commit()?;
    /// assert!(!commit.diagnostics().changed());
    /// let (package, patch) = commit.into_parts();
    /// assert!(patch.is_noop());
    /// assert_eq!(package.document()?.paragraph_count()?, 0);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns document validation, stale-source, signature, protection or
    /// package publication errors without changing the retained source.
    pub fn commit(self) -> Result<Commit> {
        let this = self;
        let source_package = this.package.clone();
        let committed = this.into_inner()?.commit()?;
        let diagnostics = committed.diagnostics();
        let (_, patch) = committed.into_parts();
        let package = source_package.apply(&patch)?.into_package();
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

impl Edit<crate::source_backed::Package> {
    /// Validate the staged document and retain a new immutable source owner.
    ///
    /// No ZIP is serialized or reopened. The source owner remains unchanged.
    ///
    /// ```
    /// use litchi_docx::{document, source_backed};
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let bytes = include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
    ///     "/../../test-data/ooxml/docx/documentProperties.docx"));
    /// let source = source_backed::Package::from_reader(bytes.as_slice())?;
    /// let edit = source.edit()?;
    /// drop(source);
    /// let commit: document::Commit<source_backed::Package> = edit.commit()?;
    /// let (package, patch) = commit.into_parts();
    /// assert!(patch.is_noop());
    /// let mut output = Vec::new();
    /// package.write_to(&mut output)?;
    /// assert_eq!(output, bytes);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns validation, source, signature, protection or admission failures.
    pub fn commit(self) -> Result<Commit<crate::source_backed::Package>> {
        let this = self;
        let source_package = this.package.clone();
        let committed = this.into_inner()?.commit()?;
        let diagnostics = committed.diagnostics();
        let (_, patch) = committed.into_parts();
        let package = source_package.apply(&patch)?.into_package();
        let patch = patch.with_restore(source_package, package.clone());
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// A validated main-document edit with its complete published package.
///
/// A standalone XML commit cannot be substituted for a package commit:
///
/// ```compile_fail
/// use litchi_docx::{Package, document};
/// let source = Package::new().unwrap().document_snapshot().unwrap();
/// let commit: document::Commit = source.edit().commit().unwrap();
/// ```
#[derive(Clone)]
pub struct Commit<P = Package> {
    package: P,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl<P> std::fmt::Debug for Commit<P> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl<P> Commit<P> {
    pub(crate) const fn from_parts(package: P, patch: Patch, diagnostics: Diagnostics) -> Self {
        Self {
            package,
            patch,
            diagnostics,
        }
    }

    /// Borrow the immutable published package.
    #[must_use]
    pub const fn package(&self) -> &P {
        &self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return content-free commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Take the published package.
    #[must_use]
    pub fn into_package(self) -> P {
        self.package
    }

    /// Take the published package and reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (P, Patch) {
        (self.package, self.patch)
    }
}
