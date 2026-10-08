//! Package-bound main-document transaction orchestration.

use super::{
    CompactionPolicy, Diagnostics, HyperlinkTextReplacement, ParagraphTextReplacement,
    ParagraphTransfer, Patch, Result, RevisionAction, RevisionKind, RevisionSelector, Snapshot,
    TableCellAddress, transaction,
};
use crate::Package;
use litchi_core::Position;

/// An isolated main-document edit retaining its complete source package.
///
/// `P` defaults to the eager [`crate::Package`]; source-backed owners use
/// [`crate::source_backed::Package`].
///
/// Dropping an edit rolls it back. Committing validates and returns a new
/// package; the source is unchanged on success and failure.
pub struct Edit<P = Package> {
    package: P,
    inner: transaction::Edit,
}

impl<P> std::fmt::Debug for Edit<P> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl<P: Clone> Edit<P> {
    pub(crate) fn new(package: P, inner: transaction::Edit) -> Self {
        Self { package, inner }
    }

    /// Borrow the source main-document projection.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.inner.source()
    }

    /// Borrow the staged main-document projection.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        self.inner.projected()
    }

    /// Choose how changed main-document XML is compacted at commit.
    #[must_use]
    pub fn with_compaction_policy(mut self, policy: CompactionPolicy) -> Self {
        self.inner = self.inner.with_compaction_policy(policy);
        self
    }

    /// Return the selected compaction policy.
    #[must_use]
    pub const fn compaction_policy(&self) -> CompactionPolicy {
        self.inner.compaction_policy()
    }

    /// Clone the staged operations while retaining the same source package.
    ///
    /// # Errors
    /// Returns an allocation or operation admission error.
    pub fn try_clone(&self) -> Result<Self> {
        Ok(Self {
            package: self.package.clone(),
            inner: self.inner.try_clone()?,
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
        self.inner.replace_body_paragraph_texts(replacements)?;
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
        self.inner.replace_paragraph_text(position, authored_text)?;
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
        self.inner
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
        self.inner.replace_body_hyperlink_texts(replacements)?;
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
        self.inner.replace_run_text(paragraph, run, authored_text)?;
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
        self.inner
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
        self.inner
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
        self.inner
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
        self.inner.apply_revision(selector, action)?;
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
        self.inner.accept_revision(paragraph, kind, revision)?;
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
        self.inner.reject_revision(paragraph, kind, revision)?;
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
        self.inner
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
        self.inner
            .replace_nested_content_control_text(paragraph, controls, authored_text)?;
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
        self.inner.replace_nested_content_control_hyperlink_text(
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
        self.inner.replace_block_content_control_paragraph_text(
            controls,
            paragraph,
            authored_text,
        )?;
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
        self.inner
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
        self.inner
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
        self.inner
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
        self.inner
            .replace_table_cell_paragraph_text(table, row, cell, paragraph, authored_text)?;
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
        self.inner
            .replace_nested_table_cell_paragraph_text(path, paragraph, authored_text)?;
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
        self.inner
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
        self.inner
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
        self.inner.insert_paragraph(position, authored_text)?;
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
        self.inner.insert_paragraph_transfer(position, plan)?;
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
        let committed = self.inner.commit()?;
        let diagnostics = committed.diagnostics();
        let (_, patch) = committed.into_parts();
        let package = self.package.apply(&patch)?;
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
    /// package.to_stream(&mut output)?;
    /// assert_eq!(output, bytes);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns validation, source, signature, protection or admission failures.
    pub fn commit(self) -> Result<Commit<crate::source_backed::Package>> {
        let committed = self.inner.commit()?;
        let diagnostics = committed.diagnostics();
        let (_, patch) = committed.into_parts();
        let package = self.package.apply(&patch)?;
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
