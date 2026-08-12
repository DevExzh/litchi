//! Immutable, source-bound transactions for packaged ODT documents.
//!
//! This is the safe package-level mutation boundary. It intentionally stages
//! only operations that keep the authoritative XML snapshot intact: callers
//! get exact no-op bytes, source-checked reversible patches, and a complete
//! compact-XML audit before publication. Broader structural `MutableDocument`
//! operations remain available separately while their opaque-content
//! preservation contracts are migrated to this transaction surface.

#![deny(
    clippy::expect_used,
    clippy::panic,
    clippy::todo,
    clippy::unimplemented,
    clippy::unwrap_used
)]

use crate::{
    Document,
    elements::field::{DynamicTextField, FieldParser},
    mutable::MutableDocument,
};
use litchi_core::{
    BlobBundle, BlobId, BlobLimits, Error, ForwardOnly, History as CoreHistory, JoinedSubEdits,
    Patch as CorePatch, PatchLimits, PatchOperation, Result, Reversible, ReversibleOperation,
    SubEdit, ThreeWayMergePlan,
};
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    sync::Arc,
};

/// Shared zero-based semantic collection position.
pub use litchi_core::Position;
pub use litchi_core::{
    CompositionLimits, ConflictSet, HistoryLimits, MergeChoice, SubEditConflict, SubEditJoinFailure,
};

const MAX_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_OPERATIONS: usize = 1_024;
const MAX_WIRE_JSON_BYTES: usize = 192 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_BYTES: usize = 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odt";
const RESTORE_OPERATION: &str = "snapshot.restore";
const NOOP_OPERATION: &str = "transaction.noop";
const SOURCE_PRECONDITION: &str = "source_sha256";
const TARGET_PRECONDITION: &str = "target_sha256";
const DEFAULT_COMPOSITION_LIMITS: CompositionLimits =
    CompositionLimits::new(1_024, 64, 4_096, 1_024);

/// Immutable, validated ODT package snapshot.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<Vec<u8>>,
}

impl Snapshot {
    /// Opens and retains an ODT package as an immutable snapshot.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        ensure_package_size(bytes.len(), "ODT transaction input")?;
        let document = Document::from_bytes(bytes)?;
        Self::from_document(&document)
    }

    /// Captures the exact bytes backing an already validated document.
    pub fn from_document(document: &Document) -> Result<Self> {
        ensure_package_size(document.original_bytes().len(), "ODT transaction package")?;
        Ok(Self {
            bytes: document.transaction_package().shared_bytes(),
        })
    }

    /// Returns the exact package bytes represented by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Reopens this immutable snapshot for semantic inspection.
    pub fn document(&self) -> Result<Document> {
        Document::from_shared_bytes(Arc::clone(&self.bytes))
    }

    /// Starts a detached, failure-atomic package edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            operations: Vec::new(),
        }
    }

    /// Starts a bounded deterministic collection of independently prepared edits.
    #[must_use]
    pub fn joined_edit(&self) -> JoinedEdit {
        JoinedEdit::new(self.clone(), DEFAULT_COMPOSITION_LIMITS)
    }

    /// Starts commit-coupled bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Reports the package envelope policy enforced by immutable edits.
    pub fn envelope_kind(&self) -> Result<EnvelopeKind> {
        envelope_kind(self)
    }
}

fn snapshot_from_final_document(source: &Snapshot, document: &Document) -> Result<Snapshot> {
    if document.original_bytes() == source.as_bytes() {
        Ok(source.clone())
    } else {
        Snapshot::from_document(document)
    }
}

/// Security envelope classification relevant to package mutation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EnvelopeKind {
    /// Credential-free package with no document or macro signature part.
    Plain,
    /// Package containing a document or macro signature part.
    Signed,
    /// Package manifest containing at least one encrypted entry.
    Encrypted,
}

/// Selector used by packaged-document edit operations.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ParagraphSelector {
    /// A checked zero-based paragraph position in semantic document order.
    Position(Position),
    /// The zero-based paragraph position in semantic document order.
    ///
    /// Prefer [`ParagraphSelector::position`] for new code.
    Index(usize),
    /// The unique paragraph whose extracted semantic text equals this value.
    ExactText(String),
}

impl ParagraphSelector {
    /// Select a paragraph by a checked semantic document position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Select a paragraph by semantic document order.
    #[must_use]
    pub const fn at(index: usize) -> Self {
        Self::Index(index)
    }

    /// Select the one paragraph matching `text` exactly.
    #[must_use]
    pub fn exact_text(text: impl Into<String>) -> Self {
        Self::ExactText(text.into())
    }
}

/// Detached, lossless packaged-ODT edit.
pub struct Edit {
    source: Snapshot,
    operations: Vec<Operation>,
}

impl Edit {
    /// Stages insertion of one plain paragraph at a semantic position.
    pub fn insert_paragraph(
        &mut self,
        position: Position,
        text: impl Into<String>,
    ) -> Result<&mut Self> {
        let text = bounded_semantic_text(text.into(), "paragraph text")?;
        self.push(Operation::InsertParagraph {
            index: position.get(),
            text,
        })
    }

    /// Stages replacement of one paragraph's content with plain text.
    pub fn replace_paragraph(
        &mut self,
        position: Position,
        text: impl Into<String>,
    ) -> Result<&mut Self> {
        let text = bounded_semantic_text(text.into(), "paragraph text")?;
        self.push(Operation::ReplaceParagraph {
            index: position.get(),
            text,
        })
    }

    /// Stages removal of one paragraph selected by semantic position.
    pub fn remove_paragraph(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveParagraph {
            index: position.get(),
        })
    }

    /// Stages one typed text run at the end of a paragraph.
    pub fn append_run(
        &mut self,
        paragraph: Position,
        text: impl Into<String>,
        style_name: Option<&str>,
    ) -> Result<&mut Self> {
        let text = bounded_semantic_text(text.into(), "run text")?;
        let style_name = style_name
            .map(|value| bounded_semantic_text(value.to_owned(), "run style name"))
            .transpose()?;
        self.push(Operation::AppendRun {
            paragraph: paragraph.get(),
            text,
            style_name,
        })
    }

    /// Stages one inert hyperlink at the end of a paragraph.
    pub fn append_hyperlink(
        &mut self,
        paragraph: Position,
        href: impl Into<String>,
        text: impl Into<String>,
    ) -> Result<&mut Self> {
        let href = bounded_semantic_text(href.into(), "hyperlink target")?;
        let text = bounded_semantic_text(text.into(), "hyperlink text")?;
        self.push(Operation::AppendHyperlink {
            paragraph: paragraph.get(),
            href,
            text,
        })
    }

    /// Stages a line break in the selected paragraph without rebuilding its
    /// inline markup. Scripts, forms, DDE declarations, and other active
    /// content are retained as inert package data and are never executed.
    pub fn append_line_break(&mut self, selector: ParagraphSelector) -> Result<&mut Self> {
        let index = match &selector {
            ParagraphSelector::Index(index) => *index,
            ParagraphSelector::Position(_) | ParagraphSelector::ExactText(_) => {
                resolve_paragraph(&self.source.document()?, &selector)?
            },
        };
        self.push(Operation::AppendLineBreak { index })
    }

    /// Stages a validated footnote or endnote at the end of one paragraph.
    pub fn insert_note(
        &mut self,
        paragraph: Position,
        note: &crate::note::Note,
    ) -> Result<&mut Self> {
        note.validate()?;
        self.push(Operation::InsertNote {
            paragraph: paragraph.get(),
            note: note.clone(),
            fragment: None,
        })
    }

    /// Stages replacement of one note in semantic document order.
    pub fn replace_note(
        &mut self,
        note: Position,
        replacement: &crate::note::Note,
    ) -> Result<&mut Self> {
        replacement.validate()?;
        self.push(Operation::ReplaceNote {
            index: note.get(),
            note: replacement.clone(),
            fragment: None,
        })
    }

    /// Stages removal of one note in semantic document order.
    pub fn remove_note(&mut self, note: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveNote { index: note.get() })
    }

    /// Stages a validated ruby annotation at the end of one paragraph.
    pub fn insert_ruby_annotation(
        &mut self,
        paragraph: Position,
        annotation: &crate::ruby_family::Annotation,
    ) -> Result<&mut Self> {
        annotation.validate()?;
        self.push(Operation::InsertRubyAnnotation {
            paragraph: paragraph.get(),
            annotation: annotation.clone(),
        })
    }

    /// Stages replacement of one ruby annotation in document order.
    pub fn replace_ruby_annotation(
        &mut self,
        annotation: Position,
        replacement: &crate::ruby_family::Annotation,
    ) -> Result<&mut Self> {
        replacement.validate()?;
        self.push(Operation::ReplaceRubyAnnotation {
            index: annotation.get(),
            annotation: replacement.clone(),
        })
    }

    /// Stages removal of one ruby annotation in document order.
    pub fn remove_ruby_annotation(&mut self, annotation: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveRubyAnnotation {
            index: annotation.get(),
        })
    }

    /// Stages a typed dynamic field at the end of one paragraph.
    pub fn insert_dynamic_text_field(
        &mut self,
        paragraph: Position,
        field: &DynamicTextField,
    ) -> Result<&mut Self> {
        field.validate()?;
        self.push(Operation::InsertDynamicTextField {
            paragraph: paragraph.get(),
            field: field.clone(),
        })
    }

    /// Stages replacement of one typed dynamic field in document order.
    pub fn replace_dynamic_text_field(
        &mut self,
        field: Position,
        replacement: &DynamicTextField,
    ) -> Result<&mut Self> {
        replacement.validate()?;
        self.push(Operation::ReplaceDynamicTextField {
            index: field.get(),
            field: replacement.clone(),
        })
    }

    /// Stages removal of one typed dynamic field in document order.
    pub fn remove_dynamic_text_field(&mut self, field: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveDynamicTextField { index: field.get() })
    }

    /// Stages insertion or replacement of one named ruby style.
    pub fn set_ruby_style(&mut self, style: &crate::ruby_family::Style) -> Result<&mut Self> {
        style.validate()?;
        self.push(Operation::SetRubyStyle {
            style: style.clone(),
        })
    }

    /// Stages removal of one named ruby style.
    pub fn remove_ruby_style(&mut self, name: &str) -> Result<&mut Self> {
        let name = bounded_semantic_text(name.to_owned(), "ruby style name")?;
        self.push(Operation::RemoveRubyStyle { name })
    }

    /// Stages inert tracked-change policy metadata without accepting a credential.
    pub fn set_tracked_change_policy(
        &mut self,
        track_changes: Option<bool>,
        protection_key: Option<&str>,
        digest_algorithm: Option<&str>,
    ) -> Result<&mut Self> {
        let protection_key = protection_key
            .map(|value| bounded_semantic_text(value.to_owned(), "tracked-change protection key"))
            .transpose()?;
        let digest_algorithm = digest_algorithm
            .map(|value| bounded_semantic_text(value.to_owned(), "tracked-change digest algorithm"))
            .transpose()?;
        self.push(Operation::SetTrackedChangePolicy {
            track_changes,
            protection_key,
            digest_algorithm,
        })
    }

    /// Stages removal of one tracked-change declaration and its correlated markers.
    pub fn remove_tracked_change(&mut self, id: &str) -> Result<&mut Self> {
        let id = bounded_semantic_text(id.to_owned(), "tracked-change ID")?;
        self.push(Operation::RemoveTrackedChange { id })
    }

    /// Stages creation of one inert RDF/XML metadata graph.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<&mut Self> {
        self.push(Operation::AddRdfGraph {
            preferred_path: preferred_path.map(str::to_owned),
            triples: triples.to_vec(),
        })
    }

    /// Stages replacement of one RDF graph selected by package path.
    pub fn replace_rdf_graph(
        &mut self,
        path: &str,
        triples: &[crate::rdf::Triple],
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceRdfGraph {
            path: path.to_owned(),
            triples: triples.to_vec(),
        })
    }

    /// Stages removal of one RDF graph selected by package path.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<&mut Self> {
        self.push(Operation::RemoveRdfGraph {
            path: path.to_owned(),
        })
    }

    /// Stages inert document-protection metadata. The policy is retained but
    /// never enforced by this library.
    pub fn set_protection(&mut self, policy: &crate::protection::Policy) -> Result<&mut Self> {
        self.push(Operation::SetProtection {
            policy: policy.clone(),
        })
    }

    /// Stages one RDF assertion in a graph selected by package path.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<&mut Self> {
        self.push(Operation::AddRdfTriple {
            path: path.to_owned(),
            triple: triple.clone(),
        })
    }

    /// Stages replacement of one RDF assertion.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceRdfTriple {
            path: path.to_owned(),
            index,
            triple: triple.clone(),
        })
    }

    /// Stages removal of one RDF assertion.
    #[deprecated(note = "use remove_rdf_triple_at with a checked Position")]
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveRdfTriple {
            path: path.to_owned(),
            index,
        })
    }

    /// Stages removal of one RDF assertion selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_rdf_triple_at(&mut self, path: &str, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveRdfTriple {
            path: path.to_owned(),
            index: position.get(),
        })
    }

    /// Stages an RDF assertion move; equal selectors are an exact no-op.
    #[deprecated(note = "use move_rdf_triple_to with checked Positions")]
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveRdfTriple {
            path: path.to_owned(),
            from,
            to,
        })
    }

    /// Stages an RDF assertion move between checked semantic positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn move_rdf_triple_to(
        &mut self,
        path: &str,
        from: Position,
        to: Position,
    ) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveRdfTriple {
            path: path.to_owned(),
            from: from.get(),
            to: to.get(),
        })
    }

    /// Stages insertion of a top-level inert form into a form group.
    pub fn add_form(
        &mut self,
        group_index: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<&mut Self> {
        self.push(Operation::AddForm {
            group_index,
            form: form.clone(),
        })
    }

    /// Stages insertion of an inert nested form.
    pub fn add_nested_form(
        &mut self,
        parent_form: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<&mut Self> {
        self.push(Operation::AddNestedForm {
            parent_form,
            form: form.clone(),
        })
    }

    /// Stages insertion of an inert form control.
    pub fn add_form_control(
        &mut self,
        form_index: usize,
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<&mut Self> {
        self.push(Operation::AddFormControl {
            form_index,
            control: control.clone(),
        })
    }

    /// Stages replacement of an inert form control.
    pub fn replace_form_control(
        &mut self,
        index: usize,
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceFormControl {
            index,
            control: control.clone(),
        })
    }

    /// Stages removal of an inert form control.
    #[deprecated(note = "use remove_form_control_at with a checked Position")]
    pub fn remove_form_control(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveFormControl { index })
    }

    /// Stages removal of a form control selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_form_control_at(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveFormControl {
            index: position.get(),
        })
    }

    /// Stages a form-control move; equal selectors are an exact no-op.
    #[deprecated(note = "use move_form_control_to with checked Positions")]
    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveFormControl { from, to })
    }

    /// Stages a form-control move between checked semantic positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn move_form_control_to(&mut self, from: Position, to: Position) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveFormControl {
            from: from.get(),
            to: to.get(),
        })
    }

    /// Stages replacement of a form selected in semantic document order.
    pub fn replace_form(
        &mut self,
        index: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceForm {
            index,
            form: form.clone(),
        })
    }

    /// Stages removal of a form selected in semantic document order.
    #[deprecated(note = "use remove_form_at with a checked Position")]
    pub fn remove_form(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveForm { index })
    }

    /// Stages removal of a form selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_form_at(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveForm {
            index: position.get(),
        })
    }

    /// Stages a form move; equal selectors are an exact no-op.
    #[deprecated(note = "use move_form_to with checked Positions")]
    pub fn move_form(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveForm { from, to })
    }

    /// Stages a form move between checked semantic positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn move_form_to(&mut self, from: Position, to: Position) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveForm {
            from: from.get(),
            to: to.get(),
        })
    }

    /// Stages creation of a packaged or inline embedded chart. Formula and
    /// calculation metadata remains inert and is never evaluated.
    pub fn add_embedded_chart(&mut self, definition: &crate::odc::Definition) -> Result<&mut Self> {
        self.push(Operation::AddEmbeddedChart {
            definition: definition.clone(),
        })
    }

    /// Stages creation of a chart with an explicit package or inline storage form.
    pub fn add_embedded_chart_with_storage(
        &mut self,
        definition: &crate::odc::Definition,
        storage: crate::package::charts::EmbeddedChartStorage,
    ) -> Result<&mut Self> {
        self.push(Operation::AddEmbeddedChartWithStorage {
            definition: definition.clone(),
            storage,
        })
    }

    /// Stages replacement of an embedded chart selected in document order.
    pub fn replace_embedded_chart(
        &mut self,
        index: usize,
        definition: &crate::odc::Definition,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceEmbeddedChart {
            index,
            definition: definition.clone(),
        })
    }

    /// Stages removal of an embedded chart selected in document order.
    #[deprecated(note = "use remove_embedded_chart_at with a checked Position")]
    pub fn remove_embedded_chart(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedChart { index })
    }

    /// Stages removal of an embedded chart selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_chart_at(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedChart {
            index: position.get(),
        })
    }

    /// Stages creation of an inert embedded object or image.
    pub fn add_embedded_resource(
        &mut self,
        resource: &crate::package::embedded::EmbeddedResource,
    ) -> Result<&mut Self> {
        self.push(Operation::AddEmbeddedResource {
            resource: resource.clone(),
        })
    }

    /// Stages up to 256 base-snapshot embedded-resource changes for one
    /// preflight and one package publication.
    ///
    /// Every selector is resolved against the document snapshot at batch
    /// execution, after all previously staged operations. Selecting one owner
    /// twice, selecting an unsupported owner, or any late resource/path/reference
    /// failure refuses the entire transaction.
    /// The corresponding [`OperationResult::Indices`] contains positions for
    /// `Add` changes only, in batch order and within each resource family.
    pub fn edit_embedded_resources(
        &mut self,
        changes: &[crate::package::embedded::EmbeddedResourceChange],
    ) -> Result<&mut Self> {
        crate::package::embedded::validate_batch_limits(changes)?;
        if changes.is_empty() {
            return Ok(self);
        }
        self.push(Operation::EmbeddedResourceBatch {
            changes: changes.to_vec(),
        })
    }

    /// Stages replacement of an embedded object selected in document order.
    pub fn replace_embedded_object(
        &mut self,
        index: usize,
        resource: &crate::package::embedded::EmbeddedResource,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceEmbeddedObject {
            index,
            resource: resource.clone(),
        })
    }

    /// Stages replacement of an embedded image selected in document order.
    pub fn replace_embedded_image(
        &mut self,
        index: usize,
        resource: &crate::package::embedded::EmbeddedResource,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceEmbeddedImage {
            index,
            resource: resource.clone(),
        })
    }

    /// Stages removal of an embedded object selected in document order.
    #[deprecated(note = "use remove_embedded_object_at with a checked Position")]
    pub fn remove_embedded_object(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedObject { index })
    }

    /// Stages removal of an embedded object selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_object_at(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedObject {
            index: position.get(),
        })
    }

    /// Stages removal of an embedded image selected in document order.
    #[deprecated(note = "use remove_embedded_image_at with a checked Position")]
    pub fn remove_embedded_image(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedImage { index })
    }

    /// Stages removal of an embedded image selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_image_at(&mut self, position: Position) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedImage {
            index: position.get(),
        })
    }

    /// Stages an embedded-object move; equal selectors are an exact no-op.
    #[deprecated(note = "use move_embedded_object_to with checked Positions")]
    pub fn move_embedded_object(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveEmbeddedObject { from, to })
    }

    /// Stages an embedded-object move between checked semantic positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn move_embedded_object_to(&mut self, from: Position, to: Position) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveEmbeddedObject {
            from: from.get(),
            to: to.get(),
        })
    }

    /// Stages an embedded-image move; equal selectors are an exact no-op.
    #[deprecated(note = "use move_embedded_image_to with checked Positions")]
    pub fn move_embedded_image(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveEmbeddedImage { from, to })
    }

    /// Stages an embedded-image move between checked semantic positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn move_embedded_image_to(&mut self, from: Position, to: Position) -> Result<&mut Self> {
        if from == to {
            return Ok(self);
        }
        self.push(Operation::MoveEmbeddedImage {
            from: from.get(),
            to: to.get(),
        })
    }

    /// Stages validated opaque script bytes. They are never interpreted,
    /// linked, loaded, or executed.
    pub fn add_script_resource(
        &mut self,
        resource: &crate::ScriptResourceSpec,
    ) -> Result<&mut Self> {
        self.push(Operation::AddScriptResource {
            resource: resource.clone(),
        })
    }

    /// Stages replacement of opaque script bytes at an exact package path.
    pub fn replace_script_resource(
        &mut self,
        path: &str,
        resource: &crate::ScriptResourceSpec,
    ) -> Result<&mut Self> {
        self.push(Operation::ReplaceScriptResource {
            path: path.to_owned(),
            resource: resource.clone(),
        })
    }

    /// Stages removal of an unreferenced opaque script resource.
    pub fn remove_script_resource(&mut self, path: &str) -> Result<&mut Self> {
        self.push(Operation::RemoveScriptResource {
            path: path.to_owned(),
        })
    }

    fn push(&mut self, operation: Operation) -> Result<&mut Self> {
        if self.operations.len() >= MAX_OPERATIONS {
            return Err(Error::InvalidFormat(format!(
                "ODT transaction exceeds {MAX_OPERATIONS} staged operations"
            )));
        }
        self.operations
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT transaction operations",
                source,
            })?;
        self.operations.push(operation);
        Ok(self)
    }

    /// Validates every staged operation and publishes one immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        if self.operations.is_empty() {
            return Ok(Commit::new(
                self.source.clone(),
                self.source,
                Vec::new(),
                Vec::new(),
            ));
        }

        self.commit_changed()
    }

    #[inline(never)]
    fn commit_changed(self) -> Result<Commit> {
        ensure_editable_envelope(&self.source)?;
        let mut document = self.source.document()?;
        let mut results = Vec::new();
        results
            .try_reserve_exact(self.operations.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT transaction results",
                source,
            })?;
        let mut operation_index = 0;
        while operation_index < self.operations.len() {
            let operation = &self.operations[operation_index];
            let before_operation = document.transaction_package().clone();
            if matches!(operation, Operation::ReplaceParagraph { .. })
                && matches!(
                    self.operations.get(operation_index + 1),
                    Some(Operation::ReplaceParagraph { .. })
                )
            {
                // Plain-text replacement cannot change paragraph topology, so
                // consecutive operations retain their scalar position and
                // last-write ordering when applied to one mutable candidate.
                // Publish and audit only that candidate: no intermediate
                // package is observable through the transaction API.
                let mut mutable = MutableDocument::from_document(document)?;
                while let Some(Operation::ReplaceParagraph { index, text }) =
                    self.operations.get(operation_index)
                {
                    mutable.replace_semantic_paragraph(*index, text)?;
                    results.push(OperationResult::Unit);
                    operation_index += 1;
                }
                document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                audit_changed_xml_is_compact(&before_operation, document.transaction_package())?;
                continue;
            }
            #[allow(
                deprecated,
                reason = "only this dispatch expression still reaches validated legacy codecs"
            )]
            let result = match operation {
                Operation::Noop => OperationResult::Unit,
                Operation::RestoreSnapshot => {
                    return Err(Error::Unsupported(
                        "snapshot restoration is available only through patch application"
                            .to_string(),
                    ));
                },
                Operation::InsertParagraph { index, text } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.insert_semantic_paragraph(*index, text)?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Index(*index)
                },
                Operation::ReplaceParagraph { index, text } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.replace_semantic_paragraph(*index, text)?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveParagraph { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_semantic_paragraph(*index)?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Unit
                },
                Operation::AppendRun {
                    paragraph,
                    text,
                    style_name,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_semantic_run(*paragraph, text, style_name.as_deref())?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Unit
                },
                Operation::AppendHyperlink {
                    paragraph,
                    href,
                    text,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_semantic_hyperlink(*paragraph, href, text)?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Unit
                },
                Operation::AppendLineBreak { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_line_break_at(Position::new(*index))?;
                    document = Document::from_bytes(mutable.to_bytes_content_only()?)?;
                    OperationResult::Unit
                },
                Operation::InsertNote {
                    paragraph,
                    note,
                    fragment,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    if let Some(fragment) = fragment {
                        mutable.insert_note_fragment(*paragraph, fragment)?;
                    } else {
                        mutable.insert_note(*paragraph, note)?;
                    }
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::ReplaceNote {
                    index,
                    note,
                    fragment,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    if let Some(fragment) = fragment {
                        mutable.replace_note_fragment(*index, fragment)?;
                    } else {
                        mutable.replace_note(*index, note)?;
                    }
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveNote { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_note(*index)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::InsertRubyAnnotation {
                    paragraph,
                    annotation,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.insert_ruby_annotation(*paragraph, annotation)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::ReplaceRubyAnnotation { index, annotation } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.replace_ruby_annotation(*index, annotation)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveRubyAnnotation { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_ruby_annotation(*index)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::InsertDynamicTextField { paragraph, field } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.insert_dynamic_text_field(*paragraph, field)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::ReplaceDynamicTextField { index, field } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.replace_dynamic_text_field(*index, field)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveDynamicTextField { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_dynamic_text_field(*index)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::SetRubyStyle { style } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.set_ruby_style(style)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveRubyStyle { name } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_ruby_style(name)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::SetTrackedChangePolicy {
                    track_changes,
                    protection_key,
                    digest_algorithm,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.set_tracked_change_policy(
                        *track_changes,
                        protection_key.clone(),
                        digest_algorithm.clone(),
                    )?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveTrackedChange { id } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_tracked_change(id)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::AddRdfGraph {
                    preferred_path,
                    triples,
                } => OperationResult::Path(
                    document.add_rdf_graph(preferred_path.as_deref(), triples)?,
                ),
                Operation::ReplaceRdfGraph { path, triples } => {
                    document.replace_rdf_graph(path, triples)?;
                    OperationResult::Unit
                },
                Operation::RemoveRdfGraph { path } => {
                    document.remove_rdf_graph(path)?;
                    OperationResult::Unit
                },
                Operation::SetProtection { policy } => {
                    let before = document.protection()?;
                    if &before != policy {
                        let mimetype = document.transaction_package().mimetype()?;
                        let bytes = crate::protection::rewrite_owned_package(
                            document.transaction_package(),
                            &mimetype,
                            policy,
                        )?;
                        document.replace_transaction_bytes(bytes)?;
                    }
                    OperationResult::Unit
                },
                Operation::AddRdfTriple { path, triple } => {
                    let index = document
                        .rdf_graphs()?
                        .into_iter()
                        .find(|graph| graph.path == *path)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!("RDF graph '{path}' was not found"))
                        })?
                        .triples
                        .len();
                    let (bytes, _) =
                        crate::rdf::add_triple(document.transaction_package(), path, triple)?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::ReplaceRdfTriple {
                    path,
                    index,
                    triple,
                } => {
                    let bytes = crate::rdf::replace_triple(
                        document.transaction_package(),
                        path,
                        *index,
                        triple,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::RemoveRdfTriple { path, index } => {
                    let bytes =
                        crate::rdf::remove_triple(document.transaction_package(), path, *index)?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::MoveRdfTriple { path, from, to } => {
                    let bytes =
                        crate::rdf::move_triple(document.transaction_package(), path, *from, *to)?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::AddForm { group_index, form } => {
                    OperationResult::Index(document.add_form(*group_index, form)?)
                },
                Operation::AddFormFragment {
                    group_index,
                    parent_form,
                    fragment,
                } => {
                    let (bytes, index) = crate::package::forms::add_form_fragment(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        crate::package::forms::FormHost::Text,
                        *group_index,
                        *parent_form,
                        fragment,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::AddNestedForm { parent_form, form } => {
                    let (bytes, index) = crate::package::forms::add_form(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        crate::package::forms::FormHost::Text,
                        0,
                        Some(*parent_form),
                        form,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::AddFormControl {
                    form_index,
                    control,
                } => {
                    let (bytes, index) = crate::package::forms::add_control(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *form_index,
                        control,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::AddFormControlFragment {
                    form_index,
                    fragment,
                } => {
                    let (bytes, index) = crate::package::forms::add_control_fragment(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *form_index,
                        fragment,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::ReplaceFormControl { index, control } => {
                    let bytes = crate::package::forms::replace_control(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *index,
                        control,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::ReplaceFormControlFragment { index, fragment } => {
                    let bytes = crate::package::forms::replace_control_fragment(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *index,
                        fragment,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::RemoveFormControl { index } => {
                    let bytes = crate::package::forms::remove_control(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *index,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::MoveFormControl { from, to } => {
                    let bytes = crate::package::forms::move_control(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *from,
                        *to,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::ReplaceForm { index, form } => {
                    document.replace_form(*index, form)?;
                    OperationResult::Unit
                },
                Operation::ReplaceFormFragment { index, fragment } => {
                    let bytes = crate::package::forms::replace_form_fragment(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *index,
                        fragment,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::RemoveForm { index } => {
                    document.remove_form(*index)?;
                    OperationResult::Unit
                },
                Operation::MoveForm { from, to } => {
                    document.move_form(*from, *to)?;
                    OperationResult::Unit
                },
                Operation::AddEmbeddedChart { definition } => {
                    OperationResult::Index(document.add_embedded_chart(definition)?)
                },
                Operation::AddEmbeddedChartWithStorage {
                    definition,
                    storage,
                } => {
                    let (bytes, index) = crate::package::charts::add_embedded_chart(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        crate::package::charts::EmbeddedChartHost::Text,
                        *storage,
                        definition,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::AddEmbeddedChartContent { content, storage } => {
                    let (bytes, index) = crate::package::charts::add_embedded_chart_content(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        crate::package::charts::EmbeddedChartHost::Text,
                        *storage,
                        content,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Index(index)
                },
                Operation::ReplaceEmbeddedChart { index, definition } => {
                    document.replace_embedded_chart(*index, definition)?;
                    OperationResult::Unit
                },
                Operation::ReplaceEmbeddedChartContent { index, content } => {
                    let bytes = crate::package::charts::replace_embedded_chart_content(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        *index,
                        content,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::RemoveEmbeddedChart { index } => {
                    document.remove_embedded_chart(*index)?;
                    OperationResult::Unit
                },
                Operation::AddEmbeddedResource { resource } => {
                    OperationResult::Index(document.add_embedded_resource(resource)?)
                },
                Operation::EmbeddedResourceBatch { changes } => {
                    let (bytes, indices) = crate::package::embedded::apply_batch(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        document.transaction_styles_xml(),
                        changes,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Indices(indices)
                },
                Operation::GarbageCollectEmbeddedResources { candidates } => {
                    let current = Snapshot::from_document(&document)?;
                    let plan = crate::package::resource_gc::plan(current, candidates)?;
                    if !plan.is_applicable() {
                        return Err(Error::Unsupported(
                            "ODT embedded-resource GC replay contains a typed refusal".to_string(),
                        ));
                    }
                    let bytes = plan.apply_to_package(document.transaction_package())?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::ReplaceEmbeddedObject { index, resource } => {
                    document.replace_embedded_object(*index, resource)?;
                    OperationResult::Unit
                },
                Operation::ReplaceEmbeddedImage { index, resource } => {
                    document.replace_embedded_image(*index, resource)?;
                    OperationResult::Unit
                },
                Operation::RemoveEmbeddedObject { index } => {
                    document.remove_embedded_object(*index)?;
                    OperationResult::Unit
                },
                Operation::RemoveEmbeddedImage { index } => {
                    document.remove_embedded_image(*index)?;
                    OperationResult::Unit
                },
                Operation::MoveEmbeddedObject { from, to } => {
                    document.move_embedded_object(*from, *to)?;
                    OperationResult::Unit
                },
                Operation::MoveEmbeddedImage { from, to } => {
                    document.move_embedded_image(*from, *to)?;
                    OperationResult::Unit
                },
                Operation::AddScriptResource { resource } => {
                    let (bytes, path) = crate::package::scripts::add_resource(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        resource,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Path(path)
                },
                Operation::ReplaceScriptResource { path, resource } => {
                    let bytes = crate::package::scripts::replace_resource(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        path,
                        resource,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
                Operation::RemoveScriptResource { path } => {
                    let bytes = crate::package::scripts::remove_resource(
                        document.transaction_package(),
                        document.transaction_content_xml(),
                        path,
                    )?;
                    document.replace_transaction_bytes(bytes)?;
                    OperationResult::Unit
                },
            };
            audit_changed_xml_is_compact(&before_operation, document.transaction_package())?;
            results.push(result);
            operation_index += 1;
        }
        let after = snapshot_from_final_document(&self.source, &document)?;
        // Reparse after the compact audit so commit never returns an invalid
        // candidate even when a writer implementation changes independently.
        after.document()?;
        Ok(Commit::new(self.source, after, results, self.operations))
    }
}

#[derive(Clone, PartialEq, Eq)]
struct Lineage(Arc<Vec<u8>>);

/// Deterministic refusal produced when independently prepared edits overlap.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct JoinConflict {
    failure: SubEditJoinFailure,
}

impl JoinConflict {
    /// Structured lineage, identifier, effect, or finite-bound refusal.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Deterministically ordered exact-effect conflicts, when overlap caused the refusal.
    #[must_use]
    pub fn conflicts(&self) -> Option<&ConflictSet<SubEditConflict>> {
        match &self.failure {
            SubEditJoinFailure::Overlap(conflicts) => Some(conflicts),
            _ => None,
        }
    }
}

/// Bounded collection of disjoint edits against one exact immutable snapshot.
pub struct JoinedEdit {
    source: Snapshot,
    limits: CompositionLimits,
    inner: JoinedSubEdits<Lineage, Vec<Operation>>,
}

impl JoinedEdit {
    /// Creates an empty composition under caller-selected finite bounds.
    #[must_use]
    pub fn new(source: Snapshot, limits: CompositionLimits) -> Self {
        let lineage = Lineage(source.bytes.clone());
        Self {
            source,
            limits,
            inner: JoinedSubEdits::new(lineage, limits),
        }
    }

    /// Joins one independently staged edit if lineage and semantic effects are disjoint.
    pub fn join(
        &mut self,
        identifier: impl Into<String>,
        edit: Edit,
    ) -> std::result::Result<&mut Self, JoinConflict> {
        let (reads, writes) = operation_effects(&edit.operations);
        let sub_edit = SubEdit::new(
            Lineage(edit.source.bytes.clone()),
            self.limits,
            identifier,
            reads,
            writes,
            edit.operations,
        )
        .map_err(|failure| JoinConflict {
            failure: SubEditJoinFailure::Limit(failure),
        })?;
        self.inner.join(sub_edit).map_err(|error| JoinConflict {
            failure: error.failure().clone(),
        })?;
        Ok(self)
    }

    /// Number of accepted independently prepared edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    /// Whether no independently prepared edit has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Atomically commits accepted work in stable identifier order.
    pub fn commit(self) -> Result<Commit> {
        let mut operations = Vec::new();
        for edit in self.inner.into_sub_edits() {
            operations.extend(edit.into_payload());
        }
        Edit {
            source: self.source,
            operations,
        }
        .commit()
    }
}

/// Non-applying three-way plan for two branches from one exact ODT snapshot.
pub struct MergePlan {
    source: Snapshot,
    limits: CompositionLimits,
    inner: ThreeWayMergePlan<Lineage, Vec<Operation>>,
}

impl MergePlan {
    /// Plans a merge without mutating or reopening the base snapshot.
    pub fn new(left: JoinedEdit, right: JoinedEdit) -> Result<Self> {
        if left.source.as_bytes() != right.source.as_bytes() {
            return Err(Error::InvalidFormat(
                "ODT merge branches do not share an exact base snapshot".to_string(),
            ));
        }
        let source = left.source.clone();
        let limits = left.limits;
        let inner = ThreeWayMergePlan::new(left.inner, right.inner).map_err(|error| {
            Error::InvalidFormat(format!(
                "ODT three-way planning failed: {:?}",
                error.failure()
            ))
        })?;
        Ok(Self {
            source,
            limits,
            inner,
        })
    }

    /// Automatically accepted disjoint edit count.
    #[must_use]
    pub fn automatic_len(&self) -> usize {
        self.inner.automatic().len()
    }

    /// Deterministically ordered unresolved conflicts.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet<SubEditConflict> {
        self.inner.conflicts()
    }

    /// Explicitly chooses the left, right, or neither conflicting branch.
    pub fn resolve(&mut self, choice: MergeChoice) -> &mut Self {
        self.inner.resolve(choice);
        self
    }

    /// Finishes only after every conflict has an explicit resolution.
    pub fn finish(self) -> std::result::Result<JoinedEdit, Box<Self>> {
        match self.inner.finish() {
            Ok(inner) => Ok(JoinedEdit {
                source: self.source,
                limits: self.limits,
                inner,
            }),
            Err(inner) => Err(Box::new(Self {
                source: self.source,
                limits: self.limits,
                inner: *inner,
            })),
        }
    }
}

/// Commit-coupled bounded undo/redo over exact immutable ODT snapshots.
pub struct History {
    inner: CoreHistory<Snapshot>,
}

impl History {
    /// Starts history at one immutable snapshot.
    #[must_use]
    pub fn new(current: Snapshot, limits: HistoryLimits) -> Self {
        Self {
            inner: CoreHistory::new(current, limits),
        }
    }

    /// Current exact immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Snapshot {
        self.inner.current()
    }

    /// Fully reopens the current snapshot for semantic inspection.
    pub fn document(&self) -> Result<Document> {
        self.current().document()
    }

    /// Starts an edit against the current history head.
    #[must_use]
    pub fn edit(&self) -> Edit {
        self.current().edit()
    }

    /// Commits and records an edit only when it targets the current exact head.
    pub fn commit(&mut self, edit: Edit) -> Result<Commit> {
        if edit.source.as_bytes() != self.current().as_bytes() {
            return Err(stale_history_error());
        }
        let commit = edit.commit()?;
        self.record(commit.snapshot().clone())?;
        Ok(commit)
    }

    /// Applies and records a durable patch only at the current exact head.
    pub fn apply(&mut self, patch: &DurablePatch) -> Result<Snapshot> {
        let snapshot = patch.apply(self.current())?;
        self.record(snapshot.clone())?;
        Ok(snapshot)
    }

    /// Moves to the exact inverse snapshot when retained.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Reapplies one retained exact snapshot when available.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }

    /// Whether an inverse snapshot is retained.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Whether a forward snapshot is retained.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    fn record(&mut self, snapshot: Snapshot) -> Result<()> {
        let weight = u64::try_from(self.current().as_bytes().len()).map_err(|_error| {
            Error::InvalidFormat("ODT history snapshot weight exceeds u64".to_string())
        })?;
        self.inner
            .record(snapshot, weight)
            .map_err(durable_wire_error)?;
        Ok(())
    }
}

#[derive(Clone)]
enum Operation {
    Noop,
    RestoreSnapshot,
    InsertParagraph {
        index: usize,
        text: String,
    },
    ReplaceParagraph {
        index: usize,
        text: String,
    },
    RemoveParagraph {
        index: usize,
    },
    AppendRun {
        paragraph: usize,
        text: String,
        style_name: Option<String>,
    },
    AppendHyperlink {
        paragraph: usize,
        href: String,
        text: String,
    },
    AppendLineBreak {
        index: usize,
    },
    InsertNote {
        paragraph: usize,
        note: crate::note::Note,
        fragment: Option<String>,
    },
    ReplaceNote {
        index: usize,
        note: crate::note::Note,
        fragment: Option<String>,
    },
    RemoveNote {
        index: usize,
    },
    InsertRubyAnnotation {
        paragraph: usize,
        annotation: crate::ruby_family::Annotation,
    },
    ReplaceRubyAnnotation {
        index: usize,
        annotation: crate::ruby_family::Annotation,
    },
    RemoveRubyAnnotation {
        index: usize,
    },
    InsertDynamicTextField {
        paragraph: usize,
        field: DynamicTextField,
    },
    ReplaceDynamicTextField {
        index: usize,
        field: DynamicTextField,
    },
    RemoveDynamicTextField {
        index: usize,
    },
    SetRubyStyle {
        style: crate::ruby_family::Style,
    },
    RemoveRubyStyle {
        name: String,
    },
    SetTrackedChangePolicy {
        track_changes: Option<bool>,
        protection_key: Option<String>,
        digest_algorithm: Option<String>,
    },
    RemoveTrackedChange {
        id: String,
    },
    AddRdfGraph {
        preferred_path: Option<String>,
        triples: Vec<crate::rdf::Triple>,
    },
    ReplaceRdfGraph {
        path: String,
        triples: Vec<crate::rdf::Triple>,
    },
    RemoveRdfGraph {
        path: String,
    },
    SetProtection {
        policy: crate::protection::Policy,
    },
    AddRdfTriple {
        path: String,
        triple: crate::rdf::Triple,
    },
    ReplaceRdfTriple {
        path: String,
        index: usize,
        triple: crate::rdf::Triple,
    },
    RemoveRdfTriple {
        path: String,
        index: usize,
    },
    MoveRdfTriple {
        path: String,
        from: usize,
        to: usize,
    },
    AddForm {
        group_index: usize,
        form: crate::package::forms::AuthoredForm,
    },
    AddFormFragment {
        group_index: usize,
        parent_form: Option<usize>,
        fragment: String,
    },
    AddNestedForm {
        parent_form: usize,
        form: crate::package::forms::AuthoredForm,
    },
    AddFormControl {
        form_index: usize,
        control: crate::package::forms::AuthoredFormControl,
    },
    AddFormControlFragment {
        form_index: usize,
        fragment: String,
    },
    ReplaceFormControl {
        index: usize,
        control: crate::package::forms::AuthoredFormControl,
    },
    ReplaceFormControlFragment {
        index: usize,
        fragment: String,
    },
    RemoveFormControl {
        index: usize,
    },
    MoveFormControl {
        from: usize,
        to: usize,
    },
    ReplaceForm {
        index: usize,
        form: crate::package::forms::AuthoredForm,
    },
    ReplaceFormFragment {
        index: usize,
        fragment: String,
    },
    RemoveForm {
        index: usize,
    },
    MoveForm {
        from: usize,
        to: usize,
    },
    AddEmbeddedChart {
        definition: crate::odc::Definition,
    },
    AddEmbeddedChartWithStorage {
        definition: crate::odc::Definition,
        storage: crate::package::charts::EmbeddedChartStorage,
    },
    AddEmbeddedChartContent {
        content: String,
        storage: crate::package::charts::EmbeddedChartStorage,
    },
    ReplaceEmbeddedChart {
        index: usize,
        definition: crate::odc::Definition,
    },
    ReplaceEmbeddedChartContent {
        index: usize,
        content: String,
    },
    RemoveEmbeddedChart {
        index: usize,
    },
    AddEmbeddedResource {
        resource: crate::package::embedded::EmbeddedResource,
    },
    EmbeddedResourceBatch {
        changes: Vec<crate::package::embedded::EmbeddedResourceChange>,
    },
    GarbageCollectEmbeddedResources {
        candidates: Vec<crate::package::resource_gc::EmbeddedResourceGcCandidate>,
    },
    ReplaceEmbeddedObject {
        index: usize,
        resource: crate::package::embedded::EmbeddedResource,
    },
    ReplaceEmbeddedImage {
        index: usize,
        resource: crate::package::embedded::EmbeddedResource,
    },
    RemoveEmbeddedObject {
        index: usize,
    },
    RemoveEmbeddedImage {
        index: usize,
    },
    MoveEmbeddedObject {
        from: usize,
        to: usize,
    },
    MoveEmbeddedImage {
        from: usize,
        to: usize,
    },
    AddScriptResource {
        resource: crate::ScriptResourceSpec,
    },
    ReplaceScriptResource {
        path: String,
        resource: crate::ScriptResourceSpec,
    },
    RemoveScriptResource {
        path: String,
    },
}

/// A validated packaged-ODT transaction result.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    results: Vec<OperationResult>,
}

impl Commit {
    fn new(
        before: Snapshot,
        snapshot: Snapshot,
        results: Vec<OperationResult>,
        operations: Vec<Operation>,
    ) -> Self {
        Self {
            patch: Patch {
                before,
                after: snapshot.clone(),
                operations,
            },
            snapshot,
            results,
        }
    }

    /// Returns the committed immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns one deterministic result for each staged operation.
    #[must_use]
    pub fn results(&self) -> &[OperationResult] {
        &self.results
    }

    /// Consumes the commit and returns its immutable snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

pub(crate) fn embedded_resource_gc_commit(
    before: Snapshot,
    snapshot: Snapshot,
    candidates: Vec<crate::package::resource_gc::EmbeddedResourceGcCandidate>,
) -> Commit {
    let operations = if candidates.is_empty() {
        Vec::new()
    } else {
        vec![Operation::GarbageCollectEmbeddedResources { candidates }]
    };
    Commit::new(before, snapshot, vec![OperationResult::Unit], operations)
}

/// The typed result produced by one staged operation.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OperationResult {
    /// Mutation completed without allocating an identity.
    Unit,
    /// Semantic index allocated by an insertion operation.
    Index(usize),
    /// Semantic indexes allocated by insertions in one bounded batch.
    Indices(Vec<usize>),
    /// Package path allocated by a resource insertion operation.
    Path(String),
}

/// Exact-byte, source-checked package patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    operations: Vec<Operation>,
}

impl Patch {
    /// Applies this patch only to the exact snapshot from which it was made.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.as_bytes() != self.before.as_bytes() {
            return Err(Error::InvalidFormat(
                "ODT patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.after.clone())
    }

    /// Returns the patch that restores the exact source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            operations: vec![Operation::RestoreSnapshot],
        }
    }

    /// Converts this exact in-memory patch into a bounded durable patch.
    ///
    /// The durable form uses deterministic JSON and content-addressed package
    /// blobs. The complete source artifact is retained only while the patch is
    /// reversible; [`DurablePatch::seal`] permanently discards it.
    ///
    /// # Errors
    ///
    /// Returns an error if either package exceeds the durable patch bounds or
    /// the format-neutral wire envelope cannot be constructed.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::from_semantic_patch(self)
    }

    /// Builds a non-mutating, dependency-checked plan for replaying portable
    /// additions and policy updates onto another package snapshot.
    ///
    /// Form, chart, and embedded-resource positions are checked against the
    /// destination before their replacements, removals, or moves may commit.
    /// Other source-local selectors and snapshot restores remain refused.
    pub fn plan_transfer(&self, destination: &Snapshot) -> Result<TransferPlan> {
        TransferPlan::new(destination.clone(), &self.operations)
    }
}

/// Kind of prerequisite checked by a cross-document transfer plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum TransferDependencyKind {
    /// A paragraph must exist at the checked semantic position.
    Paragraph,
    /// A named ruby parent style must exist or be transferred in the same plan.
    RubyStyle,
    /// A named text style referenced by rich inline content.
    TextStyle,
    /// A chart-local style reference retained inside the transferred payload.
    ChartStyle,
    /// A content-addressed embedded payload travels with the operation.
    ResourcePayload,
    /// A form group must exist, except that group zero may be created on demand.
    FormGroup,
    /// A form must exist at the selected semantic position.
    Form,
    /// A form control must exist at the selected semantic position.
    FormControl,
    /// A validated embedded chart must exist at the selected position.
    EmbeddedChart,
    /// An embedded object must exist at the selected position.
    EmbeddedObject,
    /// An embedded image must exist at the selected position.
    EmbeddedImage,
}

/// One deterministic prerequisite reported by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferDependency {
    kind: TransferDependencyKind,
    key: String,
    satisfied: bool,
}

impl TransferDependency {
    /// Returns the prerequisite family.
    #[must_use]
    pub const fn kind(&self) -> TransferDependencyKind {
        self.kind
    }

    /// Returns its stable semantic key.
    #[must_use]
    pub fn key(&self) -> &str {
        &self.key
    }

    /// Whether the destination or this plan supplies the prerequisite.
    #[must_use]
    pub const fn is_satisfied(&self) -> bool {
        self.satisfied
    }
}

/// Non-mutating, bounded plan for dependency-aware cross-document replay.
pub struct TransferPlan {
    destination: Snapshot,
    operations: Vec<Operation>,
    dependencies: Vec<TransferDependency>,
}

impl TransferPlan {
    fn new(destination: Snapshot, operations: &[Operation]) -> Result<Self> {
        if operations.len() > MAX_OPERATIONS {
            return Err(invalid_durable_patch());
        }
        let document = destination.document()?;
        let mut paragraph_count = document.paragraphs()?.len();
        let destination_forms = document.forms()?;
        let form_group_count = destination_forms.groups.len();
        let (mut form_count, mut form_control_count) = form_counts(&destination_forms);
        let embedded_objects = document.embedded_objects()?;
        let mut embedded_chart_positions: BTreeSet<usize> = embedded_objects
            .iter()
            .enumerate()
            .filter_map(|(index, _)| document.embedded_chart(index).is_ok().then_some(index))
            .collect();
        let mut next_embedded_object_position = embedded_objects
            .iter()
            .filter(|object| object.part == crate::Part::Content)
            .count();
        let mut embedded_object_count = embedded_objects
            .into_iter()
            .filter(|object| {
                object.part == crate::Part::Content
                    && matches!(
                        object.kind,
                        litchi_odf_common::embedded::Kind::Object
                            | litchi_odf_common::embedded::Kind::ObjectOle
                    )
            })
            .count();
        let mut embedded_image_count = document
            .images()?
            .into_iter()
            .filter(|image| image.part == crate::Part::Content)
            .count();
        let mut ruby_styles: BTreeSet<String> = document
            .ruby_styles()?
            .styles
            .into_iter()
            .map(|style| style.name)
            .collect();
        for operation in operations {
            if let Operation::SetRubyStyle { style } = operation {
                ruby_styles.insert(style.name.clone());
            }
        }

        let mut dependencies = BTreeMap::new();
        for operation in operations {
            match operation {
                Operation::Noop
                | Operation::AddRdfGraph { .. }
                | Operation::SetProtection { .. }
                | Operation::SetTrackedChangePolicy { .. }
                | Operation::AddScriptResource { .. } => {},
                Operation::AddEmbeddedResource { resource } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::ResourcePayload,
                        embedded_resource_dependency_key(resource),
                        true,
                    );
                    if resource.kind == crate::package::embedded::EmbeddedResourceKind::Image {
                        embedded_image_count = embedded_image_count.saturating_add(1);
                    } else {
                        embedded_object_count = embedded_object_count.saturating_add(1);
                        next_embedded_object_position =
                            next_embedded_object_position.saturating_add(1);
                    }
                },
                Operation::AddEmbeddedChart { definition }
                | Operation::AddEmbeddedChartWithStorage { definition, .. } => {
                    for name in chart_style_names(definition) {
                        merge_transfer_dependency(
                            &mut dependencies,
                            TransferDependencyKind::ChartStyle,
                            name,
                            true,
                        );
                    }
                    embedded_chart_positions.insert(next_embedded_object_position);
                    next_embedded_object_position = next_embedded_object_position.saturating_add(1);
                    embedded_object_count = embedded_object_count.saturating_add(1);
                },
                Operation::InsertParagraph { index, .. } => {
                    let satisfied = *index <= paragraph_count;
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Paragraph,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        paragraph_count = paragraph_count.saturating_add(1);
                    }
                },
                Operation::AppendRun {
                    paragraph,
                    style_name: None,
                    ..
                }
                | Operation::AppendHyperlink { paragraph, .. }
                | Operation::AppendLineBreak { index: paragraph }
                | Operation::InsertDynamicTextField { paragraph, .. }
                | Operation::InsertNote { paragraph, .. } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Paragraph,
                        paragraph.to_string(),
                        *paragraph < paragraph_count,
                    );
                },
                Operation::InsertRubyAnnotation {
                    paragraph,
                    annotation,
                } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Paragraph,
                        paragraph.to_string(),
                        *paragraph < paragraph_count,
                    );
                    if let Some(style_name) = &annotation.style_name {
                        merge_transfer_dependency(
                            &mut dependencies,
                            TransferDependencyKind::RubyStyle,
                            style_name.clone(),
                            ruby_styles.contains(style_name),
                        );
                    }
                    if let Some(style_name) = &annotation.text_style_name {
                        merge_transfer_dependency(
                            &mut dependencies,
                            TransferDependencyKind::TextStyle,
                            style_name.clone(),
                            false,
                        );
                    }
                },
                Operation::SetRubyStyle { style } => {
                    if let Some(parent) = &style.parent_style_name {
                        merge_transfer_dependency(
                            &mut dependencies,
                            TransferDependencyKind::RubyStyle,
                            parent.clone(),
                            ruby_styles.contains(parent),
                        );
                    }
                },
                Operation::AddForm { group_index, form } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::FormGroup,
                        group_index.to_string(),
                        *group_index < form_group_count
                            || (*group_index == 0 && form_group_count == 0),
                    );
                    let (forms, controls) = authored_form_counts(form);
                    form_count = form_count.saturating_add(forms);
                    form_control_count = form_control_count.saturating_add(controls);
                },
                Operation::AddNestedForm { parent_form, form } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Form,
                        parent_form.to_string(),
                        *parent_form < form_count,
                    );
                    let (forms, controls) = authored_form_counts(form);
                    form_count = form_count.saturating_add(forms);
                    form_control_count = form_control_count.saturating_add(controls);
                },
                Operation::AddFormControl { form_index, .. } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Form,
                        form_index.to_string(),
                        *form_index < form_count,
                    );
                    form_control_count = form_control_count.saturating_add(1);
                },
                Operation::ReplaceFormControl { index, .. } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::FormControl,
                        index.to_string(),
                        *index < form_control_count,
                    );
                },
                Operation::RemoveFormControl { index } => {
                    let satisfied = *index < form_control_count;
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::FormControl,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        form_control_count = form_control_count.saturating_sub(1);
                    }
                },
                Operation::MoveFormControl { from, to } => {
                    merge_position_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::FormControl,
                        *from,
                        *to,
                        form_control_count,
                    );
                },
                Operation::ReplaceForm { index, .. } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Form,
                        index.to_string(),
                        *index < form_count,
                    );
                },
                Operation::RemoveForm { index } => {
                    let satisfied = *index < form_count;
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::Form,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        form_count = form_count.saturating_sub(1);
                    }
                },
                Operation::MoveForm { from, to } => {
                    merge_position_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::Form,
                        *from,
                        *to,
                        form_count,
                    );
                },
                Operation::ReplaceEmbeddedChart { index, definition } => {
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedChart,
                        index.to_string(),
                        embedded_chart_positions.contains(index),
                    );
                    for name in chart_style_names(definition) {
                        merge_transfer_dependency(
                            &mut dependencies,
                            TransferDependencyKind::ChartStyle,
                            name,
                            true,
                        );
                    }
                },
                Operation::RemoveEmbeddedChart { index } => {
                    let satisfied = embedded_chart_positions.contains(index);
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedChart,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        embedded_chart_positions.remove(index);
                        embedded_chart_positions = embedded_chart_positions
                            .into_iter()
                            .map(|position| {
                                if position > *index {
                                    position - 1
                                } else {
                                    position
                                }
                            })
                            .collect();
                        next_embedded_object_position =
                            next_embedded_object_position.saturating_sub(1);
                        embedded_object_count = embedded_object_count.saturating_sub(1);
                    }
                },
                Operation::ReplaceEmbeddedObject { index, resource } => {
                    embedded_resource_transfer_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedObject,
                        *index,
                        embedded_object_count,
                        resource,
                    );
                },
                Operation::ReplaceEmbeddedImage { index, resource } => {
                    embedded_resource_transfer_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedImage,
                        *index,
                        embedded_image_count,
                        resource,
                    );
                },
                Operation::RemoveEmbeddedObject { index } => {
                    let satisfied = *index < embedded_object_count;
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedObject,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        embedded_object_count = embedded_object_count.saturating_sub(1);
                        next_embedded_object_position =
                            next_embedded_object_position.saturating_sub(1);
                        embedded_chart_positions.remove(index);
                        embedded_chart_positions = embedded_chart_positions
                            .into_iter()
                            .map(|position| {
                                if position > *index {
                                    position - 1
                                } else {
                                    position
                                }
                            })
                            .collect();
                    }
                },
                Operation::RemoveEmbeddedImage { index } => {
                    let satisfied = *index < embedded_image_count;
                    merge_transfer_dependency(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedImage,
                        index.to_string(),
                        satisfied,
                    );
                    if satisfied {
                        embedded_image_count = embedded_image_count.saturating_sub(1);
                    }
                },
                Operation::MoveEmbeddedObject { from, to } => {
                    merge_position_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedObject,
                        *from,
                        *to,
                        embedded_object_count,
                    );
                    if *from < embedded_object_count && *to < embedded_object_count {
                        embedded_chart_positions = embedded_chart_positions
                            .into_iter()
                            .map(|position| relocate_position(position, *from, *to))
                            .collect();
                    }
                },
                Operation::MoveEmbeddedImage { from, to } => {
                    merge_position_dependencies(
                        &mut dependencies,
                        TransferDependencyKind::EmbeddedImage,
                        *from,
                        *to,
                        embedded_image_count,
                    );
                },
                _ => {
                    return Err(Error::Unsupported(
                        "ODT cross-document transfer refuses this source-local operation"
                            .to_string(),
                    ));
                },
            }
        }
        let dependencies = dependencies
            .into_iter()
            .map(|((kind, key), satisfied)| TransferDependency {
                kind,
                key,
                satisfied,
            })
            .collect();
        Ok(Self {
            destination,
            operations: operations.to_vec(),
            dependencies,
        })
    }

    /// Returns prerequisites in stable kind/key order.
    #[must_use]
    pub fn dependencies(&self) -> &[TransferDependency] {
        &self.dependencies
    }

    /// Returns the number of semantic operations retained by this plan.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.operations.len()
    }

    /// Applies the planned edit only when every dependency is satisfied.
    pub fn commit(self) -> Result<Commit> {
        if self
            .dependencies
            .iter()
            .any(|dependency| !dependency.satisfied)
        {
            return Err(Error::InvalidFormat(
                "ODT transfer plan has unresolved dependencies".to_string(),
            ));
        }
        Edit {
            source: self.destination,
            operations: self.operations,
        }
        .commit()
    }
}

fn merge_transfer_dependency(
    dependencies: &mut BTreeMap<(TransferDependencyKind, String), bool>,
    kind: TransferDependencyKind,
    key: String,
    satisfied: bool,
) {
    dependencies
        .entry((kind, key))
        .and_modify(|existing| *existing &= satisfied)
        .or_insert(satisfied);
}

fn merge_position_dependencies(
    dependencies: &mut BTreeMap<(TransferDependencyKind, String), bool>,
    kind: TransferDependencyKind,
    from: usize,
    to: usize,
    count: usize,
) {
    for position in [from, to] {
        merge_transfer_dependency(dependencies, kind, position.to_string(), position < count);
    }
}

fn relocate_position(position: usize, from: usize, to: usize) -> usize {
    if position == from {
        to
    } else if from < to && position > from && position <= to {
        position - 1
    } else if from > to && position >= to && position < from {
        position.saturating_add(1)
    } else {
        position
    }
}

fn form_counts(forms: &crate::form::Forms) -> (usize, usize) {
    forms.groups.iter().flat_map(|group| &group.forms).fold(
        (0usize, 0usize),
        |(forms, controls), form| {
            let (nested_forms, nested_controls) = parsed_form_counts(form);
            (
                forms.saturating_add(nested_forms),
                controls.saturating_add(nested_controls),
            )
        },
    )
}

fn parsed_form_counts(form: &crate::form::Form) -> (usize, usize) {
    form.children
        .iter()
        .fold((1usize, 0usize), |(forms, controls), node| match node {
            crate::form::Node::Form(form) => {
                let (nested_forms, nested_controls) = parsed_form_counts(form);
                (
                    forms.saturating_add(nested_forms),
                    controls.saturating_add(nested_controls),
                )
            },
            crate::form::Node::Control(_) => (forms, controls.saturating_add(1)),
        })
}

fn authored_form_counts(form: &crate::package::forms::AuthoredForm) -> (usize, usize) {
    form.children
        .iter()
        .fold((1usize, 0usize), |(forms, controls), node| match node {
            crate::package::forms::AuthoredFormNode::Form(form) => {
                let (nested_forms, nested_controls) = authored_form_counts(form);
                (
                    forms.saturating_add(nested_forms),
                    controls.saturating_add(nested_controls),
                )
            },
            crate::package::forms::AuthoredFormNode::Control(_) => {
                (forms, controls.saturating_add(1))
            },
        })
}

fn embedded_resource_transfer_dependencies(
    dependencies: &mut BTreeMap<(TransferDependencyKind, String), bool>,
    target_kind: TransferDependencyKind,
    index: usize,
    count: usize,
    resource: &crate::package::embedded::EmbeddedResource,
) {
    merge_transfer_dependency(dependencies, target_kind, index.to_string(), index < count);
    merge_transfer_dependency(
        dependencies,
        TransferDependencyKind::ResourcePayload,
        embedded_resource_dependency_key(resource),
        true,
    );
}

fn embedded_resource_dependency_key(
    resource: &crate::package::embedded::EmbeddedResource,
) -> String {
    use crate::package::embedded::EmbeddedResourceSource;
    let mut material = Vec::new();
    match &resource.source {
        EmbeddedResourceSource::Linked { href } => material.extend_from_slice(href.as_bytes()),
        EmbeddedResourceSource::PackageFile { bytes, .. }
        | EmbeddedResourceSource::InlineBinary { bytes, .. } => material.extend_from_slice(bytes),
        EmbeddedResourceSource::PackageSubdocument { files, .. } => {
            for file in files {
                material.extend_from_slice(file.path.as_bytes());
                material.push(0);
                material.extend_from_slice(&file.bytes);
                material.push(0xff);
            }
        },
        EmbeddedResourceSource::InlineXml { xml, .. } => material.extend_from_slice(xml.as_bytes()),
    }
    BlobId::of(&material).as_hex()
}

fn chart_style_names(definition: &crate::odc::Definition) -> BTreeSet<String> {
    let mut names = BTreeSet::new();
    let mut add = |name: &Option<String>| {
        if let Some(name) = name {
            names.insert(name.clone());
        }
    };
    add(&definition.style_name);
    for text in [&definition.title, &definition.subtitle, &definition.footer]
        .into_iter()
        .flatten()
    {
        add(&text.style_name);
    }
    if let Some(legend) = &definition.legend {
        add(&legend.style_name);
    }
    let plot = &definition.plot_area;
    add(&plot.style_name);
    for axis in &plot.axes {
        add(&axis.style_name);
        if let Some(title) = &axis.title {
            add(&title.style_name);
        }
        for grid in &axis.grids {
            add(&grid.style_name);
        }
    }
    for series in &plot.series {
        add(&series.style_name);
        if let Some(label) = &series.data_label {
            add(&label.style_name);
        }
        for point in &series.data_points {
            add(&point.style_name);
            if let Some(label) = &point.label {
                add(&label.style_name);
            }
        }
        for style in [&series.mean_value, &series.error_indicator]
            .into_iter()
            .flatten()
        {
            add(&style.style_name);
        }
        for regression in &series.regression_curves {
            add(&regression.style_name);
            if let Some(equation) = &regression.equation {
                add(&equation.style_name);
            }
        }
    }
    for style in [
        &plot.wall,
        &plot.floor,
        &plot.stock_gain_marker,
        &plot.stock_loss_marker,
        &plot.stock_range_line,
    ]
    .into_iter()
    .flatten()
    {
        add(&style.style_name);
    }
    names
}

/// Bounded deterministic-JSON patch for cross-process ODT exchange.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    fn from_semantic_patch(patch: &Patch) -> Result<Self> {
        // Durable patches intentionally do not retain credentials. This
        // credential-free reopen makes the encrypted-package refusal explicit.
        Snapshot::from_bytes(copy_bytes(patch.before.as_bytes())?)?;
        Snapshot::from_bytes(copy_bytes(patch.after.as_bytes())?)?;
        let limits = durable_patch_limits();
        let blob_limits = limits.blobs();
        let mut forward_blobs = BlobBundle::new(blob_limits);
        let mut reverse_blobs = BlobBundle::new(blob_limits);
        let source_id = BlobId::of(patch.before.as_bytes());
        let target_id = BlobId::of(patch.after.as_bytes());
        let source_blob = reverse_blobs
            .insert(patch.before.as_bytes())
            .map_err(durable_wire_error)?;

        let mut operations = patch.operations.clone();
        if operations.is_empty() {
            operations.push(Operation::Noop);
        }
        let restoring = operations
            .iter()
            .all(|operation| matches!(operation, Operation::RestoreSnapshot));
        if restoring {
            let target_blob = forward_blobs
                .insert(patch.after.as_bytes())
                .map_err(durable_wire_error)?;
            let forward = restore_patch_operation(limits, &source_id, &target_id, &target_blob)?;
            let inverse = restore_patch_operation(limits, &target_id, &source_id, &source_blob)?;
            let inner = CorePatch::<Reversible>::new(
                limits,
                DURABLE_FORMAT,
                [ReversibleOperation::new(forward, inverse)],
                forward_blobs,
                reverse_blobs,
            )
            .map_err(durable_wire_error)?;
            return Ok(Self { inner });
        }

        let mut pairs = Vec::new();
        pairs
            .try_reserve_exact(operations.len())
            .map_err(|allocation_error| Error::Allocation {
                resource: "ODT durable semantic operations",
                source: allocation_error,
            })?;
        for operation in &operations {
            let forward = semantic_patch_operation(
                limits,
                &source_id,
                &target_id,
                operation,
                &mut forward_blobs,
            )?;
            let inverse = restore_patch_operation(limits, &target_id, &source_id, &source_blob)?;
            pairs.push(ReversibleOperation::new(forward, inverse));
        }
        let inner = CorePatch::<Reversible>::new(
            limits,
            DURABLE_FORMAT,
            pairs,
            forward_blobs,
            reverse_blobs,
        )
        .map_err(durable_wire_error)?;
        Ok(Self { inner })
    }

    /// Parses a canonical durable ODT patch under the crate's finite bounds.
    ///
    /// The retained inverse artifact and every semantic operation are validated
    /// before the patch is published. The target is fully reopened on apply.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical JSON, exceeded limits, invalid blob
    /// integrity, a foreign operation vocabulary, or an invalid ODT artifact.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_patch_limits())
            .map_err(durable_wire_error)?;
        validate_reversible_patch(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes this patch as canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if the bounded JSON output cannot be produced.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner
            .to_deterministic_json()
            .map_err(durable_wire_error)
    }

    /// Applies this patch only to its exact source package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if `source` is not byte-identical to the retained
    /// source artifact or the target package no longer validates as ODT.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        let inverse = self.inner.inverse();
        if let Some(expected_source) = restore_target_bytes(&inverse)?
            && source.as_bytes() != expected_source
        {
            return Err(durable_source_mismatch());
        }
        apply_durable_patch(&self.inner, source)
    }

    /// Returns the durable patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Permanently drops the reverse operation and source package bytes.
    #[must_use]
    pub fn seal(self) -> SealedPatch {
        SealedPatch {
            inner: self.inner.seal(),
        }
    }
}

/// Forward-only durable ODT patch with no retained inverse package.
#[derive(Clone)]
pub struct SealedPatch {
    inner: CorePatch<ForwardOnly>,
}

impl SealedPatch {
    /// Parses a canonical forward-only ODT patch under finite bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical JSON, exceeded limits, invalid blob
    /// integrity, a foreign operation vocabulary, or an invalid target ODT.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner =
            CorePatch::<ForwardOnly>::from_deterministic_json(bytes, durable_patch_limits())
                .map_err(durable_wire_error)?;
        validate_sealed_patch(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes this forward-only patch as canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if the bounded JSON output cannot be produced.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner
            .to_deterministic_json()
            .map_err(durable_wire_error)
    }

    /// Applies the sealed patch after checking the source SHA-256 precondition.
    ///
    /// Exact source bytes are intentionally unavailable after sealing; the
    /// cryptographic precondition is the remaining conflict authorization.
    ///
    /// # Errors
    ///
    /// Returns an error if the source digest differs or the target package no
    /// longer validates as ODT.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        apply_durable_patch(&self.inner, source)
    }
}

struct DurableLineage<'a> {
    source_id: &'a str,
    target_id: &'a str,
}

fn durable_patch_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(
            MAX_OPERATIONS.saturating_add(1),
            MAX_PACKAGE_BYTES,
            MAX_WIRE_JSON_BYTES,
        ),
        MAX_WIRE_JSON_BYTES,
        MAX_OPERATIONS,
        6,
        MAX_SEMANTIC_TEXT_BYTES,
        8 * 1024 * 1024,
    )
}

fn durable_preconditions(source: &BlobId, target: &BlobId) -> BTreeMap<String, Value> {
    BTreeMap::from([
        (
            SOURCE_PRECONDITION.to_string(),
            Value::String(source.as_hex()),
        ),
        (
            TARGET_PRECONDITION.to_string(),
            Value::String(target.as_hex()),
        ),
    ])
}

fn restore_patch_operation(
    limits: PatchLimits,
    source: &BlobId,
    target: &BlobId,
    target_blob: &BlobId,
) -> Result<PatchOperation> {
    PatchOperation::new(
        limits,
        RESTORE_OPERATION,
        "/",
        durable_preconditions(source, target),
        Value::String(target_blob.as_hex()),
    )
    .map_err(durable_wire_error)
}

fn semantic_patch_operation(
    limits: PatchLimits,
    source: &BlobId,
    target: &BlobId,
    operation: &Operation,
    blobs: &mut BlobBundle,
) -> Result<PatchOperation> {
    let (name, path, value) = match operation {
        Operation::Noop => (NOOP_OPERATION, "/".to_string(), Value::Null),
        Operation::InsertParagraph { index, text } => (
            "paragraph.insert",
            format!("/body/paragraphs/{index}"),
            serde_json::json!({"text": text}),
        ),
        Operation::ReplaceParagraph { index, text } => (
            "paragraph.replace",
            format!("/body/paragraphs/{index}"),
            serde_json::json!({"text": text}),
        ),
        Operation::RemoveParagraph { index } => (
            "paragraph.remove",
            format!("/body/paragraphs/{index}"),
            Value::Null,
        ),
        Operation::AppendRun {
            paragraph,
            text,
            style_name,
        } => (
            "run.append",
            format!("/body/paragraphs/{paragraph}/runs/-"),
            serde_json::json!({"style_name": style_name, "text": text}),
        ),
        Operation::AppendHyperlink {
            paragraph,
            href,
            text,
        } => (
            "hyperlink.append",
            format!("/body/paragraphs/{paragraph}/hyperlinks/-"),
            serde_json::json!({"href": href, "text": text}),
        ),
        Operation::AppendLineBreak { index } => (
            "run.append_line_break",
            format!("/body/paragraphs/{index}/runs/-"),
            Value::Null,
        ),
        Operation::InsertNote {
            paragraph,
            note,
            fragment,
        } => (
            "note.insert",
            format!("/body/paragraphs/{paragraph}/notes/-"),
            note_fragment_value(note, fragment.as_deref())?,
        ),
        Operation::ReplaceNote {
            index,
            note,
            fragment,
        } => (
            "note.replace",
            format!("/body/notes/{index}"),
            note_fragment_value(note, fragment.as_deref())?,
        ),
        Operation::RemoveNote { index } => {
            ("note.remove", format!("/body/notes/{index}"), Value::Null)
        },
        Operation::InsertRubyAnnotation {
            paragraph,
            annotation,
        } => (
            "ruby.annotation.insert",
            format!("/body/paragraphs/{paragraph}/ruby/-"),
            serde_json::json!({"xml": annotation.to_xml_fragment()?}),
        ),
        Operation::ReplaceRubyAnnotation { index, annotation } => (
            "ruby.annotation.replace",
            format!("/body/ruby/{index}"),
            serde_json::json!({"xml": annotation.to_xml_fragment()?}),
        ),
        Operation::RemoveRubyAnnotation { index } => (
            "ruby.annotation.remove",
            format!("/body/ruby/{index}"),
            Value::Null,
        ),
        Operation::InsertDynamicTextField { paragraph, field } => (
            "field.dynamic.insert",
            format!("/body/paragraphs/{paragraph}/fields/-"),
            serde_json::json!({"xml": field.to_xml_fragment()?}),
        ),
        Operation::ReplaceDynamicTextField { index, field } => (
            "field.dynamic.replace",
            format!("/body/fields/{index}"),
            serde_json::json!({"xml": field.to_xml_fragment()?}),
        ),
        Operation::RemoveDynamicTextField { index } => (
            "field.dynamic.remove",
            format!("/body/fields/{index}"),
            Value::Null,
        ),
        Operation::SetRubyStyle { style } => (
            "style.ruby.set",
            format!("/styles/ruby/{}", style.name),
            serde_json::json!({"xml": style.to_xml_fragment()?}),
        ),
        Operation::RemoveRubyStyle { name } => (
            "style.ruby.remove",
            format!("/styles/ruby/{name}"),
            Value::Null,
        ),
        Operation::SetTrackedChangePolicy {
            track_changes,
            protection_key,
            digest_algorithm,
        } => (
            "revision.policy.set",
            "/body/tracked-changes/policy".to_string(),
            serde_json::json!({
                "digest_algorithm": digest_algorithm,
                "protection_key": protection_key,
                "track_changes": track_changes,
            }),
        ),
        Operation::RemoveTrackedChange { id } => (
            "revision.remove",
            format!("/body/tracked-changes/{id}"),
            Value::Null,
        ),
        Operation::AddRdfGraph {
            preferred_path,
            triples,
        } => (
            "rdf.graph.add",
            "/package/rdf/-".to_string(),
            serde_json::json!({
                "preferred_path": preferred_path,
                "triples": rdf_triples_value(triples)?,
            }),
        ),
        Operation::ReplaceRdfGraph { path, triples } => (
            "rdf.graph.replace",
            format!("/package/rdf/{path}"),
            serde_json::json!({"triples": rdf_triples_value(triples)?}),
        ),
        Operation::RemoveRdfGraph { path } => (
            "rdf.graph.remove",
            format!("/package/rdf/{path}"),
            Value::Null,
        ),
        Operation::SetProtection { policy } => (
            "protection.set",
            "/package/protection".to_string(),
            protection_value(policy),
        ),
        Operation::AddRdfTriple { path, triple } => (
            "rdf.triple.add",
            format!("/package/rdf/{path}/triples/-"),
            rdf_triple_value(triple)?,
        ),
        Operation::ReplaceRdfTriple {
            path,
            index,
            triple,
        } => (
            "rdf.triple.replace",
            format!("/package/rdf/{path}/triples/{index}"),
            rdf_triple_value(triple)?,
        ),
        Operation::RemoveRdfTriple { path, index } => (
            "rdf.triple.remove",
            format!("/package/rdf/{path}/triples/{index}"),
            Value::Null,
        ),
        Operation::MoveRdfTriple { path, from, to } => (
            "rdf.triple.move",
            format!("/package/rdf/{path}/triples/{from}"),
            serde_json::json!({"to": to}),
        ),
        Operation::AddForm { group_index, form } => (
            "form.add",
            format!("/body/form-groups/{group_index}/forms/-"),
            form_fragment_value(&form.to_xml_fragment()?)?,
        ),
        Operation::AddNestedForm { parent_form, form } => (
            "form.add_nested",
            format!("/body/forms/{parent_form}/forms/-"),
            form_fragment_value(&form.to_xml_fragment()?)?,
        ),
        Operation::AddFormFragment {
            group_index,
            parent_form,
            fragment,
        } => match parent_form {
            Some(parent_form) => (
                "form.add_nested",
                format!("/body/forms/{parent_form}/forms/-"),
                form_fragment_value(fragment)?,
            ),
            None => (
                "form.add",
                format!("/body/form-groups/{group_index}/forms/-"),
                form_fragment_value(fragment)?,
            ),
        },
        Operation::AddFormControl {
            form_index,
            control,
        } => (
            "form.control.add",
            format!("/body/forms/{form_index}/controls/-"),
            form_control_fragment_value(&control.to_xml_fragment()?)?,
        ),
        Operation::AddFormControlFragment {
            form_index,
            fragment,
        } => (
            "form.control.add",
            format!("/body/forms/{form_index}/controls/-"),
            form_control_fragment_value(fragment)?,
        ),
        Operation::ReplaceFormControl { index, control } => (
            "form.control.replace",
            format!("/body/form-controls/{index}"),
            form_control_fragment_value(&control.to_xml_fragment()?)?,
        ),
        Operation::ReplaceFormControlFragment { index, fragment } => (
            "form.control.replace",
            format!("/body/form-controls/{index}"),
            form_control_fragment_value(fragment)?,
        ),
        Operation::RemoveFormControl { index } => (
            "form.control.remove",
            format!("/body/form-controls/{index}"),
            Value::Null,
        ),
        Operation::MoveFormControl { from, to } => (
            "form.control.move",
            format!("/body/form-controls/{from}"),
            serde_json::json!({"to": to}),
        ),
        Operation::ReplaceForm { index, form } => (
            "form.replace",
            format!("/body/forms/{index}"),
            form_fragment_value(&form.to_xml_fragment()?)?,
        ),
        Operation::ReplaceFormFragment { index, fragment } => (
            "form.replace",
            format!("/body/forms/{index}"),
            form_fragment_value(fragment)?,
        ),
        Operation::RemoveForm { index } => {
            ("form.remove", format!("/body/forms/{index}"), Value::Null)
        },
        Operation::MoveForm { from, to } => (
            "form.move",
            format!("/body/forms/{from}"),
            serde_json::json!({"to": to}),
        ),
        Operation::AddEmbeddedChart { definition } => (
            "chart.add",
            "/package/charts/-".to_string(),
            chart_content_value(
                &crate::odc::serialize_content(definition)?,
                crate::package::charts::EmbeddedChartStorage::PackageSubdocument,
                blobs,
            )?,
        ),
        Operation::AddEmbeddedChartWithStorage {
            definition,
            storage,
        } => (
            "chart.add",
            "/package/charts/-".to_string(),
            chart_content_value(&crate::odc::serialize_content(definition)?, *storage, blobs)?,
        ),
        Operation::AddEmbeddedChartContent { content, storage } => (
            "chart.add",
            "/package/charts/-".to_string(),
            chart_content_value(content, *storage, blobs)?,
        ),
        Operation::ReplaceEmbeddedChart { index, definition } => (
            "chart.replace",
            format!("/package/charts/{index}"),
            chart_content_value(
                &crate::odc::serialize_content(definition)?,
                crate::package::charts::EmbeddedChartStorage::PackageSubdocument,
                blobs,
            )?,
        ),
        Operation::ReplaceEmbeddedChartContent { index, content } => (
            "chart.replace",
            format!("/package/charts/{index}"),
            chart_content_value(
                content,
                crate::package::charts::EmbeddedChartStorage::PackageSubdocument,
                blobs,
            )?,
        ),
        Operation::AddEmbeddedResource { resource } => (
            "resource.embedded.add",
            "/package/embedded/-".to_string(),
            embedded_resource_value(resource, blobs)?,
        ),
        Operation::EmbeddedResourceBatch { changes } => (
            "resource.embedded.batch",
            "/package/embedded/batch".to_string(),
            embedded_resource_batch_value(changes, blobs)?,
        ),
        Operation::GarbageCollectEmbeddedResources { candidates } => (
            "resource.embedded.gc",
            "/package/embedded/gc".to_string(),
            embedded_resource_gc_value(candidates),
        ),
        Operation::ReplaceEmbeddedObject { index, resource } => (
            "resource.embedded.object.replace",
            format!("/package/embedded/objects/{index}"),
            embedded_resource_value(resource, blobs)?,
        ),
        Operation::ReplaceEmbeddedImage { index, resource } => (
            "resource.embedded.image.replace",
            format!("/package/embedded/images/{index}"),
            embedded_resource_value(resource, blobs)?,
        ),
        Operation::RemoveEmbeddedChart { index } => (
            "chart.remove",
            format!("/package/charts/{index}"),
            Value::Null,
        ),
        Operation::RemoveEmbeddedObject { index } => (
            "resource.embedded.object.remove",
            format!("/package/embedded/objects/{index}"),
            Value::Null,
        ),
        Operation::RemoveEmbeddedImage { index } => (
            "resource.embedded.image.remove",
            format!("/package/embedded/images/{index}"),
            Value::Null,
        ),
        Operation::MoveEmbeddedObject { from, to } => (
            "resource.embedded.object.move",
            format!("/package/embedded/objects/{from}"),
            serde_json::json!({"to": to}),
        ),
        Operation::MoveEmbeddedImage { from, to } => (
            "resource.embedded.image.move",
            format!("/package/embedded/images/{from}"),
            serde_json::json!({"to": to}),
        ),
        Operation::AddScriptResource { resource } => (
            "resource.script.add",
            "/package/scripts/-".to_string(),
            script_resource_value(resource, blobs)?,
        ),
        Operation::ReplaceScriptResource { path, resource } => (
            "resource.script.replace",
            format!("/package/scripts/{path}"),
            script_resource_value(resource, blobs)?,
        ),
        Operation::RemoveScriptResource { path } => (
            "resource.script.remove",
            format!("/package/scripts/{path}"),
            Value::Null,
        ),
        Operation::RestoreSnapshot => {
            return Err(Error::Unsupported(
                "this ODT operation has not migrated to the semantic durable vocabulary"
                    .to_string(),
            ));
        },
    };
    PatchOperation::new(
        limits,
        name,
        path,
        durable_preconditions(source, target),
        value,
    )
    .map_err(durable_wire_error)
}

fn durable_lineage<Mode>(patch: &CorePatch<Mode>) -> Result<DurableLineage<'_>> {
    if patch.format() != DURABLE_FORMAT || patch.operations().is_empty() {
        return Err(invalid_durable_patch());
    }
    let first = &patch.operations()[0];
    let source_id = first
        .preconditions
        .get(SOURCE_PRECONDITION)
        .and_then(Value::as_str)
        .ok_or_else(invalid_durable_patch)?;
    let target_id = first
        .preconditions
        .get(TARGET_PRECONDITION)
        .and_then(Value::as_str)
        .ok_or_else(invalid_durable_patch)?;
    if !is_canonical_digest(source_id) || !is_canonical_digest(target_id) {
        return Err(invalid_durable_patch());
    }
    for operation in patch.operations() {
        if operation.preconditions.len() != 2
            || operation
                .preconditions
                .get(SOURCE_PRECONDITION)
                .and_then(Value::as_str)
                != Some(source_id)
            || operation
                .preconditions
                .get(TARGET_PRECONDITION)
                .and_then(Value::as_str)
                != Some(target_id)
        {
            return Err(invalid_durable_patch());
        }
    }
    Ok(DurableLineage {
        source_id,
        target_id,
    })
}

fn validate_reversible_patch(patch: &CorePatch<Reversible>) -> Result<()> {
    validate_patch_direction(patch)?;
    let forward = durable_lineage(patch)?;
    let inverse_patch = patch.inverse();
    validate_patch_direction(&inverse_patch)?;
    let reverse = durable_lineage(&inverse_patch)?;
    if forward.source_id != reverse.target_id || forward.target_id != reverse.source_id {
        return Err(invalid_durable_patch());
    }
    Ok(())
}

fn validate_sealed_patch(patch: &CorePatch<ForwardOnly>) -> Result<()> {
    validate_patch_direction(patch)
}

fn validate_patch_direction<Mode>(patch: &CorePatch<Mode>) -> Result<()> {
    durable_lineage(patch)?;
    let restoring = patch
        .operations()
        .iter()
        .all(|operation| operation.op == RESTORE_OPERATION);
    if restoring {
        restore_target_bytes(patch)?.ok_or_else(invalid_durable_patch)?;
        return Ok(());
    }
    let mut referenced_blobs = BTreeSet::new();
    for operation in patch.operations() {
        decode_semantic_operation(operation, patch.blobs())?;
        collect_blob_references(&operation.value, &mut referenced_blobs)?;
    }
    if referenced_blobs.len() != patch.blobs().len()
        || referenced_blobs
            .iter()
            .any(|id| blob_by_hex(patch.blobs(), id).is_none())
    {
        return Err(invalid_durable_patch());
    }
    Ok(())
}

fn collect_blob_references(value: &Value, output: &mut BTreeSet<String>) -> Result<()> {
    match value {
        Value::Object(object) => {
            if let Some(blob) = object.get("blob") {
                let blob = blob.as_str().ok_or_else(invalid_durable_patch)?;
                if !is_canonical_digest(blob) {
                    return Err(invalid_durable_patch());
                }
                output.insert(blob.to_owned());
            }
            for (key, child) in object {
                if key != "blob" {
                    collect_blob_references(child, output)?;
                }
            }
        },
        Value::Array(values) => {
            for child in values {
                collect_blob_references(child, output)?;
            }
        },
        _ => {},
    }
    Ok(())
}

fn restore_target_bytes<Mode>(patch: &CorePatch<Mode>) -> Result<Option<&[u8]>> {
    if !patch
        .operations()
        .iter()
        .all(|operation| operation.op == RESTORE_OPERATION)
    {
        return Ok(None);
    }
    if patch.blobs().len() != 1 {
        return Err(invalid_durable_patch());
    }
    let first_id = patch.operations()[0]
        .value
        .as_str()
        .ok_or_else(invalid_durable_patch)?;
    if patch
        .operations()
        .iter()
        .any(|operation| operation.target != "/" || operation.value.as_str() != Some(first_id))
    {
        return Err(invalid_durable_patch());
    }
    let blob_id = patch
        .blobs()
        .ids()
        .next()
        .ok_or_else(invalid_durable_patch)?;
    if blob_id.as_hex() != first_id {
        return Err(invalid_durable_patch());
    }
    let bytes = patch
        .blobs()
        .get(blob_id)
        .ok_or_else(invalid_durable_patch)?;
    let lineage = durable_lineage(patch)?;
    if BlobId::of(bytes).as_hex() != lineage.target_id {
        return Err(invalid_durable_patch());
    }
    Snapshot::from_bytes(copy_bytes(bytes)?)?;
    Ok(Some(bytes))
}

fn apply_durable_patch<Mode>(patch: &CorePatch<Mode>, source: &Snapshot) -> Result<Snapshot> {
    validate_patch_direction(patch)?;
    let lineage = durable_lineage(patch)?;
    if BlobId::of(source.as_bytes()).as_hex() != lineage.source_id {
        return Err(durable_source_mismatch());
    }
    if let Some(bytes) = restore_target_bytes(patch)? {
        return Snapshot::from_bytes(copy_bytes(bytes)?);
    }
    let mut operations = Vec::new();
    operations
        .try_reserve_exact(patch.operations().len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODT durable semantic replay",
            source: allocation_error,
        })?;
    for operation in patch.operations() {
        let decoded = decode_semantic_operation(operation, patch.blobs())?;
        if !matches!(decoded, Operation::Noop) {
            operations.push(decoded);
        }
    }
    let committed = Edit {
        source: source.clone(),
        operations,
    }
    .commit()?;
    if BlobId::of(committed.snapshot().as_bytes()).as_hex() != lineage.target_id {
        return Err(Error::InvalidFormat(
            "ODT semantic patch replay produced an unexpected target".to_string(),
        ));
    }
    Ok(committed.into_snapshot())
}

fn decode_semantic_operation(operation: &PatchOperation, blobs: &BlobBundle) -> Result<Operation> {
    let index = |prefix: &str, suffix: &str| parse_target_index(&operation.target, prefix, suffix);
    match operation.op.as_str() {
        NOOP_OPERATION if operation.target == "/" && operation.value.is_null() => {
            Ok(Operation::Noop)
        },
        "paragraph.insert" => Ok(Operation::InsertParagraph {
            index: index("/body/paragraphs/", "")?,
            text: object_string(&operation.value, "text", 1)?,
        }),
        "paragraph.replace" => Ok(Operation::ReplaceParagraph {
            index: index("/body/paragraphs/", "")?,
            text: object_string(&operation.value, "text", 1)?,
        }),
        "paragraph.remove" if operation.value.is_null() => Ok(Operation::RemoveParagraph {
            index: index("/body/paragraphs/", "")?,
        }),
        "run.append" => {
            let value = operation
                .value
                .as_object()
                .ok_or_else(invalid_durable_patch)?;
            if value.len() != 2 {
                return Err(invalid_durable_patch());
            }
            let text = value
                .get("text")
                .and_then(Value::as_str)
                .ok_or_else(invalid_durable_patch)?;
            let style_name = match value.get("style_name") {
                Some(Value::String(style)) => {
                    Some(bounded_semantic_text(style.clone(), "run style name")?)
                },
                Some(Value::Null) => None,
                _ => return Err(invalid_durable_patch()),
            };
            Ok(Operation::AppendRun {
                paragraph: index("/body/paragraphs/", "/runs/-")?,
                text: bounded_semantic_text(text.to_owned(), "run text")?,
                style_name,
            })
        },
        "hyperlink.append" => {
            let value = operation
                .value
                .as_object()
                .ok_or_else(invalid_durable_patch)?;
            if value.len() != 2 {
                return Err(invalid_durable_patch());
            }
            Ok(Operation::AppendHyperlink {
                paragraph: index("/body/paragraphs/", "/hyperlinks/-")?,
                href: object_string(&operation.value, "href", 2)?,
                text: object_string(&operation.value, "text", 2)?,
            })
        },
        "run.append_line_break" if operation.value.is_null() => Ok(Operation::AppendLineBreak {
            index: index("/body/paragraphs/", "/runs/-")?,
        }),
        "note.insert" => {
            let (note, fragment) = note_from_value(&operation.value)?;
            Ok(Operation::InsertNote {
                paragraph: index("/body/paragraphs/", "/notes/-")?,
                note,
                fragment: Some(fragment),
            })
        },
        "note.replace" => {
            let (note, fragment) = note_from_value(&operation.value)?;
            Ok(Operation::ReplaceNote {
                index: index("/body/notes/", "")?,
                note,
                fragment: Some(fragment),
            })
        },
        "note.remove" if operation.value.is_null() => Ok(Operation::RemoveNote {
            index: index("/body/notes/", "")?,
        }),
        "ruby.annotation.insert" => Ok(Operation::InsertRubyAnnotation {
            paragraph: index("/body/paragraphs/", "/ruby/-")?,
            annotation: ruby_annotation_from_value(&operation.value)?,
        }),
        "ruby.annotation.replace" => Ok(Operation::ReplaceRubyAnnotation {
            index: index("/body/ruby/", "")?,
            annotation: ruby_annotation_from_value(&operation.value)?,
        }),
        "ruby.annotation.remove" if operation.value.is_null() => {
            Ok(Operation::RemoveRubyAnnotation {
                index: index("/body/ruby/", "")?,
            })
        },
        "field.dynamic.insert" => Ok(Operation::InsertDynamicTextField {
            paragraph: index("/body/paragraphs/", "/fields/-")?,
            field: dynamic_field_from_value(&operation.value)?,
        }),
        "field.dynamic.replace" => Ok(Operation::ReplaceDynamicTextField {
            index: index("/body/fields/", "")?,
            field: dynamic_field_from_value(&operation.value)?,
        }),
        "field.dynamic.remove" if operation.value.is_null() => {
            Ok(Operation::RemoveDynamicTextField {
                index: index("/body/fields/", "")?,
            })
        },
        "style.ruby.set" => {
            let name = target_tail(&operation.target, "/styles/ruby/")?;
            let style = ruby_style_from_value(&operation.value)?;
            if style.name != name {
                return Err(invalid_durable_patch());
            }
            Ok(Operation::SetRubyStyle { style })
        },
        "style.ruby.remove" if operation.value.is_null() => Ok(Operation::RemoveRubyStyle {
            name: target_tail(&operation.target, "/styles/ruby/")?,
        }),
        "revision.policy.set" if operation.target == "/body/tracked-changes/policy" => {
            let value = exact_object(&operation.value, 3)?;
            Ok(Operation::SetTrackedChangePolicy {
                track_changes: optional_bool(value.get("track_changes"))?,
                protection_key: optional_bounded_string(
                    value.get("protection_key"),
                    "tracked-change protection key",
                )?,
                digest_algorithm: optional_bounded_string(
                    value.get("digest_algorithm"),
                    "tracked-change digest algorithm",
                )?,
            })
        },
        "revision.remove" if operation.value.is_null() => Ok(Operation::RemoveTrackedChange {
            id: target_tail(&operation.target, "/body/tracked-changes/")?,
        }),
        "rdf.graph.add" if operation.target == "/package/rdf/-" => {
            let value = exact_object(&operation.value, 2)?;
            Ok(Operation::AddRdfGraph {
                preferred_path: optional_bounded_string(
                    value.get("preferred_path"),
                    "RDF preferred path",
                )?,
                triples: rdf_triples_from_value(
                    value.get("triples").ok_or_else(invalid_durable_patch)?,
                )?,
            })
        },
        "rdf.graph.replace" => {
            let value = exact_object(&operation.value, 1)?;
            Ok(Operation::ReplaceRdfGraph {
                path: target_tail(&operation.target, "/package/rdf/")?,
                triples: rdf_triples_from_value(
                    value.get("triples").ok_or_else(invalid_durable_patch)?,
                )?,
            })
        },
        "rdf.graph.remove" if operation.value.is_null() => Ok(Operation::RemoveRdfGraph {
            path: target_tail(&operation.target, "/package/rdf/")?,
        }),
        "protection.set" if operation.target == "/package/protection" => {
            Ok(Operation::SetProtection {
                policy: protection_from_value(&operation.value)?,
            })
        },
        "rdf.triple.add" => Ok(Operation::AddRdfTriple {
            path: target_middle(&operation.target, "/package/rdf/", "/triples/-")?,
            triple: rdf_triple_from_value(&operation.value)?,
        }),
        "rdf.triple.replace" => {
            let (path, triple_index) = rdf_target(&operation.target)?;
            Ok(Operation::ReplaceRdfTriple {
                path,
                index: triple_index,
                triple: rdf_triple_from_value(&operation.value)?,
            })
        },
        "rdf.triple.remove" if operation.value.is_null() => {
            let (path, triple_index) = rdf_target(&operation.target)?;
            Ok(Operation::RemoveRdfTriple {
                path,
                index: triple_index,
            })
        },
        "rdf.triple.move" => {
            let (path, from) = rdf_target(&operation.target)?;
            let value = exact_object(&operation.value, 1)?;
            let to = value
                .get("to")
                .and_then(Value::as_u64)
                .and_then(|value| usize::try_from(value).ok())
                .ok_or_else(invalid_durable_patch)?;
            Ok(Operation::MoveRdfTriple { path, from, to })
        },
        "form.add" => Ok(Operation::AddFormFragment {
            group_index: index("/body/form-groups/", "/forms/-")?,
            parent_form: None,
            fragment: form_fragment_from_value(&operation.value)?,
        }),
        "form.add_nested" => Ok(Operation::AddFormFragment {
            group_index: 0,
            parent_form: Some(index("/body/forms/", "/forms/-")?),
            fragment: form_fragment_from_value(&operation.value)?,
        }),
        "form.control.add" => Ok(Operation::AddFormControlFragment {
            form_index: index("/body/forms/", "/controls/-")?,
            fragment: form_control_fragment_from_value(&operation.value)?,
        }),
        "form.control.replace" => Ok(Operation::ReplaceFormControlFragment {
            index: index("/body/form-controls/", "")?,
            fragment: form_control_fragment_from_value(&operation.value)?,
        }),
        "form.control.remove" if operation.value.is_null() => Ok(Operation::RemoveFormControl {
            index: index("/body/form-controls/", "")?,
        }),
        "form.control.move" => Ok(Operation::MoveFormControl {
            from: index("/body/form-controls/", "")?,
            to: move_target_from_value(&operation.value)?,
        }),
        "form.replace" => Ok(Operation::ReplaceFormFragment {
            index: index("/body/forms/", "")?,
            fragment: form_fragment_from_value(&operation.value)?,
        }),
        "form.remove" if operation.value.is_null() => Ok(Operation::RemoveForm {
            index: index("/body/forms/", "")?,
        }),
        "form.move" => Ok(Operation::MoveForm {
            from: index("/body/forms/", "")?,
            to: move_target_from_value(&operation.value)?,
        }),
        "chart.add" if operation.target == "/package/charts/-" => {
            let (content, storage) = chart_content_from_value(&operation.value, blobs)?;
            Ok(Operation::AddEmbeddedChartContent { content, storage })
        },
        "chart.replace" => {
            let (content, _) = chart_content_from_value(&operation.value, blobs)?;
            Ok(Operation::ReplaceEmbeddedChartContent {
                index: index("/package/charts/", "")?,
                content,
            })
        },
        "chart.remove" if operation.value.is_null() => Ok(Operation::RemoveEmbeddedChart {
            index: index("/package/charts/", "")?,
        }),
        "resource.embedded.add" if operation.target == "/package/embedded/-" => {
            Ok(Operation::AddEmbeddedResource {
                resource: embedded_resource_from_value(&operation.value, blobs)?,
            })
        },
        "resource.embedded.batch" if operation.target == "/package/embedded/batch" => {
            Ok(Operation::EmbeddedResourceBatch {
                changes: embedded_resource_batch_from_value(&operation.value, blobs)?,
            })
        },
        "resource.embedded.gc" if operation.target == "/package/embedded/gc" => {
            Ok(Operation::GarbageCollectEmbeddedResources {
                candidates: embedded_resource_gc_from_value(&operation.value)?,
            })
        },
        "resource.embedded.object.replace" => Ok(Operation::ReplaceEmbeddedObject {
            index: index("/package/embedded/objects/", "")?,
            resource: embedded_resource_from_value(&operation.value, blobs)?,
        }),
        "resource.embedded.image.replace" => Ok(Operation::ReplaceEmbeddedImage {
            index: index("/package/embedded/images/", "")?,
            resource: embedded_resource_from_value(&operation.value, blobs)?,
        }),
        "resource.embedded.object.remove" if operation.value.is_null() => {
            Ok(Operation::RemoveEmbeddedObject {
                index: index("/package/embedded/objects/", "")?,
            })
        },
        "resource.embedded.image.remove" if operation.value.is_null() => {
            Ok(Operation::RemoveEmbeddedImage {
                index: index("/package/embedded/images/", "")?,
            })
        },
        "resource.embedded.object.move" => Ok(Operation::MoveEmbeddedObject {
            from: index("/package/embedded/objects/", "")?,
            to: move_target_from_value(&operation.value)?,
        }),
        "resource.embedded.image.move" => Ok(Operation::MoveEmbeddedImage {
            from: index("/package/embedded/images/", "")?,
            to: move_target_from_value(&operation.value)?,
        }),
        "resource.script.add" if operation.target == "/package/scripts/-" => {
            Ok(Operation::AddScriptResource {
                resource: script_resource_from_value(&operation.value, blobs)?,
            })
        },
        "resource.script.replace" => Ok(Operation::ReplaceScriptResource {
            path: target_tail(&operation.target, "/package/scripts/")?,
            resource: script_resource_from_value(&operation.value, blobs)?,
        }),
        "resource.script.remove" if operation.value.is_null() => {
            Ok(Operation::RemoveScriptResource {
                path: target_tail(&operation.target, "/package/scripts/")?,
            })
        },
        _ => Err(invalid_durable_patch()),
    }
}

fn note_from_value(value: &Value) -> Result<(crate::note::Note, String)> {
    let fragment = object_string(value, "xml", 1)?;
    crate::note::validate_note_fragment(&fragment)?;
    let xml = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>{fragment}</text:p></office:text></office:body></office:document-content>"#
    );
    let mut notes = crate::parse_notes(&xml)?;
    if notes.len() != 1 {
        return Err(invalid_durable_patch());
    }
    notes
        .pop()
        .map(|note| (note, fragment))
        .ok_or_else(invalid_durable_patch)
}

fn note_fragment_value(note: &crate::note::Note, fragment: Option<&str>) -> Result<Value> {
    let xml = match fragment {
        Some(fragment) => fragment.to_owned(),
        None => note.to_xml_fragment()?,
    };
    crate::note::validate_note_fragment(&xml)?;
    Ok(serde_json::json!({"xml": xml}))
}

fn ruby_annotation_from_value(value: &Value) -> Result<crate::ruby_family::Annotation> {
    crate::ruby_family::Annotation::from_xml_fragment(&object_string(value, "xml", 1)?)
        .map_err(|_error| invalid_durable_patch())
}

fn form_fragment_value(fragment: &str) -> Result<Value> {
    crate::package::forms::validate_form_fragment(fragment)?;
    Ok(serde_json::json!({"xml": fragment}))
}

fn form_fragment_from_value(value: &Value) -> Result<String> {
    let fragment = object_string(value, "xml", 1)?;
    crate::package::forms::validate_form_fragment(&fragment)?;
    Ok(fragment)
}

fn form_control_fragment_value(fragment: &str) -> Result<Value> {
    crate::package::forms::validate_control_fragment(fragment)?;
    Ok(serde_json::json!({"xml": fragment}))
}

fn form_control_fragment_from_value(value: &Value) -> Result<String> {
    let fragment = object_string(value, "xml", 1)?;
    crate::package::forms::validate_control_fragment(&fragment)?;
    Ok(fragment)
}

fn move_target_from_value(value: &Value) -> Result<usize> {
    json_usize(exact_object(value, 1)?.get("to"))
}

fn dynamic_field_from_value(value: &Value) -> Result<DynamicTextField> {
    let fragment = object_string(value, "xml", 1)?;
    let xml = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>{fragment}</text:p></office:text></office:body></office:document-content>"#
    );
    let mut fields = FieldParser::parse_dynamic_text_fields(&xml)?;
    if fields.len() != 1 {
        return Err(invalid_durable_patch());
    }
    fields.pop().ok_or_else(invalid_durable_patch)
}

fn ruby_style_from_value(value: &Value) -> Result<crate::ruby_family::Style> {
    let fragment = object_string(value, "xml", 1)?;
    let xml = format!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles>{fragment}</office:styles></office:document-styles>"#
    );
    let mut styles = crate::parse_ruby_styles(&xml)?.styles;
    if styles.len() != 1 {
        return Err(invalid_durable_patch());
    }
    styles.pop().ok_or_else(invalid_durable_patch)
}

fn rdf_triples_value(triples: &[crate::rdf::Triple]) -> Result<Value> {
    triples
        .iter()
        .map(rdf_triple_value)
        .collect::<Result<Vec<_>>>()
        .map(Value::Array)
}

fn rdf_triple_value(triple: &crate::rdf::Triple) -> Result<Value> {
    let (subject_kind, subject_value) = match &triple.subject {
        crate::rdf::Subject::Iri(value) => ("iri", value),
        crate::rdf::Subject::BlankNode(value) => ("blank", value),
        _ => {
            return Err(Error::Unsupported(
                "unsupported RDF subject variant".to_string(),
            ));
        },
    };
    let object = match &triple.object {
        crate::rdf::Object::Iri(value) => serde_json::json!({"kind": "iri", "value": value}),
        crate::rdf::Object::BlankNode(value) => {
            serde_json::json!({"kind": "blank", "value": value})
        },
        crate::rdf::Object::Literal {
            value,
            datatype,
            language,
        } => serde_json::json!({
            "datatype": datatype,
            "kind": "literal",
            "language": language,
            "value": value,
        }),
        _ => {
            return Err(Error::Unsupported(
                "unsupported RDF object variant".to_string(),
            ));
        },
    };
    Ok(serde_json::json!({
        "object": object,
        "predicate": triple.predicate,
        "subject": {"kind": subject_kind, "value": subject_value},
    }))
}

fn rdf_triples_from_value(value: &Value) -> Result<Vec<crate::rdf::Triple>> {
    let values = value.as_array().ok_or_else(invalid_durable_patch)?;
    if values.len() > MAX_OPERATIONS {
        return Err(invalid_durable_patch());
    }
    values.iter().map(rdf_triple_from_value).collect()
}

fn rdf_triple_from_value(value: &Value) -> Result<crate::rdf::Triple> {
    let value = exact_object(value, 3)?;
    let subject = exact_object(value.get("subject").ok_or_else(invalid_durable_patch)?, 2)?;
    let subject_value = bounded_semantic_text(
        object_required_string_map(subject, "value")?.to_owned(),
        "RDF subject",
    )?;
    let subject = match object_required_string_map(subject, "kind")? {
        "iri" => crate::rdf::Subject::Iri(subject_value),
        "blank" => crate::rdf::Subject::BlankNode(subject_value),
        _ => return Err(invalid_durable_patch()),
    };
    let object = exact_object(
        value.get("object").ok_or_else(invalid_durable_patch)?,
        value
            .get("object")
            .and_then(Value::as_object)
            .map_or(0, serde_json::Map::len),
    )?;
    let object_value = bounded_semantic_text(
        object_required_string_map(object, "value")?.to_owned(),
        "RDF object",
    )?;
    let object = match object_required_string_map(object, "kind")? {
        "iri" if object.len() == 2 => crate::rdf::Object::Iri(object_value),
        "blank" if object.len() == 2 => crate::rdf::Object::BlankNode(object_value),
        "literal" if object.len() == 4 => crate::rdf::Object::Literal {
            value: object_value,
            datatype: optional_bounded_string(object.get("datatype"), "RDF datatype")?,
            language: optional_bounded_string(object.get("language"), "RDF language")?,
        },
        _ => return Err(invalid_durable_patch()),
    };
    Ok(crate::rdf::Triple {
        subject,
        predicate: bounded_semantic_text(
            object_required_string_map(value, "predicate")?.to_owned(),
            "RDF predicate",
        )?,
        object,
    })
}

fn protection_value(policy: &crate::protection::Policy) -> Value {
    serde_json::json!({
        "bookmarks": policy.bookmarks,
        "forms": policy.forms,
        "read_only": policy.read_only,
        "redline_key": policy.redline_key.as_ref().map(|key| hex_encode(key.as_bytes())),
    })
}

fn protection_from_value(value: &Value) -> Result<crate::protection::Policy> {
    let value = exact_object(value, 4)?;
    let redline_key = match value.get("redline_key") {
        Some(Value::String(value)) => Some(crate::protection::Key::new(hex_decode(value)?)?),
        Some(Value::Null) => None,
        _ => return Err(invalid_durable_patch()),
    };
    let policy = crate::protection::Policy {
        forms: optional_bool(value.get("forms"))?,
        bookmarks: optional_bool(value.get("bookmarks"))?,
        read_only: optional_bool(value.get("read_only"))?,
        redline_key,
    };
    policy.validate()?;
    Ok(policy)
}

fn chart_content_value(
    content: &str,
    storage: crate::package::charts::EmbeddedChartStorage,
    blobs: &mut BlobBundle,
) -> Result<Value> {
    crate::package::charts::validate_chart_content(content)?;
    let blob = blobs
        .insert(content.as_bytes())
        .map_err(durable_wire_error)?;
    let storage = match storage {
        crate::package::charts::EmbeddedChartStorage::PackageSubdocument => "package",
        crate::package::charts::EmbeddedChartStorage::InlineXml => "inline",
    };
    Ok(serde_json::json!({"blob": blob.as_hex(), "storage": storage}))
}

fn chart_content_from_value(
    value: &Value,
    blobs: &BlobBundle,
) -> Result<(String, crate::package::charts::EmbeddedChartStorage)> {
    let value = exact_object(value, 2)?;
    let blob = object_required_string_map(value, "blob")?;
    if !is_canonical_digest(blob) {
        return Err(invalid_durable_patch());
    }
    let content = std::str::from_utf8(blob_by_hex(blobs, blob).ok_or_else(invalid_durable_patch)?)
        .map_err(|_error| invalid_durable_patch())?
        .to_owned();
    crate::package::charts::validate_chart_content(&content)?;
    let storage = match object_required_string_map(value, "storage")? {
        "package" => crate::package::charts::EmbeddedChartStorage::PackageSubdocument,
        "inline" => crate::package::charts::EmbeddedChartStorage::InlineXml,
        _ => return Err(invalid_durable_patch()),
    };
    Ok((content, storage))
}

fn embedded_resource_batch_value(
    changes: &[crate::package::embedded::EmbeddedResourceChange],
    blobs: &mut BlobBundle,
) -> Result<Value> {
    use crate::package::embedded::{EmbeddedResourceChange, EmbeddedResourceSelector};
    if changes.is_empty() || changes.len() > crate::package::embedded::MAX_BATCH_CHANGES {
        return Err(invalid_durable_patch());
    }
    let mut values = Vec::new();
    values
        .try_reserve_exact(changes.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-resource durable batch",
            source,
        })?;
    for change in changes {
        values.push(match change {
            EmbeddedResourceChange::Add(resource) => serde_json::json!({
                "action": "add",
                "resource": embedded_resource_value(resource, blobs)?,
            }),
            EmbeddedResourceChange::Replace { selector, resource } => {
                let (kind, index) = match selector {
                    EmbeddedResourceSelector::Object(position) => ("object", position.get()),
                    EmbeddedResourceSelector::Image(position) => ("image", position.get()),
                };
                serde_json::json!({
                    "action": "replace",
                    "index": index,
                    "kind": kind,
                    "resource": embedded_resource_value(resource, blobs)?,
                })
            },
            EmbeddedResourceChange::Remove(selector) => {
                let (kind, index) = match selector {
                    EmbeddedResourceSelector::Object(position) => ("object", position.get()),
                    EmbeddedResourceSelector::Image(position) => ("image", position.get()),
                };
                serde_json::json!({
                    "action": "remove",
                    "index": index,
                    "kind": kind,
                })
            },
        });
    }
    Ok(Value::Array(values))
}

fn embedded_resource_batch_from_value(
    value: &Value,
    blobs: &BlobBundle,
) -> Result<Vec<crate::package::embedded::EmbeddedResourceChange>> {
    use crate::package::embedded::{EmbeddedResourceChange, EmbeddedResourceSelector};
    let values = value.as_array().ok_or_else(invalid_durable_patch)?;
    if values.is_empty() || values.len() > crate::package::embedded::MAX_BATCH_CHANGES {
        return Err(invalid_durable_patch());
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(values.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-resource durable batch",
            source,
        })?;
    for value in values {
        let object = value.as_object().ok_or_else(invalid_durable_patch)?;
        let action = object
            .get("action")
            .and_then(Value::as_str)
            .ok_or_else(invalid_durable_patch)?;
        if action == "add" && object.len() == 2 {
            changes.push(EmbeddedResourceChange::Add(embedded_resource_from_value(
                object.get("resource").ok_or_else(invalid_durable_patch)?,
                blobs,
            )?));
            continue;
        }
        let expected_fields = if action == "replace" { 4 } else { 3 };
        if object.len() != expected_fields {
            return Err(invalid_durable_patch());
        }
        let index = object
            .get("index")
            .and_then(Value::as_u64)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(invalid_durable_patch)?;
        let selector = match object.get("kind").and_then(Value::as_str) {
            Some("object") => EmbeddedResourceSelector::Object(Position::new(index)),
            Some("image") => EmbeddedResourceSelector::Image(Position::new(index)),
            _ => return Err(invalid_durable_patch()),
        };
        match action {
            "replace" => changes.push(EmbeddedResourceChange::Replace {
                selector,
                resource: embedded_resource_from_value(
                    object.get("resource").ok_or_else(invalid_durable_patch)?,
                    blobs,
                )?,
            }),
            "remove" => changes.push(EmbeddedResourceChange::Remove(selector)),
            _ => return Err(invalid_durable_patch()),
        }
    }
    Ok(changes)
}

fn embedded_resource_gc_value(
    candidates: &[crate::package::resource_gc::EmbeddedResourceGcCandidate],
) -> Value {
    use crate::package::resource_gc::EmbeddedResourceGcCandidate;
    Value::Array(
        candidates
            .iter()
            .map(|candidate| match candidate {
                EmbeddedResourceGcCandidate::PackageFile(path) => {
                    serde_json::json!({"kind": "file", "path": path})
                },
                EmbeddedResourceGcCandidate::PackageSubdocument(path) => {
                    serde_json::json!({"kind": "subdocument", "path": path})
                },
            })
            .collect(),
    )
}

fn embedded_resource_gc_from_value(
    value: &Value,
) -> Result<Vec<crate::package::resource_gc::EmbeddedResourceGcCandidate>> {
    use crate::package::resource_gc::EmbeddedResourceGcCandidate;
    let values = value.as_array().ok_or_else(invalid_durable_patch)?;
    if values.len() > 256 {
        return Err(invalid_durable_patch());
    }
    let mut candidates = Vec::new();
    candidates
        .try_reserve_exact(values.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-resource GC durable candidates",
            source,
        })?;
    for value in values {
        let object = exact_object(value, 2)?;
        let path = object_required_string_map(object, "path")?;
        crate::package::resource_gc::validate_candidate_path_bound(path)
            .map_err(|_error| invalid_durable_patch())?;
        let path = path.to_owned();
        let candidate = match object_required_string_map(object, "kind")? {
            "file" => EmbeddedResourceGcCandidate::PackageFile(path),
            "subdocument" => EmbeddedResourceGcCandidate::PackageSubdocument(path),
            _ => return Err(invalid_durable_patch()),
        };
        candidates.push(candidate);
    }
    Ok(candidates)
}

fn embedded_resource_value(
    resource: &crate::package::embedded::EmbeddedResource,
    blobs: &mut BlobBundle,
) -> Result<Value> {
    use crate::package::embedded::{EmbeddedResourceKind, EmbeddedResourceSource};
    let kind = match resource.kind {
        EmbeddedResourceKind::Object => "object",
        EmbeddedResourceKind::ObjectOle => "object-ole",
        EmbeddedResourceKind::Image => "image",
    };
    let source = match &resource.source {
        EmbeddedResourceSource::Linked { href } => {
            serde_json::json!({"href": href, "type": "linked"})
        },
        EmbeddedResourceSource::PackageFile {
            bytes,
            media_type,
            preferred_path,
        } => {
            let blob = blobs.insert(bytes).map_err(durable_wire_error)?;
            serde_json::json!({
                "blob": blob.as_hex(),
                "media_type": media_type,
                "preferred_path": preferred_path,
                "type": "package-file",
            })
        },
        EmbeddedResourceSource::PackageSubdocument {
            files,
            media_type,
            preferred_root,
        } => {
            if files.len() > 16_384 {
                return Err(Error::InvalidFormat(
                    "embedded resource has too many durable files".to_string(),
                ));
            }
            let total = files.iter().try_fold(0usize, |total, file| {
                total.checked_add(file.bytes.len()).ok_or_else(|| {
                    Error::InvalidFormat("embedded resource byte count overflow".to_string())
                })
            })?;
            ensure_package_size(total, "embedded resource durable payload")?;
            let mut payload = Vec::new();
            payload
                .try_reserve_exact(total)
                .map_err(|source| Error::Allocation {
                    resource: "embedded resource durable payload",
                    source,
                })?;
            let mut offset = 0usize;
            let mut entries = Vec::new();
            entries
                .try_reserve_exact(files.len())
                .map_err(|source| Error::Allocation {
                    resource: "embedded resource durable file table",
                    source,
                })?;
            for file in files {
                payload.extend_from_slice(&file.bytes);
                entries.push(serde_json::json!({
                    "length": file.bytes.len(),
                    "media_type": file.media_type,
                    "offset": offset,
                    "path": file.path,
                }));
                offset = offset.saturating_add(file.bytes.len());
            }
            let blob = blobs.insert(payload).map_err(durable_wire_error)?;
            serde_json::json!({
                "blob": blob.as_hex(),
                "files": entries,
                "media_type": media_type,
                "preferred_root": preferred_root,
                "type": "package-subdocument",
            })
        },
        EmbeddedResourceSource::InlineXml { root, xml } => {
            let blob = blobs.insert(xml.as_bytes()).map_err(durable_wire_error)?;
            let root = match root {
                litchi_odf_common::embedded::Root::OpenDocument => "open-document",
                litchi_odf_common::embedded::Root::MathMl => "mathml",
                _ => {
                    return Err(Error::Unsupported(
                        "unsupported embedded XML root".to_string(),
                    ));
                },
            };
            serde_json::json!({"blob": blob.as_hex(), "root": root, "type": "inline-xml"})
        },
        EmbeddedResourceSource::InlineBinary { bytes, media_type } => {
            let blob = blobs.insert(bytes).map_err(durable_wire_error)?;
            serde_json::json!({
                "blob": blob.as_hex(),
                "media_type": media_type,
                "type": "inline-binary",
            })
        },
    };
    Ok(serde_json::json!({
        "class_id": resource.class_id,
        "frame_name": resource.frame_name,
        "kind": kind,
        "source": source,
        "xml_id": resource.xml_id,
    }))
}

fn embedded_resource_from_value(
    value: &Value,
    blobs: &BlobBundle,
) -> Result<crate::package::embedded::EmbeddedResource> {
    use crate::package::embedded::{
        EmbeddedResource, EmbeddedResourceFile, EmbeddedResourceKind, EmbeddedResourceSource,
    };
    let value = exact_object(value, 5)?;
    let kind = match object_required_string_map(value, "kind")? {
        "object" => EmbeddedResourceKind::Object,
        "object-ole" => EmbeddedResourceKind::ObjectOle,
        "image" => EmbeddedResourceKind::Image,
        _ => return Err(invalid_durable_patch()),
    };
    let source = exact_object(
        value.get("source").ok_or_else(invalid_durable_patch)?,
        value
            .get("source")
            .and_then(Value::as_object)
            .map_or(0, serde_json::Map::len),
    )?;
    let source = match object_required_string_map(source, "type")? {
        "linked" if source.len() == 2 => EmbeddedResourceSource::Linked {
            href: bounded_semantic_text(
                object_required_string_map(source, "href")?.to_owned(),
                "embedded resource link",
            )?,
        },
        "package-file" if source.len() == 4 => EmbeddedResourceSource::PackageFile {
            bytes: resource_blob(source, blobs)?,
            media_type: bounded_semantic_text(
                object_required_string_map(source, "media_type")?.to_owned(),
                "embedded resource media type",
            )?,
            preferred_path: optional_bounded_string(
                source.get("preferred_path"),
                "embedded resource path",
            )?,
        },
        "package-subdocument" if source.len() == 5 => {
            let payload = resource_blob_ref(source, blobs)?;
            let entries = source
                .get("files")
                .and_then(Value::as_array)
                .ok_or_else(invalid_durable_patch)?;
            if entries.len() > 16_384 {
                return Err(invalid_durable_patch());
            }
            let mut files = Vec::new();
            files
                .try_reserve_exact(entries.len())
                .map_err(|source| Error::Allocation {
                    resource: "embedded resource durable file table",
                    source,
                })?;
            let mut expected_offset = 0usize;
            for entry in entries {
                let entry = exact_object(entry, 4)?;
                let offset = json_usize(entry.get("offset"))?;
                let length = json_usize(entry.get("length"))?;
                if offset != expected_offset {
                    return Err(invalid_durable_patch());
                }
                let end = offset
                    .checked_add(length)
                    .ok_or_else(invalid_durable_patch)?;
                let bytes = payload.get(offset..end).ok_or_else(invalid_durable_patch)?;
                files.push(EmbeddedResourceFile {
                    path: bounded_semantic_text(
                        object_required_string_map(entry, "path")?.to_owned(),
                        "embedded resource file path",
                    )?,
                    bytes: copy_bytes(bytes)?,
                    media_type: bounded_semantic_text(
                        object_required_string_map(entry, "media_type")?.to_owned(),
                        "embedded resource file media type",
                    )?,
                });
                expected_offset = end;
            }
            if expected_offset != payload.len() {
                return Err(invalid_durable_patch());
            }
            EmbeddedResourceSource::PackageSubdocument {
                files,
                media_type: bounded_semantic_text(
                    object_required_string_map(source, "media_type")?.to_owned(),
                    "embedded resource media type",
                )?,
                preferred_root: optional_bounded_string(
                    source.get("preferred_root"),
                    "embedded resource root",
                )?,
            }
        },
        "inline-xml" if source.len() == 3 => {
            let bytes = resource_blob_ref(source, blobs)?;
            let xml = std::str::from_utf8(bytes)
                .map_err(|_error| invalid_durable_patch())?
                .to_owned();
            let root = match object_required_string_map(source, "root")? {
                "open-document" => litchi_odf_common::embedded::Root::OpenDocument,
                "mathml" => litchi_odf_common::embedded::Root::MathMl,
                _ => return Err(invalid_durable_patch()),
            };
            EmbeddedResourceSource::InlineXml { root, xml }
        },
        "inline-binary" if source.len() == 3 => EmbeddedResourceSource::InlineBinary {
            bytes: resource_blob(source, blobs)?,
            media_type: optional_bounded_string(
                source.get("media_type"),
                "embedded resource media type",
            )?,
        },
        _ => return Err(invalid_durable_patch()),
    };
    Ok(EmbeddedResource {
        kind,
        source,
        frame_name: optional_bounded_string(value.get("frame_name"), "embedded frame name")?,
        xml_id: optional_bounded_string(value.get("xml_id"), "embedded resource XML ID")?,
        class_id: optional_bounded_string(value.get("class_id"), "embedded resource class ID")?,
    })
}

fn resource_blob(source: &serde_json::Map<String, Value>, blobs: &BlobBundle) -> Result<Vec<u8>> {
    copy_bytes(resource_blob_ref(source, blobs)?)
}

fn resource_blob_ref<'a>(
    source: &serde_json::Map<String, Value>,
    blobs: &'a BlobBundle,
) -> Result<&'a [u8]> {
    let blob = object_required_string_map(source, "blob")?;
    if !is_canonical_digest(blob) {
        return Err(invalid_durable_patch());
    }
    blob_by_hex(blobs, blob).ok_or_else(invalid_durable_patch)
}

fn json_usize(value: Option<&Value>) -> Result<usize> {
    value
        .and_then(Value::as_u64)
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(invalid_durable_patch)
}

fn script_resource_value(
    resource: &crate::ScriptResourceSpec,
    blobs: &mut BlobBundle,
) -> Result<Value> {
    let blob = blobs.insert(&resource.bytes).map_err(durable_wire_error)?;
    let kind = match resource.kind {
        crate::ScriptResourceKind::BasicLibrary => "basic-library",
        crate::ScriptResourceKind::BasicModule => "basic-module",
        crate::ScriptResourceKind::Dialog => "dialog",
        crate::ScriptResourceKind::Opaque => "opaque",
    };
    Ok(serde_json::json!({
        "blob": blob.as_hex(),
        "kind": kind,
        "media_type": resource.media_type,
        "preferred_path": resource.preferred_path,
    }))
}

fn script_resource_from_value(
    value: &Value,
    blobs: &BlobBundle,
) -> Result<crate::ScriptResourceSpec> {
    let value = exact_object(value, 4)?;
    let kind = match object_required_string_map(value, "kind")? {
        "basic-library" => crate::ScriptResourceKind::BasicLibrary,
        "basic-module" => crate::ScriptResourceKind::BasicModule,
        "dialog" => crate::ScriptResourceKind::Dialog,
        "opaque" => crate::ScriptResourceKind::Opaque,
        _ => return Err(invalid_durable_patch()),
    };
    let blob = object_required_string_map(value, "blob")?;
    if !is_canonical_digest(blob) {
        return Err(invalid_durable_patch());
    }
    let bytes = blob_by_hex(blobs, blob).ok_or_else(invalid_durable_patch)?;
    Ok(crate::ScriptResourceSpec {
        kind,
        preferred_path: optional_bounded_string(value.get("preferred_path"), "script path")?,
        media_type: bounded_semantic_text(
            object_required_string_map(value, "media_type")?.to_owned(),
            "script media type",
        )?,
        bytes: copy_bytes(bytes)?,
    })
}

fn blob_by_hex<'a>(blobs: &'a BlobBundle, id: &str) -> Option<&'a [u8]> {
    blobs
        .ids()
        .find(|candidate| candidate.as_hex() == id)
        .and_then(|candidate| blobs.get(candidate))
}

fn exact_object(value: &Value, fields: usize) -> Result<&serde_json::Map<String, Value>> {
    let value = value.as_object().ok_or_else(invalid_durable_patch)?;
    if value.len() != fields {
        return Err(invalid_durable_patch());
    }
    Ok(value)
}

fn object_required_string_map<'a>(
    value: &'a serde_json::Map<String, Value>,
    key: &str,
) -> Result<&'a str> {
    value
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(invalid_durable_patch)
}

fn optional_bounded_string(value: Option<&Value>, field: &str) -> Result<Option<String>> {
    match value {
        Some(Value::String(value)) => bounded_semantic_text(value.clone(), field).map(Some),
        Some(Value::Null) => Ok(None),
        _ => Err(invalid_durable_patch()),
    }
}

fn optional_bool(value: Option<&Value>) -> Result<Option<bool>> {
    match value {
        Some(Value::Bool(value)) => Ok(Some(*value)),
        Some(Value::Null) => Ok(None),
        _ => Err(invalid_durable_patch()),
    }
}

fn target_tail(target: &str, prefix: &str) -> Result<String> {
    let value = target
        .strip_prefix(prefix)
        .ok_or_else(invalid_durable_patch)?;
    if value.is_empty() {
        return Err(invalid_durable_patch());
    }
    bounded_semantic_text(value.to_owned(), "durable target")
}

fn target_middle(target: &str, prefix: &str, suffix: &str) -> Result<String> {
    let value = target
        .strip_prefix(prefix)
        .and_then(|value| value.strip_suffix(suffix))
        .ok_or_else(invalid_durable_patch)?;
    if value.is_empty() {
        return Err(invalid_durable_patch());
    }
    bounded_semantic_text(value.to_owned(), "durable target")
}

fn rdf_target(target: &str) -> Result<(String, usize)> {
    let value = target
        .strip_prefix("/package/rdf/")
        .ok_or_else(invalid_durable_patch)?;
    let (path, index) = value
        .rsplit_once("/triples/")
        .ok_or_else(invalid_durable_patch)?;
    if path.is_empty() {
        return Err(invalid_durable_patch());
    }
    Ok((
        bounded_semantic_text(path.to_owned(), "RDF path")?,
        parse_canonical_index(index)?,
    ))
}

fn parse_canonical_index(value: &str) -> Result<usize> {
    if value.is_empty() || (value.len() > 1 && value.starts_with('0')) {
        return Err(invalid_durable_patch());
    }
    value.parse().map_err(|_error| invalid_durable_patch())
}

fn hex_encode(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        output.push(char::from(HEX[usize::from(byte >> 4)]));
        output.push(char::from(HEX[usize::from(byte & 0x0f)]));
    }
    output
}

fn hex_decode(value: &str) -> Result<Vec<u8>> {
    if !value.len().is_multiple_of(2) {
        return Err(invalid_durable_patch());
    }
    value
        .as_bytes()
        .chunks_exact(2)
        .map(|pair| {
            let high = hex_nibble(pair[0]).ok_or_else(invalid_durable_patch)?;
            let low = hex_nibble(pair[1]).ok_or_else(invalid_durable_patch)?;
            Ok((high << 4) | low)
        })
        .collect()
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        _ => None,
    }
}

fn parse_target_index(target: &str, prefix: &str, suffix: &str) -> Result<usize> {
    let value = target
        .strip_prefix(prefix)
        .and_then(|value| value.strip_suffix(suffix))
        .ok_or_else(invalid_durable_patch)?;
    parse_canonical_index(value)
}

fn object_string(value: &Value, key: &str, expected_fields: usize) -> Result<String> {
    let object = value.as_object().ok_or_else(invalid_durable_patch)?;
    if object.len() != expected_fields {
        return Err(invalid_durable_patch());
    }
    let text = object
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(invalid_durable_patch)?;
    bounded_semantic_text(text.to_owned(), key)
}

fn is_canonical_digest(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn durable_wire_error(source: litchi_core::PatchError) -> Error {
    let message = format!("invalid ODT durable patch: {source}");
    drop(source);
    Error::InvalidFormat(message)
}

fn invalid_durable_patch() -> Error {
    Error::InvalidFormat("invalid ODT durable patch vocabulary".to_string())
}

fn durable_source_mismatch() -> Error {
    Error::InvalidFormat("ODT durable patch source does not match".to_string())
}

fn stale_history_error() -> Error {
    Error::InvalidFormat("ODT history edit is stale relative to the current head".to_string())
}

fn bounded_semantic_text(value: String, field: &str) -> Result<String> {
    if value.len() > MAX_SEMANTIC_TEXT_BYTES || value.chars().any(|character| character == '\0') {
        return Err(Error::InvalidFormat(format!(
            "ODT {field} exceeds semantic patch bounds"
        )));
    }
    Ok(value)
}

fn operation_effects(operations: &[Operation]) -> (Vec<String>, Vec<String>) {
    let mut writes = Vec::new();
    for operation in operations {
        match operation {
            Operation::Noop => {},
            Operation::InsertParagraph { index, .. } | Operation::RemoveParagraph { index } => {
                writes.push("/body/paragraphs/order".to_string());
                writes.push(format!("/body/paragraphs/{index}/content"));
            },
            Operation::ReplaceParagraph { index, .. } => {
                writes.push(format!("/body/paragraphs/{index}/content"));
            },
            Operation::AppendRun { paragraph, .. }
            | Operation::AppendHyperlink { paragraph, .. }
            | Operation::AppendLineBreak { index: paragraph } => {
                writes.push(format!("/body/paragraphs/{paragraph}/content"));
            },
            Operation::InsertNote { paragraph, .. } => {
                writes.push(format!("/body/paragraphs/{paragraph}/content"));
                writes.push("/body/notes/order".to_string());
            },
            Operation::ReplaceNote { index, .. } | Operation::RemoveNote { index } => {
                writes.push(format!("/body/notes/{index}"));
            },
            Operation::InsertRubyAnnotation { paragraph, .. } => {
                writes.push(format!("/body/paragraphs/{paragraph}/content"));
                writes.push("/body/ruby/order".to_string());
            },
            Operation::ReplaceRubyAnnotation { index, .. }
            | Operation::RemoveRubyAnnotation { index } => {
                writes.push(format!("/body/ruby/{index}"));
            },
            Operation::InsertDynamicTextField { paragraph, .. } => {
                writes.push(format!("/body/paragraphs/{paragraph}/content"));
            },
            Operation::ReplaceDynamicTextField { index, .. }
            | Operation::RemoveDynamicTextField { index } => {
                writes.push(format!("/body/fields/{index}"));
            },
            Operation::SetRubyStyle { style } => {
                writes.push(format!("/styles/ruby/{}", style.name));
            },
            Operation::RemoveRubyStyle { name } => {
                writes.push(format!("/styles/ruby/{name}"));
            },
            Operation::SetTrackedChangePolicy { .. } => {
                writes.push("/body/tracked-changes/policy".to_string());
            },
            Operation::RemoveTrackedChange { id } => {
                writes.push(format!("/body/tracked-changes/{id}"));
            },
            Operation::AddRdfGraph { .. } => writes.push("/package/rdf/order".to_string()),
            Operation::ReplaceRdfGraph { path, .. } | Operation::RemoveRdfGraph { path } => {
                writes.push(format!("/package/rdf/{path}"));
            },
            Operation::SetProtection { .. } => {
                writes.push("/package/protection".to_string());
            },
            Operation::AddRdfTriple { path, .. } => {
                writes.push(format!("/package/rdf/{path}/triples/order"));
            },
            Operation::ReplaceRdfTriple { path, index, .. }
            | Operation::RemoveRdfTriple { path, index } => {
                writes.push(format!("/package/rdf/{path}/triples/{index}"));
            },
            Operation::MoveRdfTriple { path, .. } => {
                writes.push(format!("/package/rdf/{path}/triples/order"));
            },
            Operation::AddForm { group_index, .. } => {
                writes.push(format!("/body/form-groups/{group_index}/forms/order"));
            },
            Operation::AddNestedForm { parent_form, .. } => {
                writes.push(format!("/body/forms/{parent_form}/forms/order"));
            },
            Operation::AddFormFragment {
                group_index,
                parent_form,
                ..
            } => {
                if let Some(parent_form) = parent_form {
                    writes.push(format!("/body/forms/{parent_form}/forms/order"));
                } else {
                    writes.push(format!("/body/form-groups/{group_index}/forms/order"));
                }
            },
            Operation::AddFormControl { form_index, .. }
            | Operation::AddFormControlFragment { form_index, .. } => {
                writes.push(format!("/body/forms/{form_index}/controls/order"));
            },
            Operation::ReplaceFormControl { index, .. }
            | Operation::ReplaceFormControlFragment { index, .. }
            | Operation::RemoveFormControl { index } => {
                writes.push(format!("/body/form-controls/{index}"));
            },
            Operation::MoveFormControl { .. } => {
                writes.push("/body/form-controls/order".to_string());
            },
            Operation::ReplaceForm { index, .. }
            | Operation::ReplaceFormFragment { index, .. }
            | Operation::RemoveForm { index } => {
                writes.push(format!("/body/forms/{index}"));
            },
            Operation::MoveForm { .. } => {
                writes.push("/body/forms/order".to_string());
            },
            Operation::AddScriptResource { .. } => {
                writes.push("/package/scripts/order".to_string());
            },
            Operation::ReplaceScriptResource { path, .. }
            | Operation::RemoveScriptResource { path } => {
                writes.push(format!("/package/scripts/{path}"));
            },
            Operation::AddEmbeddedChart { .. }
            | Operation::AddEmbeddedChartWithStorage { .. }
            | Operation::AddEmbeddedChartContent { .. } => {
                writes.push("/package/charts/order".to_string());
            },
            Operation::ReplaceEmbeddedChart { index, .. }
            | Operation::ReplaceEmbeddedChartContent { index, .. }
            | Operation::RemoveEmbeddedChart { index } => {
                writes.push(format!("/package/charts/{index}"));
            },
            Operation::AddEmbeddedResource { .. } => {
                writes.push("/package/embedded/order".to_string());
            },
            Operation::EmbeddedResourceBatch { changes } => {
                for change in changes {
                    match change {
                        crate::package::embedded::EmbeddedResourceChange::Add(_) => {
                            writes.push("/package/embedded/order".to_string());
                        },
                        crate::package::embedded::EmbeddedResourceChange::Replace {
                            selector,
                            ..
                        } => match selector {
                            crate::package::embedded::EmbeddedResourceSelector::Object(
                                position,
                            ) => {
                                writes
                                    .push(format!("/package/embedded/objects/{}", position.get()));
                            },
                            crate::package::embedded::EmbeddedResourceSelector::Image(position) => {
                                writes.push(format!("/package/embedded/images/{}", position.get()));
                            },
                        },
                        crate::package::embedded::EmbeddedResourceChange::Remove(selector) => {
                            match selector {
                                crate::package::embedded::EmbeddedResourceSelector::Object(
                                    position,
                                ) => {
                                    writes.push(format!(
                                        "/package/embedded/objects/{}",
                                        position.get()
                                    ));
                                },
                                crate::package::embedded::EmbeddedResourceSelector::Image(
                                    position,
                                ) => {
                                    writes.push(format!(
                                        "/package/embedded/images/{}",
                                        position.get()
                                    ));
                                },
                            }
                        },
                    }
                }
            },
            Operation::GarbageCollectEmbeddedResources { candidates } => {
                for candidate in candidates {
                    writes.push(format!("/package/embedded/gc/{}", candidate.path()));
                }
            },
            Operation::ReplaceEmbeddedObject { index, .. }
            | Operation::RemoveEmbeddedObject { index } => {
                writes.push(format!("/package/embedded/objects/{index}"));
            },
            Operation::ReplaceEmbeddedImage { index, .. }
            | Operation::RemoveEmbeddedImage { index } => {
                writes.push(format!("/package/embedded/images/{index}"));
            },
            Operation::MoveEmbeddedObject { .. } => {
                writes.push("/package/embedded/objects/order".to_string());
            },
            Operation::MoveEmbeddedImage { .. } => {
                writes.push("/package/embedded/images/order".to_string());
            },
            Operation::RestoreSnapshot => writes.push("/package".to_string()),
        }
    }
    (Vec::new(), writes)
}

fn envelope_kind(snapshot: &Snapshot) -> Result<EnvelopeKind> {
    let package = envelope_package(snapshot)?;
    let archive = package.package()?;
    if archive
        .manifest()
        .entries
        .values()
        .any(|entry| entry.encryption.is_some())
    {
        return Ok(EnvelopeKind::Encrypted);
    }
    if archive.has_file("META-INF/documentsignatures.xml")
        || archive.has_file("META-INF/macrosignatures.xml")
    {
        return Ok(EnvelopeKind::Signed);
    }
    Ok(EnvelopeKind::Plain)
}

fn envelope_package(snapshot: &Snapshot) -> Result<crate::core::OwnedPackage> {
    ensure_package_size(snapshot.as_bytes().len(), "ODT transaction package")?;
    crate::core::OwnedPackage::from_shared_bytes(Arc::clone(&snapshot.bytes))
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    ensure_package_size(source.len(), "ODT transaction package")?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODT transaction package",
            source: allocation_error,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

fn ensure_package_size(size: usize, scope: &str) -> Result<()> {
    if size > MAX_PACKAGE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{scope} exceeds the {MAX_PACKAGE_BYTES}-byte package limit"
        )));
    }
    Ok(())
}

fn resolve_paragraph(document: &Document, selector: &ParagraphSelector) -> Result<usize> {
    let paragraphs = document.paragraphs()?;
    match selector {
        ParagraphSelector::Position(position) if position.get() < paragraphs.len() => {
            Ok(position.get())
        },
        ParagraphSelector::Position(position) => Err(Error::InvalidFormat(format!(
            "paragraph selector position {} is out of bounds (found {})",
            position.get(),
            paragraphs.len()
        ))),
        ParagraphSelector::Index(index) if *index < paragraphs.len() => Ok(*index),
        ParagraphSelector::Index(index) => Err(Error::InvalidFormat(format!(
            "paragraph selector index {index} is out of bounds (found {})",
            paragraphs.len()
        ))),
        ParagraphSelector::ExactText(text) => {
            let mut selected = None;
            for (index, paragraph) in paragraphs.into_iter().enumerate() {
                if paragraph.text()? == *text {
                    if selected.replace(index).is_some() {
                        return Err(Error::InvalidFormat(
                            "paragraph text selector is ambiguous".to_string(),
                        ));
                    }
                }
            }
            selected.ok_or_else(|| {
                Error::InvalidFormat("paragraph text selector did not match".to_string())
            })
        },
    }
}

fn audit_changed_xml_is_compact(
    source: &crate::core::OwnedPackage,
    candidate: &crate::core::OwnedPackage,
) -> Result<()> {
    let source_archive = source.package()?;
    let archive = candidate.package()?;
    for path in archive.files()? {
        let xml_media_type = archive
            .manifest()
            .entries
            .get(&path)
            .is_some_and(|entry| entry.media_type.contains("xml"));
        if path.ends_with(".xml") || path.ends_with(".rdf") || xml_media_type {
            let xml = candidate.get_file(&path)?;
            if source_archive.has_file(&path) && source.get_file(&path)? == xml {
                continue;
            }
            if path == crate::constants::ODF_CONTENT {
                let candidate_content = std::str::from_utf8(&xml).map_err(|error| {
                    Error::InvalidFormat(format!("changed content.xml is not valid UTF-8: {error}"))
                })?;
                if litchi_odf_common::package::content_splice_publication(
                    &source,
                    candidate_content,
                )
                .is_ok()
                {
                    continue;
                }
            }
            let limits = litchi_odf_common::compact_xml::Limits::new(MAX_PACKAGE_BYTES, 4_096)
                .map_err(Error::from)?;
            let validation_xml = numeric_reference_projection(&xml)?;
            litchi_odf_common::compact_xml::validate_with_limits(&validation_xml, limits).map_err(
                |source| {
                    Error::InvalidFormat(format!("XML publication rejected for '{path}': {source}"))
                },
            )?;
        }
    }
    Ok(())
}

fn numeric_reference_projection(xml: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT compact XML numeric-reference audit",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < xml.len() {
        if xml[cursor..].starts_with(b"&#") {
            let relative_end = xml[cursor + 2..]
                .iter()
                .position(|byte| *byte == b';')
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "ODT authored XML contains an unterminated numeric reference".to_string(),
                    )
                })?;
            let end = cursor
                .checked_add(2)
                .and_then(|value| value.checked_add(relative_end))
                .ok_or_else(invalid_durable_patch)?;
            let digits = &xml[cursor + 2..end];
            let (radix, digits) = match digits.split_first() {
                Some((b'x', digits)) => (16, digits),
                Some(_) => (10, digits),
                None => return Err(invalid_numeric_reference()),
            };
            if digits.is_empty()
                || digits.len() > 8
                || !digits.iter().all(|byte| match radix {
                    16 => byte.is_ascii_hexdigit(),
                    _ => byte.is_ascii_digit(),
                })
            {
                return Err(invalid_numeric_reference());
            }
            let digits =
                std::str::from_utf8(digits).map_err(|_error| invalid_numeric_reference())?;
            let scalar =
                u32::from_str_radix(digits, radix).map_err(|_error| invalid_numeric_reference())?;
            if !is_xml_scalar(scalar) {
                return Err(invalid_numeric_reference());
            }
            // The shared compactness checker intentionally accepts only
            // predefined named references. This sentinel preserves structure
            // after the numeric reference has been validated here.
            output.push(b'x');
            cursor = end.saturating_add(1);
        } else {
            output.push(xml[cursor]);
            cursor = cursor.saturating_add(1);
        }
    }
    Ok(output)
}

fn is_xml_scalar(value: u32) -> bool {
    matches!(
        value,
        0x9 | 0xa | 0xd | 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff
    )
}

fn invalid_numeric_reference() -> Error {
    Error::InvalidFormat("ODT authored XML contains an invalid numeric reference".to_string())
}

fn ensure_editable_envelope(snapshot: &Snapshot) -> Result<()> {
    match envelope_kind(snapshot)? {
        EnvelopeKind::Encrypted => Err(Error::Unsupported(
            "packaged ODT transactions preserve encrypted snapshots only as exact no-ops"
                .to_string(),
        )),
        EnvelopeKind::Signed => Err(Error::Unsupported(
            "packaged ODT transactions preserve signed snapshots only as exact no-ops".to_string(),
        )),
        EnvelopeKind::Plain => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn snapshot_from_document_shares_exact_package_allocation() -> Result<()> {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("shared snapshot")?;
        let document = Document::from_bytes(mutable.to_bytes()?)?;
        let package_bytes = document.transaction_package().shared_bytes();

        let snapshot = Snapshot::from_document(&document)?;

        assert!(Arc::ptr_eq(&snapshot.bytes, &package_bytes));
        assert_eq!(snapshot.as_bytes(), document.original_bytes());
        Ok(())
    }

    #[test]
    fn changed_final_document_handoff_shares_bytes_and_reopens_independently() -> Result<()> {
        let mut source_document = MutableDocument::new();
        source_document.add_paragraph("before")?;
        let source = Snapshot::from_bytes(source_document.to_bytes()?)?;

        let mut changed_document = MutableDocument::new();
        changed_document.add_paragraph("after")?;
        let changed_document = Document::from_bytes(changed_document.to_bytes()?)?;
        let changed_bytes = changed_document.transaction_package().shared_bytes();
        let copied_and_reparsed =
            Snapshot::from_bytes(copy_bytes(changed_document.original_bytes())?)?;

        let snapshot = snapshot_from_final_document(&source, &changed_document)?;

        assert!(Arc::ptr_eq(&snapshot.bytes, &changed_bytes));
        assert_eq!(snapshot.as_bytes(), copied_and_reparsed.as_bytes());
        assert_ne!(snapshot.as_bytes(), source.as_bytes());
        assert_eq!(snapshot.document()?.text()?, "after");
        Ok(())
    }

    #[test]
    fn direct_snapshot_and_reopened_document_share_exact_package_allocation() -> Result<()> {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("shared direct snapshot")?;
        let bytes = mutable.to_bytes()?;
        let source_pointer = bytes.as_ptr();

        let snapshot = Snapshot::from_bytes(bytes)?;
        let document = snapshot.document()?;
        let document_bytes = document.transaction_package().shared_bytes();

        assert_eq!(snapshot.as_bytes().as_ptr(), source_pointer);
        assert!(Arc::ptr_eq(&snapshot.bytes, &document_bytes));
        assert_eq!(snapshot.as_bytes(), document.original_bytes());
        Ok(())
    }

    #[test]
    fn envelope_classification_shares_the_snapshot_package_allocation() -> Result<()> {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("shared envelope classification")?;
        let snapshot = Snapshot::from_bytes(mutable.to_bytes()?)?;

        let package = envelope_package(&snapshot)?;

        assert!(Arc::ptr_eq(&snapshot.bytes, &package.shared_bytes()));
        assert_eq!(envelope_kind(&snapshot)?, EnvelopeKind::Plain);
        Ok(())
    }

    #[test]
    fn compact_audit_shares_the_validated_predecessor_package() -> Result<()> {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("before compact audit")?;
        let source = Document::from_bytes(mutable.to_bytes()?)?;
        let source_bytes = source.transaction_package().shared_bytes();
        let predecessor = source.transaction_package().clone();

        assert!(Arc::ptr_eq(&source_bytes, &predecessor.shared_bytes()));

        let mut mutable = MutableDocument::from_document(source)?;
        mutable.replace_semantic_paragraph(0, "after compact audit")?;
        let candidate = Document::from_bytes(mutable.to_bytes_content_only()?)?;

        audit_changed_xml_is_compact(&predecessor, candidate.transaction_package())?;
        assert!(Arc::ptr_eq(&source_bytes, &predecessor.shared_bytes()));
        Ok(())
    }
}
