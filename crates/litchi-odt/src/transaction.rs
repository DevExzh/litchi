//! Immutable, source-bound transactions for packaged ODT documents.
//!
//! This is the safe package-level mutation boundary. It intentionally stages
//! only operations that keep the authoritative XML snapshot intact: callers
//! get exact no-op bytes, source-checked reversible patches, and a complete
//! compact-XML audit before publication. Broader structural `MutableDocument`
//! operations remain available separately while their opaque-content
//! preservation contracts are migrated to this transaction surface.

#![deny(clippy::expect_used, clippy::unwrap_used)]

use crate::{Document, mutable::MutableDocument};
use litchi_core::{
    BlobBundle, BlobId, BlobLimits, Error, ForwardOnly, History as CoreHistory, JoinedSubEdits,
    Patch as CorePatch, PatchLimits, PatchOperation, Result, Reversible, ReversibleOperation,
    SubEdit, ThreeWayMergePlan,
};
use serde_json::Value;
use std::{collections::BTreeMap, sync::Arc};

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
        // Validate a bounded copy so the retained snapshot is the exact
        // caller-provided package, not a writer-normalized representation.
        Document::from_bytes(copy_bytes(&bytes)?)?;
        Ok(Self {
            bytes: Arc::new(bytes),
        })
    }

    /// Captures the exact bytes backing an already validated document.
    pub fn from_document(document: &Document) -> Result<Self> {
        let bytes = copy_bytes(document.original_bytes())?;
        Ok(Self {
            bytes: Arc::new(bytes),
        })
    }

    /// Returns the exact package bytes represented by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Reopens this immutable snapshot for semantic inspection.
    pub fn document(&self) -> Result<Document> {
        Document::from_bytes(copy_bytes(self.as_bytes())?)
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

        ensure_editable_envelope(&self.source)?;
        let mut document = self.source.document()?;
        let mut results = Vec::new();
        results
            .try_reserve_exact(self.operations.len())
            .map_err(|source| Error::Allocation {
                resource: "ODT transaction results",
                source,
            })?;
        for operation in &self.operations {
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
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Index(*index)
                },
                Operation::ReplaceParagraph { index, text } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.replace_semantic_paragraph(*index, text)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::RemoveParagraph { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.remove_semantic_paragraph(*index)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::AppendRun {
                    paragraph,
                    text,
                    style_name,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_semantic_run(*paragraph, text, style_name.as_deref())?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::AppendHyperlink {
                    paragraph,
                    href,
                    text,
                } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_semantic_hyperlink(*paragraph, href, text)?;
                    document = Document::from_bytes(mutable.to_bytes()?)?;
                    OperationResult::Unit
                },
                Operation::AppendLineBreak { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_line_break_at(Position::new(*index))?;
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
                Operation::ReplaceEmbeddedChart { index, definition } => {
                    document.replace_embedded_chart(*index, definition)?;
                    OperationResult::Unit
                },
                Operation::RemoveEmbeddedChart { index } => {
                    document.remove_embedded_chart(*index)?;
                    OperationResult::Unit
                },
                Operation::AddEmbeddedResource { resource } => {
                    OperationResult::Index(document.add_embedded_resource(resource)?)
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
            results.push(result);
        }
        let bytes = copy_bytes(document.original_bytes())?;
        ensure_package_size(bytes.len(), "ODT transaction output")?;
        audit_compact_xml(&bytes)?;
        let after = if bytes == self.source.as_bytes() {
            self.source.clone()
        } else {
            Snapshot::from_bytes(bytes)?
        };
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
        let weight = u64::try_from(self.current().as_bytes().len()).map_err(|_| {
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
    AddNestedForm {
        parent_form: usize,
        form: crate::package::forms::AuthoredForm,
    },
    AddFormControl {
        form_index: usize,
        control: crate::package::forms::AuthoredFormControl,
    },
    ReplaceFormControl {
        index: usize,
        control: crate::package::forms::AuthoredFormControl,
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
    ReplaceEmbeddedChart {
        index: usize,
        definition: crate::odc::Definition,
    },
    RemoveEmbeddedChart {
        index: usize,
    },
    AddEmbeddedResource {
        resource: crate::package::embedded::EmbeddedResource,
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

/// The typed result produced by one staged operation.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OperationResult {
    /// Mutation completed without allocating an identity.
    Unit,
    /// Semantic index allocated by an insertion operation.
    Index(usize),
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
            let forward = semantic_patch_operation(limits, &source_id, &target_id, operation)?;
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
        BlobLimits::new(1, MAX_PACKAGE_BYTES, MAX_PACKAGE_BYTES),
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
        Operation::RestoreSnapshot
        | Operation::AddRdfGraph { .. }
        | Operation::ReplaceRdfGraph { .. }
        | Operation::RemoveRdfGraph { .. }
        | Operation::SetProtection { .. }
        | Operation::AddRdfTriple { .. }
        | Operation::ReplaceRdfTriple { .. }
        | Operation::RemoveRdfTriple { .. }
        | Operation::MoveRdfTriple { .. }
        | Operation::AddForm { .. }
        | Operation::AddNestedForm { .. }
        | Operation::AddFormControl { .. }
        | Operation::ReplaceFormControl { .. }
        | Operation::RemoveFormControl { .. }
        | Operation::MoveFormControl { .. }
        | Operation::ReplaceForm { .. }
        | Operation::RemoveForm { .. }
        | Operation::MoveForm { .. }
        | Operation::AddEmbeddedChart { .. }
        | Operation::AddEmbeddedChartWithStorage { .. }
        | Operation::ReplaceEmbeddedChart { .. }
        | Operation::RemoveEmbeddedChart { .. }
        | Operation::AddEmbeddedResource { .. }
        | Operation::ReplaceEmbeddedObject { .. }
        | Operation::ReplaceEmbeddedImage { .. }
        | Operation::RemoveEmbeddedObject { .. }
        | Operation::RemoveEmbeddedImage { .. }
        | Operation::MoveEmbeddedObject { .. }
        | Operation::MoveEmbeddedImage { .. }
        | Operation::AddScriptResource { .. }
        | Operation::ReplaceScriptResource { .. }
        | Operation::RemoveScriptResource { .. } => {
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
    if !patch.blobs().is_empty() {
        return Err(invalid_durable_patch());
    }
    for operation in patch.operations() {
        decode_semantic_operation(operation)?;
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
        let decoded = decode_semantic_operation(operation)?;
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

fn decode_semantic_operation(operation: &PatchOperation) -> Result<Operation> {
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
        _ => Err(invalid_durable_patch()),
    }
}

fn parse_target_index(target: &str, prefix: &str, suffix: &str) -> Result<usize> {
    let value = target
        .strip_prefix(prefix)
        .and_then(|value| value.strip_suffix(suffix))
        .ok_or_else(invalid_durable_patch)?;
    if value.is_empty() || (value.len() > 1 && value.starts_with('0')) {
        return Err(invalid_durable_patch());
    }
    value.parse().map_err(|_| invalid_durable_patch())
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
            Operation::RestoreSnapshot
            | Operation::AddRdfGraph { .. }
            | Operation::ReplaceRdfGraph { .. }
            | Operation::RemoveRdfGraph { .. }
            | Operation::SetProtection { .. }
            | Operation::AddRdfTriple { .. }
            | Operation::ReplaceRdfTriple { .. }
            | Operation::RemoveRdfTriple { .. }
            | Operation::MoveRdfTriple { .. }
            | Operation::AddForm { .. }
            | Operation::AddNestedForm { .. }
            | Operation::AddFormControl { .. }
            | Operation::ReplaceFormControl { .. }
            | Operation::RemoveFormControl { .. }
            | Operation::MoveFormControl { .. }
            | Operation::ReplaceForm { .. }
            | Operation::RemoveForm { .. }
            | Operation::MoveForm { .. }
            | Operation::AddEmbeddedChart { .. }
            | Operation::AddEmbeddedChartWithStorage { .. }
            | Operation::ReplaceEmbeddedChart { .. }
            | Operation::RemoveEmbeddedChart { .. }
            | Operation::AddEmbeddedResource { .. }
            | Operation::ReplaceEmbeddedObject { .. }
            | Operation::ReplaceEmbeddedImage { .. }
            | Operation::RemoveEmbeddedObject { .. }
            | Operation::RemoveEmbeddedImage { .. }
            | Operation::MoveEmbeddedObject { .. }
            | Operation::MoveEmbeddedImage { .. }
            | Operation::AddScriptResource { .. }
            | Operation::ReplaceScriptResource { .. }
            | Operation::RemoveScriptResource { .. } => writes.push("/package".to_string()),
        }
    }
    (Vec::new(), writes)
}

fn envelope_kind(snapshot: &Snapshot) -> Result<EnvelopeKind> {
    let package = crate::core::OwnedPackage::from_bytes(copy_bytes(snapshot.as_bytes())?)?;
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

fn audit_compact_xml(bytes: &[u8]) -> Result<()> {
    let package = crate::core::OwnedPackage::from_bytes(copy_bytes(bytes)?)?;
    let archive = package.package()?;
    for path in archive.files()? {
        let xml_media_type = archive
            .manifest()
            .entries
            .get(&path)
            .is_some_and(|entry| entry.media_type.contains("xml"));
        if path.ends_with(".xml") || path.ends_with(".rdf") || xml_media_type {
            let xml = package.get_file(&path)?;
            let limits = litchi_odf_common::compact_xml::Limits::new(MAX_PACKAGE_BYTES, 4_096)
                .map_err(Error::from)?;
            litchi_odf_common::compact_xml::validate_with_limits(&xml, limits)
                .map_err(Error::from)?;
        }
    }
    Ok(())
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
