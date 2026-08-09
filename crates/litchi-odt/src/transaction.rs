//! Immutable, source-bound transactions for packaged ODT documents.
//!
//! This is the safe package-level mutation boundary. It intentionally stages
//! only operations that keep the authoritative XML snapshot intact: callers
//! get exact no-op bytes, source-checked reversible patches, and a complete
//! compact-XML audit before publication. Broader structural `MutableDocument`
//! operations remain available separately while their opaque-content
//! preservation contracts are migrated to this transaction surface.

use crate::{Document, mutable::MutableDocument};
use litchi_core::{Error, Result};
use std::sync::Arc;

/// Shared zero-based semantic collection position.
pub use litchi_core::Position;

const MAX_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_OPERATIONS: usize = 1_024;

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
        self.remove_rdf_triple(path, position.get())
    }

    /// Stages an RDF assertion move; equal selectors are an exact no-op.
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
        self.move_rdf_triple(path, from.get(), to.get())
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
    pub fn remove_form_control(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveFormControl { index })
    }

    /// Stages removal of a form control selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_form_control_at(&mut self, position: Position) -> Result<&mut Self> {
        self.remove_form_control(position.get())
    }

    /// Stages a form-control move; equal selectors are an exact no-op.
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
        self.move_form_control(from.get(), to.get())
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
    pub fn remove_form(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveForm { index })
    }

    /// Stages removal of a form selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_form_at(&mut self, position: Position) -> Result<&mut Self> {
        self.remove_form(position.get())
    }

    /// Stages a form move; equal selectors are an exact no-op.
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
        self.move_form(from.get(), to.get())
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
    pub fn remove_embedded_chart(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedChart { index })
    }

    /// Stages removal of an embedded chart selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_chart_at(&mut self, position: Position) -> Result<&mut Self> {
        self.remove_embedded_chart(position.get())
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
    pub fn remove_embedded_object(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedObject { index })
    }

    /// Stages removal of an embedded object selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_object_at(&mut self, position: Position) -> Result<&mut Self> {
        self.remove_embedded_object(position.get())
    }

    /// Stages removal of an embedded image selected in document order.
    pub fn remove_embedded_image(&mut self, index: usize) -> Result<&mut Self> {
        self.push(Operation::RemoveEmbeddedImage { index })
    }

    /// Stages removal of an embedded image selected by a checked position.
    ///
    /// # Errors
    ///
    /// Returns an error if the transaction operation limit has been reached.
    pub fn remove_embedded_image_at(&mut self, position: Position) -> Result<&mut Self> {
        self.remove_embedded_image(position.get())
    }

    /// Stages an embedded-object move; equal selectors are an exact no-op.
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
        self.move_embedded_object(from.get(), to.get())
    }

    /// Stages an embedded-image move; equal selectors are an exact no-op.
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
        self.move_embedded_image(from.get(), to.get())
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
    #[allow(
        deprecated,
        reason = "transaction composition temporarily routes through validated legacy codecs"
    )]
    pub fn commit(self) -> Result<Commit> {
        if self.operations.is_empty() {
            return Ok(Commit::new(self.source.clone(), self.source, Vec::new()));
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
            let result = match operation {
                Operation::AppendLineBreak { index } => {
                    let mut mutable = MutableDocument::from_document(document)?;
                    mutable.append_line_break(*index)?;
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
        Ok(Commit::new(self.source, after, results))
    }
}

#[derive(Clone)]
enum Operation {
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
    fn new(before: Snapshot, snapshot: Snapshot, results: Vec<OperationResult>) -> Self {
        Self {
            patch: Patch {
                before,
                after: snapshot.clone(),
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
        }
    }
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    ensure_package_size(source.len(), "ODT transaction package")?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT transaction package",
            source,
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
    let package = crate::core::OwnedPackage::from_bytes(copy_bytes(snapshot.as_bytes())?)?;
    let archive = package.package()?;
    if archive
        .manifest()
        .entries
        .values()
        .any(|entry| entry.encryption.is_some())
    {
        return Err(Error::Unsupported(
            "packaged ODT transactions preserve encrypted snapshots only as exact no-ops"
                .to_string(),
        ));
    }
    if archive.has_file("META-INF/documentsignatures.xml")
        || archive.has_file("META-INF/macrosignatures.xml")
    {
        return Err(Error::Unsupported(
            "packaged ODT transactions preserve signed snapshots only as exact no-ops".to_string(),
        ));
    }
    Ok(())
}
