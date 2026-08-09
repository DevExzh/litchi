//! Source-checked, failure-atomic ODP slide, shape, media, and RDF edits.

use super::mutable::MutablePresentation;
use crate::core::OwnedPackage;
use crate::{Presentation, Reference, Shape, Slide};
use litchi_core::{Error, Result};
use std::fs::File;
use std::io::Read;
use std::path::Path;
use std::sync::Arc;
use xml_minifier::audit;

const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
const MAX_DRAFT_BYTES: usize = 128 * 1024 * 1024;
const MAX_SLIDES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_PARTS: usize = 65_536;

/// An immutable presentation package and its parsed slide projection.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<Vec<u8>>,
    resource_bytes: usize,
    slides: Arc<[Slide]>,
}

impl Snapshot {
    /// Open an ODP editing snapshot from a path.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be read or is not a bounded valid ODP package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(read_bounded(path.as_ref(), MAX_PACKAGE_BYTES)?)
    }

    /// Parse owned ODP package bytes into a source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for an oversized, malformed, encrypted, or non-ODP package.
    pub fn from_bytes(source_bytes: Vec<u8>) -> Result<Self> {
        if source_bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP editing package exceeds the 128 MiB limit");
        }
        Self::from_shared_bytes(Arc::new(source_bytes))
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP editing package exceeds the 128 MiB limit");
        }
        let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
        Self::from_owned_package(bytes, package)
    }

    fn from_owned_package(bytes: Arc<Vec<u8>>, package: OwnedPackage) -> Result<Self> {
        let presentation = Presentation::from_owned_package(package)?;
        let slides = presentation.slides()?;
        if slides.len() > MAX_SLIDES {
            return invalid("ODP editing snapshot exceeds the slide-count limit");
        }
        let resource_bytes = slides_resource(&slides)?;
        if resource_bytes > MAX_DRAFT_BYTES {
            return invalid("ODP editing snapshot exceeds the aggregate draft limit");
        }
        Ok(Self {
            bytes,
            resource_bytes,
            slides: Arc::from(slides),
        })
    }

    /// Borrow the exact source package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Borrow the immutable parsed slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Select a slide by checked zero-based position or exact title.
    ///
    /// # Errors
    ///
    /// Returns an error when a title matches more than one slide.
    pub fn slide<'a, S>(&self, selector: S) -> Result<Option<&Slide>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.slides, selector.into())
            .map(|selected| selected.map(|position| &self.slides[position]))
    }

    /// Start an isolated transaction over a detached staging engine.
    ///
    /// # Errors
    ///
    /// Returns an error when the exact source package cannot be reparsed for staging.
    pub fn transaction(&self) -> Result<Transaction> {
        let presentation = Presentation::from_shared_bytes(Arc::clone(&self.bytes))?;
        ensure_editable_source(presentation.owned_package())?;
        Ok(Transaction {
            source: self.clone(),
            draft: MutablePresentation::from_presentation(&presentation)?,
            changed: false,
            rdf: None,
            media_bytes: 0,
            resource_bytes: self.resource_bytes,
            source_resource_bytes: self.resource_bytes,
        })
    }

    /// Materialize this snapshot as the ordinary read facade.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package can no longer be parsed.
    pub fn to_presentation(&self) -> Result<Presentation> {
        Presentation::from_shared_bytes(Arc::clone(&self.bytes))
    }
}

/// A semantic slide selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Checked zero-based position.
    Index(usize),
    /// Exact slide title. Duplicate titles are an ambiguity error.
    Title(&'a str),
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Title(value)
    }
}

/// An isolated mutable draft that cannot alter its source snapshot.
pub struct Transaction {
    source: Snapshot,
    draft: MutablePresentation,
    changed: bool,
    rdf: Option<RdfDraft>,
    media_bytes: usize,
    resource_bytes: usize,
    source_resource_bytes: usize,
}

#[derive(Clone)]
enum RdfOperation {
    AddGraph {
        path: String,
        triples: Vec<crate::rdf::Triple>,
    },
    ReplaceGraph {
        path: String,
        triples: Vec<crate::rdf::Triple>,
    },
    RemoveGraph {
        path: String,
    },
    AddTriple {
        path: String,
        triple: crate::rdf::Triple,
    },
    ReplaceTriple {
        path: String,
        index: usize,
        triple: crate::rdf::Triple,
    },
    RemoveTriple {
        path: String,
        index: usize,
    },
    MoveTriple {
        path: String,
        from: usize,
        to: usize,
    },
}

struct RdfDraft {
    bytes: Arc<Vec<u8>>,
    original_graphs: Vec<crate::rdf::Graph>,
    graphs: Vec<crate::rdf::Graph>,
    operations: Vec<RdfOperation>,
}

impl Transaction {
    /// Borrow the current staged slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        self.draft.slides()
    }

    /// Append a compact title/body slide.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid text, exhausted limits, or ambiguous retained bindings.
    pub fn add(&mut self, title: &str, text: &str) -> Result<()> {
        Self::check_text(title, text)?;
        self.check_structure_edit()?;
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }
        let candidate = self.resource_candidate(0, text_resource(title, text)?)?;
        self.draft
            .insert_slide(self.draft.slides().len(), title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(())
    }

    /// Insert a slide before a selected source-snapshot anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid text, exhausted limits, or retained
    /// declarations whose identity closure cannot be updated losslessly.
    pub fn add_before<'a, S>(&mut self, selector: S, title: &str, text: &str) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        Self::check_text(title, text)?;
        self.check_structure_edit()?;
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let candidate = self.resource_candidate(0, text_resource(title, text)?)?;
        self.draft.insert_slide(index, title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Replace one supported slide's title and body text.
    ///
    /// A pristine page retained from an opened package is refused because its
    /// unmodelled children cannot yet be proven lossless under regeneration.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid text, or a preservation-only page.
    pub fn replace<'a, S>(&mut self, selector: S, title: &str, text: &str) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        Self::check_text(title, text)?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let slide = &self.draft.slides()[index];
        if slide.title.as_deref() == Some(title) && slide.text == text {
            return Ok(Some(()));
        }
        self.check_slide_rewrite(index)?;
        let removed = slide_primary_text_resource(slide)?;
        let candidate = self.resource_candidate(removed, text_resource(title, text)?)?;
        self.draft.update_slide(index, title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Remove one selected slide and return its staged semantic value.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or an unresolved page/declaration reference.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Option<Slide>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_structure_edit()?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let removed_bytes = slide_resource(&self.draft.slides()[index])?;
        let candidate = self.resource_candidate(removed_bytes, 0)?;
        let removed = self.draft.remove_slide(index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(removed))
    }

    /// Append a typed shape to one supported slide.
    ///
    /// Hyperlinks, actions, event bindings, and media references remain inert
    /// metadata; this operation never follows or executes them.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid shape, or preservation-only page.
    pub fn add_shape<'a, S>(&mut self, selector: S, shape: Shape) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        self.check_slide_rewrite(index)?;
        let candidate = self.resource_candidate(0, shape_resource(&shape)?)?;
        super::builder::Builder::generate_shape_xml(
            &shape,
            self.draft.slides()[index].shapes.len(),
        )?;
        self.draft.add_shape(index, shape)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Remove a shape by checked zero-based position from a selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or preservation-only page.
    pub fn remove_shape<'a, S>(&mut self, selector: S, shape_index: usize) -> Result<Option<Shape>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(slide_index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        if shape_index >= self.draft.slides()[slide_index].shapes.len() {
            return Ok(None);
        }
        self.check_slide_rewrite(slide_index)?;
        let removed_bytes = shape_resource(&self.draft.slides()[slide_index].shapes[shape_index])?;
        let candidate = self.resource_candidate(removed_bytes, 0)?;
        let removed = self.draft.remove_shape(slide_index, shape_index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(removed))
    }

    /// Clear the title, body, and shapes of one supported slide.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or preservation-only page.
    pub fn clear<'a, S>(&mut self, selector: S) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let slide = &self.draft.slides()[index];
        if slide.title.is_none() && slide.text.is_empty() && slide.shapes.is_empty() {
            return Ok(Some(()));
        }
        self.check_slide_rewrite(index)?;
        let removed = slide_primary_resource(slide)?;
        let candidate = self.resource_candidate(removed, 0)?;
        self.draft.clear_slide(index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Add bounded package-contained media and return its inert reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or colliding path, media type, or oversized payload.
    pub fn embed_media(
        &mut self,
        path: &str,
        payload: &[u8],
        media_type: &str,
    ) -> Result<Reference> {
        crate::model::media::validate_package_media_path(path)?;
        let addition = path
            .len()
            .checked_add(media_type.len())
            .and_then(|value| value.checked_add(payload.len()))
            .ok_or_else(|| invalid_error("ODP embedded media resource size overflow"))?;
        let media_bytes = self
            .media_bytes
            .checked_add(addition)
            .ok_or_else(|| invalid_error("ODP aggregate media size overflow"))?;
        self.check_projected(self.resource_bytes, media_bytes)?;
        let reference = self.draft.embed_media(
            try_owned_str(path, "ODP media path")?,
            try_owned_bytes(payload, "ODP media payload")?,
            try_owned_str(media_type, "ODP media type")?,
        )?;
        self.media_bytes = media_bytes;
        self.changed = true;
        Ok(reference)
    }

    /// Inspect the RDF metadata graphs in the current transaction draft.
    ///
    /// The inventory is loaded lazily so slide-only transactions do not reject
    /// unrelated malformed metadata that they never touch.
    ///
    /// # Errors
    ///
    /// Returns an error when a declared RDF part is malformed, dangling, or over budget.
    pub fn rdf_graphs(&mut self) -> Result<&[crate::rdf::Graph]> {
        self.ensure_rdf()?;
        self.rdf
            .as_ref()
            .map(|draft| draft.graphs.as_slice())
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))
    }

    /// Add one RDF metadata graph to this package transaction.
    ///
    /// A missing preferred path is resolved immediately to a collision-free,
    /// deterministic package path which is retained by the transaction.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe or colliding path, invalid triples, or a package limit.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let package = self.rdf_package()?;
        let (bytes, path) = crate::rdf::add_graph(&package, preferred_path, triples)?;
        self.stage_rdf(
            bytes,
            RdfOperation::AddGraph {
                path: path.clone(),
                triples: triples.to_vec(),
            },
        )?;
        Ok(path)
    }

    /// Replace all triples in one RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, invalid triples, or a package limit.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::replace_graph(&package, path, triples)?;
        self.stage_rdf(
            bytes,
            RdfOperation::ReplaceGraph {
                path: path.to_string(),
                triples: triples.to_vec(),
            },
        )
    }

    /// Remove one RDF metadata graph after dependency validation.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph or an incoming graph reference.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::remove_graph(&package, path)?;
        self.stage_rdf(
            bytes,
            RdfOperation::RemoveGraph {
                path: path.to_string(),
            },
        )
    }

    /// Append one RDF triple and return its checked projected index.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, invalid triple, or graph limit.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let package = self.rdf_package()?;
        let (bytes, index) = crate::rdf::add_triple(&package, path, triple)?;
        self.stage_rdf(
            bytes,
            RdfOperation::AddTriple {
                path: path.to_string(),
                triple: triple.clone(),
            },
        )?;
        Ok(index)
    }

    /// Replace one RDF triple selected by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, out-of-range position, changed subject, or invalid
    /// triple.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::replace_triple(&package, path, index, triple)?;
        self.stage_rdf(
            bytes,
            RdfOperation::ReplaceTriple {
                path: path.to_string(),
                index,
                triple: triple.clone(),
            },
        )
    }

    /// Remove one RDF triple selected by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph or out-of-range position.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::remove_triple(&package, path, index)?;
        self.stage_rdf(
            bytes,
            RdfOperation::RemoveTriple {
                path: path.to_string(),
                index,
            },
        )
    }

    /// Move one RDF triple to another checked zero-based position within its subject.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, out-of-range position, or subject mismatch.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::move_triple(&package, path, from, to)?;
        self.stage_rdf(
            bytes,
            RdfOperation::MoveTriple {
                path: path.to_string(),
                from,
                to,
            },
        )
    }

    /// Validate, serialize, reparse, and atomically publish the staged draft.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, bounded serialization, package parsing, or semantic
    /// readback fails. The source snapshot is never changed.
    pub fn commit(self) -> Result<Commit> {
        let rdf_changed = self
            .rdf
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        if !self.changed && !rdf_changed {
            return Ok(Commit::unchanged(self.source));
        }
        let mut bytes = if self.changed {
            Arc::new(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?)
        } else {
            Arc::clone(&self.source.bytes)
        };
        if let Some(rdf) = &self.rdf
            && !rdf.operations.is_empty()
        {
            if self.changed {
                for operation in &rdf.operations {
                    let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
                    bytes = Arc::new(apply_rdf_operation(&package, operation)?);
                    if bytes.len() > MAX_PACKAGE_BYTES {
                        return invalid("ODP RDF transaction exceeds the 128 MiB package limit");
                    }
                }
            } else {
                bytes = Arc::clone(&rdf.bytes);
            }
        }
        let reopened = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
        validate_compact_xml_parts(&reopened)?;
        self.draft.verify_embedded_media(&reopened)?;
        if let Some(rdf) = &self.rdf
            && crate::rdf::graphs(&reopened)? != rdf.graphs
        {
            return invalid("ODP transaction RDF readback differs from the staged graph model");
        }
        let snapshot = Snapshot::from_owned_package(bytes, reopened)?;
        if snapshot.slides() != self.draft.slides() {
            return invalid("ODP transaction readback differs from the staged slide model");
        }
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn ensure_rdf(&mut self) -> Result<()> {
        if self.rdf.is_none() {
            let package = OwnedPackage::from_shared_bytes(Arc::clone(&self.source.bytes))?;
            let graphs = crate::rdf::graphs(&package)?;
            self.rdf = Some(RdfDraft {
                bytes: Arc::clone(&self.source.bytes),
                original_graphs: graphs.clone(),
                graphs,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn rdf_package(&mut self) -> Result<OwnedPackage> {
        self.ensure_rdf()?;
        let bytes = self
            .rdf
            .as_ref()
            .map(|draft| Arc::clone(&draft.bytes))
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))?;
        OwnedPackage::from_shared_bytes(bytes)
    }

    fn stage_rdf(&mut self, bytes: Vec<u8>, operation: RdfOperation) -> Result<()> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP RDF transaction exceeds the 128 MiB package limit");
        }
        let candidate = Arc::new(bytes);
        let presentation = Presentation::from_shared_bytes(Arc::clone(&candidate))?;
        let graphs = crate::rdf::graphs(presentation.owned_package())?;
        let source_bytes = Arc::clone(&self.source.bytes);
        let draft = self
            .rdf
            .as_mut()
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))?;
        if graphs == draft.graphs {
            return Ok(());
        }
        if graphs == draft.original_graphs {
            draft.bytes = source_bytes;
            draft.graphs = graphs;
            draft.operations.clear();
            return Ok(());
        }
        draft.bytes = candidate;
        draft.graphs = graphs;
        draft.operations.push(operation);
        Ok(())
    }

    fn check_text(title: &str, text: &str) -> Result<()> {
        let size = title
            .len()
            .checked_add(text.len())
            .ok_or_else(|| invalid_error("ODP slide text size overflow"))?;
        if size > MAX_TEXT_BYTES {
            return invalid("ODP slide text exceeds the 16 MiB limit");
        }
        if title.chars().chain(text.chars()).any(|value| {
            !matches!(
                value,
                '\u{9}'
                    | '\u{A}'
                    | '\u{D}'
                    | '\u{20}'..='\u{D7FF}'
                    | '\u{E000}'..='\u{FFFD}'
                    | '\u{10000}'..='\u{10FFFF}'
            )
        }) {
            return invalid("ODP slide text contains a character forbidden by XML 1.0");
        }
        Ok(())
    }

    fn check_structure_edit(&self) -> Result<()> {
        if self.draft.has_source_declarations() {
            return unsupported(
                "ODP structural editing with retained header/footer declarations is not lossless",
            );
        }
        Ok(())
    }

    fn check_slide_rewrite(&self, index: usize) -> Result<()> {
        if self.draft.retains_source_slide(index) {
            return unsupported(
                "ODP retained slide contains XML that cannot yet be proven lossless under rewrite",
            );
        }
        Ok(())
    }

    fn resource_candidate(&self, removed: usize, added: usize) -> Result<usize> {
        let candidate = bounded_candidate(self.resource_bytes, removed, added, MAX_DRAFT_BYTES)?;
        self.check_projected(candidate, self.media_bytes)?;
        Ok(candidate)
    }

    fn check_projected(&self, resource_bytes: usize, media_bytes: usize) -> Result<()> {
        let _projected = projected_size(
            self.source.bytes().len(),
            self.source_resource_bytes,
            resource_bytes,
            media_bytes,
            MAX_PACKAGE_BYTES,
        )?;
        Ok(())
    }
}

/// A validated publication result containing a snapshot and reversible patch.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        let patch = Patch {
            before: snapshot.clone(),
            after: snapshot.clone(),
        };
        Self {
            snapshot,
            patch,
            changed: false,
        }
    }

    /// Return whether publication rebuilt the package.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-byte-source-checked reversible ODP package patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Apply this patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not byte-for-byte identical to the patch source.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !same_source(&self.before, source) {
            return invalid("stale ODP presentation patch source");
        }
        Ok(self.after.clone())
    }

    /// Return the patch that restores the exact pre-commit package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return whether this patch preserves the exact package bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        same_source(&self.before, &self.after)
    }
}

fn select(slides: &[Slide], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < slides.len()).then_some(index)),
        Selector::Title(title) => {
            let mut selected = None;
            for (index, slide) in slides.iter().enumerate() {
                if slide.title.as_deref() != Some(title) {
                    continue;
                }
                if selected.is_some() {
                    return invalid("ODP slide title selector is ambiguous");
                }
                selected = Some(index);
            }
            Ok(selected)
        },
    }
}

fn same_source(left: &Snapshot, right: &Snapshot) -> bool {
    Arc::ptr_eq(&left.bytes, &right.bytes) || left.bytes == right.bytes
}

fn ensure_editable_source(package: &OwnedPackage) -> Result<()> {
    let archive = package.package()?;
    if archive.manifest().has_encrypted_entries() {
        return unsupported(
            "ODP package transactions refuse encrypted package entries; decrypt to a new unsigned package first",
        );
    }
    if archive.has_file("META-INF/documentsignatures.xml")
        || archive.has_file("META-INF/macrosignatures.xml")
    {
        return unsupported(
            "ODP package transactions refuse signed packages because mutation would invalidate their signatures",
        );
    }
    Ok(())
}

fn apply_rdf_operation(package: &OwnedPackage, operation: &RdfOperation) -> Result<Vec<u8>> {
    match operation {
        RdfOperation::AddGraph { path, triples } => {
            let (bytes, actual_path) = crate::rdf::add_graph(package, Some(path), triples)?;
            if actual_path != *path {
                return invalid("ODP RDF replay resolved a different graph path");
            }
            Ok(bytes)
        },
        RdfOperation::ReplaceGraph { path, triples } => {
            crate::rdf::replace_graph(package, path, triples)
        },
        RdfOperation::RemoveGraph { path } => crate::rdf::remove_graph(package, path),
        RdfOperation::AddTriple { path, triple } => {
            crate::rdf::add_triple(package, path, triple).map(|(bytes, _)| bytes)
        },
        RdfOperation::ReplaceTriple {
            path,
            index,
            triple,
        } => crate::rdf::replace_triple(package, path, *index, triple),
        RdfOperation::RemoveTriple { path, index } => {
            crate::rdf::remove_triple(package, path, *index)
        },
        RdfOperation::MoveTriple { path, from, to } => {
            crate::rdf::move_triple(package, path, *from, *to)
        },
    }
}

fn bounded_candidate(current: usize, removed: usize, added: usize, limit: usize) -> Result<usize> {
    let candidate = current
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| invalid_error("ODP aggregate draft accounting overflow"))?;
    if candidate > limit {
        return invalid("ODP transaction exceeds the aggregate draft limit");
    }
    Ok(candidate)
}

fn projected_size(
    source_bytes: usize,
    source_resource_bytes: usize,
    resource_bytes: usize,
    media_bytes: usize,
    limit: usize,
) -> Result<usize> {
    let growth = resource_bytes.saturating_sub(source_resource_bytes);
    let projected = source_bytes
        .checked_add(growth)
        .and_then(|value| value.checked_add(media_bytes))
        .ok_or_else(|| invalid_error("ODP projected package size overflow"))?;
    if projected > limit {
        return invalid("ODP transaction exceeds the projected package limit");
    }
    Ok(projected)
}

fn try_owned_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.extend_from_slice(value);
    Ok(owned)
}

fn try_owned_str(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.push_str(value);
    Ok(owned)
}

fn read_bounded(path: &Path, limit: usize) -> Result<Vec<u8>> {
    let mut file = File::open(path)?;
    let declared = file.metadata()?.len();
    let limit_u64 = u64::try_from(limit).unwrap_or(u64::MAX);
    if declared > limit_u64 {
        return invalid(format!(
            "ODP editing package exceeds the {limit}-byte limit"
        ));
    }
    let capacity = usize::try_from(declared).map_err(|error| {
        invalid_error(format!(
            "ODP package length does not fit this platform: {error}"
        ))
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODP bounded package input",
            source,
        })?;
    let mut buffer = [0u8; 8 * 1024];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        let candidate = bytes
            .len()
            .checked_add(read)
            .ok_or_else(|| invalid_error("ODP bounded package input size overflow"))?;
        if candidate > limit {
            return invalid(format!(
                "ODP editing package exceeds the {limit}-byte limit"
            ));
        }
        bytes
            .try_reserve_exact(read)
            .map_err(|source| Error::Allocation {
                resource: "ODP bounded package input",
                source,
            })?;
        bytes.extend_from_slice(&buffer[..read]);
    }
    Ok(bytes)
}

fn slides_resource(slides: &[Slide]) -> Result<usize> {
    slides.iter().try_fold(0usize, |total, slide| {
        total
            .checked_add(slide_resource(slide)?)
            .ok_or_else(|| invalid_error("ODP aggregate slide resource size overflow"))
    })
}

fn slide_resource(slide: &Slide) -> Result<usize> {
    let primary = slide_primary_resource(slide)?;
    primary
        .checked_add(slide.notes.as_ref().map_or(0, String::len))
        .ok_or_else(|| invalid_error("ODP slide resource size overflow"))
}

fn slide_primary_resource(slide: &Slide) -> Result<usize> {
    slide_primary_text_resource(slide)?
        .checked_add(slide.shapes.iter().try_fold(0usize, |total, shape| {
            total
                .checked_add(shape_resource(shape)?)
                .ok_or_else(|| invalid_error("ODP shape resource size overflow"))
        })?)
        .ok_or_else(|| invalid_error("ODP slide resource size overflow"))
}

fn slide_primary_text_resource(slide: &Slide) -> Result<usize> {
    text_resource(slide.title.as_deref().unwrap_or_default(), &slide.text)
}

fn text_resource(title: &str, text: &str) -> Result<usize> {
    title
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid_error("ODP text resource size overflow"))
}

fn shape_resource(root: &Shape) -> Result<usize> {
    let mut stack = Vec::new();
    stack.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODP shape resource stack",
        source,
    })?;
    stack.push(root);
    let mut total = 0usize;
    while let Some(shape) = stack.pop() {
        total = total
            .checked_add(shape.text.len())
            .and_then(|value| value.checked_add(shape.name().map_or(0, str::len)))
            .ok_or_else(|| invalid_error("ODP shape resource size overflow"))?;
        if stack.len().saturating_add(shape.children().len()) > MAX_SLIDES {
            return invalid("ODP shape resource traversal exceeds the node limit");
        }
        stack
            .try_reserve(shape.children().len())
            .map_err(|source| Error::Allocation {
                resource: "ODP shape resource stack",
                source,
            })?;
        stack.extend(shape.children());
    }
    Ok(total)
}

fn validate_compact_xml_parts(package: &OwnedPackage) -> Result<()> {
    let mut part_count = 0usize;
    let mut aggregate_bytes = 0usize;
    for path in package.files()? {
        if !path.rsplit_once('.').is_some_and(|(_, extension)| {
            extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("rdf")
        }) {
            continue;
        }
        let payload = package.get_file(&path)?;
        part_count = part_count
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP XML part count overflow"))?;
        aggregate_bytes = aggregate_bytes
            .checked_add(payload.len())
            .ok_or_else(|| invalid_error("ODP aggregate XML size overflow"))?;
        if part_count > MAX_XML_PARTS || aggregate_bytes > MAX_PACKAGE_BYTES {
            return invalid("ODP XML package audit exceeds its aggregate limit");
        }
        let limits = audit::Limits::new(
            MAX_PACKAGE_BYTES,
            512,
            1_000_000,
            250_000,
            MAX_TEXT_BYTES,
            MAX_PACKAGE_BYTES,
        )
        .map_err(|source| invalid_error(format!("invalid ODP XML audit limits: {source}")))?;
        let _report = audit::verify(&payload, limits).map_err(|source| match source {
            audit::Error::NotCompact(_) => {
                Error::Unsupported(format!("ODP XML part '{path}' is not compact: {source}"))
            },
            audit::Error::Limit { .. }
            | audit::Error::Encoding { .. }
            | audit::Error::Malformed { .. }
            | audit::Error::Doctype { .. }
            | audit::Error::Allocation
            | _ => Error::InvalidFormat(format!("ODP XML part '{path}' failed audit: {source}")),
        })?;
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Unsupported(message.into()))
}

#[cfg(test)]
mod tests {
    use super::{bounded_candidate, projected_size, read_bounded};
    use litchi_core::Result;

    #[test]
    fn bounded_file_reader_accepts_n_and_rejects_n_plus_one() -> Result<()> {
        let file = tempfile::NamedTempFile::new()?;
        std::fs::write(file.path(), b"1234")?;
        assert_eq!(read_bounded(file.path(), 4)?, b"1234");
        std::fs::write(file.path(), b"12345")?;
        assert!(read_bounded(file.path(), 4).is_err());
        Ok(())
    }

    #[test]
    fn aggregate_accounting_accepts_n_and_rejects_n_plus_one() -> Result<()> {
        assert_eq!(bounded_candidate(3, 0, 1, 4)?, 4);
        assert!(bounded_candidate(3, 0, 2, 4).is_err());
        assert_eq!(projected_size(3, 2, 2, 1, 4)?, 4);
        assert!(projected_size(3, 2, 2, 2, 4).is_err());
        Ok(())
    }
}
