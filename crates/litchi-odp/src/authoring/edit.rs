//! Unified source-checked ODP slide, media, chart, design, annotation, and RDF edits.

use super::mutable::MutablePresentation;
use crate::core::OwnedPackage;
use crate::{Presentation, Reference, Shape, Slide};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
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
const DURABLE_PATCH_MAGIC: &[u8; 16] = b"LITCHI-ODP-PATCH";
const DURABLE_PATCH_VERSION: u16 = 1;
const DURABLE_HISTORY_MAGIC: &[u8; 16] = b"LITCHI-ODP-HIST\0";
const DURABLE_HISTORY_VERSION: u16 = 1;
const MAX_DURABLE_HISTORY_BYTES: usize = 512 * 1024 * 1024;
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";

/// Semantic dependency domain touched by a root package patch.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Domain {
    /// Slide, shape, or embedded-media projection.
    Slides,
    /// RDF metadata graphs.
    Rdf,
    /// Embedded chart occurrences or parts.
    Charts,
    /// Presentation layouts, master pages, or their slide assignments.
    Design,
    /// Slide- or shape-anchored annotations.
    Annotations,
}

/// Conservative non-mutating merge assessment for two package patches.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MergePlan {
    conflicts: Vec<Domain>,
}

impl MergePlan {
    /// Return semantic domains requiring an explicit merge decision.
    #[must_use]
    pub fn conflicts(&self) -> &[Domain] {
        &self.conflicts
    }

    /// Return whether the two patches are provably independent at this API layer.
    #[must_use]
    pub fn is_independent(&self) -> bool {
        self.conflicts.is_empty()
    }
}

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
            charts: None,
            design: None,
            annotations: None,
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
    charts: Option<ChartDraft>,
    design: Option<DesignDraft>,
    annotations: Option<AnnotationDraft>,
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

#[derive(Clone)]
enum ChartSelector {
    Index(usize),
    Name(String),
}

#[derive(Clone)]
enum ChartPage {
    Index(usize),
    Name(String),
}

impl ChartSelector {
    fn from_borrowed(selector: crate::charts::Selector<'_>) -> Self {
        match selector {
            crate::charts::Selector::Index(index) => Self::Index(index),
            crate::charts::Selector::Name(name) => Self::Name(name.to_string()),
        }
    }

    fn borrowed(&self) -> crate::charts::Selector<'_> {
        match self {
            Self::Index(index) => crate::charts::Selector::Index(*index),
            Self::Name(name) => crate::charts::Selector::Name(name),
        }
    }
}

impl ChartPage {
    fn from_borrowed(page: crate::charts::Page<'_>) -> Self {
        match page {
            crate::charts::Page::Index(index) => Self::Index(index),
            crate::charts::Page::Name(name) => Self::Name(name.to_string()),
        }
    }

    fn borrowed(&self) -> crate::charts::Page<'_> {
        match self {
            Self::Index(index) => crate::charts::Page::Index(*index),
            Self::Name(name) => crate::charts::Page::Name(name),
        }
    }
}

#[derive(Clone)]
enum ChartOperation {
    Replace {
        selector: ChartSelector,
        part: crate::charts::Part,
    },
    Remove {
        selector: ChartSelector,
    },
    Add {
        page: ChartPage,
        name: String,
        storage: crate::charts::Storage,
        part: crate::charts::Part,
    },
}

struct ChartDraft {
    bytes: Arc<Vec<u8>>,
    original: Vec<crate::charts::Chart>,
    charts: Vec<crate::charts::Chart>,
    operations: Vec<ChartOperation>,
    limits: crate::charts::Limits,
}

#[derive(Clone)]
enum DesignOperation {
    AddLayout(crate::layout::Layout),
    ReplaceLayout(crate::layout::Layout),
    RemoveLayout {
        name: String,
        replacement: Option<String>,
    },
    ReorderLayouts(Vec<String>),
    AddMaster(crate::MasterPage),
    ReplaceMaster(crate::MasterPage),
    RemoveMaster {
        name: String,
        replacement: Option<String>,
    },
    ReorderMasters(Vec<String>),
    AssignSlideMaster {
        slide_index: usize,
        name: Option<String>,
    },
    AssignSlideLayout {
        slide_index: usize,
        name: Option<String>,
    },
}

struct DesignDraft {
    bytes: Arc<Vec<u8>>,
    original_layouts: crate::layout::Collection,
    layouts: crate::layout::Collection,
    original_masters: Vec<crate::MasterPage>,
    masters: Vec<crate::MasterPage>,
    original_pages: crate::page::Collection,
    pages: crate::page::Collection,
    operations: Vec<DesignOperation>,
}

#[derive(Clone)]
enum AnnotationOperation {
    Add {
        anchor: crate::annotation::Anchor,
        annotation: crate::annotation::Annotation,
    },
    Replace {
        index: usize,
        annotation: crate::annotation::Annotation,
    },
    Remove {
        index: usize,
    },
}

struct AnnotationDraft {
    bytes: Arc<Vec<u8>>,
    original: Vec<crate::annotation::Info>,
    annotations: Vec<crate::annotation::Info>,
    operations: Vec<AnnotationOperation>,
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

    /// Inspect embedded charts in the current package transaction draft.
    ///
    /// # Errors
    ///
    /// Returns an error when chart discovery encounters malformed or over-budget content.
    pub fn charts(&mut self) -> Result<&[crate::charts::Chart]> {
        self.ensure_charts()?;
        self.charts
            .as_ref()
            .map(|draft| draft.charts.as_slice())
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))
    }

    /// Replace one embedded chart part selected by exact name or checked position.
    ///
    /// Every occurrence sharing the selected package part is updated together.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector, malformed chart XML, or a limit.
    pub fn replace_chart<'a, S>(&mut self, selector: S, part: crate::charts::Part) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        edit.replace(owned_selector.borrowed(), part.clone())?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().bytes().to_vec(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Replace {
                selector: owned_selector,
                part,
            },
        )
    }

    /// Replace one chart from the complete typed ODF chart authoring model.
    ///
    /// This is the unified-root entry point for cached tables, typed series,
    /// axes, legends, plot-area details, and chart-local styles.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid definition, missing or ambiguous selector, or limit.
    pub fn replace_chart_definition<'a, S>(
        &mut self,
        selector: S,
        definition: &crate::charts::Definition,
    ) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        self.replace_chart(selector, crate::charts::Part::from_definition(definition)?)
    }

    /// Remove one embedded chart selected by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector or malformed chart package content.
    pub fn remove_chart<'a, S>(&mut self, selector: S) -> Result<crate::charts::Chart>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        let removed = edit.remove(owned_selector.borrowed())?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().bytes().to_vec(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Remove {
                selector: owned_selector,
            },
        )?;
        Ok(removed)
    }

    /// Add a named embedded chart to an exact-name or checked-position page.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous page, duplicate chart name, malformed part,
    /// or resource-limit breach.
    pub fn add_chart<'a, P>(
        &mut self,
        page: P,
        name: impl Into<String>,
        storage: crate::charts::Storage,
        part: crate::charts::Part,
    ) -> Result<usize>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let owned_page = ChartPage::from_borrowed(page.into());
        let chart_name = name.into();
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        let index = edit.add(
            owned_page.borrowed(),
            chart_name.clone(),
            storage,
            part.clone(),
        )?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().bytes().to_vec(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Add {
                page: owned_page,
                name: chart_name,
                storage,
                part,
            },
        )?;
        Ok(index)
    }

    /// Add a chart from the complete typed ODF chart authoring model.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid definition, page selector, name, or resource limit.
    pub fn add_chart_definition<'a, P>(
        &mut self,
        page: P,
        name: impl Into<String>,
        storage: crate::charts::Storage,
        definition: &crate::charts::Definition,
    ) -> Result<usize>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        self.add_chart(
            page,
            name,
            storage,
            crate::charts::Part::from_definition(definition)?,
        )
    }

    /// Copy one dependency-closed chart from another immutable presentation snapshot.
    ///
    /// The chart's complete typed part, including chart-local styles and cached data, is
    /// detached from the source. Parts with `xlink:href` dependencies are refused because
    /// this bounded operation cannot prove that their referenced package resources are owned
    /// exclusively by the selected chart. The destination always receives a fresh occurrence
    /// (and, for subdocument storage, a fresh collision-free package path).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous source chart, a dependent chart part,
    /// invalid destination selectors, identity collisions, or resource-limit violations.
    pub fn transfer_chart_from<'a, 'b, S, P>(
        &mut self,
        source: &Snapshot,
        source_chart: S,
        destination_page: P,
        destination_name: impl Into<String>,
        storage: crate::charts::Storage,
    ) -> Result<usize>
    where
        S: Into<crate::charts::Selector<'a>>,
        P: Into<crate::charts::Page<'b>>,
    {
        let inventory = crate::charts::Snapshot::from_shared_bytes(
            Arc::clone(&source.bytes),
            crate::charts::Limits::default(),
        )?;
        let selected = inventory
            .get(source_chart)?
            .ok_or_else(|| invalid_error("ODP source chart selector did not match"))?;
        ensure_chart_transfer_closed(selected.part().xml())?;
        self.add_chart(
            destination_page,
            destination_name,
            storage,
            selected.part().clone(),
        )
    }

    /// Inspect named presentation page layouts in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error when `styles.xml` is missing, malformed, or over budget.
    pub fn layouts(&mut self) -> Result<&crate::layout::Collection> {
        self.ensure_design()?;
        self.design
            .as_ref()
            .map(|draft| &draft.layouts)
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))
    }

    /// Inspect master pages in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error when `styles.xml` is missing, malformed, or over budget.
    pub fn master_pages(&mut self) -> Result<&[crate::MasterPage]> {
        self.ensure_design()?;
        self.design
            .as_ref()
            .map(|draft| draft.masters.as_slice())
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))
    }

    /// Add one named presentation page layout.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or duplicate layout or malformed package dependencies.
    pub fn add_layout(&mut self, layout: &crate::layout::Layout) -> Result<()> {
        self.stage_design_operation(DesignOperation::AddLayout(layout.clone()))
    }

    /// Replace one named presentation page layout.
    ///
    /// # Errors
    ///
    /// Returns an error when the layout is invalid or does not exist.
    pub fn replace_page_layout(&mut self, layout: &crate::layout::Layout) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReplaceLayout(layout.clone()))
    }

    /// Remove one layout and optionally retarget all modeled incoming references.
    ///
    /// # Errors
    ///
    /// Returns an error when either name is invalid or the replacement does not exist.
    pub fn remove_page_layout(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        self.stage_design_operation(DesignOperation::RemoveLayout {
            name: name.to_string(),
            replacement: replacement.map(str::to_string),
        })
    }

    /// Reorder every named presentation layout using an exact dependency-checked name list.
    ///
    /// # Errors
    ///
    /// Returns an error when the list is incomplete, duplicated, or contains an unknown name.
    pub fn reorder_layouts(&mut self, names: &[String]) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReorderLayouts(names.to_vec()))
    }

    /// Add one named master page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML, duplicate identity, or dangling dependencies.
    pub fn add_master_page(&mut self, master: &crate::MasterPage) -> Result<()> {
        self.stage_design_operation(DesignOperation::AddMaster(master.clone()))
    }

    /// Replace one named master page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML, missing identity, or dangling dependencies.
    pub fn replace_master_page(&mut self, master: &crate::MasterPage) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReplaceMaster(master.clone()))
    }

    /// Remove one master page and optionally retarget modeled incoming references.
    ///
    /// # Errors
    ///
    /// Returns an error when either name is invalid or the replacement does not exist.
    pub fn remove_master_page(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        self.stage_design_operation(DesignOperation::RemoveMaster {
            name: name.to_string(),
            replacement: replacement.map(str::to_string),
        })
    }

    /// Reorder every master page using an exact dependency-checked name list.
    ///
    /// # Errors
    ///
    /// Returns an error when the list is incomplete, duplicated, or contains an unknown name.
    pub fn reorder_master_pages(&mut self, names: &[String]) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReorderMasters(names.to_vec()))
    }

    /// Assign or clear a slide's master-page dependency by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide or missing master.
    pub fn assign_slide_master_page(
        &mut self,
        slide_index: usize,
        master_name: Option<&str>,
    ) -> Result<()> {
        self.stage_design_operation(DesignOperation::AssignSlideMaster {
            slide_index,
            name: master_name.map(str::to_string),
        })
    }

    /// Assign or clear a slide's presentation-layout dependency by checked position.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide or missing layout.
    pub fn assign_slide_page_layout(
        &mut self,
        slide_index: usize,
        layout_name: Option<&str>,
    ) -> Result<()> {
        self.stage_design_operation(DesignOperation::AssignSlideLayout {
            slide_index,
            name: layout_name.map(str::to_string),
        })
    }

    /// Inspect slide- and shape-anchored annotations in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed annotation XML, anchors, or resource limits.
    pub fn annotations(&mut self) -> Result<&[crate::annotation::Info]> {
        self.ensure_annotations()?;
        self.annotations
            .as_ref()
            .map(|draft| draft.annotations.as_slice())
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))
    }

    /// Add an annotation at a checked page or uniquely named shape anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous anchor, duplicate name, or invalid body.
    pub fn add_annotation(
        &mut self,
        anchor: &crate::annotation::Anchor,
        annotation: &crate::annotation::Annotation,
    ) -> Result<usize> {
        self.ensure_annotations()?;
        let current = self.annotation_bytes()?;
        let mut presentation = Presentation::from_shared_bytes(current)?;
        let index = presentation.add_annotation(anchor, annotation)?;
        self.stage_annotation(
            Arc::new(presentation.to_bytes()?),
            AnnotationOperation::Add {
                anchor: anchor.clone(),
                annotation: annotation.clone(),
            },
        )?;
        Ok(index)
    }

    /// Replace one annotation selected by checked zero-based document order.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range position, duplicate name, or invalid body.
    pub fn replace_annotation(
        &mut self,
        index: usize,
        annotation: &crate::annotation::Annotation,
    ) -> Result<()> {
        self.ensure_annotations()?;
        let current = self.annotation_bytes()?;
        let mut presentation = Presentation::from_shared_bytes(current)?;
        presentation.replace_annotation(index, annotation)?;
        self.stage_annotation(
            Arc::new(presentation.to_bytes()?),
            AnnotationOperation::Replace {
                index,
                annotation: annotation.clone(),
            },
        )
    }

    /// Remove one annotation selected by checked zero-based document order.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is out of range or package content is malformed.
    pub fn remove_annotation(&mut self, index: usize) -> Result<()> {
        self.ensure_annotations()?;
        let current = self.annotation_bytes()?;
        let mut presentation = Presentation::from_shared_bytes(current)?;
        presentation.remove_annotation(index)?;
        self.stage_annotation(
            Arc::new(presentation.to_bytes()?),
            AnnotationOperation::Remove { index },
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
        let charts_changed = self
            .charts
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let design_changed = self
            .design
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let annotations_changed = self
            .annotations
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let mut domains = Vec::new();
        if self.changed {
            domains.push(Domain::Slides);
        }
        if rdf_changed {
            domains.push(Domain::Rdf);
        }
        if charts_changed {
            domains.push(Domain::Charts);
        }
        if design_changed {
            domains.push(Domain::Design);
        }
        if annotations_changed {
            domains.push(Domain::Annotations);
        }
        if !self.changed
            && !rdf_changed
            && !charts_changed
            && !design_changed
            && !annotations_changed
        {
            return Ok(Commit::unchanged(self.source));
        }
        let mut bytes = if self.changed {
            Arc::new(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?)
        } else {
            Arc::clone(&self.source.bytes)
        };
        if self.changed {
            let slide_candidate = Snapshot::from_shared_bytes(Arc::clone(&bytes))?;
            if slide_candidate.slides() != self.draft.slides() {
                return invalid("ODP transaction readback differs from the staged slide model");
            }
        }
        if let Some(design) = &self.design
            && !design.operations.is_empty()
        {
            if self.changed {
                for operation in &design.operations {
                    bytes = Arc::new(apply_design_operation(Arc::clone(&bytes), operation)?);
                }
            } else {
                bytes = Arc::clone(&design.bytes);
            }
        }
        if let Some(annotations) = &self.annotations
            && !annotations.operations.is_empty()
        {
            if self.changed || design_changed {
                for operation in &annotations.operations {
                    bytes = Arc::new(apply_annotation_operation(Arc::clone(&bytes), operation)?);
                }
            } else {
                bytes = Arc::clone(&annotations.bytes);
            }
        }
        if let Some(rdf) = &self.rdf
            && !rdf.operations.is_empty()
        {
            if self.changed || design_changed || annotations_changed {
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
        if let Some(charts) = &self.charts
            && !charts.operations.is_empty()
        {
            if self.changed || design_changed || annotations_changed || rdf_changed {
                for operation in &charts.operations {
                    bytes = Arc::new(apply_chart_operation(
                        Arc::clone(&bytes),
                        charts.limits,
                        operation,
                    )?);
                    if bytes.len() > MAX_PACKAGE_BYTES {
                        return invalid("ODP chart transaction exceeds the 128 MiB package limit");
                    }
                }
            } else {
                bytes = Arc::clone(&charts.bytes);
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
        if let Some(charts) = &self.charts {
            let reopened_charts =
                crate::charts::Snapshot::from_shared_bytes(Arc::clone(&bytes), charts.limits)?;
            if !root_charts_equal(reopened_charts.charts(), &charts.charts) {
                return invalid("ODP transaction chart readback differs from the staged model");
            }
        }
        if let Some(design) = &self.design {
            let presentation = Presentation::from_shared_bytes(Arc::clone(&bytes))?;
            if presentation.layouts()? != design.layouts
                || !root_masters_equal(&presentation.master_pages()?, &design.masters)
                || !root_design_pages_equal(&presentation.pages()?, &design.pages)
            {
                return invalid("ODP transaction design readback differs from the staged model");
            }
        }
        if let Some(annotations) = &self.annotations {
            let presentation = Presentation::from_shared_bytes(Arc::clone(&bytes))?;
            if !root_annotations_equal(&presentation.annotations()?, &annotations.annotations) {
                return invalid(
                    "ODP transaction annotation readback differs from the staged model",
                );
            }
        }
        let snapshot = Snapshot::from_owned_package(bytes, reopened)?;
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
            domains: Arc::from(domains),
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

    fn ensure_charts(&mut self) -> Result<()> {
        if self.charts.is_none() {
            let limits = crate::charts::Limits::default();
            let snapshot =
                crate::charts::Snapshot::from_shared_bytes(Arc::clone(&self.source.bytes), limits)?;
            let charts = snapshot.charts().to_vec();
            self.charts = Some(ChartDraft {
                bytes: Arc::clone(&self.source.bytes),
                original: charts.clone(),
                charts,
                operations: Vec::new(),
                limits,
            });
        }
        Ok(())
    }

    fn chart_snapshot(&mut self) -> Result<crate::charts::Snapshot> {
        self.ensure_charts()?;
        let (current_bytes, limits, operations) = self
            .charts
            .as_ref()
            .map(|draft| {
                (
                    Arc::clone(&draft.bytes),
                    draft.limits,
                    draft.operations.clone(),
                )
            })
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))?;
        if !self.changed {
            return crate::charts::Snapshot::from_shared_bytes(current_bytes, limits);
        }
        let mut bytes = Arc::new(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?);
        for operation in &operations {
            bytes = Arc::new(apply_chart_operation(bytes, limits, operation)?);
        }
        crate::charts::Snapshot::from_shared_bytes(bytes, limits)
    }

    fn stage_chart(
        &mut self,
        bytes: Vec<u8>,
        charts: Vec<crate::charts::Chart>,
        operation: ChartOperation,
    ) -> Result<()> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP chart transaction exceeds the 128 MiB package limit");
        }
        let source_bytes = Arc::clone(&self.source.bytes);
        let draft = self
            .charts
            .as_mut()
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))?;
        if root_charts_equal(&charts, &draft.charts) {
            return Ok(());
        }
        if root_charts_equal(&charts, &draft.original) {
            draft.bytes = source_bytes;
            draft.charts = charts;
            draft.operations.clear();
            return Ok(());
        }
        draft.bytes = Arc::new(bytes);
        draft.charts = charts;
        draft.operations.push(operation);
        Ok(())
    }

    fn ensure_design(&mut self) -> Result<()> {
        if self.design.is_none() {
            let presentation = Presentation::from_shared_bytes(Arc::clone(&self.source.bytes))?;
            let layouts = presentation.layouts()?;
            let masters = presentation.master_pages()?;
            let pages = presentation.pages()?;
            self.design = Some(DesignDraft {
                bytes: Arc::clone(&self.source.bytes),
                original_layouts: layouts.clone(),
                layouts,
                original_masters: masters.clone(),
                masters,
                original_pages: pages.clone(),
                pages,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn stage_design_operation(&mut self, operation: DesignOperation) -> Result<()> {
        self.ensure_design()?;
        let current = self
            .design
            .as_ref()
            .map(|draft| Arc::clone(&draft.bytes))
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))?;
        let bytes = Arc::new(apply_design_operation(current, &operation)?);
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP design transaction exceeds the 128 MiB package limit");
        }
        let presentation = Presentation::from_shared_bytes(Arc::clone(&bytes))?;
        let layouts = presentation.layouts()?;
        let masters = presentation.master_pages()?;
        let pages = presentation.pages()?;
        let source_bytes = Arc::clone(&self.source.bytes);
        let draft = self
            .design
            .as_mut()
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))?;
        if layouts == draft.layouts
            && root_masters_equal(&masters, &draft.masters)
            && pages == draft.pages
        {
            return Ok(());
        }
        if layouts == draft.original_layouts
            && root_masters_equal(&masters, &draft.original_masters)
            && pages == draft.original_pages
        {
            draft.bytes = source_bytes;
            draft.layouts = layouts;
            draft.masters = masters;
            draft.pages = pages;
            draft.operations.clear();
            return Ok(());
        }
        draft.bytes = bytes;
        draft.layouts = layouts;
        draft.masters = masters;
        draft.pages = pages;
        draft.operations.push(operation);
        Ok(())
    }

    fn ensure_annotations(&mut self) -> Result<()> {
        if self.annotations.is_none() {
            let presentation = Presentation::from_shared_bytes(Arc::clone(&self.source.bytes))?;
            let annotations = presentation.annotations()?;
            self.annotations = Some(AnnotationDraft {
                bytes: Arc::clone(&self.source.bytes),
                original: annotations.clone(),
                annotations,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn annotation_bytes(&self) -> Result<Arc<Vec<u8>>> {
        self.annotations
            .as_ref()
            .map(|draft| Arc::clone(&draft.bytes))
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))
    }

    fn stage_annotation(
        &mut self,
        bytes: Arc<Vec<u8>>,
        operation: AnnotationOperation,
    ) -> Result<()> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP annotation transaction exceeds the 128 MiB package limit");
        }
        let presentation = Presentation::from_shared_bytes(Arc::clone(&bytes))?;
        let annotations = presentation.annotations()?;
        let source_bytes = Arc::clone(&self.source.bytes);
        let draft = self
            .annotations
            .as_mut()
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))?;
        if annotations == draft.annotations {
            return Ok(());
        }
        if annotations == draft.original {
            draft.bytes = source_bytes;
            draft.annotations = annotations;
            draft.operations.clear();
            return Ok(());
        }
        draft.bytes = bytes;
        draft.annotations = annotations;
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
            domains: Arc::from([]),
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
    domains: Arc<[Domain]>,
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
            domains: Arc::clone(&self.domains),
        }
    }

    /// Return whether this patch preserves the exact package bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        same_source(&self.before, &self.after)
    }

    /// Borrow the sorted semantic dependency domains changed by this patch.
    #[must_use]
    pub fn domains(&self) -> &[Domain] {
        &self.domains
    }

    /// Produce a conservative, non-mutating join plan against another patch.
    ///
    /// RDF-only work is independent from other modeled domains. All edits that
    /// can rewrite `content.xml` are conservatively reported as conflicts until
    /// a semantic operation compositor is available.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches accept the exact same source package.
    pub fn plan_join(&self, other: &Self) -> Result<MergePlan> {
        if !same_source(&self.before, &other.before) {
            return invalid("ODP patch join requires an exact common source");
        }
        let mut conflicts = Vec::new();
        for domain in self.domains.iter().copied() {
            if other.domains.contains(&domain)
                || (domain != Domain::Rdf
                    && other
                        .domains
                        .iter()
                        .any(|other_domain| *other_domain != Domain::Rdf))
            {
                conflicts.push(domain);
            }
        }
        conflicts.sort_unstable();
        conflicts.dedup();
        Ok(MergePlan { conflicts })
    }

    /// Join two patch intents into a non-mutating merge plan.
    ///
    /// This deliberately does not publish a package: even an independent plan
    /// must be replayed from semantic operations so neither target archive wins
    /// by accident.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches accept the exact same source package.
    pub fn join(&self, other: &Self) -> Result<MergePlan> {
        self.plan_join(other)
    }

    /// Materialize two patch intents when the semantic planner proves them independent.
    ///
    /// The current compositor has one deliberately narrow safe case: an RDF-only patch can
    /// be replayed over a patch that leaves RDF untouched. Content-coupled domains continue
    /// to be refused rather than selecting one complete target archive as an accidental winner.
    /// Durable patches retain enough source and target state for this operation after reload.
    ///
    /// # Errors
    ///
    /// Returns an error for different sources, a reported conflict, or an RDF delta that cannot
    /// be replayed and verified over the other target.
    pub fn join_snapshot(&self, other: &Self) -> Result<Snapshot> {
        let plan = self.plan_join(other)?;
        if !plan.is_independent() {
            return unsupported("ODP patch join requires resolution of semantic conflicts");
        }
        if self.is_noop() {
            return Ok(other.after.clone());
        }
        if other.is_noop() {
            return Ok(self.after.clone());
        }
        let (rdf_patch, target_patch) = if self.domains.as_ref() == [Domain::Rdf]
            && !other.domains.contains(&Domain::Rdf)
        {
            (self, other)
        } else if other.domains.as_ref() == [Domain::Rdf] && !self.domains.contains(&Domain::Rdf) {
            (other, self)
        } else {
            return unsupported("ODP independent join has no bounded semantic compositor");
        };
        materialize_rdf_join(rdf_patch, target_patch)
    }

    /// Plan a conservative three-way merge rooted at an exact base snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when either side was not authored against `base`.
    pub fn plan_three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<MergePlan> {
        if !same_source(base, &left.before) || !same_source(base, &right.before) {
            return invalid("ODP three-way merge patches do not share the supplied base");
        }
        left.plan_join(right)
    }

    /// Build a non-mutating three-way merge plan from an exact common base.
    ///
    /// # Errors
    ///
    /// Returns an error when either side was not authored against `base`.
    pub fn three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<MergePlan> {
        Self::plan_three_way(base, left, right)
    }

    /// Materialize a checked three-way merge for the independent cases supported by
    /// [`Self::join_snapshot`].
    ///
    /// # Errors
    ///
    /// Returns an error when either patch is not rooted at `base`, the planner reports a
    /// conflict, or the bounded semantic compositor cannot verify the merged package.
    pub fn three_way_snapshot(base: &Snapshot, left: &Self, right: &Self) -> Result<Snapshot> {
        Self::plan_three_way(base, left, right)?;
        left.join_snapshot(right)
    }

    /// Serialize this exact reversible patch into a deterministic bounded binary envelope.
    ///
    /// The envelope retains both complete package artifacts so stale-source authorization and
    /// byte-exact inversion remain available after process boundaries.
    ///
    /// # Errors
    ///
    /// Returns an allocation or size error when the bounded envelope cannot be materialized.
    pub fn to_durable_bytes(&self) -> Result<Vec<u8>> {
        let before_len = self.before.bytes().len();
        let after_len = self.after.bytes().len();
        let capacity = DURABLE_PATCH_MAGIC
            .len()
            .checked_add(2 + 1 + 8 + 8)
            .and_then(|size| size.checked_add(before_len))
            .and_then(|size| size.checked_add(after_len))
            .ok_or_else(|| invalid_error("ODP durable patch size overflow"))?;
        let maximum = MAX_PACKAGE_BYTES
            .checked_mul(2)
            .and_then(|size| size.checked_add(64))
            .ok_or_else(|| invalid_error("ODP durable patch limit overflow"))?;
        if capacity > maximum {
            return invalid("ODP durable patch exceeds its package-derived size limit");
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable patch envelope",
                source,
            })?;
        output.extend_from_slice(DURABLE_PATCH_MAGIC);
        output.extend_from_slice(&DURABLE_PATCH_VERSION.to_le_bytes());
        output.push(domain_bits(&self.domains));
        output.extend_from_slice(&u64::try_from(before_len).unwrap_or(u64::MAX).to_le_bytes());
        output.extend_from_slice(&u64::try_from(after_len).unwrap_or(u64::MAX).to_le_bytes());
        output.extend_from_slice(self.before.bytes());
        output.extend_from_slice(self.after.bytes());
        Ok(output)
    }

    /// Rehydrate a deterministic durable patch with full package validation.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed version, unknown domain, oversized artifact, trailing
    /// bytes, or invalid ODP source/target package.
    pub fn from_durable_bytes(bytes: &[u8]) -> Result<Self> {
        let header_len = DURABLE_PATCH_MAGIC.len() + 2 + 1 + 8 + 8;
        if bytes.len() < header_len || &bytes[..DURABLE_PATCH_MAGIC.len()] != DURABLE_PATCH_MAGIC {
            return invalid("invalid ODP durable patch magic or truncated header");
        }
        let mut offset = DURABLE_PATCH_MAGIC.len();
        let version = read_u16(bytes, &mut offset)?;
        if version != DURABLE_PATCH_VERSION {
            return invalid(format!("unsupported ODP durable patch version {version}"));
        }
        let bits = *bytes
            .get(offset)
            .ok_or_else(|| invalid_error("truncated ODP durable patch domains"))?;
        offset += 1;
        let domains = domains_from_bits(bits)?;
        let before_len = read_len(bytes, &mut offset)?;
        let after_len = read_len(bytes, &mut offset)?;
        if before_len > MAX_PACKAGE_BYTES || after_len > MAX_PACKAGE_BYTES {
            return invalid("ODP durable patch contains an oversized package");
        }
        let expected = offset
            .checked_add(before_len)
            .and_then(|size| size.checked_add(after_len))
            .ok_or_else(|| invalid_error("ODP durable patch length overflow"))?;
        if expected != bytes.len() {
            return invalid("ODP durable patch length does not match its envelope");
        }
        let before_end = offset + before_len;
        let before = Snapshot::from_bytes(bytes[offset..before_end].to_vec())?;
        let after = Snapshot::from_bytes(bytes[before_end..].to_vec())?;
        Ok(Self {
            before,
            after,
            domains: Arc::from(domains),
        })
    }
}

/// Entry- and byte-bounded undo/redo history for immutable ODP snapshots.
pub struct History {
    entries: Vec<Snapshot>,
    cursor: usize,
    max_entries: usize,
    max_bytes: usize,
    retained_bytes: usize,
}

impl History {
    /// Create history rooted at one immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for zero limits or a byte budget smaller than the root package.
    pub fn new(initial: Snapshot, max_entries: usize, max_bytes: usize) -> Result<Self> {
        if max_entries == 0 || max_bytes == 0 {
            return invalid("ODP history limits must be positive");
        }
        if initial.bytes().len() > max_bytes {
            return invalid("ODP history byte budget cannot retain its initial snapshot");
        }
        let retained_bytes = initial.bytes().len();
        Ok(Self {
            entries: vec![initial],
            cursor: 0,
            max_entries,
            max_bytes,
            retained_bytes,
        })
    }

    /// Borrow the current immutable snapshot.
    #[must_use]
    pub fn current(&self) -> &Snapshot {
        &self.entries[self.cursor]
    }

    /// Record a commit whose exact source is the current history snapshot.
    ///
    /// Redo entries are discarded only after source validation succeeds. Oldest
    /// entries are evicted deterministically to enforce both configured limits.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale commit or a target package larger than the byte budget.
    pub fn record(&mut self, commit: &Commit) -> Result<()> {
        if !same_source(self.current(), &commit.patch.before) {
            return invalid("ODP history commit source is not current");
        }
        if !commit.changed {
            return Ok(());
        }
        let target_bytes = commit.snapshot.bytes().len();
        if target_bytes > self.max_bytes {
            return invalid("ODP history target exceeds the byte budget");
        }
        if self.cursor + 1 < self.entries.len() {
            for removed in self.entries.drain(self.cursor + 1..) {
                self.retained_bytes = self.retained_bytes.saturating_sub(removed.bytes().len());
            }
        }
        self.entries.push(commit.snapshot.clone());
        self.retained_bytes = self
            .retained_bytes
            .checked_add(target_bytes)
            .ok_or_else(|| invalid_error("ODP history byte count overflow"))?;
        self.cursor = self.entries.len() - 1;
        while self.entries.len() > self.max_entries || self.retained_bytes > self.max_bytes {
            let removed = self.entries.remove(0);
            self.retained_bytes = self.retained_bytes.saturating_sub(removed.bytes().len());
            self.cursor = self.cursor.saturating_sub(1);
        }
        Ok(())
    }

    /// Move to the previous retained snapshot.
    #[must_use]
    pub fn undo(&mut self) -> Option<&Snapshot> {
        if self.cursor == 0 {
            return None;
        }
        self.cursor -= 1;
        Some(self.current())
    }

    /// Move to the next retained snapshot.
    #[must_use]
    pub fn redo(&mut self) -> Option<&Snapshot> {
        if self.cursor + 1 >= self.entries.len() {
            return None;
        }
        self.cursor += 1;
        Some(self.current())
    }

    /// Return the number of retained immutable snapshots.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether no snapshots are retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return the exact package-byte accounting used by the history budget.
    #[must_use]
    pub const fn retained_bytes(&self) -> usize {
        self.retained_bytes
    }

    /// Serialize the complete bounded undo/redo timeline into a deterministic envelope.
    ///
    /// Every retained package is included so cursor position, redo state, and exact package
    /// bytes survive a process boundary. The envelope is independently capped at 512 MiB.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained history exceeds the durable bound or a native limit
    /// cannot be represented in the envelope.
    pub fn to_durable_bytes(&self) -> Result<Vec<u8>> {
        let header_len = DURABLE_HISTORY_MAGIC.len() + 2 + (8 * 4);
        let capacity = self.entries.iter().try_fold(header_len, |size, snapshot| {
            size.checked_add(8)
                .and_then(|value| value.checked_add(snapshot.bytes().len()))
                .ok_or_else(|| invalid_error("ODP durable history size overflow"))
        })?;
        if capacity > MAX_DURABLE_HISTORY_BYTES {
            return invalid("ODP durable history exceeds the 512 MiB envelope limit");
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable history envelope",
                source,
            })?;
        output.extend_from_slice(DURABLE_HISTORY_MAGIC);
        output.extend_from_slice(&DURABLE_HISTORY_VERSION.to_le_bytes());
        write_len(&mut output, self.max_entries)?;
        write_len(&mut output, self.max_bytes)?;
        write_len(&mut output, self.cursor)?;
        write_len(&mut output, self.entries.len())?;
        for snapshot in &self.entries {
            write_len(&mut output, snapshot.bytes().len())?;
            output.extend_from_slice(snapshot.bytes());
        }
        Ok(output)
    }

    /// Rehydrate a bounded undo/redo timeline with full package validation.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed bounds, cursor/count inconsistencies, trailing bytes,
    /// oversized packages, budget violations, or invalid retained ODP artifacts.
    pub fn from_durable_bytes(bytes: &[u8]) -> Result<Self> {
        let header_len = DURABLE_HISTORY_MAGIC.len() + 2 + (8 * 4);
        if bytes.len() < header_len
            || bytes.len() > MAX_DURABLE_HISTORY_BYTES
            || &bytes[..DURABLE_HISTORY_MAGIC.len()] != DURABLE_HISTORY_MAGIC
        {
            return invalid("invalid ODP durable history magic, size, or truncated header");
        }
        let mut offset = DURABLE_HISTORY_MAGIC.len();
        let version = read_u16(bytes, &mut offset)?;
        if version != DURABLE_HISTORY_VERSION {
            return invalid(format!("unsupported ODP durable history version {version}"));
        }
        let max_entries = read_len(bytes, &mut offset)?;
        let max_bytes = read_len(bytes, &mut offset)?;
        let cursor = read_len(bytes, &mut offset)?;
        let count = read_len(bytes, &mut offset)?;
        if max_entries == 0
            || max_bytes == 0
            || count == 0
            || count > max_entries
            || cursor >= count
        {
            return invalid("ODP durable history contains inconsistent bounds or cursor state");
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable history entries",
                source,
            })?;
        let mut retained_bytes = 0usize;
        for _ in 0..count {
            let length = read_len(bytes, &mut offset)?;
            if length > MAX_PACKAGE_BYTES {
                return invalid("ODP durable history contains an oversized package");
            }
            let end = offset
                .checked_add(length)
                .ok_or_else(|| invalid_error("ODP durable history package offset overflow"))?;
            let package = bytes
                .get(offset..end)
                .ok_or_else(|| invalid_error("truncated ODP durable history package"))?;
            entries.push(Snapshot::from_bytes(package.to_vec())?);
            retained_bytes = retained_bytes
                .checked_add(length)
                .ok_or_else(|| invalid_error("ODP durable history byte count overflow"))?;
            offset = end;
        }
        if offset != bytes.len() || retained_bytes > max_bytes {
            return invalid(
                "ODP durable history length or byte budget does not match its envelope",
            );
        }
        Ok(Self {
            entries,
            cursor,
            max_entries,
            max_bytes,
            retained_bytes,
        })
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

fn domain_bits(domains: &[Domain]) -> u8 {
    domains.iter().fold(0u8, |bits, domain| {
        bits | match domain {
            Domain::Slides => 1 << 0,
            Domain::Rdf => 1 << 1,
            Domain::Charts => 1 << 2,
            Domain::Design => 1 << 3,
            Domain::Annotations => 1 << 4,
        }
    })
}

fn domains_from_bits(bits: u8) -> Result<Vec<Domain>> {
    if bits & !0b1_1111 != 0 {
        return invalid("ODP durable patch contains an unknown semantic domain");
    }
    let mut domains = Vec::new();
    for (mask, domain) in [
        (1 << 0, Domain::Slides),
        (1 << 1, Domain::Rdf),
        (1 << 2, Domain::Charts),
        (1 << 3, Domain::Design),
        (1 << 4, Domain::Annotations),
    ] {
        if bits & mask != 0 {
            domains.push(domain);
        }
    }
    Ok(domains)
}

fn read_u16(bytes: &[u8], offset: &mut usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid_error("ODP durable patch offset overflow"))?;
    let value = bytes
        .get(*offset..end)
        .ok_or_else(|| invalid_error("truncated ODP durable patch version"))?;
    *offset = end;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_len(bytes: &[u8], offset: &mut usize) -> Result<usize> {
    let end = offset
        .checked_add(8)
        .ok_or_else(|| invalid_error("ODP durable patch offset overflow"))?;
    let value = bytes
        .get(*offset..end)
        .ok_or_else(|| invalid_error("truncated ODP durable patch length"))?;
    *offset = end;
    let decoded = u64::from_le_bytes([
        value[0], value[1], value[2], value[3], value[4], value[5], value[6], value[7],
    ]);
    usize::try_from(decoded)
        .map_err(|error| invalid_error(format!("ODP durable patch length is not native: {error}")))
}

fn write_len(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let encoded = u64::try_from(value).map_err(|source| {
        invalid_error(format!(
            "ODP durable envelope length is not representable: {source}"
        ))
    })?;
    output.extend_from_slice(&encoded.to_le_bytes());
    Ok(())
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

fn materialize_rdf_join(rdf_patch: &Patch, target_patch: &Patch) -> Result<Snapshot> {
    let base_package = OwnedPackage::from_shared_bytes(Arc::clone(&rdf_patch.before.bytes))?;
    let rdf_target_package = OwnedPackage::from_shared_bytes(Arc::clone(&rdf_patch.after.bytes))?;
    let base_graphs = crate::rdf::graphs(&base_package)?;
    let expected_graphs = crate::rdf::graphs(&rdf_target_package)?;
    let mut bytes = Arc::clone(&target_patch.after.bytes);

    for expected in &expected_graphs {
        let Some(before) = base_graphs.iter().find(|graph| graph.path == expected.path) else {
            continue;
        };
        if before != expected {
            let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
            bytes = Arc::new(crate::rdf::replace_graph(
                &package,
                &expected.path,
                &expected.triples,
            )?);
        }
    }

    let mut removals = base_graphs
        .iter()
        .filter(|graph| {
            !expected_graphs
                .iter()
                .any(|expected| expected.path == graph.path)
        })
        .map(|graph| graph.path.clone())
        .collect::<Vec<_>>();
    while !removals.is_empty() {
        let mut progress = false;
        let mut retained = Vec::new();
        for path in removals {
            let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
            match crate::rdf::remove_graph(&package, &path) {
                Ok(updated) => {
                    bytes = Arc::new(updated);
                    progress = true;
                },
                Err(_error) => retained.push(path),
            }
        }
        if !progress {
            return unsupported("ODP RDF join cannot close graph-removal dependencies");
        }
        removals = retained;
    }

    for expected in &expected_graphs {
        if base_graphs.iter().any(|graph| graph.path == expected.path) {
            continue;
        }
        let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
        let (updated, actual_path) =
            crate::rdf::add_graph(&package, Some(&expected.path), &expected.triples)?;
        if actual_path != expected.path {
            return invalid("ODP RDF join resolved a different metadata graph path");
        }
        bytes = Arc::new(updated);
    }

    if bytes.len() > MAX_PACKAGE_BYTES {
        return invalid("ODP joined package exceeds the 128 MiB package limit");
    }
    let joined = Snapshot::from_shared_bytes(bytes)?;
    let joined_package = OwnedPackage::from_shared_bytes(Arc::clone(&joined.bytes))?;
    if crate::rdf::graphs(&joined_package)? != expected_graphs {
        return invalid("ODP joined package RDF readback differs from the expected graph model");
    }
    Ok(joined)
}

fn ensure_chart_transfer_closed(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|source| invalid_error(format!("invalid ODP chart transfer XML: {source}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|source| {
                        invalid_error(format!("invalid ODP chart transfer attribute: {source}"))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *XLINK_NS)
                        && local.as_ref() == b"href"
                    {
                        return unsupported(
                            "ODP chart transfer refuses unresolved xlink:href dependencies",
                        );
                    }
                }
            },
            Event::Eof => break,
            Event::DocType(_) => return invalid("ODP chart transfer refuses document types"),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::End(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
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

fn apply_chart_operation(
    bytes: Arc<Vec<u8>>,
    limits: crate::charts::Limits,
    operation: &ChartOperation,
) -> Result<Vec<u8>> {
    let snapshot = crate::charts::Snapshot::from_shared_bytes(bytes, limits)?;
    let mut edit = snapshot.edit();
    match operation {
        ChartOperation::Replace { selector, part } => {
            edit.replace(selector.borrowed(), part.clone())?;
        },
        ChartOperation::Remove { selector } => {
            let _removed = edit.remove(selector.borrowed())?;
        },
        ChartOperation::Add {
            page,
            name,
            storage,
            part,
        } => {
            let _index = edit.add(page.borrowed(), name.clone(), *storage, part.clone())?;
        },
    }
    edit.commit()
        .map(|commit| commit.snapshot().bytes().to_vec())
}

fn root_charts_equal(left: &[crate::charts::Chart], right: &[crate::charts::Chart]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_chart, right_chart)| {
            left_chart.frame() == right_chart.frame()
                && left_chart.storage() == right_chart.storage()
                && left_chart.part() == right_chart.part()
        })
}

fn apply_design_operation(bytes: Arc<Vec<u8>>, operation: &DesignOperation) -> Result<Vec<u8>> {
    let mut presentation = Presentation::from_shared_bytes(bytes)?;
    match operation {
        DesignOperation::AddLayout(layout) => presentation.add_layout(layout)?,
        DesignOperation::ReplaceLayout(layout) => presentation.replace_page_layout(layout)?,
        DesignOperation::RemoveLayout { name, replacement } => {
            presentation.remove_page_layout(name, replacement.as_deref())?;
        },
        DesignOperation::ReorderLayouts(names) => presentation.reorder_layouts(names)?,
        DesignOperation::AddMaster(master) => presentation.add_master_page(master)?,
        DesignOperation::ReplaceMaster(master) => presentation.replace_master_page(master)?,
        DesignOperation::RemoveMaster { name, replacement } => {
            presentation.remove_master_page(name, replacement.as_deref())?;
        },
        DesignOperation::ReorderMasters(names) => presentation.reorder_master_pages(names)?,
        DesignOperation::AssignSlideMaster { slide_index, name } => {
            presentation.assign_slide_master_page(*slide_index, name.as_deref())?;
        },
        DesignOperation::AssignSlideLayout { slide_index, name } => {
            presentation.assign_slide_page_layout(*slide_index, name.as_deref())?;
        },
    }
    presentation.to_bytes()
}

fn root_masters_equal(left: &[crate::MasterPage], right: &[crate::MasterPage]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_master, right_master)| {
            left_master.master_page == right_master.master_page
                && left_master.page_layout_name == right_master.page_layout_name
                && left_master.header_name == right_master.header_name
                && left_master.footer_name == right_master.footer_name
                && left_master.date_time_name == right_master.date_time_name
        })
}

fn root_design_pages_equal(
    actual: &crate::page::Collection,
    expected: &crate::page::Collection,
) -> bool {
    expected.pages().iter().all(|expected_page| {
        actual
            .page(expected_page.slide_index)
            .is_some_and(|actual_page| actual_page == expected_page)
    })
}

fn apply_annotation_operation(
    bytes: Arc<Vec<u8>>,
    operation: &AnnotationOperation,
) -> Result<Vec<u8>> {
    let mut presentation = Presentation::from_shared_bytes(bytes)?;
    match operation {
        AnnotationOperation::Add { anchor, annotation } => {
            let _index = presentation.add_annotation(anchor, annotation)?;
        },
        AnnotationOperation::Replace { index, annotation } => {
            presentation.replace_annotation(*index, annotation)?;
        },
        AnnotationOperation::Remove { index } => presentation.remove_annotation(*index)?,
    }
    presentation.to_bytes()
}

fn root_annotations_equal(
    left: &[crate::annotation::Info],
    right: &[crate::annotation::Info],
) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_info, right_info)| {
            left_info.index == right_info.index
                && left_info.anchor == right_info.anchor
                && left_info.annotation.attributes() == right_info.annotation.attributes()
                && left_info.annotation.children() == right_info.annotation.children()
        })
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
