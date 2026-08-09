//! Immutable ODG package snapshots and lossless semantic shape edits.

use crate::model::{
    layer::Layer,
    page::Page,
    shape::{Properties as ShapeProperties, Shape, ShapeKind},
};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
    drawing::Frame,
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{ops::Range, path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";
const BODY_MARKER: &str = "<office:drawing";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_DEPTH: usize = 256;
const MAX_PAGES: usize = 16_384;
const MAX_LAYERS: usize = 16_384;
const MAX_SHAPES: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Svg,
    Text,
    Other,
}

type TextSpans = Vec<Vec<Vec<Option<Range<usize>>>>>;
type NameSpans = Vec<Vec<Option<Range<usize>>>>;
type LayerSpans = Vec<Vec<Option<Range<usize>>>>;

struct State {
    package: Package,
    pages: Vec<Page>,
    layers: Vec<Layer>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
}

/// An immutable, source-owning ODG package snapshot.
///
/// Unknown package members and unmodeled XML remain in the retained source
/// bytes. Semantic inspection never evaluates controls, scripts, actions, DDE,
/// links, or embedded payloads.
#[derive(Clone)]
pub struct Snapshot(Arc<State>);

impl Snapshot {
    /// Opens a package from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package cannot be read or is not a structurally valid ODG.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(Package::open(path, MIMETYPE, BODY_MARKER, "ODG")?)
    }

    /// Opens a package from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid ODG.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODG")?)
    }

    fn from_package(package: Package) -> Result<Self> {
        let parsed = parse_content(package.content_xml())?;
        let layers = package
            .styles_xml()
            .map(parse_declared_layers)
            .transpose()?
            .unwrap_or_default();
        if parsed.layer_count.saturating_add(layers.len()) > MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        Ok(Self(Arc::new(State {
            package,
            pages: parsed.pages,
            layers,
            text_spans: parsed.text_spans,
            name_spans: parsed.name_spans,
            layer_spans: parsed.layer_spans,
        })))
    }

    /// Returns the exact `content.xml` source.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    /// Returns exact `styles.xml`, when present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    /// Returns common document metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    /// Returns bounded pages in source order.
    #[must_use]
    pub fn pages(&self) -> &[Page] {
        &self.0.pages
    }

    /// Selects one page by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact name is ambiguous.
    pub fn page<'selector>(
        &self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Option<&Page>> {
        let resolved_selector = selector.into();
        match resolved_selector {
            crate::page::Selector::Position(position) => Ok(self.pages().get(position.get())),
            crate::page::Selector::Name(name) => {
                let mut matches = self
                    .pages()
                    .iter()
                    .filter(|page| page.name() == Some(name.as_ref()));
                let selected = matches.next();
                if selected.is_some() && matches.next().is_some() {
                    return invalid("ODG page name selector is ambiguous");
                }
                Ok(selected)
            },
        }
    }

    /// Returns global drawing layers declared by `styles.xml` in source order.
    ///
    /// Page-local declarations are available from [`Page::layers`](crate::page::Page::layers).
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.0.layers
    }

    /// Returns original package bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    /// Lists safe package entry names.
    ///
    /// # Errors
    ///
    /// Returns an error when package member validation fails.
    pub fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    /// Starts a source-bound semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            change: None,
        }
    }

    /// Consumes the snapshot and returns its source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }
}

/// A staged source-bound package shape edit.
pub struct Transaction {
    source: Snapshot,
    change: Option<ShapeChange>,
}

impl Transaction {
    /// Replaces one shape's sole plain paragraph character-data span.
    ///
    /// Split, mixed, CDATA, and entity-reference text is refused rather than
    /// serialized through a lossy XML model. A transaction owns one edit;
    /// restaging the same selector replaces its pending value.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, unsupported source span, or limit violation.
    pub fn set_shape_text(
        &mut self,
        page: usize,
        shape: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape text exceeds the limit");
        }
        self.ensure_selector(page, shape)?;
        let selected = self
            .source
            .pages()
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let spans = self
            .source
            .0
            .text_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape source span is missing".to_string()))?;
        if !matches!(spans.as_slice(), [Some(_)]) {
            return invalid("ODG shape text is not one losslessly replaceable XML span");
        }
        if selected.text() == after {
            self.change = None;
            return Ok(());
        }
        self.change = Some(ShapeChange::Text(TextChange {
            page,
            shape,
            before: selected.text().to_string(),
            after,
        }));
        Ok(())
    }

    /// Renames one shape through its existing `draw:name` attribute.
    ///
    /// ODF 1.4 Part 3 §19.197 defines `draw:name` as the reference name for
    /// graphical elements. This preserves the original start tag and attribute
    /// spelling, replacing only the validated attribute-value span.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-bounds selector, an unnamed shape, a
    /// name over the bounded size, or a shape whose source attribute cannot be
    /// losslessly addressed.
    pub fn set_shape_name(
        &mut self,
        page: usize,
        shape: usize,
        name: impl Into<String>,
    ) -> Result<()> {
        let after = name.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape name exceeds the limit");
        }
        self.ensure_selector(page, shape)?;
        let selected = self
            .source
            .pages()
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let before = selected.name().ok_or_else(|| {
            Error::Unsupported(
                "ODG shape rename requires an existing losslessly addressable draw:name"
                    .to_string(),
            )
        })?;
        if self
            .source
            .0
            .name_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .is_none()
        {
            return invalid("ODG shape name source span is missing");
        }
        if before == after {
            self.change = None;
            return Ok(());
        }
        self.change = Some(ShapeChange::Name(NameChange {
            page,
            shape,
            before: before.to_string(),
            after,
        }));
        Ok(())
    }

    /// Changes a shape's existing layer assignment without normalizing its tag.
    ///
    /// ODF 1.4 Part 3 §§10.2.2-10.2.3 and 19.189 define drawing layers and
    /// their shape assignment. The destination must be one of the declarations
    /// visible through [`Snapshot::layers`].
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, undeclared layer, absent source attribute, or
    /// limit violation.
    pub fn set_shape_layer(
        &mut self,
        page: usize,
        shape: usize,
        layer: impl Into<String>,
    ) -> Result<()> {
        let after = layer.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement layer name exceeds the limit");
        }
        self.ensure_selector(page, shape)?;
        let selected_page = self.source.pages().get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG page selector is out of bounds".to_string())
        })?;
        let selected = selected_page.shapes().get(shape).ok_or_else(|| {
            Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
        })?;
        let visible_layers = if selected_page.has_layer_set() {
            selected_page.layers()
        } else {
            self.source.layers()
        };
        if !visible_layers
            .iter()
            .any(|declared_layer| declared_layer.name() == after)
        {
            return invalid("ODG destination layer is not declared");
        }
        let before = selected.layer().ok_or_else(|| {
            Error::Unsupported(
                "ODG layer change requires an existing losslessly addressable draw:layer"
                    .to_string(),
            )
        })?;
        if self
            .source
            .0
            .layer_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .is_none()
        {
            return invalid("ODG shape layer source span is missing");
        }
        if before == after {
            self.change = None;
            return Ok(());
        }
        self.change = Some(ShapeChange::Layer(LayerChange {
            page,
            shape,
            before: before.to_string(),
            after,
        }));
        Ok(())
    }

    fn ensure_selector(&self, page: usize, shape: usize) -> Result<()> {
        if self
            .change
            .as_ref()
            .is_some_and(|change| change.page() != page || change.shape() != shape)
        {
            return invalid("an ODG package transaction supports one semantic edit target");
        }
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the edited package.
    ///
    /// # Errors
    ///
    /// Returns an error when source policy, rebuilding, parsing, or typed readback fails.
    pub fn commit(self) -> Result<Commit> {
        let Some(staged) = self.change else {
            return Ok(Commit::unchanged(self.source));
        };
        ensure_compact_rewrite_source(&self.source)?;
        let content = match &staged {
            ShapeChange::Text(text_change) => {
                let span = self
                    .source
                    .0
                    .text_spans
                    .get(text_change.page)
                    .and_then(|shapes| shapes.get(text_change.shape))
                    .and_then(|spans| spans.first())
                    .and_then(Option::as_ref)
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG shape text source span is missing".to_string())
                    })?;
                replace_xml_value(self.source.content_xml(), span, &text_change.after)?
            },
            ShapeChange::Name(name_change) => {
                let span = self
                    .source
                    .0
                    .name_spans
                    .get(name_change.page)
                    .and_then(|shapes| shapes.get(name_change.shape))
                    .and_then(Option::as_ref)
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG shape name source span is missing".to_string())
                    })?;
                replace_xml_value(self.source.content_xml(), span, &name_change.after)?
            },
            ShapeChange::Layer(layer_change) => {
                let span = self
                    .source
                    .0
                    .layer_spans
                    .get(layer_change.page)
                    .and_then(|shapes| shapes.get(layer_change.shape))
                    .and_then(Option::as_ref)
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG shape layer source span is missing".to_string())
                    })?;
                replace_xml_value(self.source.content_xml(), span, &layer_change.after)?
            },
        };
        compact_xml::validate(content.as_bytes()).map_err(Error::from)?;
        let snapshot = Snapshot::from_bytes(rebuild(&self.source, &content)?)?;
        let actual_page = snapshot.pages().get(staged.page()).ok_or_else(|| {
            Error::InvalidFormat("ODG edited page disappeared during readback".to_string())
        })?;
        let actual_shape = actual_page.shapes().get(staged.shape()).ok_or_else(|| {
            Error::InvalidFormat("ODG edited shape disappeared during readback".to_string())
        })?;
        match &staged {
            ShapeChange::Text(text_change) if actual_shape.text() != text_change.after => {
                return invalid("ODG package text edit failed typed readback");
            },
            ShapeChange::Name(name_change)
                if actual_shape.name() != Some(name_change.after.as_str()) =>
            {
                return invalid("ODG package name edit failed typed readback");
            },
            ShapeChange::Layer(layer_change)
                if actual_shape.layer() != Some(layer_change.after.as_str()) =>
            {
                return invalid("ODG package layer edit failed typed readback");
            },
            ShapeChange::Text(_) | ShapeChange::Name(_) | ShapeChange::Layer(_) => {},
        }
        let (text_change, name_change, layer_change) = match staged {
            ShapeChange::Text(change) => (Some(change), None, None),
            ShapeChange::Name(change) => (None, Some(change), None),
            ShapeChange::Layer(change) => (None, None, Some(change)),
        };
        Ok(Commit {
            patch: Patch {
                source: self.source,
                target: snapshot.clone(),
                text_change,
                name_change,
                layer_change,
            },
            snapshot,
            changed: true,
        })
    }
}

enum ShapeChange {
    Text(TextChange),
    Name(NameChange),
    Layer(LayerChange),
}

impl ShapeChange {
    const fn page(&self) -> usize {
        match self {
            Self::Text(change) => change.page,
            Self::Name(change) => change.page,
            Self::Layer(change) => change.page,
        }
    }

    const fn shape(&self) -> usize {
        match self {
            Self::Text(change) => change.shape,
            Self::Name(change) => change.shape,
            Self::Layer(change) => change.shape,
        }
    }
}

/// One reversible semantic shape-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl TextChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Text expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Text produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible `draw:name` change for a shape.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl NameChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible drawing-layer assignment change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LayerChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl LayerChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// The layer name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The layer name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A committed package publication and its exact-source patch.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                text_change: None,
                name_change: None,
                layer_change: None,
            },
            snapshot,
            changed: false,
        }
    }

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The published immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-source-checked reversible package patch.
#[derive(Clone)]
pub struct Patch {
    source: Snapshot,
    target: Snapshot,
    text_change: Option<TextChange>,
    name_change: Option<NameChange>,
    layer_change: Option<LayerChange>,
}

impl Patch {
    /// Whether this patch authorizes the supplied exact source bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Snapshot) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not the exact source artifact.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !self.is_applicable_to(source) {
            return invalid("ODG package patch source does not match");
        }
        Ok(self.target.clone())
    }

    /// The semantic change represented by this patch.
    #[must_use]
    pub fn change(&self) -> Option<&TextChange> {
        self.text_change.as_ref()
    }

    /// The semantic `draw:name` change, when this is a name patch.
    #[must_use]
    pub fn name_change(&self) -> Option<&NameChange> {
        self.name_change.as_ref()
    }

    /// The semantic drawing-layer change, when present.
    #[must_use]
    pub fn layer_change(&self) -> Option<&LayerChange> {
        self.layer_change.as_ref()
    }

    /// An exact-source patch restoring the original package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            text_change: self.text_change.as_ref().map(|change| TextChange {
                page: change.page,
                shape: change.shape,
                before: change.after.clone(),
                after: change.before.clone(),
            }),
            name_change: self.name_change.as_ref().map(|change| NameChange {
                page: change.page,
                shape: change.shape,
                before: change.after.clone(),
                after: change.before.clone(),
            }),
            layer_change: self.layer_change.as_ref().map(|change| LayerChange {
                page: change.page,
                shape: change.shape,
                before: change.after.clone(),
                after: change.before.clone(),
            }),
        }
    }
}

struct Parsed {
    pages: Vec<Page>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    layer_count: usize,
}

struct ActiveShape {
    depth: usize,
    page: usize,
    shape: usize,
}

#[derive(Clone, Copy)]
enum AccessibilityKind {
    Description,
    Title,
}

struct ActiveAccessibility {
    depth: usize,
    page: usize,
    shape: usize,
    kind: AccessibilityKind,
}

struct Scanner {
    depth: usize,
    root_seen: bool,
    body_seen: bool,
    drawing_seen: bool,
    body_depth: Option<usize>,
    drawing_depth: Option<usize>,
    pages: Vec<Page>,
    page_depths: Vec<usize>,
    layer_sets: Vec<(usize, Option<usize>)>,
    active_shapes: Vec<ActiveShape>,
    active_accessibility: Option<ActiveAccessibility>,
    paragraph_depths: Vec<usize>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    layer_count: usize,
    shape_count: usize,
    text_bytes: usize,
}

impl Scanner {
    fn new() -> Self {
        Self {
            depth: 0,
            root_seen: false,
            body_seen: false,
            drawing_seen: false,
            body_depth: None,
            drawing_depth: None,
            pages: Vec::new(),
            page_depths: Vec::new(),
            layer_sets: Vec::new(),
            active_shapes: Vec::new(),
            active_accessibility: None,
            paragraph_depths: Vec::new(),
            text_spans: Vec::new(),
            name_spans: Vec::new(),
            layer_spans: Vec::new(),
            layer_count: 0,
            shape_count: 0,
            text_bytes: 0,
        }
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        self.depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
        if self.depth > MAX_DEPTH {
            return invalid("ODG XML nesting exceeds the limit");
        }
        self.observe(reader, namespace, element, tag, tag_start, empty)?;
        if empty {
            self.depth = self.depth.saturating_sub(1);
        }
        Ok(())
    }

    fn observe(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        let local_name = element.local_name();
        let local = local_name.as_ref();
        if self.depth == 1 {
            if self.root_seen
                || namespace != NamespaceKind::Office
                || local != b"document-content"
                || empty
            {
                return invalid("ODG content.xml requires one office:document-content root");
            }
            self.root_seen = true;
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"body" {
            if self.body_seen || self.depth != 2 || empty {
                return invalid("ODG content.xml requires one non-empty office:body");
            }
            self.body_seen = true;
            self.body_depth = Some(self.depth);
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"drawing" {
            if self.drawing_seen || self.body_depth != Some(self.depth - 1) {
                return invalid("ODG office:drawing is misplaced or duplicated");
            }
            self.drawing_seen = true;
            if !empty {
                self.drawing_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"layer-set" {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer-set is outside draw:page".to_string())
            })?;
            self.pages[page].mark_layer_set();
            if !empty {
                self.layer_sets.push((self.depth, Some(page)));
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer"
            && self
                .layer_sets
                .last()
                .is_some_and(|(depth, _)| *depth + 1 == self.depth)
        {
            self.add_layer(reader, element)?;
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"page" {
            if self.drawing_depth != Some(self.depth - 1) {
                return invalid("ODG draw:page is outside office:drawing");
            }
            if self.pages.len() >= MAX_PAGES {
                return invalid("ODG page count exceeds the limit");
            }
            self.pages.push(Page::parsed(
                attribute(reader, element, DRAW, b"name")?,
                attribute(reader, element, XML, b"id")?,
                attribute(reader, element, DRAW, b"style-name")?,
                attribute(reader, element, DRAW, b"master-page-name")?,
            ));
            self.text_spans.push(Vec::new());
            self.name_spans.push(Vec::new());
            self.layer_spans.push(Vec::new());
            if !empty {
                self.page_depths.push(self.depth);
            }
            return Ok(());
        }
        if let Some(kind) = shape_kind(namespace, local) {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG drawing shape is outside draw:page".to_string())
            })?;
            if self.shape_count >= MAX_SHAPES {
                return invalid("ODG shape count exceeds the limit");
            }
            self.shape_count += 1;
            let name = attribute(reader, element, DRAW, b"name")?;
            let page_name = self.pages[page].name().map(str::to_string);
            let frame = if kind == ShapeKind::Frame {
                Some(frame(reader, element, name.clone(), page_name)?)
            } else {
                None
            };
            let z_index = optional_u32_attribute(reader, element, DRAW, b"z-index")?;
            let geometry = [
                attribute(reader, element, SVG, b"x")?,
                attribute(reader, element, SVG, b"y")?,
                attribute(reader, element, SVG, b"width")?,
                attribute(reader, element, SVG, b"height")?,
            ];
            let shape = self.pages[page].shapes().len();
            self.pages[page].push_shape(Shape::new(
                ShapeProperties {
                    geometry,
                    layer: attribute(reader, element, DRAW, b"layer")?,
                    name,
                    style_name: attribute(reader, element, DRAW, b"style-name")?,
                    text_style_name: attribute(reader, element, DRAW, b"text-style-name")?,
                    z_index,
                },
                kind,
                frame,
            ));
            self.text_spans[page].push(Vec::new());
            self.name_spans[page].push(shape_name_span(reader, element, tag, tag_start)?);
            self.layer_spans[page].push(attribute_source_span(
                reader, element, tag, tag_start, DRAW, b"layer",
            )?);
            if !empty {
                self.active_shapes.push(ActiveShape {
                    depth: self.depth,
                    page,
                    shape,
                });
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Svg
            && matches!(local, b"title" | b"desc")
            && !empty
            && let Some(active) = self.active_shapes.last()
            && active.depth + 1 == self.depth
        {
            self.active_accessibility = Some(ActiveAccessibility {
                depth: self.depth,
                page: active.page,
                shape: active.shape,
                kind: if local == b"title" {
                    AccessibilityKind::Title
                } else {
                    AccessibilityKind::Description
                },
            });
            return Ok(());
        }
        if !self.active_shapes.is_empty()
            && namespace == NamespaceKind::Text
            && local == b"p"
            && !empty
        {
            self.paragraph_depths.push(self.depth);
        }
        Ok(())
    }

    fn add_layer(&mut self, reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
        if self.layer_count >= MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        self.layer_count += 1;
        let name = required_attribute(reader, element, DRAW, b"name", "draw:layer")?;
        let protected = optional_bool_attribute(reader, element, DRAW, b"protected")?;
        let layer = Layer::parsed(
            name,
            attribute(reader, element, DRAW, b"display")?,
            protected,
        );
        if let Some((_, Some(page))) = self.layer_sets.last() {
            self.pages[*page].push_layer(layer.clone());
        }
        Ok(())
    }

    fn text(&mut self, span: Option<Range<usize>>, value: &str) -> Result<()> {
        if let Some(accessibility) = &self.active_accessibility {
            self.text_bytes = self.text_bytes.checked_add(value.len()).ok_or_else(|| {
                Error::InvalidFormat("ODG text extraction size overflow".to_string())
            })?;
            if self.text_bytes > MAX_TEXT_BYTES {
                return invalid("ODG text extraction exceeds the limit");
            }
            let shape = self.pages[accessibility.page]
                .shape_mut(accessibility.shape)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG active accessibility shape disappeared".to_string())
                })?;
            match accessibility.kind {
                AccessibilityKind::Description => shape.push_description(value),
                AccessibilityKind::Title => shape.push_title(value),
            }
            return Ok(());
        }
        if self.paragraph_depths.is_empty() {
            return Ok(());
        }
        let Some(active) = self.active_shapes.last() else {
            return Ok(());
        };
        let (page, shape_index) = (active.page, active.shape);
        self.text_bytes = self
            .text_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("ODG text extraction size overflow".to_string()))?;
        if self.text_bytes > MAX_TEXT_BYTES {
            return invalid("ODG text extraction exceeds the limit");
        }
        let shape = self.pages[page]
            .shape_mut(shape_index)
            .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
        shape.push_text(value);
        self.text_spans[page][shape_index].push(span);
        Ok(())
    }

    fn end(&mut self, namespace: NamespaceKind, local: &[u8]) -> Result<()> {
        if self
            .active_accessibility
            .as_ref()
            .is_some_and(|active| active.depth == self.depth)
        {
            self.active_accessibility = None;
        }
        if self.paragraph_depths.last() == Some(&self.depth)
            && namespace == NamespaceKind::Text
            && local == b"p"
        {
            self.paragraph_depths.pop();
        }
        if self
            .active_shapes
            .last()
            .is_some_and(|shape| shape.depth == self.depth)
        {
            self.active_shapes.pop();
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer-set"
            && self
                .layer_sets
                .last()
                .is_some_and(|set| set.0 == self.depth)
        {
            self.layer_sets.pop();
        }
        if namespace == NamespaceKind::Draw
            && local == b"page"
            && self.page_depths.last() == Some(&self.depth)
        {
            self.page_depths.pop();
        }
        if self.drawing_depth == Some(self.depth) {
            self.drawing_depth = None;
        }
        if self.body_depth == Some(self.depth) {
            self.body_depth = None;
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth underflow".to_string()))?;
        Ok(())
    }

    fn current_page(&self) -> Option<usize> {
        self.page_depths
            .last()
            .filter(|depth| self.depth > **depth)
            .map(|_| self.pages.len() - 1)
    }

    fn finish(self) -> Result<Parsed> {
        if self.depth != 0
            || !self.root_seen
            || !self.body_seen
            || !self.drawing_seen
            || self.body_depth.is_some()
            || self.drawing_depth.is_some()
            || !self.page_depths.is_empty()
            || !self.layer_sets.is_empty()
            || self.active_accessibility.is_some()
        {
            return invalid("ODG content.xml has an incomplete drawing structure");
        }
        Ok(Parsed {
            pages: self.pages,
            text_spans: self.text_spans,
            name_spans: self.name_spans,
            layer_spans: self.layer_spans,
            layer_count: self.layer_count,
        })
    }
}

fn parse_content(xml: &str) -> Result<Parsed> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG content.xml: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event = borrowed_event.into_owned();
        let end = position(&reader)?;
        match event {
            Event::Start(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                false,
            )?,
            Event::Empty(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                true,
            )?,
            Event::End(element) => scanner.end(namespace, element.local_name().as_ref())?,
            Event::Text(text) => {
                let value = text_value(&text)?;
                scanner.text(Some(start..end), &value)?;
            },
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::InvalidFormat(format!("invalid ODG CDATA: {error}")))?;
                scanner.text(None, &value)?;
            },
            Event::GeneralRef(reference) => {
                let value = reference_value(&reference)?;
                scanner.text(None, &value)?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG content.xml"),
            Event::Eof => return scanner.finish(),
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
}

fn parse_declared_layers(xml: &str) -> Result<Vec<Layer>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut layer_sets = Vec::<usize>::new();
    let mut layers = Vec::new();
    loop {
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG styles.xml: {error}")))?;
        let namespace = classify(&resolved_namespace);
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"layer-set" {
                    layer_sets.push(depth);
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets.last().is_some_and(|set| *set + 1 == depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::Empty(element) => {
                let virtual_depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets
                        .last()
                        .is_some_and(|set| *set + 1 == virtual_depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Draw
                    && element.local_name().as_ref() == b"layer-set"
                    && layer_sets.last() == Some(&depth)
                {
                    layer_sets.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG styles XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG styles.xml"),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || !layer_sets.is_empty() {
        return invalid("ODG styles.xml has incomplete layer declarations");
    }
    Ok(layers)
}

fn checked_xml_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
    if next_depth > MAX_DEPTH {
        return invalid("ODG XML nesting exceeds the limit");
    }
    Ok(next_depth)
}

fn push_declared_layer(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    layers: &mut Vec<Layer>,
) -> Result<()> {
    if layers.len() >= MAX_LAYERS {
        return invalid("ODG declared layer count exceeds the limit");
    }
    layers.push(Layer::parsed(
        required_attribute(reader, element, DRAW, b"name", "draw:layer")?,
        attribute(reader, element, DRAW, b"display")?,
        optional_bool_attribute(reader, element, DRAW, b"protected")?,
    ));
    Ok(())
}

fn text_value(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
    let decoded = text
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text: {error}")))?;
    quick_xml::escape::unescape(&decoded)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text escape: {error}")))
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODG character reference: {error}"))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "apos" => Ok("'".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("ODG custom entities are not allowed"),
    }
}

fn rebuild(source: &Snapshot, content: &str) -> Result<Vec<u8>> {
    let files = source.files()?;
    if files.iter().any(|path| {
        matches!(
            path.as_str(),
            "META-INF/documentsignatures.xml" | "META-INF/macrosignatures.xml"
        )
    }) {
        return invalid("ODG package edits refuse signed packages");
    }
    let archive = source.0.package.package();
    let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
    writer.set_mimetype(MIMETYPE)?;
    writer.add_file("content.xml", content.as_bytes())?;
    for path in ["styles.xml", "meta.xml", "settings.xml"] {
        if archive.has_file(path)? {
            writer.add_file(path, &archive.get_file(path)?)?;
        }
    }
    writer.copy_auxiliary_files_from(archive)?;
    writer.finish_to_bounded_bytes()
}

fn ensure_compact_rewrite_source(source: &Snapshot) -> Result<()> {
    let archive = source.0.package.package();
    for path in source.files()? {
        let is_xml = Path::new(&path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"));
        if is_xml && path != "META-INF/manifest.xml" {
            compact_xml::validate(&archive.get_file(&path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}

fn replace_xml_value(source: &str, span: &Range<usize>, replacement: &str) -> Result<String> {
    if span.start > span.end || span.end > source.len() {
        return invalid("ODG text source span is invalid");
    }
    let escaped_replacement = quick_xml::escape::escape(replacement);
    let capacity = source
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|size| size.checked_add(escaped_replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".to_string()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..span.start]);
    output.push_str(&escaped_replacement);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn frame(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: Option<String>,
    page_name: Option<String>,
) -> Result<Frame> {
    Ok(Frame {
        name,
        xml_id: attribute(reader, element, XML, b"id")?,
        title: None,
        description: None,
        anchor_type: attribute(reader, element, TEXT, b"anchor-type")?,
        x: attribute(reader, element, SVG, b"x")?,
        y: attribute(reader, element, SVG, b"y")?,
        width: attribute(reader, element, SVG, b"width")?,
        height: attribute(reader, element, SVG, b"height")?,
        end_cell_address: None,
        page_name,
        sheet_name: None,
        sheet_shape: false,
    })
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected) && name.as_ref() == local {
            let decoded = parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODG attribute value: {error}"))
                })?
                .into_owned();
            if value.replace(decoded).is_some() {
                return invalid("ODG element has a duplicate namespaced attribute");
            }
        }
    }
    Ok(value)
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
    owner: &str,
) -> Result<String> {
    attribute(reader, element, expected, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODG {owner} requires a namespaced attribute")))
}

fn optional_bool_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<bool>> {
    attribute(reader, element, expected, local)?
        .map(|value| match value.as_str() {
            "false" => Ok(false),
            "true" => Ok(true),
            _ => invalid("ODG Boolean attribute is not true or false"),
        })
        .transpose()
}

fn optional_u32_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<u32>> {
    attribute(reader, element, expected, local)?
        .map(|value| {
            value.parse::<u32>().map_err(|_error| {
                Error::InvalidFormat("ODG integer attribute is invalid".to_string())
            })
        })
        .transpose()
}

fn shape_name_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<Option<Range<usize>>> {
    attribute_source_span(reader, element, tag, tag_start, DRAW, b"name")
}

fn attribute_source_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
    expected: &[u8],
    wanted_local: &[u8],
) -> Result<Option<Range<usize>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected)
            && local.as_ref() == wanted_local
            && key
                .replace(parsed_attribute.key.as_ref().to_vec())
                .is_some()
        {
            return invalid("ODG element has duplicate namespaced attributes");
        }
    }
    let Some(name_key) = key else {
        return Ok(None);
    };
    let (start, end) = attribute_value_span(tag, &name_key)?;
    Ok(Some(tag_start + start..tag_start + end))
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return invalid("ODG shape attribute is missing '='");
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| Error::InvalidFormat("ODG shape attribute is not quoted".to_string()))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= tag.len() {
            return invalid("ODG shape attribute is unterminated");
        }
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    invalid("ODG shape name span was not found")
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| {
        Error::InvalidFormat("ODG XML position exceeds platform limits".to_string())
    })
}

fn shape_kind(namespace: NamespaceKind, local: &[u8]) -> Option<ShapeKind> {
    (namespace == NamespaceKind::Draw).then_some(match local {
        b"caption" => ShapeKind::Caption,
        b"circle" => ShapeKind::Circle,
        b"connector" => ShapeKind::Connector,
        b"control" => ShapeKind::Control,
        b"custom-shape" => ShapeKind::Custom,
        b"ellipse" => ShapeKind::Ellipse,
        b"frame" => ShapeKind::Frame,
        b"g" => ShapeKind::Group,
        b"line" => ShapeKind::Line,
        b"measure" => ShapeKind::Measure,
        b"path" => ShapeKind::Path,
        b"page-thumbnail" => ShapeKind::PageThumbnail,
        b"polygon" => ShapeKind::Polygon,
        b"polyline" => ShapeKind::Polyline,
        b"rect" => ShapeKind::Rectangle,
        b"regular-polygon" => ShapeKind::RegularPolygon,
        _ => return None,
    })
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DRAW => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Bound(Namespace(uri)) if *uri == SVG => NamespaceKind::Svg,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
