//! Immutable ODG package snapshots and lossless semantic shape edits.

use crate::model::{
    layer::Layer,
    page::Page,
    resource::Resource,
    shape::{Properties as ShapeProperties, Shape, ShapeKind},
};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
    drawing::Frame,
    media,
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
type GeometrySpans = Vec<Vec<[Option<Range<usize>>; 4]>>;

struct State {
    package: Package,
    pages: Vec<Page>,
    layers: Vec<Layer>,
    resources: Vec<Resource>,
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
        let resources = scan_resources(&package)?;
        Ok(Self(Arc::new(State {
            package,
            pages: parsed.pages,
            layers,
            resources,
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

    /// Returns package-local image resources referenced by drawing XML.
    #[must_use]
    pub fn resources(&self) -> &[Resource] {
        &self.0.resources
    }

    /// Reads one inventoried package-local resource without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable package member.
    pub fn resource_bytes(&self, resource: usize) -> Result<Option<Vec<u8>>> {
        let selected = self.resources().get(resource).ok_or_else(|| {
            Error::InvalidFormat("ODG resource selector is out of bounds".to_string())
        })?;
        if !selected.is_present() {
            return Ok(None);
        }
        self.0.package.package().get_file(selected.path()).map(Some)
    }

    /// Starts a source-bound semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            content: self.content_xml().to_string(),
            changes: Vec::new(),
            resource_edits: Vec::new(),
        }
    }

    /// Starts explicit bounded undo/redo history at this immutable snapshot.
    #[must_use]
    pub fn history(&self, limits: litchi_core::patch::HistoryLimits) -> super::History {
        super::History::new(self.clone(), limits)
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
    content: String,
    changes: Vec<Change>,
    resource_edits: Vec<ResourceEdit>,
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
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let spans = parsed
            .text_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape source span is missing".to_string()))?;
        if !matches!(spans.as_slice(), [Some(_)]) {
            return invalid("ODG shape text is not one losslessly replaceable XML span");
        }
        if selected.text() == after {
            return Ok(());
        }
        let span = spans[0].as_ref().ok_or_else(|| {
            Error::InvalidFormat("ODG shape text source span is missing".to_string())
        })?;
        let before = selected.text().to_string();
        self.content = replace_xml_value(&self.content, span, &after)?;
        self.changes.push(Change::Text(TextChange {
            page,
            shape,
            before,
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
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
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
        let span = parsed
            .name_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape name source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.content = replace_xml_value(&self.content, span, &after)?;
        self.changes.push(Change::Name(NameChange {
            page,
            shape,
            before: before_owned,
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
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed.pages.get(page).ok_or_else(|| {
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
        let span = parsed
            .layer_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape layer source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.content = replace_xml_value(&self.content, span, &after)?;
        self.changes.push(Change::Layer(LayerChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Replaces all four existing SVG geometry attributes as one operation.
    ///
    /// # Errors
    ///
    /// Returns an error if the checked selectors fail or the shape does not
    /// own four losslessly addressable geometry attributes.
    pub fn set_shape_geometry(
        &mut self,
        page: usize,
        shape: usize,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Result<()> {
        let after = [x.into(), y.into(), width.into(), height.into()];
        validate_geometry(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = [
            selected.x(),
            selected.y(),
            selected.width(),
            selected.height(),
        ]
        .map(|value| value.map(str::to_owned));
        let spans = parsed
            .geometry_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape geometry spans are missing".into()))?;
        let ranges = spans
            .iter()
            .map(|span| {
                span.as_ref().ok_or_else(|| {
                    Error::Unsupported(
                        "ODG geometry edit requires existing x, y, width, and height attributes"
                            .into(),
                    )
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if before
            .iter()
            .zip(&after)
            .all(|(source_value, target_value)| {
                source_value.as_deref() == Some(target_value.as_str())
            })
        {
            return Ok(());
        }
        self.content = replace_xml_values(&self.content, &ranges, &after)?;
        self.changes.push(Change::Geometry(GeometryChange {
            page,
            shape,
            before: before.map(Option::unwrap_or_default),
            after,
        }));
        Ok(())
    }

    /// Changes an existing graphic style reference without normalizing XML.
    ///
    /// # Errors
    ///
    /// Returns an error for a checked-selector failure or missing source attribute.
    pub fn set_shape_style_name(
        &mut self,
        page: usize,
        shape: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let after = style_name.into();
        validate_bounded_value(&after, "ODG shape style name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = selected.style_name().ok_or_else(|| {
            Error::Unsupported("ODG style edit requires an existing draw:style-name".into())
        })?;
        let span = parsed
            .style_name_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape style span is missing".into()))?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.content = replace_xml_value(&self.content, span, &after)?;
        self.changes.push(Change::Style(StyleChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Inserts a detached page at a checked source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid position, duplicate page identity, or limit violation.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached page values transfer ownership into the transaction"
    )]
    pub fn insert_page(&mut self, position: usize, page: Page) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if position > parsed.pages.len() {
            return invalid("ODG page insertion position is out of bounds");
        }
        if parsed.pages.len() >= MAX_PAGES {
            return invalid("ODG page count exceeds the limit");
        }
        if let Some(name) = page.name()
            && parsed.pages.iter().any(|value| value.name() == Some(name))
        {
            return invalid("ODG inserted page name is already present");
        }
        let at = if position == parsed.pages.len() {
            parsed.drawing_insert_position
        } else {
            parsed.page_spans[position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?
                .start
        };
        let xml = serialize_page(&page)?;
        self.content = insert_child_xml(&self.content, at, &xml)?;
        self.changes
            .push(Change::Structure(StructureChange::PageInserted {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(())
    }

    /// Appends a detached page.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate identity or a resource limit.
    pub fn add_page(&mut self, page: Page) -> Result<()> {
        let position = parse_content(&self.content)?.pages.len();
        self.insert_page(position, page)
    }

    /// Removes one page selected by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is absent, ambiguous, or unaddressable.
    pub fn remove_page<'selector>(
        &mut self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Page> {
        let parsed = parse_content(&self.content)?;
        let position = resolve_page_position(&parsed.pages, selector.into())?;
        let page = parsed.pages[position].clone();
        let span = parsed.page_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
        self.content = remove_xml(&self.content, span)?;
        self.changes
            .push(Change::Structure(StructureChange::PageRemoved {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(page)
    }

    /// Inserts a detached shape at a checked page shape position.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, undeclared layers, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached shape values transfer ownership into the transaction"
    )]
    pub fn insert_shape(&mut self, page: usize, position: usize, shape: Shape) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if position > selected_page.shapes().len() {
            return invalid("ODG shape insertion position is out of bounds");
        }
        validate_shape_layer(selected_page, self.source.layers(), &shape)?;
        if parsed
            .pages
            .iter()
            .map(|value| value.shapes().len())
            .sum::<usize>()
            >= MAX_SHAPES
        {
            return invalid("ODG shape count exceeds the limit");
        }
        let at = if position == selected_page.shapes().len() {
            parsed.page_insert_positions[page]
                .ok_or_else(|| Error::InvalidFormat("ODG page insertion point is missing".into()))?
        } else {
            parsed.shape_spans[page][position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?
                .start
        };
        let xml = serialize_shape(&shape)?;
        self.content = insert_child_xml(&self.content, at, &xml)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeInserted {
                page,
                position,
                kind: shape.kind(),
            }));
        Ok(())
    }

    /// Appends a detached shape to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, dependencies, or limits.
    pub fn add_shape(&mut self, page: usize, shape: Shape) -> Result<()> {
        let position = parse_content(&self.content)?
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?
            .shapes()
            .len();
        self.insert_shape(page, position, shape)
    }

    /// Appends an empty structural group to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page selector or limit violation.
    pub fn add_group(&mut self, page: usize, name: impl Into<String>) -> Result<()> {
        self.add_shape(page, Shape::new(ShapeKind::Group).with_name(name))
    }

    /// Removes one shape; removing a group owns and removes its complete subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid checked selector or missing source span.
    pub fn remove_shape(&mut self, page: usize, shape: usize) -> Result<Shape> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let span = parsed.shape_spans[page][shape]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?;
        self.content = remove_xml(&self.content, span)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeRemoved {
                page,
                position: shape,
                kind: selected.kind(),
            }));
        Ok(selected)
    }

    /// Adds a page-local layer declaration.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, duplicate names, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached layer values transfer ownership into the transaction"
    )]
    pub fn add_layer(&mut self, page: usize, layer: Layer) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .layers()
            .iter()
            .any(|value| value.name() == layer.name())
        {
            return invalid("ODG page-local layer name is already present");
        }
        let layer_xml = serialize_layer(&layer)?;
        if selected.has_layer_set() {
            let at = parsed.layer_set_insert_positions[page].ok_or_else(|| {
                Error::InvalidFormat("ODG page-local layer-set insertion point is missing".into())
            })?;
            self.content = insert_child_xml(&self.content, at, &layer_xml)?;
        } else {
            let page_span = parsed.page_spans[page]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
            let at = start_tag_end(&self.content, page_span.start)?;
            let xml = format!(
                "<draw:layer-set xmlns:draw=\"{}\">{layer_xml}</draw:layer-set>",
                std::str::from_utf8(DRAW).unwrap_or_default()
            );
            self.content = insert_xml(&self.content, at, &xml)?;
        }
        self.changes
            .push(Change::Structure(StructureChange::LayerInserted {
                page,
                name: layer.name().to_owned(),
            }));
        Ok(())
    }

    /// Removes an unreferenced page-local layer by exact name.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous name or a live shape dependency.
    pub fn remove_layer(&mut self, page: usize, name: &str) -> Result<Layer> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .shapes()
            .iter()
            .any(|shape| shape.layer() == Some(name))
        {
            return Err(Error::Unsupported(
                "ODG layer removal is blocked by a shape assignment".into(),
            ));
        }
        let mut matches = selected
            .layers()
            .iter()
            .enumerate()
            .filter(|(_, layer)| layer.name() == name);
        let (position, matched_layer) = matches
            .next()
            .ok_or_else(|| Error::InvalidFormat("ODG layer selector did not match".into()))?;
        if matches.next().is_some() {
            return invalid("ODG layer name selector is ambiguous");
        }
        let removed_layer = matched_layer.clone();
        let span = parsed.layer_element_spans[page][position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG layer span is missing".into()))?;
        self.content = remove_xml(&self.content, span)?;
        self.changes
            .push(Change::Structure(StructureChange::LayerRemoved {
                page,
                name: name.to_owned(),
            }));
        Ok(removed_layer)
    }

    /// Adds or replaces one referenced package-local resource and manifest entry.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, media type, or size limit.
    pub fn set_resource(
        &mut self,
        resource: usize,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let target_media_type = media_type.into();
        validate_media_type(&target_media_type)?;
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG resource exceeds the output limit");
        }
        self.stage_resource(resource, Some(target_media_type), Some(bytes))
    }

    /// Removes one package-local resource while retaining its inert reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid resource selector.
    pub fn remove_resource(&mut self, resource: usize) -> Result<()> {
        self.stage_resource(resource, None, None)
    }

    fn stage_resource(
        &mut self,
        resource: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let selected =
            self.source.resources().get(resource).ok_or_else(|| {
                Error::InvalidFormat("ODG resource selector is out of bounds".into())
            })?;
        let before_bytes = self.source.resource_bytes(resource)?;
        let before_media_type = selected.media_type().map(str::to_owned);
        if let Some(edit) = self
            .resource_edits
            .iter_mut()
            .find(|edit| edit.resource == resource)
        {
            edit.after_media_type = after_media_type;
            edit.after_bytes = after_bytes;
        } else {
            self.resource_edits.push(ResourceEdit {
                resource,
                path: selected.path().to_owned(),
                before_media_type: before_media_type.clone(),
                after_media_type,
                before_bytes: before_bytes.clone(),
                after_bytes,
            });
        }
        self.resource_edits.retain(|edit| {
            edit.before_media_type != edit.after_media_type || edit.before_bytes != edit.after_bytes
        });
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the edited package.
    ///
    /// # Errors
    ///
    /// Returns an error when source policy, rebuilding, parsing, or typed readback fails.
    pub fn commit(self) -> Result<Commit> {
        if self.content == self.source.content_xml() && self.resource_edits.is_empty() {
            return Ok(Commit::unchanged(self.source));
        }
        ensure_compact_rewrite_source(&self.source)?;
        compact_xml::validate(self.content.as_bytes()).map_err(Error::from)?;
        let replacements = self
            .resource_edits
            .iter()
            .map(|edit| ResourceReplacement {
                path: &edit.path,
                media_type: edit.after_media_type.as_deref().unwrap_or_default(),
                bytes: edit.after_bytes.as_deref(),
            })
            .collect::<Vec<_>>();
        let snapshot = Snapshot::from_bytes(rebuild(&self.source, &self.content, &replacements)?)?;
        if snapshot.content_xml() != self.content {
            return invalid("ODG package edit failed exact content readback");
        }
        for edit in &self.resource_edits {
            let present = snapshot
                .resources()
                .iter()
                .find(|resource| resource.path() == edit.path);
            if present.and_then(Resource::media_type) != edit.after_media_type.as_deref() {
                return invalid("ODG resource edit failed manifest readback");
            }
            let actual = if snapshot.0.package.package().has_file(&edit.path)? {
                Some(snapshot.0.package.package().get_file(&edit.path)?)
            } else {
                None
            };
            if actual != edit.after_bytes {
                return invalid("ODG resource edit failed byte readback");
            }
        }
        let resource_changes = self
            .resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        Ok(Commit {
            patch: Patch {
                source: self.source,
                target: snapshot.clone(),
                changes: self.changes,
                resource_changes,
            },
            snapshot,
            changed: true,
        })
    }
}

/// One semantic operation published by a unified ODG package transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    Text(TextChange),
    Name(NameChange),
    Layer(LayerChange),
    Geometry(GeometryChange),
    Style(StyleChange),
    Structure(StructureChange),
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

/// One reversible four-attribute geometry change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GeometryChange {
    page: usize,
    shape: usize,
    before: [String; 4],
    after: [String; 4],
}

impl GeometryChange {
    /// Page position at the time of this operation.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// Shape position at the time of this operation.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Source `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn before(&self) -> &[String; 4] {
        &self.before
    }

    /// Target `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn after(&self) -> &[String; 4] {
        &self.after
    }
}

/// One reversible graphic-style reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl StyleChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A structural page, layer, shape, or group operation.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum StructureChange {
    PageInserted {
        position: usize,
        name: Option<String>,
    },
    PageRemoved {
        position: usize,
        name: Option<String>,
    },
    LayerInserted {
        page: usize,
        name: String,
    },
    LayerRemoved {
        page: usize,
        name: String,
    },
    ShapeInserted {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
    ShapeRemoved {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
}

/// One package-local resource replacement or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    #[must_use]
    pub const fn resource(&self) -> usize {
        self.resource
    }

    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn before_media_type(&self) -> Option<&str> {
        self.before_media_type.as_deref()
    }

    #[must_use]
    pub fn after_media_type(&self) -> Option<&str> {
        self.after_media_type.as_deref()
    }

    #[must_use]
    pub const fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    #[must_use]
    pub const fn after_size(&self) -> Option<usize> {
        self.after_size
    }
}

struct ResourceEdit {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_bytes: Option<Vec<u8>>,
    after_bytes: Option<Vec<u8>>,
}

impl ResourceEdit {
    fn change(&self) -> ResourceChange {
        ResourceChange {
            resource: self.resource,
            path: self.path.clone(),
            before_media_type: self.before_media_type.clone(),
            after_media_type: self.after_media_type.clone(),
            before_size: self.before_bytes.as_ref().map(Vec::len),
            after_size: self.after_bytes.as_ref().map(Vec::len),
        }
    }
}

struct ResourceReplacement<'a> {
    path: &'a str,
    media_type: &'a str,
    bytes: Option<&'a [u8]>,
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
                changes: Vec::new(),
                resource_changes: Vec::new(),
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
    changes: Vec<Change>,
    resource_changes: Vec<ResourceChange>,
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
        self.changes.iter().find_map(|change| match change {
            Change::Text(value) => Some(value),
            Change::Name(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic `draw:name` change, when this is a name patch.
    #[must_use]
    pub fn name_change(&self) -> Option<&NameChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Name(value) => Some(value),
            Change::Text(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic drawing-layer change, when present.
    #[must_use]
    pub fn layer_change(&self) -> Option<&LayerChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Layer(value) => Some(value),
            Change::Text(_)
            | Change::Name(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Structure(_) => None,
        })
    }

    /// All semantic operations in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Package-local resource changes in source selector order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Composes adjacent exact-lineage patches.
    ///
    /// # Errors
    ///
    /// Returns an error unless this target is byte-identical to `next`'s source.
    pub fn then(&self, next: &Self) -> Result<Self> {
        if self.target.as_bytes() != next.source.as_bytes() {
            return invalid("ODG patch composition lineage does not match");
        }
        let mut changes = self.changes.clone();
        changes.extend_from_slice(&next.changes);
        let mut resource_changes = self.resource_changes.clone();
        resource_changes.extend_from_slice(&next.resource_changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
            resource_changes,
        })
    }

    /// An exact-source patch restoring the original package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().rev().map(inverse_change).collect(),
            resource_changes: self
                .resource_changes
                .iter()
                .rev()
                .map(inverse_resource_change)
                .collect(),
        }
    }
}

struct Parsed {
    pages: Vec<Page>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    layer_count: usize,
    geometry_spans: GeometrySpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    page_spans: Vec<Option<Range<usize>>>,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_insert_positions: Vec<Option<usize>>,
    drawing_insert_position: usize,
}

struct ActiveShape {
    depth: usize,
    page: usize,
    shape: usize,
    start: usize,
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
    page_starts: Vec<usize>,
    layer_sets: Vec<(usize, Option<usize>)>,
    active_shapes: Vec<ActiveShape>,
    active_accessibility: Option<ActiveAccessibility>,
    paragraph_depths: Vec<usize>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    geometry_spans: GeometrySpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_count: usize,
    shape_count: usize,
    text_bytes: usize,
    page_spans: Vec<Option<Range<usize>>>,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_starts: Vec<(usize, usize, usize)>,
    layer_set_insert_positions: Vec<Option<usize>>,
    active_layers: Vec<(usize, usize, usize, usize)>,
    drawing_insert_position: Option<usize>,
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
            page_starts: Vec::new(),
            layer_sets: Vec::new(),
            active_shapes: Vec::new(),
            active_accessibility: None,
            paragraph_depths: Vec::new(),
            text_spans: Vec::new(),
            name_spans: Vec::new(),
            layer_spans: Vec::new(),
            geometry_spans: Vec::new(),
            style_name_spans: Vec::new(),
            layer_count: 0,
            shape_count: 0,
            text_bytes: 0,
            page_spans: Vec::new(),
            page_insert_positions: Vec::new(),
            shape_spans: Vec::new(),
            layer_element_spans: Vec::new(),
            layer_set_starts: Vec::new(),
            layer_set_insert_positions: Vec::new(),
            active_layers: Vec::new(),
            drawing_insert_position: None,
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
            if empty {
                self.drawing_insert_position = Some(tag_start + tag.len() - 2);
            } else {
                self.drawing_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"layer-set" {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer-set is outside draw:page".to_string())
            })?;
            self.pages[page].mark_layer_set();
            if empty {
                self.layer_set_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.layer_sets.push((self.depth, Some(page)));
                self.layer_set_starts.push((self.depth, page, tag_start));
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
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer is outside draw:page".to_string())
            })?;
            let layer = self.pages[page].layers().len() - 1;
            if empty {
                self.layer_element_spans[page][layer] = Some(tag_start..tag_start + tag.len());
            } else {
                self.active_layers
                    .push((self.depth, page, layer, tag_start));
            }
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
            self.geometry_spans.push(Vec::new());
            self.style_name_spans.push(Vec::new());
            self.page_spans.push(None);
            self.page_insert_positions.push(None);
            self.shape_spans.push(Vec::new());
            self.layer_element_spans.push(Vec::new());
            self.layer_set_insert_positions.push(None);
            if empty {
                let page = self.pages.len() - 1;
                self.page_spans[page] = Some(tag_start..tag_start + tag.len());
                self.page_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.page_depths.push(self.depth);
                self.page_starts.push(tag_start);
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
            self.pages[page].push_shape(Shape::parsed(
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
            self.geometry_spans[page].push([
                attribute_source_span(reader, element, tag, tag_start, SVG, b"x")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"y")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"width")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"height")?,
            ]);
            self.style_name_spans[page].push(attribute_source_span(
                reader,
                element,
                tag,
                tag_start,
                DRAW,
                b"style-name",
            )?);
            self.shape_spans[page].push(empty.then_some(tag_start..tag_start + tag.len()));
            if !empty {
                self.active_shapes.push(ActiveShape {
                    depth: self.depth,
                    page,
                    shape,
                    start: tag_start,
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
            self.layer_element_spans[*page].push(None);
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

    fn end(
        &mut self,
        namespace: NamespaceKind,
        local: &[u8],
        tag_start: usize,
        tag_end: usize,
    ) -> Result<()> {
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
            let active = self
                .active_shapes
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
            self.shape_spans[active.page][active.shape] = Some(active.start..tag_end);
        }
        if self
            .active_layers
            .last()
            .is_some_and(|layer| layer.0 == self.depth)
        {
            let (_, page, layer, start) = self
                .active_layers
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active layer disappeared".to_string()))?;
            self.layer_element_spans[page][layer] = Some(start..tag_end);
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer-set"
            && self
                .layer_sets
                .last()
                .is_some_and(|set| set.0 == self.depth)
        {
            self.layer_sets.pop();
            let (_, page, _) = self.layer_set_starts.pop().ok_or_else(|| {
                Error::InvalidFormat("ODG active layer-set disappeared".to_string())
            })?;
            self.layer_set_insert_positions[page] = Some(tag_start);
        }
        if namespace == NamespaceKind::Draw
            && local == b"page"
            && self.page_depths.last() == Some(&self.depth)
        {
            self.page_depths.pop();
            let start = self
                .page_starts
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active page disappeared".to_string()))?;
            let page = self.pages.len() - 1;
            self.page_spans[page] = Some(start..tag_end);
            self.page_insert_positions[page] = Some(tag_start);
        }
        if self.drawing_depth == Some(self.depth) {
            self.drawing_depth = None;
            self.drawing_insert_position = Some(tag_start);
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
            geometry_spans: self.geometry_spans,
            style_name_spans: self.style_name_spans,
            layer_count: self.layer_count,
            page_spans: self.page_spans,
            page_insert_positions: self.page_insert_positions,
            shape_spans: self.shape_spans,
            layer_element_spans: self.layer_element_spans,
            layer_set_insert_positions: self.layer_set_insert_positions,
            drawing_insert_position: self.drawing_insert_position.ok_or_else(|| {
                Error::InvalidFormat("ODG drawing insertion point is missing".to_string())
            })?,
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
            Event::End(element) => {
                scanner.end(namespace, element.local_name().as_ref(), start, end)?;
            },
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

fn scan_resources(package: &Package) -> Result<Vec<Resource>> {
    let archive = package.package().package()?;
    let images = media::scan_package(package.content_xml(), package.styles_xml(), &archive)?;
    let mut resources = Vec::new();
    for (occurrence, image) in images.into_iter().enumerate() {
        match image.source {
            media::Source::PackagePart {
                href,
                path,
                manifest_media_type,
            } => resources.push(Resource::new(
                occurrence,
                href,
                path,
                manifest_media_type,
                true,
            )),
            media::Source::MissingPackagePart {
                href,
                resolved_path,
            } => resources.push(Resource::new(occurrence, href, resolved_path, None, false)),
            media::Source::Inline { .. }
            | media::Source::Linked { .. }
            | media::Source::Missing
            | _ => {},
        }
    }
    Ok(resources)
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

fn rebuild(
    source: &Snapshot,
    content: &str,
    replacements: &[ResourceReplacement<'_>],
) -> Result<Vec<u8>> {
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
    let excluded = replacements
        .iter()
        .map(|replacement| replacement.path.to_owned())
        .collect::<Vec<_>>();
    writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
    for replacement in replacements {
        if let Some(bytes) = replacement.bytes {
            writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
        }
    }
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

fn replace_xml_values(
    source: &str,
    spans: &[&Range<usize>],
    replacements: &[String; 4],
) -> Result<String> {
    let mut edits = spans
        .iter()
        .zip(replacements)
        .map(|(span, value)| ((*span).clone(), value.as_str()))
        .collect::<Vec<_>>();
    edits.sort_unstable_by_key(|(span, _)| std::cmp::Reverse(span.start));
    let mut output = source.to_owned();
    for (span, replacement) in edits {
        output = replace_xml_value(&output, &span, replacement)?;
    }
    Ok(output)
}

fn insert_xml(source: &str, at: usize, xml: &str) -> Result<String> {
    if at > source.len() || !source.is_char_boundary(at) {
        return invalid("ODG XML insertion point is invalid");
    }
    let capacity = source
        .len()
        .checked_add(xml.len())
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".into()))?;
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
    output.push_str(&source[..at]);
    output.push_str(xml);
    output.push_str(&source[at..]);
    Ok(output)
}

fn insert_child_xml(source: &str, at: usize, child: &str) -> Result<String> {
    if source.as_bytes().get(at..at.saturating_add(2)) != Some(b"/>") {
        return insert_xml(source, at, child);
    }
    let element_start = source
        .get(..at)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner start is missing".into()))?;
    let name_start = element_start + 1;
    let name_end = source
        .as_bytes()
        .get(name_start..at)
        .and_then(|bytes| {
            bytes
                .iter()
                .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
                .map(|offset| name_start + offset)
        })
        .unwrap_or(at);
    let name = source
        .get(name_start..name_end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner name is missing".into()))?;
    let replacement = format!(">{child}</{name}>");
    let mut output = String::with_capacity(
        source
            .len()
            .saturating_sub(2)
            .saturating_add(replacement.len()),
    );
    output.push_str(&source[..at]);
    output.push_str(&replacement);
    output.push_str(&source[at + 2..]);
    if output.len() > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    Ok(output)
}

fn remove_xml(source: &str, span: &Range<usize>) -> Result<String> {
    if span.start > span.end
        || span.end > source.len()
        || !source.is_char_boundary(span.start)
        || !source.is_char_boundary(span.end)
    {
        return invalid("ODG XML removal span is invalid");
    }
    let mut output = String::with_capacity(source.len() - (span.end - span.start));
    output.push_str(&source[..span.start]);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn start_tag_end(source: &str, start: usize) -> Result<usize> {
    source
        .get(start..)
        .and_then(|tail| tail.find('>').map(|offset| start + offset + 1))
        .ok_or_else(|| Error::InvalidFormat("ODG element start tag is unterminated".into()))
}

fn serialize_page(page: &Page) -> Result<String> {
    let mut xml = format!(
        "<draw:page xmlns:draw=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default()
    );
    push_attribute(&mut xml, "draw:name", page.name())?;
    push_attribute(&mut xml, "xml:id", page.xml_id())?;
    push_attribute(&mut xml, "draw:style-name", page.style_name())?;
    push_attribute(&mut xml, "draw:master-page-name", page.master_page_name())?;
    xml.push_str("></draw:page>");
    Ok(xml)
}

fn serialize_layer(layer: &Layer) -> Result<String> {
    validate_bounded_value(layer.name(), "ODG layer name")?;
    let mut xml = String::from("<draw:layer");
    push_attribute(&mut xml, "draw:name", Some(layer.name()))?;
    push_attribute(&mut xml, "draw:display", layer.display())?;
    if let Some(protected) = layer.protected() {
        push_attribute(
            &mut xml,
            "draw:protected",
            Some(if protected { "true" } else { "false" }),
        )?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn serialize_shape(shape: &Shape) -> Result<String> {
    let element = shape.kind().element_name();
    let mut xml = format!(
        "<draw:{element} xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(TEXT).unwrap_or_default()
    );
    push_attribute(&mut xml, "draw:name", shape.name())?;
    push_attribute(&mut xml, "draw:layer", shape.layer())?;
    push_attribute(&mut xml, "draw:style-name", shape.style_name())?;
    push_attribute(&mut xml, "draw:text-style-name", shape.text_style_name())?;
    if let Some(z_index) = shape.z_index() {
        push_attribute(&mut xml, "draw:z-index", Some(&z_index.to_string()))?;
    }
    push_attribute(&mut xml, "svg:x", shape.x())?;
    push_attribute(&mut xml, "svg:y", shape.y())?;
    push_attribute(&mut xml, "svg:width", shape.width())?;
    push_attribute(&mut xml, "svg:height", shape.height())?;
    if shape.title().is_none() && shape.description().is_none() && shape.text().is_empty() {
        xml.push_str("/>");
        return Ok(xml);
    }
    xml.push('>');
    if let Some(title) = shape.title() {
        xml.push_str("<svg:title>");
        xml.push_str(&quick_xml::escape::escape(title));
        xml.push_str("</svg:title>");
    }
    if let Some(description) = shape.description() {
        xml.push_str("<svg:desc>");
        xml.push_str(&quick_xml::escape::escape(description));
        xml.push_str("</svg:desc>");
    }
    if !shape.text().is_empty() {
        xml.push_str("<text:p>");
        xml.push_str(&quick_xml::escape::escape(shape.text()));
        xml.push_str("</text:p>");
    }
    xml.push_str("</draw:");
    xml.push_str(element);
    xml.push('>');
    Ok(xml)
}

fn push_attribute(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(attribute_value) = value else {
        return Ok(());
    };
    validate_bounded_value(attribute_value, "ODG XML attribute value")?;
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&quick_xml::escape::escape(attribute_value));
    output.push('"');
    Ok(())
}

fn validate_bounded_value(value: &str, owner: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} is invalid")));
    }
    Ok(())
}

fn validate_geometry(values: &[String; 4]) -> Result<()> {
    for value in values {
        validate_bounded_value(value, "ODG geometry value")?;
        if value.bytes().any(|byte| byte.is_ascii_whitespace()) {
            return invalid("ODG geometry value contains whitespace");
        }
    }
    Ok(())
}

fn validate_shape_layer(page: &Page, global_layers: &[Layer], shape: &Shape) -> Result<()> {
    let Some(layer) = shape.layer() else {
        return Ok(());
    };
    let visible = if page.has_layer_set() {
        page.layers()
    } else {
        global_layers
    };
    if !visible.iter().any(|value| value.name() == layer) {
        return invalid("ODG inserted shape references an undeclared layer");
    }
    Ok(())
}

fn resolve_page_position(pages: &[Page], selector: crate::page::Selector<'_>) -> Result<usize> {
    match selector {
        crate::page::Selector::Position(position) => pages
            .get(position.get())
            .map(|_| position.get())
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into())),
        crate::page::Selector::Name(name) => {
            let mut matches = pages
                .iter()
                .enumerate()
                .filter(|(_, page)| page.name() == Some(name.as_ref()));
            let selected = matches
                .next()
                .ok_or_else(|| Error::InvalidFormat("ODG page selector did not match".into()))?;
            if matches.next().is_some() {
                return invalid("ODG page name selector is ambiguous");
            }
            Ok(selected.0)
        },
    }
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || !media_type.contains('/')
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return invalid("ODG resource media type is invalid");
    }
    Ok(())
}

fn inverse_change(change: &Change) -> Change {
    match change {
        Change::Text(value) => Change::Text(TextChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Name(value) => Change::Name(NameChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Layer(value) => Change::Layer(LayerChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Geometry(value) => Change::Geometry(GeometryChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Style(value) => Change::Style(StyleChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Structure(value) => Change::Structure(match value {
            StructureChange::PageInserted { position, name } => StructureChange::PageRemoved {
                position: *position,
                name: name.clone(),
            },
            StructureChange::PageRemoved { position, name } => StructureChange::PageInserted {
                position: *position,
                name: name.clone(),
            },
            StructureChange::LayerInserted { page, name } => StructureChange::LayerRemoved {
                page: *page,
                name: name.clone(),
            },
            StructureChange::LayerRemoved { page, name } => StructureChange::LayerInserted {
                page: *page,
                name: name.clone(),
            },
            StructureChange::ShapeInserted {
                page,
                position,
                kind,
            } => StructureChange::ShapeRemoved {
                page: *page,
                position: *position,
                kind: *kind,
            },
            StructureChange::ShapeRemoved {
                page,
                position,
                kind,
            } => StructureChange::ShapeInserted {
                page: *page,
                position: *position,
                kind: *kind,
            },
        }),
    }
}

fn inverse_resource_change(change: &ResourceChange) -> ResourceChange {
    ResourceChange {
        resource: change.resource,
        path: change.path.clone(),
        before_media_type: change.after_media_type.clone(),
        after_media_type: change.before_media_type.clone(),
        before_size: change.after_size,
        after_size: change.before_size,
    }
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
