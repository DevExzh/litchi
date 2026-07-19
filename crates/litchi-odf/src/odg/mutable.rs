//! Mutable and builder APIs for packaged OpenDocument drawings.

use super::{
    DrawingDocument, DrawingLayer, DrawingLayerDisplay, DrawingPage, DrawingPageProperties,
};
use crate::constants;
use crate::core::{OdfStructure, OwnedPackage, PackageWriter};
use crate::odp::{DrawingAttributeNamespace, PresentationBuilder};
use crate::Shape;
use litchi_core::{Error, Metadata, Result, xml::escape_xml};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeSet;
use std::io::Write;
use std::ops::Range;
use std::path::Path;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_PAGES: usize = 1_000_000;
const MAX_LAYERS_PER_PAGE: usize = 65_536;
const MAX_SHAPES_PER_PAGE: usize = 65_536;
const MAX_SHAPE_DEPTH: usize = 64;
const MAX_STRING_BYTES: usize = 1_048_576;

#[derive(Clone)]
struct MutablePage {
    page: DrawingPage,
    raw_xml: Option<String>,
    dirty: bool,
}

struct SourceDrawing {
    bytes: Vec<u8>,
    package: OwnedPackage,
    content_prefix: String,
    content_suffix: String,
}

/// Atomic mutable authoring model for `.odg` and `.otg` packages.
pub struct MutableDrawing {
    pages: Vec<MutablePage>,
    metadata: Metadata,
    metadata_dirty: bool,
    mimetype: String,
    source: Option<SourceDrawing>,
}

impl MutableDrawing {
    /// Create an empty packaged drawing.
    pub fn new() -> Self {
        Self::with_mimetype(constants::ODF_DRAWING)
    }

    /// Create an empty packaged drawing template.
    pub fn new_template() -> Self {
        Self::with_mimetype(constants::ODF_DRAWING_TEMPLATE)
    }

    fn with_mimetype(mimetype: &str) -> Self {
        Self {
            pages: Vec::new(),
            metadata: Metadata::default(),
            metadata_dirty: true,
            mimetype: mimetype.to_string(),
            source: None,
        }
    }

    /// Open a drawing or drawing template for mutation.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        DrawingDocument::open(path)?.into_mutable()
    }

    /// Parse packaged drawing bytes for mutation.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        DrawingDocument::from_bytes(bytes)?.into_mutable()
    }

    pub(crate) fn from_document_parts(
        source_bytes: Vec<u8>,
        mimetype: String,
        content: String,
        metadata: Metadata,
        pages: Vec<DrawingPage>,
    ) -> Result<Self> {
        let (inner, ranges) = drawing_page_ranges(&content)?;
        if ranges.len() != pages.len() {
            return Err(Error::InvalidFormat(
                "drawing page inventory does not match content.xml".to_string(),
            ));
        }
        let mutable_pages = pages
            .into_iter()
            .zip(ranges)
            .map(|(page, range)| MutablePage {
                page,
                raw_xml: Some(content[range].to_string()),
                dirty: false,
            })
            .collect::<Vec<_>>();
        validate_pages(&mutable_pages)?;
        let package = OwnedPackage::from_bytes(source_bytes.clone())?;
        Ok(Self {
            pages: mutable_pages,
            metadata,
            metadata_dirty: false,
            mimetype,
            source: Some(SourceDrawing {
                bytes: source_bytes,
                package,
                content_prefix: content[..inner.start].to_string(),
                content_suffix: content[inner.end..].to_string(),
            }),
        })
    }

    /// Return the package MIME type.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Whether this drawing uses the `.otg` template MIME type.
    pub fn is_template(&self) -> bool {
        self.mimetype == constants::ODF_DRAWING_TEMPLATE
    }

    /// Return common document metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Replace common document metadata.
    pub fn set_metadata(&mut self, metadata: Metadata) -> Result<()> {
        validate_metadata(&metadata)?;
        self.metadata = metadata;
        self.metadata_dirty = true;
        Ok(())
    }

    /// Return the number of drawing pages.
    pub fn page_count(&self) -> usize {
        self.pages.len()
    }

    /// Return a drawing page by zero-based index.
    pub fn page(&self, index: usize) -> Option<&DrawingPage> {
        self.pages.get(index).map(|page| &page.page)
    }

    /// Iterate over drawing pages in package order.
    pub fn pages(&self) -> impl ExactSizeIterator<Item = &DrawingPage> {
        self.pages.iter().map(|page| &page.page)
    }

    /// Append an empty page and return its index.
    pub fn add_page(&mut self, properties: DrawingPageProperties) -> Result<usize> {
        let index = self.pages.len();
        self.insert_page(index, DrawingPage::new(properties))?;
        Ok(index)
    }

    /// Insert a complete page at a zero-based index.
    pub fn insert_page(&mut self, index: usize, page: DrawingPage) -> Result<()> {
        self.update_pages(|pages| {
            if index > pages.len() {
                return Err(bounds("page", index, pages.len()));
            }
            pages.insert(
                index,
                MutablePage {
                    page,
                    raw_xml: None,
                    dirty: true,
                },
            );
            Ok(())
        })
    }

    /// Remove and return a page.
    pub fn remove_page(&mut self, index: usize) -> Result<DrawingPage> {
        self.update_pages(|pages| {
            if index >= pages.len() {
                return Err(bounds("page", index, pages.len()));
            }
            Ok(pages.remove(index).page)
        })
    }

    /// Move a page, preserving its exact source XML when it was otherwise untouched.
    pub fn move_page(&mut self, from: usize, to: usize) -> Result<()> {
        self.update_pages(|pages| {
            if from >= pages.len() || to >= pages.len() {
                return Err(bounds("page", from.max(to), pages.len()));
            }
            let page = pages.remove(from);
            pages.insert(to, page);
            Ok(())
        })
    }

    /// Replace all standard properties on one page.
    pub fn set_page_properties(
        &mut self,
        page_index: usize,
        properties: DrawingPageProperties,
    ) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            page.page.properties = properties;
            page.dirty = true;
            Ok(())
        })
    }

    /// Append a layer declaration and return its index.
    pub fn add_layer(&mut self, page_index: usize, layer: DrawingLayer) -> Result<usize> {
        let index = self
            .pages
            .get(page_index)
            .ok_or_else(|| bounds("page", page_index, self.pages.len()))?
            .page
            .layers
            .len();
        self.insert_layer(page_index, index, layer)?;
        Ok(index)
    }

    /// Insert a layer declaration.
    pub fn insert_layer(
        &mut self,
        page_index: usize,
        layer_index: usize,
        layer: DrawingLayer,
    ) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if layer_index > page.page.layers.len() {
                return Err(bounds("layer", layer_index, page.page.layers.len()));
            }
            page.page.layers.insert(layer_index, layer);
            page.dirty = true;
            Ok(())
        })
    }

    /// Replace a layer declaration, including an atomic rename.
    pub fn set_layer(
        &mut self,
        page_index: usize,
        layer_index: usize,
        layer: DrawingLayer,
    ) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            let length = page.page.layers.len();
            let slot = page
                .page
                .layers
                .get_mut(layer_index)
                .ok_or_else(|| bounds("layer", layer_index, length))?;
            *slot = layer;
            page.dirty = true;
            Ok(())
        })
    }

    /// Remove and return an unreferenced layer declaration.
    pub fn remove_layer(
        &mut self,
        page_index: usize,
        layer_index: usize,
    ) -> Result<DrawingLayer> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if layer_index >= page.page.layers.len() {
                return Err(bounds("layer", layer_index, page.page.layers.len()));
            }
            page.dirty = true;
            Ok(page.page.layers.remove(layer_index))
        })
    }

    /// Reorder a layer declaration.
    pub fn move_layer(&mut self, page_index: usize, from: usize, to: usize) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if from >= page.page.layers.len() || to >= page.page.layers.len() {
                return Err(bounds("layer", from.max(to), page.page.layers.len()));
            }
            let layer = page.page.layers.remove(from);
            page.page.layers.insert(to, layer);
            page.dirty = true;
            Ok(())
        })
    }

    /// Append a shape and return its index.
    pub fn add_shape(&mut self, page_index: usize, shape: Shape) -> Result<usize> {
        let index = self
            .pages
            .get(page_index)
            .ok_or_else(|| bounds("page", page_index, self.pages.len()))?
            .page
            .shapes()
            .len();
        self.insert_shape(page_index, index, shape)?;
        Ok(index)
    }

    /// Insert a top-level shape.
    pub fn insert_shape(
        &mut self,
        page_index: usize,
        shape_index: usize,
        shape: Shape,
    ) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if shape_index > page.page.page.shapes.len() {
                return Err(bounds("shape", shape_index, page.page.page.shapes.len()));
            }
            page.page.page.shapes.insert(shape_index, shape);
            page.dirty = true;
            Ok(())
        })
    }

    /// Replace a top-level shape.
    pub fn set_shape(
        &mut self,
        page_index: usize,
        shape_index: usize,
        shape: Shape,
    ) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            let length = page.page.page.shapes.len();
            let slot = page
                .page
                .page
                .shapes
                .get_mut(shape_index)
                .ok_or_else(|| bounds("shape", shape_index, length))?;
            *slot = shape;
            page.dirty = true;
            Ok(())
        })
    }

    /// Remove and return a top-level shape.
    pub fn remove_shape(&mut self, page_index: usize, shape_index: usize) -> Result<Shape> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if shape_index >= page.page.page.shapes.len() {
                return Err(bounds("shape", shape_index, page.page.page.shapes.len()));
            }
            page.dirty = true;
            Ok(page.page.page.shapes.remove(shape_index))
        })
    }

    /// Reorder a top-level shape.
    pub fn move_shape(&mut self, page_index: usize, from: usize, to: usize) -> Result<()> {
        self.update_pages(|pages| {
            let page = get_page_mut(pages, page_index)?;
            if from >= page.page.page.shapes.len() || to >= page.page.page.shapes.len() {
                return Err(bounds("shape", from.max(to), page.page.page.shapes.len()));
            }
            let shape = page.page.page.shapes.remove(from);
            page.page.page.shapes.insert(to, shape);
            page.dirty = true;
            Ok(())
        })
    }

    fn update_pages<T>(
        &mut self,
        operation: impl FnOnce(&mut Vec<MutablePage>) -> Result<T>,
    ) -> Result<T> {
        let mut staged = self.pages.clone();
        let result = operation(&mut staged)?;
        for (index, page) in staged.iter_mut().enumerate() {
            page.page.page.index = index;
        }
        validate_pages(&staged)?;
        self.pages = staged;
        Ok(result)
    }

    fn generate_content_xml(&self) -> Result<String> {
        validate_pages(&self.pages)?;
        let mut pages_xml = String::new();
        for page in &self.pages {
            if !page.dirty
                && let Some(raw) = &page.raw_xml
            {
                pages_xml.push_str(raw);
            } else {
                pages_xml.push_str(&serialize_page(&page.page)?);
            }
        }
        if let Some(source) = &self.source {
            return Ok(format!(
                "{}{}{}",
                source.content_prefix, pages_xml, source.content_suffix
            ));
        }
        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" office:version="1.3"><office:automatic-styles/><office:body><office:drawing>{}</office:drawing></office:body></office:document-content>"#,
            pages_xml
        ))
    }

    /// Serialize a validated package. The generated bytes are reparsed before return.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.source.is_some()
            && !self.metadata_dirty
            && self.pages.iter().all(|page| !page.dirty)
        {
            return Ok(self.source.as_ref().expect("checked").bytes.clone());
        }
        let content = self.generate_content_xml()?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(&self.mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        if let Some(source) = &self.source {
            if source.package.has_file(constants::ODF_STYLES)? {
                writer.add_file(
                    constants::ODF_STYLES,
                    &source.package.get_file(constants::ODF_STYLES)?,
                )?;
            }
        } else {
            writer.add_file(constants::ODF_STYLES, OdfStructure::default_styles_xml().as_bytes())?;
        }
        if self.metadata_dirty || self.source.is_none() {
            writer.add_file(constants::ODF_META, generate_meta_xml(&self.metadata).as_bytes())?;
        } else if let Some(source) = &self.source
            && source.package.has_file(constants::ODF_META)?
        {
            writer.add_file(
                constants::ODF_META,
                &source.package.get_file(constants::ODF_META)?,
            )?;
        }
        if let Some(source) = &self.source {
            writer.copy_auxiliary_files_from(&source.package)?;
        }
        let bytes = writer.finish_to_bytes()?;
        DrawingDocument::from_bytes(bytes.clone())?;
        Ok(bytes)
    }

    /// Save through a same-directory temporary file and atomic rename.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let bytes = self.to_bytes()?;
        atomic_write(path.as_ref(), &bytes)
    }
}

impl Default for MutableDrawing {
    fn default() -> Self {
        Self::new()
    }
}

/// Fluent packaged drawing builder.
pub struct DrawingBuilder {
    drawing: MutableDrawing,
}

impl DrawingBuilder {
    /// Start a `.odg` drawing.
    pub fn new() -> Self {
        Self {
            drawing: MutableDrawing::new(),
        }
    }

    /// Start an `.otg` drawing template.
    pub fn template() -> Self {
        Self {
            drawing: MutableDrawing::new_template(),
        }
    }

    /// Append a page.
    pub fn add_page(&mut self, properties: DrawingPageProperties) -> Result<&mut Self> {
        self.drawing.add_page(properties)?;
        Ok(self)
    }

    /// Replace common metadata.
    pub fn metadata(&mut self, metadata: Metadata) -> Result<&mut Self> {
        self.drawing.set_metadata(metadata)?;
        Ok(self)
    }

    /// Finish into the mutable model for layer and shape authoring.
    pub fn build(self) -> MutableDrawing {
        self.drawing
    }

    /// Serialize the current drawing package.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.drawing.to_bytes()
    }

    /// Save the current drawing package atomically.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.drawing.save(path)
    }
}

impl Default for DrawingBuilder {
    fn default() -> Self {
        Self::new()
    }
}

fn get_page_mut(pages: &mut [MutablePage], index: usize) -> Result<&mut MutablePage> {
    let length = pages.len();
    pages
        .get_mut(index)
        .ok_or_else(|| bounds("page", index, length))
}

fn bounds(kind: &str, index: usize, length: usize) -> Error {
    Error::InvalidFormat(format!("{kind} index {index} out of bounds for length {length}"))
}

fn validate_string(label: &str, value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODG {label} exceeds {MAX_STRING_BYTES} bytes"
        )));
    }
    if value.chars().any(|character| character == '\0') {
        return Err(Error::InvalidFormat(format!("ODG {label} contains NUL")));
    }
    Ok(())
}

fn validate_xml_id(label: &str, value: &str) -> Result<()> {
    validate_string(label, value)?;
    let mut chars = value.chars();
    if !chars
        .next()
        .is_some_and(|character| character == '_' || character.is_ascii_alphabetic())
        || !chars.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric()
        })
    {
        return Err(Error::InvalidFormat(format!("invalid ODG {label} '{value}'")));
    }
    Ok(())
}

fn validate_pages(pages: &[MutablePage]) -> Result<()> {
    if pages.len() > MAX_PAGES {
        return Err(Error::InvalidFormat(format!(
            "ODG exceeds {MAX_PAGES} pages"
        )));
    }
    let mut document_ids = BTreeSet::new();
    for page in pages {
        for value in page.page.properties.values().into_iter().flatten() {
            validate_string("page property", value)?;
        }
        for (label, id) in [
            ("draw:id", page.page.properties.draw_id()),
            ("xml:id", page.page.properties.xml_id()),
        ] {
            if let Some(id) = id {
                validate_xml_id(label, id)?;
                if !document_ids.insert(id.to_string()) {
                    return Err(Error::InvalidFormat(format!(
                        "duplicate ODG identifier '{id}'"
                    )));
                }
            }
        }
        validate_page(page, &mut document_ids)?;
    }
    Ok(())
}

fn validate_page(page: &MutablePage, document_ids: &mut BTreeSet<String>) -> Result<()> {
    if page.page.layers.len() > MAX_LAYERS_PER_PAGE {
        return Err(Error::InvalidFormat(format!(
            "ODG page exceeds {MAX_LAYERS_PER_PAGE} layers"
        )));
    }
    let mut layers = BTreeSet::new();
    for layer in &page.page.layers {
        validate_string("layer name", layer.name())?;
        if layer.name().is_empty() || !layers.insert(layer.name().to_string()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate or empty ODG layer name '{}'",
                layer.name()
            )));
        }
    }
    let mut shape_count = 0usize;
    let mut page_shape_ids = BTreeSet::new();
    for shape in page.page.shapes() {
        validate_shape(
            shape,
            0,
            &mut shape_count,
            &layers,
            &mut page_shape_ids,
            document_ids,
        )?;
    }
    if let Some(order) = page.page.properties.navigation_order() {
        let mut seen = BTreeSet::new();
        for id in order.split_whitespace() {
            validate_xml_id("navigation IDREF", id)?;
            if !seen.insert(id) || !page_shape_ids.contains(id) {
                return Err(Error::InvalidFormat(format!(
                    "ODG navigation order contains duplicate or unresolved ID '{id}'"
                )));
            }
        }
    }
    Ok(())
}

fn validate_shape(
    shape: &Shape,
    depth: usize,
    count: &mut usize,
    layers: &BTreeSet<String>,
    page_ids: &mut BTreeSet<String>,
    document_ids: &mut BTreeSet<String>,
) -> Result<()> {
    if depth > MAX_SHAPE_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "ODG shape nesting exceeds {MAX_SHAPE_DEPTH} levels"
        )));
    }
    *count = count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG shape count overflow".to_string()))?;
    if *count > MAX_SHAPES_PER_PAGE {
        return Err(Error::InvalidFormat(format!(
            "ODG page exceeds {MAX_SHAPES_PER_PAGE} shapes"
        )));
    }
    for value in [
        shape.name.as_deref(),
        shape.text.is_empty().then_some("").or(Some(shape.text.as_str())),
        shape.x.as_deref(),
        shape.y.as_deref(),
        shape.width.as_deref(),
        shape.height.as_deref(),
        shape.style_name.as_deref(),
        shape.layer.as_deref(),
        shape.z_index.as_deref(),
        shape.transform.as_deref(),
        shape.image_href.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_string("shape value", value)?;
    }
    let effective_layer = shape.layer.as_deref().or_else(|| {
        (shape.drawing_kind.is_none() && shape.shape_type != litchi_core::ShapeType::Group)
            .then_some("layout")
    });
    if let Some(layer) = effective_layer
        && !layers.contains(layer)
    {
        return Err(Error::InvalidFormat(format!(
            "ODG shape references undeclared layer '{layer}'"
        )));
    }
    for attribute in shape.drawing_attributes() {
        validate_string("shape attribute", attribute.value())?;
        if attribute.namespace() == DrawingAttributeNamespace::Drawing
            && attribute.local_name() == "id"
        {
            let id = attribute.value();
            validate_xml_id("shape draw:id", id)?;
            if !page_ids.insert(id.to_string()) || !document_ids.insert(id.to_string()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate ODG shape identifier '{id}'"
                )));
            }
        }
    }
    for child in shape.children() {
        validate_shape(
            child,
            depth + 1,
            count,
            layers,
            page_ids,
            document_ids,
        )?;
    }
    Ok(())
}

fn validate_metadata(metadata: &Metadata) -> Result<()> {
    for value in [
        metadata.title.as_deref(),
        metadata.author.as_deref(),
        metadata.subject.as_deref(),
        metadata.description.as_deref(),
        metadata.keywords.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_string("metadata value", value)?;
    }
    Ok(())
}

fn push_attribute(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
}

fn serialize_page(page: &DrawingPage) -> Result<String> {
    let mut xml = String::from("<draw:page");
    let properties = page.properties();
    push_attribute(&mut xml, "draw:name", properties.name());
    push_attribute(&mut xml, "draw:id", properties.draw_id());
    push_attribute(&mut xml, "xml:id", properties.xml_id());
    push_attribute(&mut xml, "draw:style-name", properties.style_name());
    push_attribute(
        &mut xml,
        "draw:master-page-name",
        properties.master_page_name(),
    );
    push_attribute(&mut xml, "draw:nav-order", properties.navigation_order());
    push_attribute(
        &mut xml,
        "presentation:presentation-page-layout-name",
        properties.presentation_layout_name(),
    );
    push_attribute(
        &mut xml,
        "presentation:use-header-name",
        properties.header_name(),
    );
    push_attribute(
        &mut xml,
        "presentation:use-footer-name",
        properties.footer_name(),
    );
    push_attribute(
        &mut xml,
        "presentation:use-date-time-name",
        properties.date_time_name(),
    );
    xml.push('>');
    if !page.layers().is_empty() {
        xml.push_str("<draw:layer-set>");
        for layer in page.layers() {
            xml.push_str("<draw:layer");
            push_attribute(&mut xml, "draw:name", Some(layer.name()));
            if let Some(protected) = layer.protected() {
                push_attribute(
                    &mut xml,
                    "draw:protected",
                    Some(if protected { "true" } else { "false" }),
                );
            }
            let display = layer.display().map(|display| match display {
                DrawingLayerDisplay::Always => "always",
                DrawingLayerDisplay::Screen => "screen",
                DrawingLayerDisplay::Printer => "printer",
                DrawingLayerDisplay::None => "none",
            });
            push_attribute(&mut xml, "draw:display", display);
            xml.push_str("/>");
        }
        xml.push_str("</draw:layer-set>");
    }
    for (index, shape) in page.shapes().iter().enumerate() {
        xml.push_str(&PresentationBuilder::generate_shape_xml(shape, index)?);
    }
    xml.push_str("</draw:page>");
    Ok(xml)
}

fn generate_meta_xml(metadata: &Metadata) -> String {
    let mut fields = String::new();
    for (name, value) in [
        ("dc:title", metadata.title.as_deref()),
        ("dc:creator", metadata.author.as_deref()),
        ("dc:subject", metadata.subject.as_deref()),
        ("dc:description", metadata.description.as_deref()),
        ("meta:keyword", metadata.keywords.as_deref()),
    ] {
        if let Some(value) = value {
            fields.push('<');
            fields.push_str(name);
            fields.push('>');
            fields.push_str(&escape_xml(value));
            fields.push_str("</");
            fields.push_str(name);
            fields.push('>');
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi</meta:generator>{fields}</office:meta></office:document-meta>"#
    )
}

fn drawing_page_ranges(content: &str) -> Result<(Range<usize>, Vec<Range<usize>>)> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut drawing_depth = None;
    let mut inner_start = None;
    let mut inner_end = None;
    let mut page_start = None;
    let mut page_depth = None;
    let mut pages = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG content.xml while staging: {error}"))
        })?;
        let office_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE);
        let draw_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DRAW_NAMESPACE);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let office_drawing =
                    office_namespace && element.local_name().as_ref() == b"drawing";
                let draw_page = draw_namespace && element.local_name().as_ref() == b"page";
                if office_drawing {
                    drawing_depth = Some(depth);
                    inner_start = Some(event_end);
                } else if draw_page && drawing_depth.is_some_and(|drawing| depth == drawing + 1) {
                    page_start = Some(event_start);
                    page_depth = Some(depth);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML nesting overflow".to_string())
                })?;
            },
            Event::Empty(element) => {
                if draw_namespace
                    && element.local_name().as_ref() == b"page"
                    && drawing_depth.is_some_and(|drawing| depth == drawing + 1)
                {
                    pages.push(event_start..event_end);
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML nesting underflow".to_string())
                })?;
                if page_depth == Some(depth)
                    && draw_namespace
                    && element.local_name().as_ref() == b"page"
                {
                    pages.push(page_start.take().expect("page start recorded")..event_end);
                    page_depth = None;
                }
                if drawing_depth == Some(depth)
                    && office_namespace
                    && element.local_name().as_ref() == b"drawing"
                {
                    inner_end = Some(event_start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    let start = inner_start.ok_or_else(|| {
        Error::InvalidFormat("ODG content.xml has no office:drawing body".to_string())
    })?;
    let end = inner_end.ok_or_else(|| {
        Error::InvalidFormat("ODG content.xml has no complete office:drawing body".to_string())
    })?;
    Ok((start..end, pages))
}

fn atomic_write(path: &Path, bytes: &[u8]) -> Result<()> {
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let stem = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("drawing");
    for attempt in 0..1000u32 {
        let temporary = parent.join(format!(".{stem}.litchi-{}-{attempt}.tmp", std::process::id()));
        match std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(mut file) => {
                let result = (|| -> std::io::Result<()> {
                    file.write_all(bytes)?;
                    file.sync_all()?;
                    std::fs::rename(&temporary, path)?;
                    Ok(())
                })();
                if result.is_err() {
                    let _ = std::fs::remove_file(&temporary);
                }
                result?;
                return Ok(());
            },
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
    Err(Error::InvalidFormat(
        "unable to allocate atomic ODG temporary file".to_string(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp::{DrawingAttribute, DrawingAttributeNamespace};

    fn page(name: &str) -> DrawingPageProperties {
        let mut properties = DrawingPageProperties::new();
        properties.set_name(Some(name));
        properties
    }

    fn drawing_with_shape() -> MutableDrawing {
        let mut drawing = MutableDrawing::new();
        drawing.add_page(page("Page 1")).unwrap();
        drawing
            .add_layer(0, DrawingLayer::new("layout"))
            .unwrap();
        let mut shape = Shape::new();
        shape.name = Some("Box".to_string());
        shape.text = "hello".to_string();
        shape.layer = Some("layout".to_string());
        drawing.add_shape(0, shape).unwrap();
        drawing
    }

    #[test]
    fn creates_packaged_drawing_and_template() {
        let bytes = drawing_with_shape().to_bytes().unwrap();
        let parsed = DrawingDocument::from_bytes(bytes).unwrap();
        assert_eq!(parsed.pages().len(), 1);
        assert_eq!(parsed.pages()[0].text(), "hello");

        let mut template = DrawingBuilder::template();
        template.add_page(page("Template Page")).unwrap();
        let parsed = DrawingDocument::from_bytes(template.to_bytes().unwrap()).unwrap();
        assert!(parsed.is_template());
    }

    #[test]
    fn unmodified_conversion_is_byte_exact() {
        let bytes = drawing_with_shape().to_bytes().unwrap();
        let mutable = MutableDrawing::from_bytes(bytes.clone()).unwrap();
        assert_eq!(mutable.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn page_layer_and_shape_crud_roundtrip() {
        let mut drawing = drawing_with_shape();
        drawing.add_page(page("Page 2")).unwrap();
        drawing.move_page(1, 0).unwrap();
        drawing.move_layer(1, 0, 0).unwrap();
        let shape = drawing.remove_shape(1, 0).unwrap();
        drawing.add_shape(1, shape).unwrap();
        let parsed = DrawingDocument::from_bytes(drawing.to_bytes().unwrap()).unwrap();
        assert_eq!(parsed.pages()[0].name(), Some("Page 2"));
        assert_eq!(parsed.pages()[1].layers()[0].name(), "layout");
        assert_eq!(parsed.pages()[1].shapes().len(), 1);
    }

    #[test]
    fn duplicate_and_referenced_layer_fail_without_mutation() {
        let mut drawing = drawing_with_shape();
        let before = drawing.to_bytes().unwrap();
        assert!(drawing.add_layer(0, DrawingLayer::new("layout")).is_err());
        assert_eq!(drawing.to_bytes().unwrap(), before);
        assert!(drawing.remove_layer(0, 0).is_err());
        assert_eq!(drawing.to_bytes().unwrap(), before);
    }

    #[test]
    fn navigation_ids_are_validated_atomically() {
        let mut drawing = drawing_with_shape();
        let mut shape = drawing.page(0).unwrap().shapes()[0].clone();
        shape.drawing_attributes.push(
            DrawingAttribute::new(DrawingAttributeNamespace::Drawing, "id", "shape1").unwrap(),
        );
        drawing.set_shape(0, 0, shape).unwrap();
        let before = drawing.to_bytes().unwrap();
        let mut invalid = drawing.page(0).unwrap().properties().clone();
        invalid.set_navigation_order(Some("missing"));
        assert!(drawing.set_page_properties(0, invalid).is_err());
        assert_eq!(drawing.to_bytes().unwrap(), before);
    }

    #[test]
    fn preserves_auxiliary_package_parts_after_mutation() {
        let original = drawing_with_shape().to_bytes().unwrap();
        let source = OwnedPackage::from_bytes(original).unwrap();
        let content = String::from_utf8(source.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
        let content = content.replace(
            "office:version=\"1.3\"",
            "xmlns:vendor=\"urn:vendor\" vendor:flag=\"keep\" office:version=\"1.3\"",
        );
        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_DRAWING).unwrap();
        writer.add_file(constants::ODF_CONTENT, content.as_bytes()).unwrap();
        writer
            .add_file(constants::ODF_STYLES, b"<vendor-styles/>")
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/p.bin", b"media", "application/octet-stream")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let mut drawing = MutableDrawing::from_bytes(bytes).unwrap();
        drawing.add_page(page("Added")).unwrap();
        let output = drawing.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(output).unwrap();
        assert_eq!(package.get_file("Pictures/p.bin").unwrap(), b"media");
        assert_eq!(package.get_file(constants::ODF_STYLES).unwrap(), b"<vendor-styles/>");
        let content = String::from_utf8(package.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
        assert!(content.contains("vendor:flag=\"keep\""));
    }

    #[test]
    fn oversized_metadata_rolls_back() {
        let mut drawing = drawing_with_shape();
        let before = drawing.to_bytes().unwrap();
        let mut metadata = Metadata::default();
        metadata.title = Some("x".repeat(MAX_STRING_BYTES + 1));
        assert!(drawing.set_metadata(metadata).is_err());
        assert_eq!(drawing.to_bytes().unwrap(), before);
    }
}
