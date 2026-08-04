//! Namespace-aware, inert access to standalone OpenDocument images.

use crate::{OpenDocumentFamily, OpenDocumentPackage};
use base64::Engine as _;
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 128 * 1_048_576;

/// A recognized element in the standard ODF frame vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ImageElementKind {
    Frame,
    Image,
    BinaryData,
    TextBox,
    Object,
    ObjectOle,
    Applet,
    FloatingFrame,
    Plugin,
    Table,
    EventListeners,
    EventListener,
    GluePoint,
    ImageMap,
    AreaRectangle,
    AreaCircle,
    AreaPolygon,
    Title,
    Description,
    ContourPolygon,
    ContourPath,
    Paragraph,
    Span,
    /// A future ODF element, embedded vocabulary, or vendor extension.
    Other,
}

/// One decoded attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageAttribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl ImageAttribute {
    /// Return the expanded namespace URI, or `None` for an unqualified attribute.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the decoded and normalized XML attribute value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within an image-frame element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageContent {
    /// Decoded character content. Unknown declared entities remain in `&name;` notation.
    Text(String),
    /// A child element.
    Element(ImageElement),
}

/// A complete element in the image document's required `draw:frame` subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageElement {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<ImageAttribute>,
    content: Vec<ImageContent>,
}

impl ImageElement {
    /// Return the element's expanded namespace URI.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Classify common frame elements without discarding unknown vocabularies.
    pub fn kind(&self) -> ImageElementKind {
        match (self.namespace_uri(), self.local_name.as_str()) {
            (Some(DRAW_NAMESPACE), "frame") => ImageElementKind::Frame,
            (Some(DRAW_NAMESPACE), "image") => ImageElementKind::Image,
            (Some(OFFICE_NAMESPACE), "binary-data") => ImageElementKind::BinaryData,
            (Some(DRAW_NAMESPACE), "text-box") => ImageElementKind::TextBox,
            (Some(DRAW_NAMESPACE), "object") => ImageElementKind::Object,
            (Some(DRAW_NAMESPACE), "object-ole") => ImageElementKind::ObjectOle,
            (Some(DRAW_NAMESPACE), "applet") => ImageElementKind::Applet,
            (Some(DRAW_NAMESPACE), "floating-frame") => ImageElementKind::FloatingFrame,
            (Some(DRAW_NAMESPACE), "plugin") => ImageElementKind::Plugin,
            (Some(TABLE_NAMESPACE), "table") => ImageElementKind::Table,
            (Some(OFFICE_NAMESPACE), "event-listeners") => ImageElementKind::EventListeners,
            (Some(SCRIPT_NAMESPACE), "event-listener") => ImageElementKind::EventListener,
            (Some(DRAW_NAMESPACE), "glue-point") => ImageElementKind::GluePoint,
            (Some(DRAW_NAMESPACE), "image-map") => ImageElementKind::ImageMap,
            (Some(DRAW_NAMESPACE), "area-rectangle") => ImageElementKind::AreaRectangle,
            (Some(DRAW_NAMESPACE), "area-circle") => ImageElementKind::AreaCircle,
            (Some(DRAW_NAMESPACE), "area-polygon") => ImageElementKind::AreaPolygon,
            (Some(SVG_NAMESPACE), "title") => ImageElementKind::Title,
            (Some(SVG_NAMESPACE), "desc") => ImageElementKind::Description,
            (Some(DRAW_NAMESPACE), "contour-polygon") => ImageElementKind::ContourPolygon,
            (Some(DRAW_NAMESPACE), "contour-path") => ImageElementKind::ContourPath,
            (Some(TEXT_NAMESPACE), "p") => ImageElementKind::Paragraph,
            (Some(TEXT_NAMESPACE), "span") => ImageElementKind::Span,
            _ => ImageElementKind::Other,
        }
    }

    /// Return all attributes in document order.
    pub fn attributes(&self) -> &[ImageAttribute] {
        &self.attributes
    }

    /// Find an attribute by expanded name.
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(ImageAttribute::value)
    }

    /// Return ordered mixed content.
    pub fn content(&self) -> &[ImageContent] {
        &self.content
    }

    /// Iterate direct child elements.
    pub fn children(&self) -> impl Iterator<Item = &ImageElement> {
        self.content.iter().filter_map(|content| match content {
            ImageContent::Element(element) => Some(element),
            ImageContent::Text(_) => None,
        })
    }

    /// Iterate direct children with the requested standard kind.
    pub fn children_of_kind(&self, kind: ImageElementKind) -> impl Iterator<Item = &ImageElement> {
        self.children().filter(move |child| child.kind() == kind)
    }

    /// Compose all descendant character content in exact element/text order.
    pub fn all_text(&self) -> String {
        fn append(element: &ImageElement, output: &mut String) {
            for content in &element.content {
                match content {
                    ImageContent::Text(text) => output.push_str(text),
                    ImageContent::Element(child) => append(child, output),
                }
            }
        }
        let mut output = String::new();
        append(self, &mut output);
        output
    }

    fn collect_kind<'a>(&'a self, kind: ImageElementKind, output: &mut Vec<&'a Self>) {
        if self.kind() == kind {
            output.push(self);
        }
        for child in self.children() {
            child.collect_kind(kind, output);
        }
    }
}

/// A validated OpenDocument Image (`.odi`) or image template (`.oti`).
///
/// The required frame and all embedded markup are inert. External links are
/// never fetched, and unmodified saves return the original package exactly.
pub struct ImageDocument {
    package: OpenDocumentPackage,
    frame: ImageElement,
}

impl ImageDocument {
    /// Open an image document from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read an image document from a stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate an image document from owned package bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OpenDocumentPackage::from_bytes(bytes)?;
        if package.family() != OpenDocumentFamily::Image {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument image: MIME type is '{}'",
                package.mimetype()
            )));
        }
        let frame = parse_image_content(&package.content_xml()?)?;
        validate_image_payloads(&frame)?;
        Ok(Self { package, frame })
    }

    /// Whether this package uses the image-template MIME type.
    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    /// Return the required complete `draw:frame` subtree.
    pub fn frame(&self) -> &ImageElement {
        &self.frame
    }

    /// Return all `draw:image` elements in document order.
    pub fn images(&self) -> Vec<&ImageElement> {
        let mut images = Vec::new();
        self.frame
            .collect_kind(ImageElementKind::Image, &mut images);
        images
    }

    /// Return an image's XLink target, when it uses linked data.
    pub fn image_href<'a>(&self, image: &'a ImageElement) -> Option<&'a str> {
        (image.kind() == ImageElementKind::Image)
            .then(|| image.attribute(Some(XLINK_NAMESPACE), "href"))
            .flatten()
    }

    /// Read an embedded or package-local image without following external URLs.
    pub fn image_data(&self, image: &ImageElement) -> Result<Option<Vec<u8>>> {
        if image.kind() != ImageElementKind::Image {
            return Err(Error::InvalidFormat(
                "image data requires a draw:image element".to_string(),
            ));
        }
        if let Some(binary) = image.children_of_kind(ImageElementKind::BinaryData).next() {
            return decode_binary_data(binary).map(Some);
        }
        let Some(href) = image.attribute(Some(XLINK_NAMESPACE), "href") else {
            return Ok(None);
        };
        let Some(path) = package_path(href) else {
            return Ok(None);
        };
        if !self.package.has_file(path)? {
            return Ok(None);
        }
        self.package.get_file(path).map(Some)
    }

    /// Extract common package metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    /// Extract complete OpenDocument metadata.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.package.odf_metadata()
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact original package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    /// Consume this document and return the exact original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save without reconstructing XML or the ZIP package.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn parse_image_content(xml: &str) -> Result<ImageElement> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut body_depth = None;
    let mut image_seen = false;
    let mut image_depth = None;
    let mut frame_depth = None;
    let mut frame_complete = None;
    let mut stack = Vec::new();
    let mut node_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid image XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                validate_container_start(
                    namespace_uri.as_deref(),
                    &local,
                    depth,
                    &mut root_seen,
                    root_closed,
                    &mut body_seen,
                    &mut body_depth,
                    &mut image_seen,
                    &mut image_depth,
                    &mut frame_depth,
                    frame_complete.is_some(),
                )?;
                if frame_depth.is_some() {
                    let node =
                        make_element(&reader, namespace_uri, local, element, &mut node_count)?;
                    stack.push(node);
                    if stack.len() > MAX_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "image frame nesting exceeds {MAX_DEPTH} levels"
                        )));
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("image XML nesting overflow".to_string())
                })?;
            },
            Event::Empty(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "image content root cannot be empty".to_string(),
                    ));
                }
                if depth == 1
                    && namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                {
                    if body_seen {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                } else if depth == 2 && body_depth == Some(2) {
                    if namespace_uri.as_deref() != Some(OFFICE_NAMESPACE)
                        || local != "image"
                        || image_seen
                    {
                        return Err(Error::InvalidFormat(
                            "image body must contain exactly one office:image".to_string(),
                        ));
                    }
                    image_seen = true;
                } else if depth == 3 && image_depth == Some(3) {
                    if namespace_uri.as_deref() != Some(DRAW_NAMESPACE)
                        || local != "frame"
                        || frame_complete.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "office:image must contain exactly one draw:frame".to_string(),
                        ));
                    }
                    frame_complete = Some(make_element(
                        &reader,
                        namespace_uri,
                        local,
                        element,
                        &mut node_count,
                    )?);
                } else if frame_depth.is_some() {
                    let node =
                        make_element(&reader, namespace_uri, local, element, &mut node_count)?;
                    stack
                        .last_mut()
                        .expect("frame parent exists")
                        .content
                        .push(ImageContent::Element(node));
                }
            },
            Event::End(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected image XML closing tag".to_string())
                })?;
                if frame_depth.is_some() {
                    let node = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("image frame element stack underflow".to_string())
                    })?;
                    if stack.is_empty() {
                        frame_complete = Some(node);
                        frame_depth = None;
                    } else {
                        stack
                            .last_mut()
                            .expect("parent exists")
                            .content
                            .push(ImageContent::Element(node));
                    }
                }
                if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "image"
                    && depth == 2
                {
                    image_depth = None;
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                    && depth == 1
                {
                    body_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid image text: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid image CDATA: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                push_text(
                    stack.last_mut().expect("element exists"),
                    decode_reference(reference)?,
                    &mut text_bytes,
                )?;
            },
            Event::Text(ref text)
                if (depth == 0 || body_depth.is_some() || image_depth.is_some())
                    && !text.iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the image frame".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_)
                if depth == 0 || body_depth.is_some() || image_depth.is_some() =>
            {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the image frame".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || body_depth.is_some()
        || !image_seen
        || image_depth.is_some()
        || frame_depth.is_some()
        || !stack.is_empty()
    {
        return Err(Error::InvalidFormat(
            "incomplete standalone image structure".to_string(),
        ));
    }
    frame_complete
        .ok_or_else(|| Error::InvalidFormat("standalone image has no draw:frame".to_string()))
}

#[allow(clippy::too_many_arguments)]
fn validate_container_start(
    namespace_uri: Option<&str>,
    local: &str,
    depth: usize,
    root_seen: &mut bool,
    root_closed: bool,
    body_seen: &mut bool,
    body_depth: &mut Option<usize>,
    image_seen: &mut bool,
    image_depth: &mut Option<usize>,
    frame_depth: &mut Option<usize>,
    frame_complete: bool,
) -> Result<()> {
    if depth == 0 {
        if *root_seen
            || root_closed
            || namespace_uri != Some(OFFICE_NAMESPACE)
            || local != "document-content"
        {
            return Err(Error::InvalidFormat(
                "image content must have one office:document-content root".to_string(),
            ));
        }
        *root_seen = true;
    } else if depth == 1 && namespace_uri == Some(OFFICE_NAMESPACE) && local == "body" {
        if *body_seen || body_depth.is_some() {
            return Err(Error::InvalidFormat("duplicate office:body".to_string()));
        }
        *body_seen = true;
        *body_depth = Some(2);
    } else if depth == 2 && *body_depth == Some(2) {
        if namespace_uri != Some(OFFICE_NAMESPACE) || local != "image" || *image_seen {
            return Err(Error::InvalidFormat(
                "image body must contain exactly one office:image".to_string(),
            ));
        }
        *image_seen = true;
        *image_depth = Some(3);
    } else if depth == 3 && *image_depth == Some(3) {
        if namespace_uri != Some(DRAW_NAMESPACE) || local != "frame" || frame_complete {
            return Err(Error::InvalidFormat(
                "office:image must contain exactly one draw:frame".to_string(),
            ));
        }
        *frame_depth = Some(4);
    }
    Ok(())
}

fn make_element(
    reader: &NsReader<&[u8]>,
    resolved_namespace_uri: Option<String>,
    local_name: String,
    element: &BytesStart<'_>,
    node_count: &mut usize,
) -> Result<ImageElement> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("image node count overflow".to_string()))?;
    if *node_count > MAX_NODES {
        return Err(Error::InvalidFormat(format!(
            "image frame exceeds {MAX_NODES} elements"
        )));
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "image element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid image attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let local_name = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &ImageAttribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded image attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid image attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "image attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(ImageAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(ImageElement {
        namespace_uri: resolved_namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

fn push_text(element: &mut ImageElement, value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("image text size overflow".to_string()))?;
    if *total > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(
            "image frame exceeds 128 MiB of XML text".to_string(),
        ));
    }
    if let Some(ImageContent::Text(existing)) = element.content.last_mut() {
        existing.push_str(&value);
    } else {
        element.content.push(ImageContent::Text(value));
    }
    Ok(())
}

fn validate_image_payloads(frame: &ImageElement) -> Result<()> {
    let mut images = Vec::new();
    frame.collect_kind(ImageElementKind::Image, &mut images);
    for image in images {
        let href = image.attribute(Some(XLINK_NAMESPACE), "href");
        let link_type = image.attribute(Some(XLINK_NAMESPACE), "type");
        let binaries: Vec<_> = image
            .children_of_kind(ImageElementKind::BinaryData)
            .collect();
        if let Some(link_type) = link_type
            && link_type != "simple"
        {
            return Err(Error::InvalidFormat(
                "draw:image xlink:type must be 'simple'".to_string(),
            ));
        }
        if let Some(show) = image.attribute(Some(XLINK_NAMESPACE), "show")
            && show != "embed"
        {
            return Err(Error::InvalidFormat(
                "draw:image xlink:show must be 'embed'".to_string(),
            ));
        }
        if let Some(actuate) = image.attribute(Some(XLINK_NAMESPACE), "actuate")
            && actuate != "onLoad"
        {
            return Err(Error::InvalidFormat(
                "draw:image xlink:actuate must be 'onLoad'".to_string(),
            ));
        }
        match (href, link_type, binaries.as_slice()) {
            (Some(_), Some("simple"), []) => {},
            (None, None, [binary]) => {
                if binary.children().next().is_some() {
                    return Err(Error::InvalidFormat(
                        "office:binary-data cannot contain elements".to_string(),
                    ));
                }
                decode_binary_data(binary)?;
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "draw:image must contain either a simple XLink or one office:binary-data"
                        .to_string(),
                ));
            },
        }
    }
    Ok(())
}

fn decode_binary_data(binary: &ImageElement) -> Result<Vec<u8>> {
    let compact: String = binary
        .all_text()
        .chars()
        .filter(|character| !character.is_ascii_whitespace())
        .collect();
    base64::engine::general_purpose::STANDARD
        .decode(compact)
        .map_err(|error| Error::InvalidFormat(format!("invalid embedded image base64: {error}")))
}

fn package_path(href: &str) -> Option<&str> {
    let path = href.strip_prefix("./").unwrap_or(href);
    let first = path.split('/').next()?;
    if path.is_empty()
        || path.starts_with('/')
        || path.ends_with('/')
        || path.contains('\\')
        || path.contains('?')
        || path.contains('#')
        || first.contains(':')
        || path
            .split('/')
            .any(|component| component.is_empty() || component == "." || component == "..")
    {
        None
    } else {
        Some(path)
    }
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown image namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 image {kind}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid image character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid image entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Ok(format!("&{name};")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str, media: Option<(&str, &[u8])>) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        if let Some((path, bytes)) = media {
            writer.add_file(path, bytes).unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn image_xml() -> &'static str {
        r#"<?xml version="1.0"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:x="http://www.w3.org/1999/xlink"
 xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:v="urn:vendor:image">
 <o:automatic-styles/>
 <o:body><o:image><d:frame d:name="Canvas" s:x="1cm" s:y="2cm" s:width="10cm" s:height="8cm">
  <d:image x:type="simple" x:href="Pictures/pixel.png" x:show="embed" x:actuate="onLoad"><t:p>linked</t:p></d:image>
  <d:image><o:binary-data>
    AQIDBA==
  </o:binary-data></d:image>
  <d:image x:type="simple" x:href="https://example.invalid/image.png"/>
  <d:text-box><t:p>before<t:span>inside</t:span>after</t:p></d:text-box>
  <v:effect v:mode="inert">extension</v:effect>
  <s:title>A &amp; B</s:title><s:desc><![CDATA[exact <description>]]></s:desc>
 </d:frame></o:image></o:body>
</o:document-content>"#
    }

    #[test]
    fn parses_complete_frame_and_reads_only_local_or_embedded_images() {
        let bytes = package(
            constants::ODF_IMAGE,
            image_xml(),
            Some(("Pictures/pixel.png", b"PNG")),
        );
        let document = ImageDocument::from_bytes(bytes.clone()).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.frame().kind(), ImageElementKind::Frame);
        assert_eq!(
            document.frame().attribute(Some(SVG_NAMESPACE), "width"),
            Some("10cm")
        );
        let images = document.images();
        assert_eq!(images.len(), 3);
        assert_eq!(document.image_href(images[0]), Some("Pictures/pixel.png"));
        assert_eq!(document.image_data(images[0]).unwrap().unwrap(), b"PNG");
        assert_eq!(
            document.image_data(images[1]).unwrap().unwrap(),
            [1, 2, 3, 4]
        );
        assert!(document.image_data(images[2]).unwrap().is_none());
        assert!(document.frame().all_text().contains("beforeinsideafter"));
        assert!(document.frame().all_text().contains("A & B"));
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_templates_readers_empty_frames_and_exact_mixed_content() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:image><d:frame/></o:image></o:body></o:document-content>"#;
        let bytes = package(constants::ODF_IMAGE_TEMPLATE, xml, None);
        let document = ImageDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert!(document.is_template());
        assert!(document.frame().content().is_empty());
        assert_eq!(document.into_bytes(), bytes);

        let document =
            ImageDocument::from_bytes(package(constants::ODF_IMAGE, image_xml(), None)).unwrap();
        let text_box = document
            .frame()
            .children_of_kind(ImageElementKind::TextBox)
            .next()
            .unwrap();
        let paragraph = text_box.children().next().unwrap();
        assert!(
            matches!(paragraph.content()[0], ImageContent::Text(ref value) if value == "before")
        );
        assert!(matches!(paragraph.content()[1], ImageContent::Element(_)));
        assert!(
            matches!(paragraph.content()[2], ImageContent::Text(ref value) if value == "after")
        );
    }

    #[test]
    fn rejects_other_families_and_invalid_container_hierarchy() {
        assert!(
            ImageDocument::from_bytes(package(constants::ODF_DRAWING, image_xml(), None)).is_err()
        );
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:drawing><d:frame/></o:drawing></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:image/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:image><d:frame/><d:frame/></o:image></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:image><d:image/></o:image></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:image><d:frame/></o:image></o:body>"#,
        ] {
            assert!(
                ImageDocument::from_bytes(package(constants::ODF_IMAGE, xml, None)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_invalid_link_and_binary_payload_choices() {
        for payload in [
            r#"<d:image/>"#,
            r#"<d:image x:type="extended" x:href="Pictures/a.png"/>"#,
            r#"<d:image x:type="simple"/>"#,
            r#"<d:image x:type="simple" x:href="Pictures/a.png"><o:binary-data>AQ==</o:binary-data></d:image>"#,
            r#"<d:image><o:binary-data>***</o:binary-data></d:image>"#,
            r#"<d:image><o:binary-data>AQ==</o:binary-data><o:binary-data>Ag==</o:binary-data></d:image>"#,
        ] {
            let xml = format!(
                r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DRAW_NAMESPACE}" xmlns:x="{XLINK_NAMESPACE}"><o:body><o:image><d:frame>{payload}</d:frame></o:image></o:body></o:document-content>"#
            );
            assert!(
                ImageDocument::from_bytes(package(constants::ODF_IMAGE, &xml, None)).is_err(),
                "accepted {payload}"
            );
        }
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_and_excessive_nesting() {
        let duplicate = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DRAW_NAMESPACE}" xmlns:a="urn:test" xmlns:b="urn:test"><o:body><o:image><d:frame a:x="one" b:x="two"/></o:image></o:body></o:document-content>"#
        );
        assert!(
            ImageDocument::from_bytes(package(constants::ODF_IMAGE, &duplicate, None)).is_err()
        );

        let nested = "<v:x>".repeat(MAX_DEPTH) + &"</v:x>".repeat(MAX_DEPTH);
        let deep = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DRAW_NAMESPACE}" xmlns:v="urn:vendor"><o:body><o:image><d:frame>{nested}</d:frame></o:image></o:body></o:document-content>"#
        );
        assert!(ImageDocument::from_bytes(package(constants::ODF_IMAGE, &deep, None)).is_err());
    }
}
