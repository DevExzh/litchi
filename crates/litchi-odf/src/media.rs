//! Semantic image discovery shared by packaged and flat OpenDocument families.

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use crate::elements::xml::namespaced_attribute;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";

const MAX_IMAGE_DEPTH: usize = 4_096;
const MAX_IMAGES: usize = 100_000;
const MAX_INLINE_ENCODED_BYTES: usize = 24 * 1024 * 1024;
const MAX_INLINE_IMAGE_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_INLINE_IMAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_ACCESSIBILITY_TEXT_BYTES: usize = 64 * 1024;
const MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES: usize = 8 * 1024 * 1024;

/// XML part containing an image occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfImagePart {
    Content,
    Styles,
    FlatDocument,
}

/// The inert source of an OpenDocument image.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfImageSource {
    /// Base64 data stored in an `office:binary-data` child.
    Inline {
        bytes: Vec<u8>,
        /// An href present on the parent is ignored by ODF when inline data exists.
        ignored_href: Option<String>,
    },
    /// A verified file in the same OpenDocument package.
    PackagePart {
        href: String,
        path: String,
        manifest_media_type: Option<String>,
    },
    /// A safe package path which is referenced but absent from the archive.
    MissingPackagePart {
        href: String,
        resolved_path: String,
    },
    /// An inert external, filesystem, fragment, query-bearing, or flat-document link.
    Linked { href: String },
    /// A malformed producer omitted both href and inline data.
    Missing,
}

/// Drawing-frame context for an image or embedded-object occurrence.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfImageFrame {
    pub name: Option<String>,
    pub xml_id: Option<String>,
    /// Short alternative title from a direct `svg:title` child.
    pub title: Option<String>,
    /// Prose alternative description from a direct `svg:desc` child.
    pub description: Option<String>,
    pub anchor_type: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub page_name: Option<String>,
    pub sheet_name: Option<String>,
    /// Whether this frame is a direct child of a spreadsheet `table:shapes` container.
    pub sheet_shape: bool,
}

/// One `draw:image` occurrence and its safely classified source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfImage {
    pub part: OdfImagePart,
    pub source: OdfImageSource,
    pub frame: Option<OdfImageFrame>,
    pub xml_id: Option<String>,
    pub filter_name: Option<String>,
    pub declared_media_type: Option<String>,
    pub link_type: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
    /// Zero-based position among alternative images in the same frame.
    pub alternative_index: usize,
}

impl OdfImage {
    /// Return inline bytes without copying, if this is an inline image.
    pub fn inline_bytes(&self) -> Option<&[u8]> {
        match &self.source {
            OdfImageSource::Inline { bytes, .. } => Some(bytes),
            _ => None,
        }
    }

    /// Return the resolved package path, if this references an existing package part.
    pub fn package_path(&self) -> Option<&str> {
        match &self.source {
            OdfImageSource::PackagePart { path, .. } => Some(path),
            _ => None,
        }
    }
}

#[derive(Clone, Copy)]
struct PackageLookup<'a> {
    has_file: &'a dyn Fn(&str) -> bool,
    media_type: &'a dyn Fn(&str) -> Option<String>,
}

#[derive(Clone)]
struct FrameState {
    depth: usize,
    frame: OdfImageFrame,
    image_count: usize,
    image_indices: Vec<usize>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum AccessibilityKind {
    Title,
    Description,
}

struct AccessibilityText {
    depth: usize,
    kind: AccessibilityKind,
    value: String,
}

struct NamedContext {
    depth: usize,
    name: Option<String>,
}

struct ImageBuilder {
    depth: usize,
    href: Option<String>,
    frame: Option<OdfImageFrame>,
    xml_id: Option<String>,
    filter_name: Option<String>,
    declared_media_type: Option<String>,
    link_type: Option<String>,
    show: Option<String>,
    actuate: Option<String>,
    alternative_index: usize,
    inline_present: bool,
    inline_depth: Option<usize>,
    inline_encoded: String,
}

pub(crate) fn scan_packaged_images(
    content_xml: &str,
    styles_xml: Option<&str>,
    has_file: impl Fn(&str) -> bool,
    media_type: impl Fn(&str) -> Option<String>,
) -> Result<Vec<OdfImage>> {
    let lookup = PackageLookup {
        has_file: &has_file,
        media_type: &media_type,
    };
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        content_xml,
        OdfImagePart::Content,
        Some(lookup),
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    if let Some(styles_xml) = styles_xml {
        scan_xml(
            styles_xml,
            OdfImagePart::Styles,
            Some(lookup),
            &mut images,
            &mut inline_bytes,
            &mut accessibility_bytes,
        )?;
    }
    Ok(images)
}

pub(crate) fn scan_flat_images(xml: &str) -> Result<Vec<OdfImage>> {
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        xml,
        OdfImagePart::FlatDocument,
        None,
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    Ok(images)
}

pub(crate) fn scan_content_images(xml: &str) -> Result<Vec<OdfImage>> {
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        xml,
        OdfImagePart::Content,
        None,
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    Ok(images)
}

fn scan_xml(
    xml: &str,
    part: OdfImagePart,
    package: Option<PackageLookup<'_>>,
    images: &mut Vec<OdfImage>,
    total_inline_bytes: &mut usize,
    total_accessibility_bytes: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut frames = Vec::<FrameState>::new();
    let mut pages = Vec::<NamedContext>::new();
    let mut sheets = Vec::<NamedContext>::new();
    let mut sheet_shapes = Vec::<usize>::new();
    let mut sheets_with_shapes = HashSet::<usize>::new();
    let mut active: Option<ImageBuilder> = None;
    let mut accessibility: Option<AccessibilityText> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF image XML: {error}")))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF image nesting overflow".to_string())
                })?;
                if depth > MAX_IMAGE_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF image nesting exceeds {MAX_IMAGE_DEPTH}"
                    )));
                }
                if active
                    .as_ref()
                    .and_then(|image| image.inline_depth)
                    .is_some()
                {
                    return Err(Error::InvalidFormat(
                        "office:binary-data must not contain elements".to_string(),
                    ));
                }
                if accessibility.is_some() {
                    return Err(Error::InvalidFormat(
                        "svg:title and svg:desc must not contain elements".to_string(),
                    ));
                }
                if frames
                    .last()
                    .is_some_and(|frame| depth == frame.depth + 1)
                    && let Some(kind) =
                        accessibility_kind(&namespace, element.local_name().as_ref())
                {
                    begin_accessibility(
                        frames.last_mut().expect("direct frame child"),
                        kind,
                        depth,
                        &mut accessibility,
                    )?;
                } else if bound_to(&namespace, DRAW_NAMESPACE)
                    && element.local_name().as_ref() == b"page"
                {
                    pages.push(NamedContext {
                        depth,
                        name: attribute(&reader, &element, DRAW_NAMESPACE, b"name")?,
                    });
                } else if bound_to(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"table"
                {
                    sheets.push(NamedContext {
                        depth,
                        name: attribute(&reader, &element, TABLE_NAMESPACE, b"name")?,
                    });
                } else if element.local_name().as_ref() == b"shapes"
                    && sheets.last().is_some_and(|sheet| depth == sheet.depth + 1)
                    && !bound_to(&namespace, TABLE_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "spoofed table:shapes namespace".to_string(),
                    ));
                } else if bound_to(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"shapes"
                    && sheets.last().is_some_and(|sheet| depth == sheet.depth + 1)
                {
                    let sheet_depth = sheets.last().expect("checked sheet").depth;
                    if !sheets_with_shapes.insert(sheet_depth) {
                        return Err(Error::InvalidFormat(
                            "a table must not contain multiple table:shapes elements".to_string(),
                        ));
                    }
                    sheet_shapes.push(depth);
                } else if element.local_name().as_ref() == b"frame"
                    && sheet_shapes.last().is_some_and(|shape_depth| depth == *shape_depth + 1)
                    && !bound_to(&namespace, DRAW_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "spoofed draw:frame namespace in table:shapes".to_string(),
                    ));
                } else if bound_to(&namespace, DRAW_NAMESPACE)
                    && element.local_name().as_ref() == b"frame"
                {
                    frames.push(FrameState {
                        depth,
                        frame: parse_frame(
                            &reader,
                            &element,
                            &pages,
                            &sheets,
                            sheet_shapes.last().is_some_and(|shape_depth| depth == *shape_depth + 1),
                        )?,
                        image_count: 0,
                        image_indices: Vec::new(),
                    });
                } else if element.local_name().as_ref() == b"image"
                    && frames.last().is_some_and(|frame| {
                        frame.frame.sheet_shape && depth == frame.depth + 1
                    })
                    && !bound_to(&namespace, DRAW_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "spoofed draw:image namespace in sheet frame".to_string(),
                    ));
                } else if bound_to(&namespace, DRAW_NAMESPACE)
                    && element.local_name().as_ref() == b"image"
                {
                    if active.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested draw:image elements are not allowed".to_string(),
                        ));
                    }
                    ensure_image_capacity(images.len())?;
                    active = Some(start_image(
                        &reader,
                        &element,
                        depth,
                        images.len(),
                        frames.last_mut(),
                    )?);
                } else if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"binary-data"
                    && let Some(image) = active.as_mut()
                {
                    if image.inline_present {
                        return Err(Error::InvalidFormat(
                            "draw:image contains multiple office:binary-data elements".to_string(),
                        ));
                    }
                    image.inline_present = true;
                    image.inline_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                if active
                    .as_ref()
                    .and_then(|image| image.inline_depth)
                    .is_some()
                {
                    return Err(Error::InvalidFormat(
                        "office:binary-data must not contain elements".to_string(),
                    ));
                }
                if accessibility.is_some() {
                    return Err(Error::InvalidFormat(
                        "svg:title and svg:desc must not contain elements".to_string(),
                    ));
                }
                if element.local_name().as_ref() == b"shapes"
                    && sheets.last().is_some_and(|sheet| depth == sheet.depth)
                    && !bound_to(&namespace, TABLE_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "spoofed table:shapes namespace".to_string(),
                    ));
                } else if bound_to(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"shapes"
                    && sheets.last().is_some_and(|sheet| depth == sheet.depth)
                {
                    let sheet_depth = sheets.last().expect("checked sheet").depth;
                    if !sheets_with_shapes.insert(sheet_depth) {
                        return Err(Error::InvalidFormat(
                            "a table must not contain multiple table:shapes elements".to_string(),
                        ));
                    }
                } else if frames.last().is_some_and(|frame| depth == frame.depth)
                    && let Some(kind) =
                        accessibility_kind(&namespace, element.local_name().as_ref())
                {
                    set_empty_accessibility(
                        frames.last_mut().expect("direct frame child"),
                        kind,
                    )?;
                } else if bound_to(&namespace, DRAW_NAMESPACE)
                    && element.local_name().as_ref() == b"image"
                {
                    if active.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested draw:image elements are not allowed".to_string(),
                        ));
                    }
                    ensure_image_capacity(images.len())?;
                    let image = start_image(
                        &reader,
                        &element,
                        depth + 1,
                        images.len(),
                        frames.last_mut(),
                    )?;
                    images.push(finish_image(image, part, package, total_inline_bytes)?);
                } else if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"binary-data"
                    && let Some(image) = active.as_mut()
                {
                    if image.inline_present {
                        return Err(Error::InvalidFormat(
                            "draw:image contains multiple office:binary-data elements".to_string(),
                        ));
                    }
                    image.inline_present = true;
                }
            },
            Event::Text(value) if active.as_ref().and_then(|image| image.inline_depth).is_some() => {
                let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid inline image text: {error}"))
                })?;
                append_inline(active.as_mut().expect("active inline image"), &value)?;
            },
            Event::CData(value) if active.as_ref().and_then(|image| image.inline_depth).is_some() => {
                let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid inline image CDATA: {error}"))
                })?;
                append_inline(active.as_mut().expect("active inline image"), &value)?;
            },
            Event::Text(value) if accessibility.is_some() => {
                let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid image accessibility text: {error}"))
                })?;
                append_accessibility(
                    accessibility.as_mut().expect("active accessibility text"),
                    &value,
                    total_accessibility_bytes,
                )?;
            },
            Event::CData(value) if accessibility.is_some() => {
                let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid image accessibility CDATA: {error}"))
                })?;
                append_accessibility(
                    accessibility.as_mut().expect("active accessibility text"),
                    &value,
                    total_accessibility_bytes,
                )?;
            },
            Event::GeneralRef(_) if active.as_ref().and_then(|image| image.inline_depth).is_some() => {
                return Err(Error::InvalidFormat(
                    "XML references are not allowed in office:binary-data".to_string(),
                ));
            },
            Event::GeneralRef(value) if accessibility.is_some() => {
                let value = resolve_accessibility_reference(&value)?;
                append_accessibility(
                    accessibility.as_mut().expect("active accessibility text"),
                    &value,
                    total_accessibility_bytes,
                )?;
            },
            Event::End(element) => {
                if accessibility.as_ref().map(|text| text.depth) == Some(depth) {
                    let text = accessibility.take().expect("active accessibility text");
                    if accessibility_kind(&namespace, element.local_name().as_ref())
                        != Some(text.kind)
                    {
                        return Err(Error::InvalidFormat(
                            "malformed image accessibility element".to_string(),
                        ));
                    }
                    finish_accessibility(
                        frames.last_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "image accessibility text has no enclosing frame".to_string(),
                            )
                        })?,
                        text,
                    )?;
                } else if let Some(image) = active.as_mut()
                    && image.inline_depth == Some(depth)
                {
                    if !bound_to(&namespace, OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"binary-data"
                    {
                        return Err(Error::InvalidFormat(
                            "malformed office:binary-data element".to_string(),
                        ));
                    }
                    image.inline_depth = None;
                } else if active.as_ref().map(|image| image.depth) == Some(depth) {
                    if !bound_to(&namespace, DRAW_NAMESPACE)
                        || element.local_name().as_ref() != b"image"
                    {
                        return Err(Error::InvalidFormat(
                            "malformed draw:image element".to_string(),
                        ));
                    }
                    let image = active.take().expect("active image");
                    images.push(finish_image(image, part, package, total_inline_bytes)?);
                }

                if frames.last().map(|frame| frame.depth) == Some(depth) {
                    let frame = frames.pop().expect("closing frame");
                    for image_index in frame.image_indices {
                        let image = images.get_mut(image_index).ok_or_else(|| {
                            Error::InvalidFormat(
                                "image frame occurrence index is invalid".to_string(),
                            )
                        })?;
                        image.frame = Some(frame.frame.clone());
                    }
                }
                if pages.last().map(|page| page.depth) == Some(depth) {
                    pages.pop();
                }
                if sheets.last().map(|sheet| sheet.depth) == Some(depth) {
                    sheets_with_shapes.remove(&depth);
                    sheets.pop();
                }
                if sheet_shapes.last().copied() == Some(depth) {
                    sheet_shapes.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected ODF image closing tag".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs are not allowed while scanning ODF images".to_string(),
                ));
            },
            Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not allowed while scanning ODF images"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if depth != 0
        || active.is_some()
        || accessibility.is_some()
        || !frames.is_empty()
        || !pages.is_empty()
        || !sheets.is_empty()
        || !sheet_shapes.is_empty()
    {
        return Err(Error::InvalidFormat(
            "unterminated ODF image XML".to_string(),
        ));
    }
    Ok(())
}

fn ensure_image_capacity(current: usize) -> Result<()> {
    if current >= MAX_IMAGES {
        return Err(Error::InvalidFormat(format!(
            "ODF image count exceeds {MAX_IMAGES}"
        )));
    }
    Ok(())
}

fn parse_frame(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    pages: &[NamedContext],
    sheets: &[NamedContext],
    sheet_shape: bool,
) -> Result<OdfImageFrame> {
    Ok(OdfImageFrame {
        name: attribute(reader, element, DRAW_NAMESPACE, b"name")?,
        xml_id: attribute(reader, element, XML_NAMESPACE, b"id")?,
        title: None,
        description: None,
        anchor_type: attribute(reader, element, TEXT_NAMESPACE, b"anchor-type")?,
        x: attribute(reader, element, SVG_NAMESPACE, b"x")?,
        y: attribute(reader, element, SVG_NAMESPACE, b"y")?,
        width: attribute(reader, element, SVG_NAMESPACE, b"width")?,
        height: attribute(reader, element, SVG_NAMESPACE, b"height")?,
        page_name: pages.last().and_then(|page| page.name.clone()),
        sheet_name: sheets.last().and_then(|sheet| sheet.name.clone()),
        sheet_shape,
    })
}

fn start_image(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    image_index: usize,
    frame: Option<&mut FrameState>,
) -> Result<ImageBuilder> {
    let (frame_context, alternative_index) = match frame {
        Some(frame) => {
            let alternative_index = frame.image_count;
            frame.image_count = frame.image_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("image alternative count overflow".to_string())
            })?;
            frame.image_indices.push(image_index);
            (Some(frame.frame.clone()), alternative_index)
        },
        None => (None, 0),
    };
    Ok(ImageBuilder {
        depth,
        href: attribute(reader, element, XLINK_NAMESPACE, b"href")?,
        frame: frame_context,
        xml_id: attribute(reader, element, XML_NAMESPACE, b"id")?,
        filter_name: attribute(reader, element, DRAW_NAMESPACE, b"filter-name")?,
        declared_media_type: attribute(reader, element, DRAW_NAMESPACE, b"mime-type")?,
        link_type: attribute(reader, element, XLINK_NAMESPACE, b"type")?,
        show: attribute(reader, element, XLINK_NAMESPACE, b"show")?,
        actuate: attribute(reader, element, XLINK_NAMESPACE, b"actuate")?,
        alternative_index,
        inline_present: false,
        inline_depth: None,
        inline_encoded: String::new(),
    })
}

fn accessibility_kind(
    namespace: &ResolveResult<'_>,
    local_name: &[u8],
) -> Option<AccessibilityKind> {
    if !bound_to(namespace, SVG_NAMESPACE) {
        return None;
    }
    match local_name {
        b"title" => Some(AccessibilityKind::Title),
        b"desc" => Some(AccessibilityKind::Description),
        _ => None,
    }
}

fn begin_accessibility(
    frame: &mut FrameState,
    kind: AccessibilityKind,
    depth: usize,
    active: &mut Option<AccessibilityText>,
) -> Result<()> {
    ensure_accessibility_absent(frame, kind)?;
    *active = Some(AccessibilityText {
        depth,
        kind,
        value: String::new(),
    });
    Ok(())
}

fn set_empty_accessibility(frame: &mut FrameState, kind: AccessibilityKind) -> Result<()> {
    ensure_accessibility_absent(frame, kind)?;
    match kind {
        AccessibilityKind::Title => frame.frame.title = Some(String::new()),
        AccessibilityKind::Description => frame.frame.description = Some(String::new()),
    }
    Ok(())
}

fn finish_accessibility(frame: &mut FrameState, text: AccessibilityText) -> Result<()> {
    ensure_accessibility_absent(frame, text.kind)?;
    match text.kind {
        AccessibilityKind::Title => frame.frame.title = Some(text.value),
        AccessibilityKind::Description => frame.frame.description = Some(text.value),
    }
    Ok(())
}

fn ensure_accessibility_absent(frame: &FrameState, kind: AccessibilityKind) -> Result<()> {
    let duplicate = match kind {
        AccessibilityKind::Title => frame.frame.title.is_some(),
        AccessibilityKind::Description => frame.frame.description.is_some(),
    };
    if duplicate {
        let element = match kind {
            AccessibilityKind::Title => "svg:title",
            AccessibilityKind::Description => "svg:desc",
        };
        return Err(Error::InvalidFormat(format!(
            "draw:frame contains multiple {element} elements"
        )));
    }
    Ok(())
}

fn append_accessibility(
    active: &mut AccessibilityText,
    value: &str,
    total_accessibility_bytes: &mut usize,
) -> Result<()> {
    let field_len = active
        .value
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("image accessibility text size overflow".to_string()))?;
    if field_len > MAX_ACCESSIBILITY_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "image accessibility text exceeds {MAX_ACCESSIBILITY_TEXT_BYTES} bytes"
        )));
    }
    *total_accessibility_bytes = total_accessibility_bytes
        .checked_add(value.len())
        .ok_or_else(|| {
            Error::InvalidFormat("total image accessibility text size overflow".to_string())
        })?;
    if *total_accessibility_bytes > MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "total image accessibility text exceeds {MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES} bytes"
        )));
    }
    active.value.push_str(value);
    Ok(())
}

fn resolve_accessibility_reference(value: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = value.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid image accessibility character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let entity_name: &[u8] = value.as_ref();
    let character = match entity_name {
        b"amp" => '&',
        b"lt" => '<',
        b"gt" => '>',
        b"apos" => '\'',
        b"quot" => '"',
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported entity in image accessibility text".to_string(),
            ));
        },
    };
    Ok(character.to_string())
}

fn append_inline(image: &mut ImageBuilder, value: &str) -> Result<()> {
    let new_len = image
        .inline_encoded
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("inline image size overflow".to_string()))?;
    if new_len > MAX_INLINE_ENCODED_BYTES {
        return Err(Error::InvalidFormat(format!(
            "inline image encoding exceeds {MAX_INLINE_ENCODED_BYTES} bytes"
        )));
    }
    image.inline_encoded.push_str(value);
    Ok(())
}

fn finish_image(
    image: ImageBuilder,
    part: OdfImagePart,
    package: Option<PackageLookup<'_>>,
    total_inline_bytes: &mut usize,
) -> Result<OdfImage> {
    let source = if image.inline_present {
        let mut compact = Vec::with_capacity(image.inline_encoded.len());
        for byte in image.inline_encoded.bytes() {
            if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
                continue;
            }
            if !byte.is_ascii() {
                return Err(Error::InvalidFormat(
                    "non-ASCII data in office:binary-data".to_string(),
                ));
            }
            compact.push(byte);
        }
        let bytes = BASE64_STANDARD.decode(&compact).map_err(|error| {
            Error::InvalidFormat(format!("invalid office:binary-data base64: {error}"))
        })?;
        if bytes.len() > MAX_INLINE_IMAGE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "inline image exceeds {MAX_INLINE_IMAGE_BYTES} decoded bytes"
            )));
        }
        *total_inline_bytes = total_inline_bytes.checked_add(bytes.len()).ok_or_else(|| {
            Error::InvalidFormat("total inline image size overflow".to_string())
        })?;
        if *total_inline_bytes > MAX_TOTAL_INLINE_IMAGE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "total inline image data exceeds {MAX_TOTAL_INLINE_IMAGE_BYTES} bytes"
            )));
        }
        OdfImageSource::Inline {
            bytes,
            ignored_href: image.href.clone(),
        }
    } else if let Some(href) = image.href.clone().filter(|href| !href.is_empty()) {
        match package {
            None => OdfImageSource::Linked { href },
            Some(_) if is_linked_href(&href) => OdfImageSource::Linked { href },
            Some(package) => {
                let path = resolve_package_path(&href)?;
                if (package.has_file)(&path) {
                    OdfImageSource::PackagePart {
                        href,
                        manifest_media_type: (package.media_type)(&path),
                        path,
                    }
                } else {
                    OdfImageSource::MissingPackagePart {
                        href,
                        resolved_path: path,
                    }
                }
            },
        }
    } else {
        OdfImageSource::Missing
    };

    Ok(OdfImage {
        part,
        source,
        frame: image.frame,
        xml_id: image.xml_id,
        filter_name: image.filter_name,
        declared_media_type: image.declared_media_type,
        link_type: image.link_type,
        show: image.show,
        actuate: image.actuate,
        alternative_index: image.alternative_index,
    })
}

pub(crate) fn is_linked_href(href: &str) -> bool {
    if href.starts_with('/')
        || href.starts_with('\\')
        || href.starts_with('#')
        || href.contains('\\')
        || href.contains('?')
        || href.contains('#')
    {
        return true;
    }
    let Some(colon) = href.find(':') else {
        return false;
    };
    let scheme = &href[..colon];
    !scheme.is_empty()
        && scheme.as_bytes()[0].is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
}

pub(crate) fn resolve_package_path(href: &str) -> Result<String> {
    let decoded = percent_decode(href)?;
    if decoded.starts_with('/') || decoded.contains('\\') {
        return Err(Error::InvalidFormat(format!(
            "unsafe package image href '{href}'"
        )));
    }
    let mut segments = Vec::new();
    for segment in decoded.split('/') {
        if segment.is_empty() || segment == "." {
            continue;
        }
        if segment == ".." {
            if segments.pop().is_none() {
                return Err(Error::InvalidFormat(format!(
                    "package image href escapes the package root: '{href}'"
                )));
            }
            continue;
        }
        if segment
            .chars()
            .any(|character| character == '\0' || character.is_control())
        {
            return Err(Error::InvalidFormat(format!(
                "invalid character in package image href '{href}'"
            )));
        }
        segments.push(segment);
    }
    if segments.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "package image href has no file path: '{href}'"
        )));
    }
    let path = segments.join("/");
    if path == "mimetype" || path == "META-INF" || path.starts_with("META-INF/") {
        return Err(Error::InvalidFormat(format!(
            "package image href targets an administrative entry: '{href}'"
        )));
    }
    Ok(path)
}

fn percent_decode(value: &str) -> Result<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Err(Error::InvalidFormat(format!(
                "invalid percent escape in package image href '{value}'"
            )));
        }
        let high = hex_value(bytes[index + 1]);
        let low = hex_value(bytes[index + 2]);
        let (Some(high), Some(low)) = (high, low) else {
            return Err(Error::InvalidFormat(format!(
                "invalid percent escape in package image href '{value}'"
            )));
        };
        decoded.push((high << 4) | low);
        index += 3;
    }
    String::from_utf8(decoded)
        .map_err(|_| Error::InvalidFormat("package image href is not valid UTF-8".to_string()))
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> Result<Option<String>> {
    namespaced_attribute(
        reader,
        element,
        expected_namespace,
        expected_local_name,
        "ODF image",
    )
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(namespace)) if *namespace == expected)
}
