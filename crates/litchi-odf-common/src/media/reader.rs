//! Bounded XML scanning and package-path safety for image media.

use crate::drawing::{Frame, Part};
use crate::namespace::namespaced_attribute;
use crate::package::{PackageLookup, is_linked_href, resolve_package_path};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::{Image, Source};

macro_rules! required {
    ($value:expr, $message:literal) => {
        $value.ok_or_else(|| Error::InvalidFormat($message.to_string()))?
    };
}

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

#[derive(Clone)]
struct FrameState {
    depth: usize,
    frame: Frame,
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
    frame: Option<Frame>,
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

/// Scan package-backed `draw:image` occurrences in content and styles XML.
///
/// # Errors
///
/// Returns an error when either XML part is malformed, violates the accepted image grammar, or
/// exceeds a configured resource limit.
pub fn scan_package(
    content_xml: &str,
    styles_xml: Option<&str>,
    package: &impl PackageLookup,
) -> Result<Vec<Image>> {
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        content_xml,
        Part::Content,
        Some(package),
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    if let Some(styles_document) = styles_xml {
        scan_xml(
            styles_document,
            Part::Styles,
            Some(package),
            &mut images,
            &mut inline_bytes,
            &mut accessibility_bytes,
        )?;
    }
    Ok(images)
}

/// Scans a flat `OpenDocument` XML document for inert image occurrences.
///
/// # Errors
///
/// Returns an error when the XML is malformed, violates the accepted image grammar, or exceeds a
/// configured resource limit.
pub fn scan_flat(xml: &str) -> Result<Vec<Image>> {
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        xml,
        Part::FlatDocument,
        None,
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    Ok(images)
}

/// Scan one package XML part without resolving package-local references.
///
/// # Errors
///
/// Returns an error when the XML is malformed, violates the accepted image grammar, or exceeds a
/// configured resource limit.
pub fn scan_content(xml: &str) -> Result<Vec<Image>> {
    let mut images = Vec::new();
    let mut inline_bytes = 0usize;
    let mut accessibility_bytes = 0usize;
    scan_xml(
        xml,
        Part::Content,
        None,
        &mut images,
        &mut inline_bytes,
        &mut accessibility_bytes,
    )?;
    Ok(images)
}

fn scan_xml(
    xml: &str,
    part: Part,
    package: Option<&dyn PackageLookup>,
    images: &mut Vec<Image>,
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
                if frames.last().is_some_and(|frame| depth == frame.depth + 1)
                    && let Some(kind) =
                        accessibility_kind(&namespace, element.local_name().as_ref())
                {
                    begin_accessibility(
                        required!(frames.last_mut(), "accessibility element has no frame"),
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
                    let sheet_depth = required!(sheets.last(), "table:shapes has no sheet").depth;
                    if !sheets_with_shapes.insert(sheet_depth) {
                        return Err(Error::InvalidFormat(
                            "a table must not contain multiple table:shapes elements".to_string(),
                        ));
                    }
                    sheet_shapes.push(depth);
                } else if element.local_name().as_ref() == b"frame"
                    && sheet_shapes
                        .last()
                        .is_some_and(|shape_depth| depth == *shape_depth + 1)
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
                            sheet_shapes
                                .last()
                                .is_some_and(|shape_depth| depth == *shape_depth + 1),
                        )?,
                        image_count: 0,
                        image_indices: Vec::new(),
                    });
                } else if element.local_name().as_ref() == b"image"
                    && frames
                        .last()
                        .is_some_and(|frame| frame.frame.sheet_shape && depth == frame.depth + 1)
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
                    let sheet_depth = required!(sheets.last(), "table:shapes has no sheet").depth;
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
                        required!(frames.last_mut(), "accessibility element has no frame"),
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
            Event::Text(text)
                if active
                    .as_ref()
                    .and_then(|image| image.inline_depth)
                    .is_some() =>
            {
                let decoded_text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid inline image text: {error}"))
                })?;
                append_inline(
                    required!(active.as_mut(), "inline image has no active image"),
                    &decoded_text,
                )?;
            },
            Event::CData(cdata)
                if active
                    .as_ref()
                    .and_then(|image| image.inline_depth)
                    .is_some() =>
            {
                let decoded_text = cdata
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid inline image CDATA: {error}"))
                    })?;
                append_inline(
                    required!(active.as_mut(), "inline image has no active image"),
                    &decoded_text,
                )?;
            },
            Event::Text(text) if accessibility.is_some() => {
                let decoded_text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid image accessibility text: {error}"))
                })?;
                append_accessibility(
                    required!(accessibility.as_mut(), "accessibility text is missing"),
                    &decoded_text,
                    total_accessibility_bytes,
                )?;
            },
            Event::CData(cdata) if accessibility.is_some() => {
                let decoded_text = cdata
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid image accessibility CDATA: {error}"))
                    })?;
                append_accessibility(
                    required!(accessibility.as_mut(), "accessibility text is missing"),
                    &decoded_text,
                    total_accessibility_bytes,
                )?;
            },
            Event::GeneralRef(_)
                if active
                    .as_ref()
                    .and_then(|image| image.inline_depth)
                    .is_some() =>
            {
                return Err(Error::InvalidFormat(
                    "XML references are not allowed in office:binary-data".to_string(),
                ));
            },
            Event::GeneralRef(reference) if accessibility.is_some() => {
                let resolved_reference = resolve_accessibility_reference(&reference)?;
                append_accessibility(
                    required!(accessibility.as_mut(), "accessibility text is missing"),
                    &resolved_reference,
                    total_accessibility_bytes,
                )?;
            },
            Event::End(element) => {
                if accessibility.as_ref().map(|text| text.depth) == Some(depth) {
                    let text = required!(accessibility.take(), "accessibility text is missing");
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
                    let image = required!(active.take(), "embedded image is missing");
                    images.push(finish_image(image, part, package, total_inline_bytes)?);
                }

                if frames.last().map(|frame| frame.depth) == Some(depth) {
                    let frame = required!(frames.pop(), "image frame is missing");
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
                    "processing instructions are not allowed while scanning ODF images".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
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
) -> Result<Frame> {
    Ok(Frame {
        name: attribute(reader, element, DRAW_NAMESPACE, b"name")?,
        xml_id: attribute(reader, element, XML_NAMESPACE, b"id")?,
        title: None,
        description: None,
        anchor_type: attribute(reader, element, TEXT_NAMESPACE, b"anchor-type")?,
        x: attribute(reader, element, SVG_NAMESPACE, b"x")?,
        y: attribute(reader, element, SVG_NAMESPACE, b"y")?,
        width: attribute(reader, element, SVG_NAMESPACE, b"width")?,
        height: attribute(reader, element, SVG_NAMESPACE, b"height")?,
        end_cell_address: attribute(reader, element, TABLE_NAMESPACE, b"end-cell-address")?,
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
    frame_state: Option<&mut FrameState>,
) -> Result<ImageBuilder> {
    let (frame_context, alternative_index) = match frame_state {
        Some(state) => {
            let alternative_index = state.image_count;
            state.image_count = state.image_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("image alternative count overflow".to_string())
            })?;
            state.image_indices.push(image_index);
            (Some(state.frame.clone()), alternative_index)
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
    let field_len = active.value.len().checked_add(value.len()).ok_or_else(|| {
        Error::InvalidFormat("image accessibility text size overflow".to_string())
    })?;
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
        Error::InvalidFormat(format!(
            "invalid image accessibility character reference: {error}"
        ))
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
    part: Part,
    package: Option<&dyn PackageLookup>,
    total_inline_bytes: &mut usize,
) -> Result<Image> {
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
        *total_inline_bytes = total_inline_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("total inline image size overflow".to_string()))?;
        if *total_inline_bytes > MAX_TOTAL_INLINE_IMAGE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "total inline image data exceeds {MAX_TOTAL_INLINE_IMAGE_BYTES} bytes"
            )));
        }
        Source::Inline {
            bytes,
            ignored_href: image.href.clone(),
        }
    } else if let Some(href) = image.href.clone().filter(|href| !href.is_empty()) {
        match package {
            None => Source::Linked { href },
            Some(_) if is_linked_href(&href) => Source::Linked { href },
            Some(package_lookup) => {
                let path = resolve_package_path(&href)?;
                if package_lookup.has_file(&path) {
                    Source::PackagePart {
                        href,
                        manifest_media_type: package_lookup.media_type(&path).map(str::to_owned),
                        path,
                    }
                } else {
                    Source::MissingPackagePart {
                        href,
                        resolved_path: path,
                    }
                }
            },
        }
    } else {
        Source::Missing
    };

    Ok(Image {
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
    matches!(namespace, ResolveResult::Bound(Namespace(bound_namespace)) if *bound_namespace == expected)
}
