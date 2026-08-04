//! Inert semantic discovery of embedded ODF and OLE objects.

use crate::ImageFrame;
use crate::elements::xml::namespaced_attribute;
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use litchi_odf_common::package::{is_linked_href, resolve_package_path};
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MATH_NAMESPACE: &[u8] = b"http://www.w3.org/1998/Math/MathML";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";

const MAX_OBJECT_DEPTH: usize = 4_096;
const MAX_OBJECTS: usize = 100_000;
const MAX_OBJECT_PARAMETERS: usize = 65_536;
const MAX_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_INLINE_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_INLINE_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_INLINE_ENCODED_BYTES: usize = 24 * 1024 * 1024;
const MAX_INLINE_BINARY_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_INLINE_BINARY_BYTES: usize = 64 * 1024 * 1024;
const MAX_ACCESSIBILITY_TEXT_BYTES: usize = 64 * 1024;
const MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES: usize = 8 * 1024 * 1024;

/// XML part containing an embedded-object occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedObjectPart {
    Content,
    Styles,
    FlatDocument,
}

/// Normative embedded-object element kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedObjectKind {
    Object,
    ObjectOle,
    Applet,
    Plugin,
    FloatingFrame,
}

/// One ordered, inert applet or plugin parameter.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedObjectParameter {
    pub name: String,
    pub value: String,
}

/// Root kind of an inline XML object payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum InlineObjectRoot {
    OpenDocument,
    MathMl,
}

/// Inert storage classification for an embedded object.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedObjectSource {
    InlineXml {
        root: InlineObjectRoot,
        xml: String,
        ignored_href: Option<String>,
    },
    InlineBinary {
        bytes: Vec<u8>,
        ignored_href: Option<String>,
    },
    PackageFile {
        href: String,
        path: String,
        manifest_media_type: Option<String>,
    },
    PackageSubdocument {
        href: String,
        root_path: String,
        content_path: String,
        manifest_media_type: Option<String>,
    },
    MissingPackagePart {
        href: String,
        resolved_path: String,
    },
    Linked {
        href: String,
    },
    Missing,
}

/// One inert `draw:object` or `draw:object-ole` occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedObject {
    pub part: EmbeddedObjectPart,
    pub kind: EmbeddedObjectKind,
    pub source: EmbeddedObjectSource,
    pub frame: Option<ImageFrame>,
    pub xml_id: Option<String>,
    pub class_id: Option<String>,
    pub notify_on_update_of_ranges: Option<String>,
    pub link_type: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
    pub code: Option<String>,
    pub object_name: Option<String>,
    pub archive: Option<String>,
    /// Stored applet scripting intent. No script or applet is ever started.
    pub may_script: Option<bool>,
    pub applet_name: Option<String>,
    pub mime_type: Option<String>,
    pub frame_name: Option<String>,
    pub parameters: Vec<EmbeddedObjectParameter>,
}

#[derive(Clone, Copy)]
struct PackageLookup<'a> {
    has_file: &'a dyn Fn(&str) -> bool,
    media_type: &'a dyn Fn(&str) -> Option<String>,
}

struct NamedContext {
    depth: usize,
    name: Option<String>,
}

struct FrameState {
    depth: usize,
    frame: ImageFrame,
    object_indices: Vec<usize>,
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

struct InlineXmlCapture {
    depth: usize,
    root: InlineObjectRoot,
    writer: Writer<Vec<u8>>,
}

struct ObjectBuilder {
    depth: usize,
    kind: EmbeddedObjectKind,
    href: Option<String>,
    frame: Option<ImageFrame>,
    xml_id: Option<String>,
    class_id: Option<String>,
    notify_on_update_of_ranges: Option<String>,
    link_type: Option<String>,
    show: Option<String>,
    actuate: Option<String>,
    code: Option<String>,
    object_name: Option<String>,
    archive: Option<String>,
    may_script: Option<bool>,
    applet_name: Option<String>,
    mime_type: Option<String>,
    frame_name: Option<String>,
    parameters: Vec<EmbeddedObjectParameter>,
    inline_xml: Option<(InlineObjectRoot, String)>,
    binary_present: bool,
    binary_depth: Option<usize>,
    binary_encoded: String,
}

pub(crate) fn scan_packaged_objects(
    content_xml: &str,
    styles_xml: Option<&str>,
    has_file: impl Fn(&str) -> bool,
    media_type: impl Fn(&str) -> Option<String>,
) -> Result<Vec<EmbeddedObject>> {
    let lookup = PackageLookup {
        has_file: &has_file,
        media_type: &media_type,
    };
    let mut objects = Vec::new();
    let mut total_xml = 0usize;
    let mut total_binary = 0usize;
    let mut total_accessibility = 0usize;
    scan_xml(
        content_xml,
        EmbeddedObjectPart::Content,
        Some(lookup),
        &mut objects,
        &mut total_xml,
        &mut total_binary,
        &mut total_accessibility,
    )?;
    if let Some(styles_xml) = styles_xml {
        scan_xml(
            styles_xml,
            EmbeddedObjectPart::Styles,
            Some(lookup),
            &mut objects,
            &mut total_xml,
            &mut total_binary,
            &mut total_accessibility,
        )?;
    }
    Ok(objects)
}

pub(crate) fn scan_flat_objects(xml: &str) -> Result<Vec<EmbeddedObject>> {
    let mut objects = Vec::new();
    let mut total_xml = 0usize;
    let mut total_binary = 0usize;
    let mut total_accessibility = 0usize;
    scan_xml(
        xml,
        EmbeddedObjectPart::FlatDocument,
        None,
        &mut objects,
        &mut total_xml,
        &mut total_binary,
        &mut total_accessibility,
    )?;
    Ok(objects)
}

#[allow(clippy::too_many_arguments)]
fn scan_xml(
    xml: &str,
    part: EmbeddedObjectPart,
    package: Option<PackageLookup<'_>>,
    objects: &mut Vec<EmbeddedObject>,
    total_xml_bytes: &mut usize,
    total_binary_bytes: &mut usize,
    total_accessibility_bytes: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut frames = Vec::<FrameState>::new();
    let mut pages = Vec::<NamedContext>::new();
    let mut sheets = Vec::<NamedContext>::new();
    let mut active: Option<ObjectBuilder> = None;
    let mut inline_capture: Option<InlineXmlCapture> = None;
    let mut accessibility: Option<AccessibilityText> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid embedded-object XML: {error}"))
            })?;

        if matches!(&event, Event::DocType(_)) {
            return Err(Error::InvalidFormat(
                "DTDs are not allowed while scanning embedded objects".to_string(),
            ));
        }

        if inline_capture.is_some() {
            match event {
                Event::Start(element) => {
                    depth = checked_depth(depth)?;
                    write_inline_event(
                        inline_capture.as_mut().expect("active inline XML"),
                        Event::Start(element),
                    )?;
                },
                Event::Empty(element) => write_inline_event(
                    inline_capture.as_mut().expect("active inline XML"),
                    Event::Empty(element),
                )?,
                Event::End(element) => {
                    let closing_root = inline_capture
                        .as_ref()
                        .is_some_and(|capture| capture.depth == depth);
                    write_inline_event(
                        inline_capture.as_mut().expect("active inline XML"),
                        Event::End(element),
                    )?;
                    if closing_root {
                        finish_inline_xml(
                            inline_capture.take().expect("closing inline XML"),
                            active.as_mut().expect("inline XML object"),
                            total_xml_bytes,
                        )?;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("embedded-object XML stack underflow".to_string())
                    })?;
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated inline embedded-object XML".to_string(),
                    ));
                },
                event => {
                    write_inline_event(inline_capture.as_mut().expect("active inline XML"), event)?
                },
            }
            buffer.clear();
            continue;
        }

        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                if active
                    .as_ref()
                    .and_then(|object| object.binary_depth)
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

                if let Some(object) = active.as_mut() {
                    if object_kind(&namespace, &element).is_some() {
                        return Err(Error::InvalidFormat(
                            "nested embedded-object elements are not allowed".to_string(),
                        ));
                    }
                    if depth == object.depth + 1
                        && bound_to(&namespace, DRAW_NAMESPACE)
                        && element.local_name().as_ref() == b"param"
                    {
                        push_parameter(&reader, &element, object)?;
                    } else if depth == object.depth + 1
                        && let Some(root) = inline_root(&namespace, &element)
                    {
                        if object.kind != EmbeddedObjectKind::Object {
                            return Err(Error::InvalidFormat(
                                "draw:object-ole must not contain inline XML objects".to_string(),
                            ));
                        }
                        if object.inline_xml.is_some() {
                            return Err(Error::InvalidFormat(
                                "draw:object contains multiple inline payloads".to_string(),
                            ));
                        }
                        let mut capture = InlineXmlCapture {
                            depth,
                            root,
                            writer: Writer::new(Vec::new()),
                        };
                        write_inline_event(&mut capture, Event::Start(element))?;
                        inline_capture = Some(capture);
                    } else if depth == object.depth + 1
                        && bound_to(&namespace, OFFICE_NAMESPACE)
                        && element.local_name().as_ref() == b"binary-data"
                    {
                        if object.kind != EmbeddedObjectKind::ObjectOle {
                            return Err(Error::InvalidFormat(
                                "draw:object must not contain office:binary-data".to_string(),
                            ));
                        }
                        if object.binary_present {
                            return Err(Error::InvalidFormat(
                                "draw:object-ole contains multiple office:binary-data elements"
                                    .to_string(),
                            ));
                        }
                        object.binary_present = true;
                        object.binary_depth = Some(depth);
                    }
                } else if frames.last().is_some_and(|frame| depth == frame.depth + 1)
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
                } else if bound_to(&namespace, DRAW_NAMESPACE)
                    && element.local_name().as_ref() == b"frame"
                {
                    frames.push(FrameState {
                        depth,
                        frame: parse_frame(&reader, &element, &pages, &sheets)?,
                        object_indices: Vec::new(),
                    });
                } else if let Some(kind) = object_kind(&namespace, &element) {
                    ensure_object_capacity(objects.len())?;
                    active = Some(start_object(
                        &reader,
                        &element,
                        depth,
                        kind,
                        objects.len(),
                        frames.last_mut(),
                    )?);
                }
            },
            Event::Empty(element) => {
                if active
                    .as_ref()
                    .and_then(|object| object.binary_depth)
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
                if let Some(object) = active.as_mut() {
                    if depth == object.depth
                        && bound_to(&namespace, DRAW_NAMESPACE)
                        && element.local_name().as_ref() == b"param"
                    {
                        push_parameter(&reader, &element, object)?;
                    } else if depth == object.depth
                        && let Some(root) = inline_root(&namespace, &element)
                    {
                        if object.kind != EmbeddedObjectKind::Object || object.inline_xml.is_some()
                        {
                            return Err(Error::InvalidFormat(
                                "invalid duplicate inline embedded-object payload".to_string(),
                            ));
                        }
                        let mut capture = InlineXmlCapture {
                            depth: depth + 1,
                            root,
                            writer: Writer::new(Vec::new()),
                        };
                        write_inline_event(&mut capture, Event::Empty(element))?;
                        finish_inline_xml(capture, object, total_xml_bytes)?;
                    } else if depth == object.depth
                        && bound_to(&namespace, OFFICE_NAMESPACE)
                        && element.local_name().as_ref() == b"binary-data"
                    {
                        if object.kind != EmbeddedObjectKind::ObjectOle || object.binary_present {
                            return Err(Error::InvalidFormat(
                                "invalid duplicate inline OLE payload".to_string(),
                            ));
                        }
                        object.binary_present = true;
                    } else if object_kind(&namespace, &element).is_some() {
                        return Err(Error::InvalidFormat(
                            "nested embedded-object elements are not allowed".to_string(),
                        ));
                    }
                } else if frames.last().is_some_and(|frame| depth == frame.depth)
                    && let Some(kind) =
                        accessibility_kind(&namespace, element.local_name().as_ref())
                {
                    set_empty_accessibility(frames.last_mut().expect("direct frame child"), kind)?;
                } else if let Some(kind) = object_kind(&namespace, &element) {
                    ensure_object_capacity(objects.len())?;
                    let object = start_object(
                        &reader,
                        &element,
                        depth + 1,
                        kind,
                        objects.len(),
                        frames.last_mut(),
                    )?;
                    objects.push(finish_object(object, part, package, total_binary_bytes)?);
                }
            },
            Event::Text(value)
                if active
                    .as_ref()
                    .and_then(|object| object.binary_depth)
                    .is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid inline OLE text: {error}"))
                    })?;
                append_binary(active.as_mut().expect("active OLE object"), &value)?;
            },
            Event::CData(value)
                if active
                    .as_ref()
                    .and_then(|object| object.binary_depth)
                    .is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid inline OLE CDATA: {error}"))
                    })?;
                append_binary(active.as_mut().expect("active OLE object"), &value)?;
            },
            Event::GeneralRef(_)
                if active
                    .as_ref()
                    .and_then(|object| object.binary_depth)
                    .is_some() =>
            {
                return Err(Error::InvalidFormat(
                    "XML references are not allowed in office:binary-data".to_string(),
                ));
            },
            Event::Text(value) if accessibility.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid object accessibility text: {error}"))
                    })?;
                append_accessibility(
                    accessibility.as_mut().expect("active accessibility text"),
                    &value,
                    total_accessibility_bytes,
                )?;
            },
            Event::CData(value) if accessibility.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid object accessibility CDATA: {error}"))
                    })?;
                append_accessibility(
                    accessibility.as_mut().expect("active accessibility text"),
                    &value,
                    total_accessibility_bytes,
                )?;
            },
            Event::GeneralRef(value) if accessibility.is_some() => {
                let value = decode_reference(&value)?;
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
                            "malformed object accessibility element".to_string(),
                        ));
                    }
                    finish_accessibility(
                        frames.last_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "object accessibility text has no enclosing frame".to_string(),
                            )
                        })?,
                        text,
                    )?;
                } else if let Some(object) = active.as_mut()
                    && object.binary_depth == Some(depth)
                {
                    if !bound_to(&namespace, OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"binary-data"
                    {
                        return Err(Error::InvalidFormat(
                            "malformed office:binary-data element".to_string(),
                        ));
                    }
                    object.binary_depth = None;
                } else if active.as_ref().map(|object| object.depth) == Some(depth) {
                    let object = active.take().expect("closing embedded object");
                    objects.push(finish_object(object, part, package, total_binary_bytes)?);
                }

                if frames.last().map(|frame| frame.depth) == Some(depth) {
                    let frame = frames.pop().expect("closing frame");
                    for object_index in frame.object_indices {
                        let object = objects.get_mut(object_index).ok_or_else(|| {
                            Error::InvalidFormat(
                                "embedded-object frame occurrence index is invalid".to_string(),
                            )
                        })?;
                        object.frame = Some(frame.frame.clone());
                    }
                }
                if pages.last().map(|page| page.depth) == Some(depth) {
                    pages.pop();
                }
                if sheets.last().map(|sheet| sheet.depth) == Some(depth) {
                    sheets.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("embedded-object XML stack underflow".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if depth != 0
        || active.is_some()
        || inline_capture.is_some()
        || accessibility.is_some()
        || !frames.is_empty()
        || !pages.is_empty()
        || !sheets.is_empty()
    {
        return Err(Error::InvalidFormat(
            "unterminated embedded-object XML".to_string(),
        ));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("embedded-object nesting overflow".to_string()))?;
    if depth > MAX_OBJECT_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "embedded-object nesting exceeds {MAX_OBJECT_DEPTH}"
        )));
    }
    Ok(depth)
}

fn ensure_object_capacity(current: usize) -> Result<()> {
    if current >= MAX_OBJECTS {
        return Err(Error::InvalidFormat(format!(
            "embedded-object count exceeds {MAX_OBJECTS}"
        )));
    }
    Ok(())
}

fn object_kind(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Option<EmbeddedObjectKind> {
    if !bound_to(namespace, DRAW_NAMESPACE) {
        return None;
    }
    match element.local_name().as_ref() {
        b"object" => Some(EmbeddedObjectKind::Object),
        b"object-ole" => Some(EmbeddedObjectKind::ObjectOle),
        b"applet" => Some(EmbeddedObjectKind::Applet),
        b"plugin" => Some(EmbeddedObjectKind::Plugin),
        b"floating-frame" => Some(EmbeddedObjectKind::FloatingFrame),
        _ => None,
    }
}

fn push_parameter(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    object: &mut ObjectBuilder,
) -> Result<()> {
    if !matches!(
        object.kind,
        EmbeddedObjectKind::Applet | EmbeddedObjectKind::Plugin
    ) {
        return Err(Error::InvalidFormat(
            "draw:param is allowed only in draw:applet or draw:plugin".to_string(),
        ));
    }
    if object.parameters.len() >= MAX_OBJECT_PARAMETERS {
        return Err(Error::InvalidFormat(format!(
            "embedded object exceeds {MAX_OBJECT_PARAMETERS} parameters"
        )));
    }
    let name = limited_attribute(reader, element, DRAW_NAMESPACE, b"name")?
        .ok_or_else(|| Error::InvalidFormat("draw:param requires draw:name".to_string()))?;
    let value = limited_attribute(reader, element, DRAW_NAMESPACE, b"value")?
        .ok_or_else(|| Error::InvalidFormat("draw:param requires draw:value".to_string()))?;
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "draw:param name must not be empty".to_string(),
        ));
    }
    object
        .parameters
        .push(EmbeddedObjectParameter { name, value });
    Ok(())
}

fn inline_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Option<InlineObjectRoot> {
    if bound_to(namespace, OFFICE_NAMESPACE) && element.local_name().as_ref() == b"document" {
        Some(InlineObjectRoot::OpenDocument)
    } else if bound_to(namespace, MATH_NAMESPACE) && element.local_name().as_ref() == b"math" {
        Some(InlineObjectRoot::MathMl)
    } else {
        None
    }
}

fn start_object(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    kind: EmbeddedObjectKind,
    object_index: usize,
    frame: Option<&mut FrameState>,
) -> Result<ObjectBuilder> {
    let href = limited_attribute(reader, element, XLINK_NAMESPACE, b"href")?;
    let class_id = limited_attribute(reader, element, DRAW_NAMESPACE, b"class-id")?;
    let notify_on_update_of_ranges = limited_attribute(
        reader,
        element,
        DRAW_NAMESPACE,
        b"notify-on-update-of-ranges",
    )?;
    let frame_context = match frame {
        Some(frame) => {
            frame.object_indices.push(object_index);
            Some(frame.frame.clone())
        },
        None => None,
    };
    Ok(ObjectBuilder {
        depth,
        kind,
        href,
        frame: frame_context,
        xml_id: attribute(reader, element, XML_NAMESPACE, b"id")?,
        class_id,
        notify_on_update_of_ranges,
        link_type: limited_attribute(reader, element, XLINK_NAMESPACE, b"type")?,
        show: limited_attribute(reader, element, XLINK_NAMESPACE, b"show")?,
        actuate: limited_attribute(reader, element, XLINK_NAMESPACE, b"actuate")?,
        code: limited_attribute(reader, element, DRAW_NAMESPACE, b"code")?,
        object_name: limited_attribute(reader, element, DRAW_NAMESPACE, b"object")?,
        archive: limited_attribute(reader, element, DRAW_NAMESPACE, b"archive")?,
        may_script: limited_attribute(reader, element, DRAW_NAMESPACE, b"may-script")?
            .map(|value| parse_object_bool("draw:may-script", &value))
            .transpose()?,
        applet_name: limited_attribute(reader, element, DRAW_NAMESPACE, b"applet-name")?,
        mime_type: limited_attribute(reader, element, DRAW_NAMESPACE, b"mime-type")?,
        frame_name: limited_attribute(reader, element, DRAW_NAMESPACE, b"frame-name")?,
        parameters: Vec::new(),
        inline_xml: None,
        binary_present: false,
        binary_depth: None,
        binary_encoded: String::new(),
    })
}

fn finish_object(
    object: ObjectBuilder,
    part: EmbeddedObjectPart,
    package: Option<PackageLookup<'_>>,
    total_binary_bytes: &mut usize,
) -> Result<EmbeddedObject> {
    let source = if let Some((root, xml)) = object.inline_xml {
        EmbeddedObjectSource::InlineXml {
            root,
            xml,
            ignored_href: object.href.clone(),
        }
    } else if object.binary_present {
        let mut compact = Vec::with_capacity(object.binary_encoded.len());
        for byte in object.binary_encoded.bytes() {
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
            Error::InvalidFormat(format!("invalid embedded OLE base64: {error}"))
        })?;
        if bytes.len() > MAX_INLINE_BINARY_BYTES {
            return Err(Error::InvalidFormat(format!(
                "inline OLE exceeds {MAX_INLINE_BINARY_BYTES} decoded bytes"
            )));
        }
        *total_binary_bytes = total_binary_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("total inline OLE size overflow".to_string()))?;
        if *total_binary_bytes > MAX_TOTAL_INLINE_BINARY_BYTES {
            return Err(Error::InvalidFormat(format!(
                "total inline OLE data exceeds {MAX_TOTAL_INLINE_BINARY_BYTES} bytes"
            )));
        }
        EmbeddedObjectSource::InlineBinary {
            bytes,
            ignored_href: object.href.clone(),
        }
    } else if let Some(href) = object.href.clone().filter(|href| !href.is_empty()) {
        if matches!(
            object.kind,
            EmbeddedObjectKind::Applet
                | EmbeddedObjectKind::Plugin
                | EmbeddedObjectKind::FloatingFrame
        ) {
            EmbeddedObjectSource::Linked { href }
        } else {
            match package {
                None => EmbeddedObjectSource::Linked { href },
                Some(_) if is_linked_href(&href) => EmbeddedObjectSource::Linked { href },
                Some(package) => {
                    let path = resolve_package_path(&href)?;
                    if (package.has_file)(&path) {
                        EmbeddedObjectSource::PackageFile {
                            href,
                            manifest_media_type: (package.media_type)(&path),
                            path,
                        }
                    } else {
                        let content_path = format!("{path}/content.xml");
                        if (package.has_file)(&content_path) {
                            let root_path = format!("{path}/");
                            EmbeddedObjectSource::PackageSubdocument {
                                href,
                                manifest_media_type: (package.media_type)(&root_path)
                                    .or_else(|| (package.media_type)(&path)),
                                root_path,
                                content_path,
                            }
                        } else {
                            EmbeddedObjectSource::MissingPackagePart {
                                href,
                                resolved_path: path,
                            }
                        }
                    }
                },
            }
        }
    } else {
        EmbeddedObjectSource::Missing
    };

    Ok(EmbeddedObject {
        part,
        kind: object.kind,
        source,
        frame: object.frame,
        xml_id: object.xml_id,
        class_id: object.class_id,
        notify_on_update_of_ranges: object.notify_on_update_of_ranges,
        link_type: object.link_type,
        show: object.show,
        actuate: object.actuate,
        code: object.code,
        object_name: object.object_name,
        archive: object.archive,
        may_script: object.may_script,
        applet_name: object.applet_name,
        mime_type: object.mime_type,
        frame_name: object.frame_name,
        parameters: object.parameters,
    })
}

fn parse_object_bool(name: &str, value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "{name} is not an XML Schema boolean"
        ))),
    }
}

fn write_inline_event(capture: &mut InlineXmlCapture, event: Event<'_>) -> Result<()> {
    capture.writer.write_event(event).map_err(|error| {
        Error::InvalidFormat(format!("cannot retain inline embedded-object XML: {error}"))
    })?;
    if capture.writer.get_ref().len() > MAX_INLINE_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "inline embedded-object XML exceeds {MAX_INLINE_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

fn finish_inline_xml(
    capture: InlineXmlCapture,
    object: &mut ObjectBuilder,
    total_xml_bytes: &mut usize,
) -> Result<()> {
    if object.inline_xml.is_some() {
        return Err(Error::InvalidFormat(
            "draw:object contains multiple inline payloads".to_string(),
        ));
    }
    let bytes = capture.writer.into_inner();
    *total_xml_bytes = total_xml_bytes.checked_add(bytes.len()).ok_or_else(|| {
        Error::InvalidFormat("total inline embedded-object XML size overflow".to_string())
    })?;
    if *total_xml_bytes > MAX_TOTAL_INLINE_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "total inline embedded-object XML exceeds {MAX_TOTAL_INLINE_XML_BYTES} bytes"
        )));
    }
    let xml = String::from_utf8(bytes).map_err(|_| {
        Error::InvalidFormat("inline embedded-object XML is not valid UTF-8".to_string())
    })?;
    object.inline_xml = Some((capture.root, xml));
    Ok(())
}

fn append_binary(object: &mut ObjectBuilder, value: &str) -> Result<()> {
    let new_len = object
        .binary_encoded
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("inline OLE size overflow".to_string()))?;
    if new_len > MAX_INLINE_ENCODED_BYTES {
        return Err(Error::InvalidFormat(format!(
            "inline OLE encoding exceeds {MAX_INLINE_ENCODED_BYTES} bytes"
        )));
    }
    object.binary_encoded.push_str(value);
    Ok(())
}

fn parse_frame(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    pages: &[NamedContext],
    sheets: &[NamedContext],
) -> Result<ImageFrame> {
    Ok(ImageFrame {
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
        sheet_shape: false,
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
        return Err(Error::InvalidFormat(
            "draw:frame contains duplicate accessibility metadata".to_string(),
        ));
    }
    Ok(())
}

fn append_accessibility(
    active: &mut AccessibilityText,
    value: &str,
    total_accessibility_bytes: &mut usize,
) -> Result<()> {
    let field_len =
        active.value.len().checked_add(value.len()).ok_or_else(|| {
            Error::InvalidFormat("object accessibility size overflow".to_string())
        })?;
    if field_len > MAX_ACCESSIBILITY_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "object accessibility text exceeds {MAX_ACCESSIBILITY_TEXT_BYTES} bytes"
        )));
    }
    *total_accessibility_bytes = total_accessibility_bytes
        .checked_add(value.len())
        .ok_or_else(|| {
            Error::InvalidFormat("total object accessibility size overflow".to_string())
        })?;
    if *total_accessibility_bytes > MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "total object accessibility text exceeds {MAX_TOTAL_ACCESSIBILITY_TEXT_BYTES} bytes"
        )));
    }
    active.value.push_str(value);
    Ok(())
}

fn decode_reference(value: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = value.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid object accessibility reference: {error}"))
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
                "unsupported entity in object accessibility text".to_string(),
            ));
        },
    };
    Ok(character.to_string())
}

fn limited_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    let value = attribute(reader, element, namespace, local_name)?;
    if value
        .as_ref()
        .is_some_and(|value| value.len() > MAX_ATTRIBUTE_BYTES)
    {
        return Err(Error::InvalidFormat(format!(
            "embedded-object attribute exceeds {MAX_ATTRIBUTE_BYTES} bytes"
        )));
    }
    Ok(value)
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    namespaced_attribute(reader, element, namespace, local_name, "embedded object")
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(namespace)) if *namespace == expected)
}

#[cfg(test)]
mod active_object_tests {
    use super::*;

    const PREFIX: &str = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:drawing>"#;
    const SUFFIX: &str = "</o:drawing></o:body></o:document>";

    #[test]
    fn retains_applets_plugins_and_floating_frames_inertly() {
        let xml = format!(
            r#"{PREFIX}<d:frame d:name="Applet"><d:applet x:href="https://example.invalid/app" d:code="Main" d:archive="app.jar" d:may-script="true"><d:param d:name="theme" d:value="dark"/></d:applet></d:frame><d:frame><d:plugin x:href="media.bin" d:mime-type="application/x-example"><d:param d:name="quality" d:value="high"></d:param></d:plugin></d:frame><d:frame><d:floating-frame x:href="https://example.invalid/frame" d:frame-name="preview"/></d:frame>{SUFFIX}"#
        );
        let objects = scan_flat_objects(&xml).unwrap();
        assert_eq!(objects.len(), 3);
        assert_eq!(objects[0].kind, EmbeddedObjectKind::Applet);
        assert_eq!(objects[0].code.as_deref(), Some("Main"));
        assert_eq!(objects[0].may_script, Some(true));
        assert_eq!(objects[0].parameters[0].name, "theme");
        assert!(matches!(
            objects[0].source,
            EmbeddedObjectSource::Linked { .. }
        ));
        assert_eq!(objects[1].kind, EmbeddedObjectKind::Plugin);
        assert_eq!(
            objects[1].mime_type.as_deref(),
            Some("application/x-example")
        );
        assert_eq!(objects[2].kind, EmbeddedObjectKind::FloatingFrame);
        assert_eq!(objects[2].frame_name.as_deref(), Some("preview"));
    }

    #[test]
    fn rejects_invalid_active_object_metadata_and_nesting() {
        for body in [
            r#"<d:applet d:may-script="yes"/>"#,
            r#"<d:plugin><d:param d:value="x"/></d:plugin>"#,
            r#"<d:floating-frame><d:param d:name="x" d:value="y"/></d:floating-frame>"#,
            r#"<d:plugin><d:applet/></d:plugin>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(scan_flat_objects(&xml).is_err(), "accepted {body}");
        }
    }
}
