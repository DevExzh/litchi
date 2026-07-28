//! Atomic authoring of inert embedded OpenDocument chart objects.

use crate::core::{OwnedPackage, PackageWriter};
use crate::elements::xml::namespaced_attribute;
use crate::{
    ChartDefinition, ChartDocument, OdfEmbeddedObject, OdfEmbeddedObjectKind,
    OdfEmbeddedObjectPart, OdfEmbeddedObjectSource, OdfInlineObjectRoot, constants,
    serialize_chart_content,
};
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const DRAW_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TABLE_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_EMBEDDED_FILES: usize = 16_384;
const MAX_EMBEDDED_BYTES: usize = 64 * 1024 * 1024;
const MAX_CONTENT_BYTES: usize = 16 * 1024 * 1024;

/// Storage form for a newly created embedded chart.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfEmbeddedChartStorage {
    /// A referenced `Object_N/` subdocument with manifest entries.
    #[default]
    PackageSubdocument,
    /// An inline flat `office:document` child of `draw:object`.
    InlineXml,
}

pub(crate) enum EmbeddedChartHost<'a> {
    Text,
    Sheet(&'a str),
    Page(&'a str),
}

pub(crate) struct Addition {
    pub(crate) path: String,
    pub(crate) bytes: Vec<u8>,
    pub(crate) media_type: String,
}

#[derive(Clone)]
pub(crate) struct ObjectSpan {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) inline_payload: Option<(usize, usize)>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Table,
    Other,
}

pub(crate) fn open_embedded_chart(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
) -> Result<ChartDocument> {
    let current_objects = objects(package, content, styles)?;
    let object = select_chart_object(&current_objects, index)?;
    match &object.source {
        OdfEmbeddedObjectSource::PackageSubdocument {
            root_path,
            manifest_media_type,
            ..
        } => open_subdocument(package, root_path, manifest_media_type.as_deref()),
        OdfEmbeddedObjectSource::InlineXml {
            root: OdfInlineObjectRoot::OpenDocument,
            xml,
            ..
        } => open_inline(xml),
        _ => invalid("selected object is not an embedded OpenDocument chart"),
    }
}

pub(crate) fn replace_embedded_chart(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    definition: &ChartDefinition,
) -> Result<Vec<u8>> {
    let objects = objects(package, content, styles)?;
    let object = select_chart_object(&objects, index)?;
    // Opening first validates MIME type and the existing chart hierarchy.
    let _ = open_embedded_chart(package, content, styles, index)?;
    let chart_content = serialize_chart_content(definition)?;
    match &object.source {
        OdfEmbeddedObjectSource::PackageSubdocument {
            content_path,
            ..
        } => rebuild_package(
            package,
            content,
            vec![Addition {
                path: content_path.clone(),
                bytes: chart_content.into_bytes(),
                media_type: "text/xml".to_string(),
            }],
            Vec::new(),
            vec![content_path.clone()],
            Vec::new(),
        ),
        OdfEmbeddedObjectSource::InlineXml { .. } => {
            let spans = locate_objects(content)?;
            let span = spans.get(index).ok_or_else(|| {
                Error::InvalidFormat("embedded-object scanner/span mismatch".to_string())
            })?;
            let (start, end) = span.inline_payload.ok_or_else(|| {
                Error::InvalidFormat("inline chart payload span is missing".to_string())
            })?;
            let inline = content_to_inline(&chart_content)?;
            let updated = splice(content, start, end, &inline)?;
            rebuild_package(package, &updated, Vec::new(), Vec::new(), Vec::new(), Vec::new())
        },
        _ => invalid("selected object is not a replaceable embedded chart"),
    }
}

pub(crate) fn remove_embedded_chart(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
) -> Result<Vec<u8>> {
    let current_objects = objects(package, content, styles)?;
    let object = select_chart_object(&current_objects, index)?;
    let _ = open_embedded_chart(package, content, styles, index)?;
    let spans = locate_objects(content)?;
    let span = spans.get(index).ok_or_else(|| {
        Error::InvalidFormat("embedded-object scanner/span mismatch".to_string())
    })?;
    let updated = splice(content, span.start, span.end, "")?;
    let mut excluded_prefixes = Vec::new();
    if let OdfEmbeddedObjectSource::PackageSubdocument { root_path, .. } = &object.source {
        let remaining = objects(package, &updated, styles)?;
        let still_referenced = remaining.iter().any(|candidate| {
            matches!(&candidate.source, OdfEmbeddedObjectSource::PackageSubdocument { root_path: candidate_root, .. } if candidate_root.as_str() == root_path.as_str())
        });
        if !still_referenced {
            excluded_prefixes.push(root_path.clone());
        }
    }
    rebuild_package(
        package,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        excluded_prefixes,
    )
}

pub(crate) fn add_embedded_chart(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    host: EmbeddedChartHost<'_>,
    storage: OdfEmbeddedChartStorage,
    definition: &ChartDefinition,
) -> Result<(Vec<u8>, usize)> {
    let current = objects(package, content, styles)?;
    let index = current
        .iter()
        .filter(|object| object.part == OdfEmbeddedObjectPart::Content)
        .count();
    let chart_content = serialize_chart_content(definition)?;
    let root = unused_object_root(package)?;
    let (object_xml, additions, directories) = match storage {
        OdfEmbeddedChartStorage::PackageSubdocument => {
            let href = root.trim_end_matches('/');
            (
                format!(
                    "<draw:object xlink:href=\"./{href}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>"
                ),
                vec![Addition {
                    path: format!("{root}content.xml"),
                    bytes: chart_content.into_bytes(),
                    media_type: "text/xml".to_string(),
                }],
                vec![(root.clone(), constants::ODF_CHART.to_string())],
            )
        },
        OdfEmbeddedChartStorage::InlineXml => (
            format!("<draw:object>{}</draw:object>", content_to_inline(&chart_content)?),
            Vec::new(),
            Vec::new(),
        ),
    };
    let frame = format!(
        "<draw:frame xmlns:draw=\"{DRAW_URI}\" xmlns:xlink=\"{XLINK_NS}\" draw:name=\"Embedded Chart {}\">{object_xml}</draw:frame>",
        index + 1
    );
    let updated = insert_at_host(content, host, &frame)?;
    let bytes = rebuild_package(
        package,
        &updated,
        additions,
        directories,
        Vec::new(),
        Vec::new(),
    )?;
    Ok((bytes, index))
}

fn objects(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<OdfEmbeddedObject>> {
    let borrowed = package.package()?;
    crate::embedded_object::scan_packaged_objects(
        content,
        styles,
        |path| borrowed.has_file(path),
        |path| borrowed.manifest().get_media_type(path).map(str::to_string),
    )
}

fn select_chart_object(objects: &[OdfEmbeddedObject], index: usize) -> Result<&OdfEmbeddedObject> {
    let object = objects
        .get(index)
        .ok_or_else(|| Error::InvalidFormat(format!("embedded-object index {index} is out of bounds")))?;
    if object.part != OdfEmbeddedObjectPart::Content
        || object.kind != OdfEmbeddedObjectKind::Object
        || object.class_id.is_some()
        || object.code.is_some()
        || object.archive.is_some()
        || object.may_script.is_some()
        || object.applet_name.is_some()
        || object.mime_type.is_some()
        || !object.parameters.is_empty()
        || object.link_type.as_deref().is_some_and(|value| value != "simple")
        || object.show.as_deref().is_some_and(|value| value != "embed")
        || object.actuate.as_deref().is_some_and(|value| value != "onLoad")
    {
        return invalid("selected object has active, external, or unsupported metadata");
    }
    Ok(object)
}

fn open_subdocument(
    source: &OwnedPackage,
    root: &str,
    media_type: Option<&str>,
) -> Result<ChartDocument> {
    let media_type = media_type.ok_or_else(|| {
        Error::InvalidFormat("embedded chart root has no manifest media type".to_string())
    })?;
    if !matches!(media_type, constants::ODF_CHART | constants::ODF_CHART_TEMPLATE) {
        return invalid(format!("embedded chart root has invalid media type '{media_type}'"));
    }
    let package = source.package()?;
    if package.manifest().has_encrypted_entries() {
        return invalid("encrypted embedded chart packages cannot be opened for mutation");
    }
    let mut paths: Vec<String> = package
        .files()?
        .into_iter()
        .filter(|path| path.starts_with(root) && !path.ends_with('/'))
        .collect();
    if paths.len() > MAX_EMBEDDED_FILES {
        return invalid("embedded chart contains too many package entries");
    }
    paths.sort();
    let content_path = format!("{root}content.xml");
    if !paths.iter().any(|path| path == &content_path) {
        return invalid("embedded chart has no content.xml");
    }
    let mut writer = PackageWriter::new();
    writer.set_mimetype(media_type)?;
    let mut total = 0usize;
    let mut ordered = vec![content_path.clone()];
    ordered.extend(paths.into_iter().filter(|path| path != &content_path));
    for path in ordered {
        let relative = &path[root.len()..];
        if relative == "mimetype" || relative == "META-INF/manifest.xml" || relative.is_empty() {
            continue;
        }
        let bytes = package.get_file(&path)?;
        total = total.checked_add(bytes.len()).ok_or_else(|| {
            Error::InvalidFormat("embedded chart byte count overflow".to_string())
        })?;
        if total > MAX_EMBEDDED_BYTES || (relative == "content.xml" && bytes.len() > MAX_CONTENT_BYTES) {
            return invalid("embedded chart exceeds package size limits");
        }
        let entry_media = package.manifest().get_media_type(&path).unwrap_or_else(|| {
            if relative.ends_with(".xml") { "text/xml" } else { "application/octet-stream" }
        });
        writer.add_file_with_media_type(relative, &bytes, entry_media)?;
    }
    ChartDocument::from_bytes(writer.finish_to_bytes()?)
}

fn open_inline(xml: &str) -> Result<ChartDocument> {
    if xml.len() > MAX_CONTENT_BYTES {
        return invalid("inline chart exceeds content size limit");
    }
    let media_type = inline_chart_mimetype(xml)?;
    let content = rename_document_root(xml, "document", "document-content", None)?;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(&media_type)?;
    writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
    ChartDocument::from_bytes(writer.finish_to_bytes()?)
}

fn inline_chart_mimetype(xml: &str) -> Result<String> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid inline chart XML: {error}"))
        })?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element) => {
                if namespace != NamespaceKind::Office || element.local_name().as_ref() != b"document" {
                    return invalid("inline chart root must be office:document");
                }
                let media_type = namespaced_attribute(
                    &reader,
                    &element,
                    OFFICE_NS,
                    b"mimetype",
                    "inline chart root",
                )?
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "inline chart root is missing office:mimetype".to_string(),
                    )
                })?;
                if !matches!(media_type.as_str(), constants::ODF_CHART | constants::ODF_CHART_TEMPLATE) {
                    return invalid(format!(
                        "inline chart root has invalid office:mimetype '{media_type}'"
                    ));
                }
                return Ok(media_type);
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in inline charts"),
            Event::Empty(_) => return invalid("inline chart document root cannot be empty"),
            Event::Text(value) if {
                let bytes: &[u8] = value.as_ref();
                bytes.iter().all(u8::is_ascii_whitespace)
            } => {},
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("inline chart has no document root"),
            _ => return invalid("invalid content before inline chart root"),
        }
        buffer.clear();
    }
}

fn content_to_inline(content: &str) -> Result<String> {
    rename_document_root(
        content,
        "document-content",
        "document",
        Some(("office:mimetype", constants::ODF_CHART)),
    )
}

fn rename_document_root(
    xml: &str,
    expected_local: &str,
    replacement_local: &str,
    added_attribute: Option<(&str, &str)>,
) -> Result<String> {
    let mut root_start = xml.find('<').ok_or_else(|| Error::InvalidFormat("chart XML has no root".to_string()))?;
    if xml[root_start..].starts_with("<?xml") {
        let declaration_end = xml[root_start..]
            .find("?>")
            .ok_or_else(|| Error::InvalidFormat("unterminated XML declaration".to_string()))?
            + root_start
            + 2;
        root_start = xml[declaration_end..]
            .find('<')
            .map(|offset| declaration_end + offset)
            .ok_or_else(|| Error::InvalidFormat("chart XML has no document root".to_string()))?;
    }
    let name_end = xml[root_start + 1..]
        .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
        .map(|offset| root_start + 1 + offset)
        .ok_or_else(|| Error::InvalidFormat("invalid chart root start tag".to_string()))?;
    let qname = &xml[root_start + 1..name_end];
    let (prefix, local) = qname.rsplit_once(':').unwrap_or(("", qname));
    if local != expected_local {
        return invalid(format!("expected office:{expected_local} inline chart root"));
    }
    let close_start = xml.rfind("</").ok_or_else(|| Error::InvalidFormat("chart root is not closed".to_string()))?;
    let close_name_end = xml[close_start + 2..]
        .find('>')
        .map(|offset| close_start + 2 + offset)
        .ok_or_else(|| Error::InvalidFormat("invalid chart root closing tag".to_string()))?;
    let close_name = &xml[close_start + 2..close_name_end];
    if close_name != qname {
        return invalid("chart root start/end names do not match");
    }
    let replacement = if prefix.is_empty() {
        replacement_local.to_string()
    } else {
        format!("{prefix}:{replacement_local}")
    };
    let mut out = String::with_capacity(xml.len() + 96);
    out.push_str(&xml[root_start..root_start + 1]);
    out.push_str(&replacement);
    if let Some((name, value)) = added_attribute {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(value);
        out.push('"');
    }
    out.push_str(&xml[name_end..close_start + 2]);
    out.push_str(&replacement);
    out.push_str(&xml[close_name_end..]);
    Ok(out)
}

pub(crate) fn rebuild_package(
    source: &OwnedPackage,
    content: &str,
    additions: Vec<Addition>,
    directories: Vec<(String, String)>,
    excluded_paths: Vec<String>,
    excluded_prefixes: Vec<String>,
) -> Result<Vec<u8>> {
    if content.len() > MAX_CONTENT_BYTES {
        return invalid("outer content.xml exceeds embedded-chart mutation limit");
    }
    let mut writer = PackageWriter::new();
    writer.set_mimetype(&source.mimetype()?)?;
    writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
    if source.has_file(constants::ODF_STYLES)? {
        writer.add_file(constants::ODF_STYLES, &source.get_file(constants::ODF_STYLES)?)?;
    }
    if source.has_file(constants::ODF_META)? {
        writer.add_file(constants::ODF_META, &source.get_file(constants::ODF_META)?)?;
    }
    for (path, media_type) in directories {
        writer.add_manifest_directory(&path, &media_type)?;
    }
    let mut addition_bytes = 0usize;
    for addition in additions {
        addition_bytes = addition_bytes.checked_add(addition.bytes.len()).ok_or_else(|| {
            Error::InvalidFormat("embedded chart addition size overflow".to_string())
        })?;
        if addition_bytes > MAX_EMBEDDED_BYTES {
            return invalid("embedded chart additions exceed size limit");
        }
        writer.add_file_with_media_type(&addition.path, &addition.bytes, &addition.media_type)?;
    }
    writer.copy_auxiliary_files_from_except(source, &excluded_paths, &excluded_prefixes)?;
    writer.finish_to_bytes()
}

pub(crate) fn unused_object_root(package: &OwnedPackage) -> Result<String> {
    let files = package.files()?;
    let manifest = package.package()?;
    for index in 1..=100_000usize {
        let root = format!("Object_{index}/");
        if !files.iter().any(|path| path.starts_with(&root))
            && manifest.manifest().get_media_type(&root).is_none()
        {
            return Ok(root);
        }
    }
    invalid("no embedded chart object path is available")
}

pub(crate) fn locate_objects(xml: &str) -> Result<Vec<ObjectSpan>> {
    struct Active {
        depth: usize,
        start: usize,
        payload: Option<(usize, usize, usize)>,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<Active> = None;
    let mut spans = Vec::new();
    loop {
        let event_start = position(&reader)?;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid embedded chart host XML: {error}"))
        })?;
        let namespace = namespace_kind(&namespace);
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                if is_object(namespace, element.local_name().as_ref()) {
                    if active.is_some() { return invalid("nested embedded objects are not supported"); }
                    active = Some(Active { depth, start: event_start, payload: None });
                } else if let Some(object) = active.as_mut()
                    && depth == object.depth + 1
                    && namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"document"
                {
                    object.payload = Some((depth, event_start, 0));
                }
                depth = depth.checked_add(1).ok_or_else(|| Error::InvalidFormat("XML depth overflow".to_string()))?;
            },
            Event::Empty(element)
                if is_object(namespace, element.local_name().as_ref()) => {
                    spans.push(ObjectSpan { start: event_start, end: event_end, inline_payload: None });
                },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| Error::InvalidFormat("XML depth underflow".to_string()))?;
                if let Some(object) = active.as_mut()
                    && object.payload.is_some_and(|(payload_depth, _, _)| payload_depth == depth)
                    && namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"document"
                    && let Some((payload_depth, start, _)) = object.payload
                {
                    object.payload = Some((payload_depth, start, event_end));
                }
                if active.as_ref().is_some_and(|object| object.depth == depth)
                    && is_object(namespace, element.local_name().as_ref())
                {
                    let object = active.take().expect("active object");
                    spans.push(ObjectSpan {
                        start: object.start,
                        end: event_end,
                        inline_payload: object.payload.and_then(|(_, start, end)| (end != 0).then_some((start, end))),
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() { return invalid("unterminated embedded object"); }
    Ok(spans)
}

pub(crate) fn insert_at_host(xml: &str, host: EmbeddedChartHost<'_>, frame: &str) -> Result<String> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut target_depth = None;
    let mut target_end = None;
    let mut shapes_depth = None;
    let mut shapes_end = None;
    let mut matches = 0usize;
    loop {
        let event_start = position(&reader)?;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid embedded chart host XML: {error}"))
        })?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element) => {
                if host_matches(&reader, namespace, &element, &host)? {
                    matches += 1;
                    if matches > 1 { return invalid("embedded chart host name is ambiguous"); }
                    target_depth = Some(depth);
                } else if matches == 1
                    && target_depth.is_some_and(|value| depth == value + 1)
                    && namespace == NamespaceKind::Table
                    && element.local_name().as_ref() == b"shapes"
                {
                    shapes_depth = Some(depth);
                }
                depth = depth.checked_add(1).ok_or_else(|| Error::InvalidFormat("XML depth overflow".to_string()))?;
            },
            Event::Empty(element)
                if host_matches(&reader, namespace, &element, &host)? => {
                    return invalid("embedded chart host must not be self-closing");
                },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| Error::InvalidFormat("XML depth underflow".to_string()))?;
                if shapes_depth == Some(depth)
                    && namespace == NamespaceKind::Table
                    && element.local_name().as_ref() == b"shapes"
                {
                    shapes_end = Some(event_start);
                    shapes_depth = None;
                }
                if target_depth == Some(depth) {
                    target_end = Some(event_start);
                    target_depth = None;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if matches != 1 { return invalid("embedded chart host was not found"); }
    match host {
        EmbeddedChartHost::Sheet(_) => {
            if let Some(position) = shapes_end {
                splice(xml, position, position, frame)
            } else {
                let wrapper = format!("<table:shapes xmlns:table=\"{TABLE_URI}\">{frame}</table:shapes>");
                let position = target_end.ok_or_else(|| Error::InvalidFormat("sheet host end is missing".to_string()))?;
                splice(xml, position, position, &wrapper)
            }
        },
        EmbeddedChartHost::Text => {
            let wrapper = format!("<text:p xmlns:text=\"{TEXT_URI}\">{frame}</text:p>");
            let position = target_end.ok_or_else(|| Error::InvalidFormat("text host end is missing".to_string()))?;
            splice(xml, position, position, &wrapper)
        },
        EmbeddedChartHost::Page(_) => {
            let position = target_end.ok_or_else(|| Error::InvalidFormat("page host end is missing".to_string()))?;
            splice(xml, position, position, frame)
        },
    }
}

fn host_matches(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    host: &EmbeddedChartHost<'_>,
) -> Result<bool> {
    match host {
        EmbeddedChartHost::Text => Ok(namespace == NamespaceKind::Office && element.local_name().as_ref() == b"text"),
        EmbeddedChartHost::Sheet(name) => Ok(namespace == NamespaceKind::Table
            && element.local_name().as_ref() == b"table"
            && attribute(reader, element, TABLE_NS, b"name")?.as_deref() == Some(*name)),
        EmbeddedChartHost::Page(name) => Ok(namespace == NamespaceKind::Draw
            && element.local_name().as_ref() == b"page"
            && attribute(reader, element, DRAW_NS, b"name")?.as_deref() == Some(*name)),
    }
}

fn attribute(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, ns: &[u8], local: &[u8]) -> Result<Option<String>> {
    namespaced_attribute(reader, element, ns, local, "embedded chart host")
}
fn is_object(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Draw
        && matches!(local, b"object" | b"object-ole" | b"applet" | b"plugin" | b"floating-frame")
}
fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(value)) if *value == TABLE_NS => NamespaceKind::Table,
        _ => NamespaceKind::Other,
    }
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| Error::InvalidFormat("XML position exceeds platform limits".to_string()))
}
pub(crate) fn splice(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end) {
        return invalid("invalid embedded chart XML splice range");
    }
    let mut out = String::with_capacity(xml.len() - (end - start) + replacement.len());
    out.push_str(&xml[..start]); out.push_str(replacement); out.push_str(&xml[end..]); Ok(out)
}
fn invalid<T>(message: impl Into<String>) -> Result<T> { Err(Error::InvalidFormat(message.into())) }
