//! ODS chart discovery, inline-root adaptation, and bounded XML spans.

use super::model::{Chart, Limits, Location, Part, Storage};
use crate::package::Package;
use litchi_core::{Error, Result};
use litchi_odf_common::chart::authoring::{Definition, serialize_content};
use litchi_odf_common::chart::read;
use litchi_odf_common::constants::{ODF_CHART, ODF_CHART_TEMPLATE};
use litchi_odf_common::drawing::Part as DrawingPart;
use litchi_odf_common::embedded::{Kind, Root, Source, scan_package};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::Arc;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

/// Read all content-level embedded chart occurrences from an ODS package.
pub(crate) fn inventory(package: &Package, limits: Limits) -> Result<Vec<Chart>> {
    let owned = package.package();
    let borrowed = owned.package()?;
    let objects = scan_package(package.content_xml(), None, &borrowed)?;
    let spans = locate_objects(package.content_xml())?;
    let mut content_objects = 0usize;
    let mut total_bytes = 0usize;
    let mut charts = Vec::new();

    for object in objects {
        if object.part != DrawingPart::Content {
            continue;
        }
        let span = spans.get(content_objects).ok_or_else(|| {
            invalid_error("ODS chart object scanner and XML span scanner disagree")
        })?;
        content_objects = content_objects
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS chart object count overflow"))?;

        if !is_chart_object(&object) {
            continue;
        }
        if charts.len() >= limits.max_charts() {
            return invalid(format!(
                "ODS embedded chart count exceeds {}",
                limits.max_charts()
            ));
        }

        let (part, storage, location) = match object.source {
            Source::PackageSubdocument {
                content_path,
                manifest_media_type,
                ..
            } => {
                if !is_chart_media_type(manifest_media_type.as_deref()) {
                    continue;
                }
                let bytes = owned.get_file(&content_path)?;
                let xml = String::from_utf8(bytes).map_err(|_error| {
                    invalid_error("embedded ODS chart content.xml is not UTF-8")
                })?;
                let part = Part::from_xml_with_limit(xml, limits.max_part_bytes())?;
                (
                    part,
                    Storage::PackageSubdocument,
                    Location::Package { content_path },
                )
            },
            Source::InlineXml {
                root: Root::OpenDocument,
                xml,
                ..
            } => {
                let Some(media_type) = inline_mimetype(&xml)? else {
                    continue;
                };
                if !is_chart_media_type(Some(media_type.as_str())) {
                    continue;
                }
                let part = Part::from_inline_with_limit(xml, limits.max_part_bytes())?;
                let (payload_start, payload_end) = span
                    .payload
                    .ok_or_else(|| invalid_error("inline ODS chart payload span is missing"))?;
                (
                    part,
                    Storage::InlineXml,
                    Location::Inline {
                        payload_start,
                        payload_end,
                    },
                )
            },
            Source::InlineXml { .. }
            | Source::InlineBinary { .. }
            | Source::PackageFile { .. }
            | Source::MissingPackagePart { .. }
            | Source::Linked { .. }
            | Source::Missing
            | _ => continue,
        };

        total_bytes = total_bytes
            .checked_add(part.xml().len())
            .ok_or_else(|| invalid_error("ODS chart byte count overflow"))?;
        if total_bytes > limits.max_total_bytes() {
            return invalid(format!(
                "ODS embedded chart content exceeds {} bytes",
                limits.max_total_bytes()
            ));
        }
        charts.push(Chart {
            frame: object.frame,
            storage,
            part,
            location,
        });
    }

    if content_objects != spans.len() {
        return invalid("ODS chart object scanner found an inconsistent object count");
    }
    Ok(charts)
}

impl Part {
    /// Parse an existing standalone chart `content.xml` part.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_xml(xml: impl Into<String>) -> Result<Self> {
        Self::from_xml_with_limit(xml.into(), Limits::default().max_part_bytes())
    }

    /// Serialize and validate a typed common ODF chart definition.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_definition(definition: &Definition) -> Result<Self> {
        Self::from_xml(serialize_content(definition)?)
    }

    /// Parse an inline `office:document` chart payload.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_inline_xml(xml: impl Into<String>) -> Result<Self> {
        Self::from_inline_with_limit(xml.into(), Limits::default().max_part_bytes())
    }

    fn from_xml_with_limit(xml: String, max_bytes: usize) -> Result<Self> {
        if xml.is_empty() || xml.len() > max_bytes {
            return invalid("ODS chart content is empty or exceeds its byte limit");
        }
        let chart = read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml),
            chart: Arc::new(chart),
        })
    }

    fn from_inline_with_limit(xml: String, max_bytes: usize) -> Result<Self> {
        let Some(media_type) = inline_mimetype(&xml)? else {
            return invalid("inline ODS chart has no office:mimetype");
        };
        if !is_chart_media_type(Some(media_type.as_str())) {
            return invalid("inline ODS object is not an OpenDocument chart");
        }
        let content = rename_document_root(&xml, "document", "document-content", None)?;
        Self::from_xml_with_limit(content, max_bytes)
    }
}

fn is_chart_object(object: &litchi_odf_common::embedded::Object) -> bool {
    object.part == DrawingPart::Content
        && object.kind == Kind::Object
        && object.class_id.is_none()
        && object.code.is_none()
        && object.archive.is_none()
        && object.may_script.is_none()
        && object.applet_name.is_none()
        && object.mime_type.is_none()
        && object.parameters.is_empty()
        && object
            .link_type
            .as_deref()
            .is_none_or(|value| value == "simple")
        && object.show.as_deref().is_none_or(|value| value == "embed")
        && object
            .actuate
            .as_deref()
            .is_none_or(|value| value == "onLoad")
}

fn is_chart_media_type(value: Option<&str>) -> bool {
    matches!(value, Some(ODF_CHART | ODF_CHART_TEMPLATE))
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct Span {
    pub(crate) payload: Option<(usize, usize)>,
}

/// Locate direct embedded-object payload spans without rebuilding the host XML.
pub(crate) fn locate_objects(xml: &str) -> Result<Vec<Span>> {
    struct Active {
        depth: usize,
        payload: Option<(usize, usize, usize)>,
    }

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active = None;
    let mut spans = Vec::new();

    enum Token {
        Start {
            object: bool,
            document: bool,
        },
        Empty {
            object: bool,
            document: bool,
        },
        End {
            kind: NamespaceKind,
            object: bool,
            document: bool,
        },
        Eof,
        Other,
    }

    loop {
        let start = position(&reader)?;
        let token = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid_error(format!("invalid ODS drawing XML: {error}")))?;
            let kind = namespace_kind(&namespace);
            match event {
                Event::Start(element) => Token::Start {
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::Empty(element) => Token::Empty {
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::End(element) => Token::End {
                    kind,
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::Eof => Token::Eof,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => Token::Other,
            }
        };
        let end = position(&reader)?;
        match token {
            Token::Start { object, document } => {
                if object {
                    if active.is_some() {
                        return invalid("nested ODS embedded objects are not supported");
                    }
                    active = Some(Active {
                        depth,
                        payload: None,
                    });
                } else if let Some(object) = active.as_mut()
                    && depth == object.depth + 1
                    && document
                {
                    object.payload = Some((depth, start, 0));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS drawing XML depth overflow"))?;
            },
            Token::Empty { object, document } => {
                if object {
                    if active.is_some() {
                        return invalid("nested ODS embedded objects are not supported");
                    }
                    spans.push(Span { payload: None });
                } else if active.is_some()
                    && depth
                        == active
                            .as_ref()
                            .map(|object| object.depth + 1)
                            .ok_or_else(|| invalid_error("ODS embedded object state disappeared"))?
                    && document
                {
                    return invalid("inline ODS chart document cannot be empty");
                }
            },
            Token::End {
                kind,
                object,
                document,
            } => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODS drawing XML depth underflow"))?;
                if let Some(object) = active.as_mut()
                    && object
                        .payload
                        .is_some_and(|(payload_depth, _, _)| payload_depth == depth)
                    && document
                    && let Some((payload_depth, payload_start, _)) = object.payload
                {
                    object.payload = Some((payload_depth, payload_start, end));
                }
                if active.as_ref().is_some_and(|object| object.depth == depth)
                    && kind == NamespaceKind::Draw
                    && object
                {
                    let object = active
                        .take()
                        .ok_or_else(|| invalid_error("ODS embedded object state disappeared"))?;
                    spans.push(Span {
                        payload: object
                            .payload
                            .and_then(|(_, start, end)| (end != 0).then_some((start, end))),
                    });
                }
            },
            Token::Eof => break,
            Token::Other => {},
        }
        buffer.clear();
    }

    if active.is_some() || depth != 0 {
        return invalid("unterminated ODS embedded-object XML");
    }
    Ok(spans)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Other,
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn is_object_name(local: &[u8]) -> bool {
    matches!(
        local,
        b"object" | b"object-ole" | b"applet" | b"plugin" | b"floating-frame"
    )
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("ODS XML position exceeds platform limits"))
}

fn inline_mimetype(xml: &str) -> Result<Option<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid inline ODS object XML: {error}")))?;
        match event {
            Event::Start(element)
                if namespace_kind(&namespace) == NamespaceKind::Office
                    && element.local_name().as_ref() == b"document" =>
            {
                let mut value = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        invalid_error(format!("invalid inline ODS object attribute: {error}"))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == *OFFICE_NS)
                        && local.as_ref() == b"mimetype"
                    {
                        if value.is_some() {
                            return invalid("duplicate inline office:mimetype");
                        }
                        value = Some(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| {
                                    invalid_error(format!(
                                        "invalid inline office:mimetype value: {error}"
                                    ))
                                })?
                                .into_owned(),
                        );
                    }
                }
                return Ok(value);
            },
            Event::Empty(_) => return invalid("inline ODS object document root cannot be empty"),
            Event::Text(value) if value.iter().all(u8::is_ascii_whitespace) => {},
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("inline ODS object has no office:document root"),
            Event::Start(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return invalid("inline ODS object has invalid content before its root");
            },
        }
        buffer.clear();
    }
}

fn rename_document_root(
    xml: &str,
    expected_local: &str,
    replacement_local: &str,
    added_attribute: Option<(&str, &str)>,
) -> Result<String> {
    let mut root_start = xml
        .find('<')
        .ok_or_else(|| invalid_error("ODS chart XML has no root"))?;
    if xml[root_start..].starts_with("<?xml") {
        let declaration_end = xml[root_start..]
            .find("?>")
            .ok_or_else(|| invalid_error("unterminated ODS chart XML declaration"))?
            + root_start
            + 2;
        root_start = xml[declaration_end..]
            .find('<')
            .map(|offset| declaration_end + offset)
            .ok_or_else(|| invalid_error("ODS chart XML has no document root"))?;
    }
    let name_end = xml[root_start + 1..]
        .find(|character: char| character.is_whitespace() || character == '>' || character == '/')
        .map(|offset| root_start + 1 + offset)
        .ok_or_else(|| invalid_error("invalid ODS chart root start tag"))?;
    let qname = &xml[root_start + 1..name_end];
    let (prefix, local) = qname.rsplit_once(':').unwrap_or(("", qname));
    if local != expected_local || prefix.is_empty() {
        return invalid(format!("expected office:{expected_local} chart root"));
    }
    let close_start = xml
        .rfind("</")
        .ok_or_else(|| invalid_error("ODS chart root is not closed"))?;
    let close_name_end = xml[close_start + 2..]
        .find('>')
        .map(|offset| close_start + 2 + offset)
        .ok_or_else(|| invalid_error("invalid ODS chart root closing tag"))?;
    if xml[close_start + 2..close_name_end].trim() != qname {
        return invalid("ODS chart root start/end names do not match");
    }
    let replacement = format!("{prefix}:{replacement_local}");
    let mut output = String::with_capacity(xml.len() + 96);
    output.push_str(&xml[..=root_start]);
    output.push_str(&replacement);
    if let Some((name, value)) = added_attribute {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(value);
        output.push('"');
    }
    output.push_str(&xml[name_end..close_start + 2]);
    output.push_str(&replacement);
    output.push_str(&xml[close_name_end..]);
    Ok(output)
}

#[cfg(test)]
pub(crate) fn inline_content(xml: &str) -> Result<String> {
    rename_document_root(xml, "document", "document-content", None)
}

pub(crate) fn content_inline(xml: &str) -> Result<String> {
    let mimetype = (!root_has_office_mimetype(xml)?).then_some(("office:mimetype", ODF_CHART));
    rename_document_root(xml, "document-content", "document", mimetype)
}

fn root_has_office_mimetype(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS chart XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let mut found = false;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        invalid_error(format!("invalid ODS chart root attribute: {error}"))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == *OFFICE_NS)
                        && local.as_ref() == b"mimetype"
                    {
                        if found {
                            return invalid("duplicate inline office:mimetype");
                        }
                        found = true;
                    }
                }
                return Ok(found);
            },
            Event::Empty(_) => return invalid("ODS chart document root cannot be empty"),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(value) if value.iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return invalid("ODS chart XML has no document root"),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return invalid("ODS chart XML has invalid content before its root");
            },
        }
        buffer.clear();
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
