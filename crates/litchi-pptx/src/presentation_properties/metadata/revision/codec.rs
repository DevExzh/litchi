//! `PowerPoint` 2016 Revision Information parts.
//!
//! Revision extension payloads are retained as XML and never interpreted or
//! used to resolve relationships.

use super::model::{Client, Info, Namespace, Part};
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub const CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.revisioninfo+xml";
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo";

const P1510: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2015/10/main";
const P1510_TEXT: &str = "http://schemas.microsoft.com/office/powerpoint/2015/10/main";
const P: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_CLIENTS: usize = 65_536;
const MAX_EXTENSIONS: usize = 4_096;
const MAX_STRING_BYTES: usize = 1024 * 1024;

impl Info {
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        parse_revision_information(xml)
    }

    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_model(self)?;
        let root_prefix = unused_root_prefix(&self.namespace_declarations);
        let mut out = Vec::new();
        out.extend_from_slice(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        out.extend_from_slice(b"<");
        out.extend_from_slice(root_prefix.as_bytes());
        out.extend_from_slice(b":revInfo xmlns:");
        out.extend_from_slice(root_prefix.as_bytes());
        out.extend_from_slice(b"=\"");
        escape(&mut out, P1510_TEXT);
        out.push(b'"');
        for declaration in &self.namespace_declarations {
            out.extend_from_slice(b" xmlns");
            if !declaration.prefix.is_empty() {
                out.push(b':');
                out.extend_from_slice(declaration.prefix.as_bytes());
            }
            out.extend_from_slice(b"=\"");
            escape(&mut out, &declaration.uri);
            out.push(b'"');
        }
        if self.clients.is_empty() && self.extension_xml.is_none() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            if !self.clients.is_empty() {
                out.extend_from_slice(b"<");
                out.extend_from_slice(root_prefix.as_bytes());
                out.extend_from_slice(b":revLst>");
                for client in &self.clients {
                    out.extend_from_slice(b"<");
                    out.extend_from_slice(root_prefix.as_bytes());
                    out.extend_from_slice(b":client id=\"");
                    escape(&mut out, &client.client_id);
                    out.extend_from_slice(b"\" dt=\"");
                    escape(&mut out, &client.date_time);
                    out.push(b'"');
                    if let Some(value) = client.revision {
                        write_u32_attr(&mut out, "v", value);
                    }
                    if let Some(value) = client.wet_revision {
                        write_u32_attr(&mut out, "vWet", value);
                    }
                    out.extend_from_slice(b"/>");
                }
                out.extend_from_slice(b"</");
                out.extend_from_slice(root_prefix.as_bytes());
                out.extend_from_slice(b":revLst>");
            }
            if let Some(extension) = &self.extension_xml {
                out.extend_from_slice(extension);
            }
            out.extend_from_slice(b"</");
            out.extend_from_slice(root_prefix.as_bytes());
            out.extend_from_slice(b":revInfo>");
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized Revision Information bytes"));
        }
        // Validate programmatically supplied opaque XML in its inherited context.
        parse_revision_information(&out)?;
        Ok(out)
    }
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage) -> Result<Option<Part>> {
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().to_string();

    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "Revision Information relationship cannot originate at the package root",
        ));
    }
    for source in package.iter_parts() {
        if source.partname().as_str() != presentation_name.as_str()
            && source
                .rels()
                .iter()
                .any(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
        {
            return Err(invalid(
                "Revision Information relationship has a non-Presentation source",
            ));
        }
    }

    let relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
        .collect();
    if relationships.len() > 1 {
        return Err(invalid(
            "Presentation has multiple Revision Information relationships",
        ));
    }
    let Some(relationship) = relationships.first().copied() else {
        if package
            .iter_parts()
            .any(|part| part.content_type() == CONTENT_TYPE)
        {
            return Err(invalid(
                "package contains an orphan Revision Information part",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid(
            "Revision Information relationship cannot be external",
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.into(),
            actual: part.content_type().into(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Revision Information part cannot have outbound relationships",
        ));
    }
    if package.iter_parts().any(|candidate| {
        candidate.content_type() == CONTENT_TYPE && candidate.partname() != &target
    }) {
        return Err(invalid(
            "package contains an orphan Revision Information part",
        ));
    }
    Ok(Some(Part {
        relationship_id: relationship.r_id().to_string(),
        part_name: target.to_string(),
        revision_information: Info::parse(part.blob())?,
    }))
}

/// Add a new Revision Information part after validating the complete graph.
/// Existing Revision Information parts are deliberately not overwritten.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store(package: &mut OpcPackage, value: &Part) -> Result<()> {
    if load(package)?.is_some() {
        return Err(invalid("package already contains Revision Information"));
    }
    validate_relationship_id(&value.relationship_id)?;
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    if presentation.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(
            "Revision Information relationship ID already exists",
        ));
    }
    let presentation_name = presentation.partname().clone();
    let part_name = PackURI::new(&value.part_name).map_err(Error::Uri)?;
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!("part '{part_name}' already exists")));
    }
    let xml = value.revision_information.to_xml()?;
    let target = part_name.relative_ref(presentation_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(part_name, CONTENT_TYPE.into(), xml)))?;
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

#[derive(Debug)]
enum Frame {
    Root,
    RevisionList,
    Client,
    ExtensionList,
    Extension { payloads: u8 },
    Payload,
    Opaque,
}

fn parse_revision_information(xml: &[u8]) -> Result<Info> {
    if xml.len() > MAX_BYTES {
        return Err(limit("Revision Information part bytes"));
    }
    let selected = process_ooxml(xml)?;
    if selected.len() > MAX_BYTES {
        return Err(limit("MCE-processed Revision Information bytes"));
    }
    let bytes = selected.as_ref();
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut seen_list = false;
    let mut seen_extensions = false;
    let mut clients = Vec::new();
    let mut namespaces = Vec::new();
    let mut extension_xml = None;
    let mut extension_start = None;
    let mut extension_finish_pending = None;
    let mut extension_count = 0usize;
    let mut node_count = 0usize;

    loop {
        let start = reader.buffer_position() as usize;
        if let Some(from) = extension_finish_pending.take() {
            extension_xml = Some(bytes[from..start].to_vec());
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                node_count = node_count
                    .checked_add(1)
                    .ok_or_else(|| limit("Revision Information nodes"))?;
                if node_count > MAX_NODES {
                    return Err(limit("Revision Information nodes"));
                }
                let local = element.local_name();
                let frame = if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("Revision Information has multiple roots"));
                    }
                    expect_namespace_name(&namespace, P1510, local.as_ref(), b"revInfo")?;
                    let root_prefix = element_prefix(&element)?;
                    namespaces = root_namespaces(&element, reader.decoder(), &root_prefix)?;
                    root_seen = true;
                    Frame::Root
                } else {
                    match stack.last_mut().expect("nonempty stack") {
                        Frame::Root => {
                            if namespace_name(&namespace, P1510, local.as_ref(), b"revLst") {
                                if seen_list || seen_extensions {
                                    return Err(invalid("revLst is duplicated or out of order"));
                                }
                                no_attributes(&element, reader.decoder())?;
                                seen_list = true;
                                Frame::RevisionList
                            } else if namespace_name(&namespace, P, local.as_ref(), b"extLst") {
                                if seen_extensions {
                                    return Err(invalid(
                                        "Revision Information has duplicate extLst",
                                    ));
                                }
                                no_attributes(&element, reader.decoder())?;
                                seen_extensions = true;
                                extension_start = Some(start);
                                Frame::ExtensionList
                            } else {
                                return Err(invalid("unexpected Revision Information child"));
                            }
                        },
                        Frame::RevisionList => {
                            expect_namespace_name(&namespace, P1510, local.as_ref(), b"client")?;
                            if clients.len() >= MAX_CLIENTS {
                                return Err(limit("Revision Information clients"));
                            }
                            clients.push(parse_client(&element, reader.decoder())?);
                            Frame::Client
                        },
                        Frame::Client => {
                            return Err(invalid("client revision must be empty"));
                        },
                        Frame::ExtensionList => {
                            expect_namespace_name(&namespace, P, local.as_ref(), b"ext")?;
                            extension_count += 1;
                            if extension_count > MAX_EXTENSIONS {
                                return Err(limit("Revision Information extensions"));
                            }
                            extension_attributes(&element, reader.decoder())?;
                            Frame::Extension { payloads: 0 }
                        },
                        Frame::Extension { payloads } => {
                            if *payloads != 0 || !other_namespace(&namespace) {
                                return Err(invalid(
                                    "p:ext requires exactly one foreign-namespace payload",
                                ));
                            }
                            *payloads = 1;
                            validate_any_attributes(&element, reader.decoder())?;
                            Frame::Payload
                        },
                        Frame::Payload | Frame::Opaque => {
                            validate_any_attributes(&element, reader.decoder())?;
                            Frame::Opaque
                        },
                    }
                };

                if stack.len() + 1 > MAX_DEPTH {
                    return Err(limit("Revision Information XML depth"));
                }
                if empty_event {
                    close_empty_frame(&frame)?;
                    if matches!(frame, Frame::ExtensionList) {
                        let from = extension_start
                            .take()
                            .ok_or_else(|| invalid("missing extLst start offset"))?;
                        extension_finish_pending = Some(from);
                    }
                    if matches!(frame, Frame::Root) {
                        root_closed = true;
                    }
                } else {
                    stack.push(frame);
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected Revision Information closing element"))?;
                validate_end(&namespace, element.local_name().as_ref(), &frame)?;
                match &frame {
                    Frame::Extension { payloads } if *payloads != 1 => {
                        return Err(invalid(
                            "p:ext requires exactly one foreign-namespace payload",
                        ));
                    },
                    Frame::ExtensionList => {
                        let from = extension_start
                            .take()
                            .ok_or_else(|| invalid("missing extLst start offset"))?;
                        extension_finish_pending = Some(from);
                    },
                    Frame::Root => root_closed = true,
                    _ => {},
                }
            },
            Event::Text(text) => {
                if !matches!(stack.last(), Some(Frame::Payload | Frame::Opaque)) {
                    let decoded = text.decode().map_err(xml_error)?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !unescaped.trim().is_empty() {
                        return Err(invalid("unexpected Revision Information text"));
                    }
                }
            },
            Event::CData(text) => {
                if !matches!(stack.last(), Some(Frame::Payload | Frame::Opaque))
                    && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid("unexpected Revision Information CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD, processing instructions, and general references are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() || extension_start.is_some() {
        return Err(invalid("unterminated Revision Information part"));
    }
    let value = Info {
        clients,
        namespace_declarations: namespaces,
        extension_xml,
    };
    validate_model(&value)?;
    Ok(value)
}

fn parse_client(element: &BytesStart<'_>, decoder: Decoder) -> Result<Client> {
    let attributes = known_attributes(element, decoder, &["id", "v", "vWet", "dt"])?;
    let required = |name: &str| {
        attributes
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone())
            .ok_or_else(|| invalid(format!("client revision is missing '{name}'")))
    };
    let client_id = required("id")?;
    let date_time = required("dt")?;
    bounded(&client_id)?;
    validate_date_time(&date_time)?;
    Ok(Client {
        client_id,
        revision: optional_u32(&attributes, "v")?,
        wet_revision: optional_u32(&attributes, "vWet")?,
        date_time,
    })
}

fn optional_u32(attributes: &[(String, String)], name: &str) -> Result<Option<u32>> {
    attributes
        .iter()
        .find(|(key, _)| key == name)
        .map(|(_, value)| {
            value
                .parse()
                .map_err(|_err| invalid(format!("invalid unsigned client revision '{value}'")))
        })
        .transpose()
}

fn root_namespaces(
    element: &BytesStart<'_>,
    decoder: Decoder,
    root_prefix: &str,
) -> Result<Vec<Namespace>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if key != "xmlns" && !key.starts_with("xmlns:") {
            return Err(invalid(format!("unexpected revInfo attribute '{key}'")));
        }
        let prefix = key.strip_prefix("xmlns:").unwrap_or("").to_string();
        if !seen.insert(prefix.clone()) {
            return Err(invalid("duplicate root namespace declaration"));
        }
        let uri = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&uri)?;
        if prefix != root_prefix {
            output.push(Namespace { prefix, uri });
        }
    }
    Ok(output)
}

fn known_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    known: &[&str],
) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        if key.contains(':') || !known.contains(&key) || !seen.insert(key.to_string()) {
            return Err(invalid(format!(
                "unexpected or duplicate attribute '{key}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&value)?;
        output.push((key.to_string(), value));
    }
    Ok(output)
}

fn no_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    known_attributes(element, decoder, &[]).map(|_| ())
}

fn extension_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    let values = known_attributes(element, decoder, &["uri"])?;
    let uri = values
        .iter()
        .find(|(key, _)| key == "uri")
        .map(|(_, value)| value.as_str())
        .ok_or_else(|| invalid("p:ext is missing required uri"))?;
    if uri.trim().is_empty() {
        return Err(invalid("p:ext uri cannot be empty"));
    }
    Ok(())
}

fn validate_any_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        bounded(&value)?;
    }
    Ok(())
}

fn close_empty_frame(frame: &Frame) -> Result<()> {
    if matches!(frame, Frame::Extension { .. }) {
        Err(invalid(
            "p:ext requires exactly one foreign-namespace payload",
        ))
    } else {
        Ok(())
    }
}

fn validate_end(namespace: &ResolveResult<'_>, local: &[u8], frame: &Frame) -> Result<()> {
    match frame {
        Frame::Root => expect_namespace_name(namespace, P1510, local, b"revInfo"),
        Frame::RevisionList => expect_namespace_name(namespace, P1510, local, b"revLst"),
        Frame::Client => expect_namespace_name(namespace, P1510, local, b"client"),
        Frame::ExtensionList => expect_namespace_name(namespace, P, local, b"extLst"),
        Frame::Extension { .. } => expect_namespace_name(namespace, P, local, b"ext"),
        Frame::Payload | Frame::Opaque => Ok(()),
    }
}

fn validate_model(value: &Info) -> Result<()> {
    if value.clients.len() > MAX_CLIENTS {
        return Err(limit("Revision Information clients"));
    }
    for client in &value.clients {
        bounded(&client.client_id)?;
        validate_date_time(&client.date_time)?;
    }
    let mut prefixes = HashSet::new();
    for declaration in &value.namespace_declarations {
        if (!declaration.prefix.is_empty() && !ncname(&declaration.prefix))
            || matches!(declaration.prefix.as_str(), "xml" | "xmlns")
            || declaration.uri.is_empty()
            || !prefixes.insert(declaration.prefix.as_str())
        {
            return Err(invalid(
                "invalid or duplicate preserved namespace declaration",
            ));
        }
        bounded(&declaration.uri)?;
    }
    if value
        .extension_xml
        .as_ref()
        .is_some_and(|extension| extension.len() > MAX_BYTES)
    {
        return Err(limit("Revision Information extension bytes"));
    }
    Ok(())
}

fn validate_date_time(value: &str) -> Result<()> {
    bounded(value)?;
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{value}'")))
    }
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if ncname(value) {
        Ok(())
    } else {
        Err(invalid(
            "Revision Information relationship ID is not an XML NCName",
        ))
    }
}

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main document has non-Presentation content type '{value}'"
        )))
    }
}

fn element_prefix(element: &BytesStart<'_>) -> Result<String> {
    let name = element.name();
    let qualified = std::str::from_utf8(name.as_ref()).map_err(xml_error)?;
    Ok(qualified
        .rsplit_once(':')
        .map_or("", |(prefix, _)| prefix)
        .to_string())
}

fn unused_root_prefix(namespaces: &[Namespace]) -> String {
    let used: HashSet<_> = namespaces
        .iter()
        .map(|value| value.prefix.as_str())
        .collect();
    let mut prefix = "p1510".to_string();
    while used.contains(prefix.as_str()) {
        prefix.push('r');
    }
    prefix
}

fn namespace_name(
    namespace: &ResolveResult<'_>,
    expected_namespace: &[u8],
    local: &[u8],
    expected_local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected_namespace)
        && local == expected_local
}

fn expect_namespace_name(
    namespace: &ResolveResult<'_>,
    expected_namespace: &[u8],
    local: &[u8],
    expected_local: &[u8],
) -> Result<()> {
    if namespace_name(namespace, expected_namespace, local, expected_local) {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected namespaced element '{}'",
            String::from_utf8_lossy(expected_local)
        )))
    }
}

fn other_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() != P)
}

fn ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
}

fn bounded(value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        Err(limit("Revision Information string bytes"))
    } else {
        Ok(())
    }
}

fn write_u32_attr(output: &mut Vec<u8>, name: &str, value: u32) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(value.to_string().as_bytes());
    output.push(b'"');
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(label: &str) -> Error {
    invalid(format!("{label} exceed configured limit"))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const PRESENTATION_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    fn base_package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            PRESENTATION_CONTENT_TYPE.into(),
            b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>".to_vec(),
        )));
        package
    }

    fn value() -> Part {
        Part {
            relationship_id: "rIdRevision".into(),
            part_name: "/ppt/revisionInfo.xml".into(),
            revision_information: Info {
                clients: vec![Client {
                    client_id: "{793478A3-5DEC-486D-815F-463E73860F83}".into(),
                    revision: Some(12),
                    wet_revision: None,
                    date_time: "2021-07-26T09:34:00.336".into(),
                }],
                namespace_declarations: vec![Namespace {
                    prefix: "p".into(),
                    uri: "http://schemas.openxmlformats.org/presentationml/2006/main".into(),
                }],
                extension_xml: Some(br#"<p:extLst><p:ext uri="urn:producer"><v:payload xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst>"#.to_vec()),
            },
        }
    }

    #[test]
    fn loads_real_powerpoint_and_libreoffice_revision_parts() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let powerpoint =
            OpcPackage::open(root.join("test-data/poi/test-data/slideshow/bug65551.pptx")).unwrap();
        let loaded = load(&powerpoint).unwrap().unwrap();
        assert_eq!(loaded.part_name, "/ppt/revisionInfo.xml");
        assert_eq!(loaded.revision_information.clients.len(), 2);
        assert_eq!(loaded.revision_information.clients[0].revision, Some(12));
        assert_eq!(loaded.revision_information.clients[1].revision, Some(18));

        let libreoffice = OpcPackage::open(
            root.join("test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf114821.pptx"),
        )
        .unwrap();
        let loaded = load(&libreoffice).unwrap().unwrap();
        assert!(loaded.revision_information.clients.is_empty());
    }

    #[test]
    fn package_store_load_round_trip_keeps_extensions_inert() {
        let expected = value();
        let mut package = base_package();
        store(&mut package, &expected).unwrap();
        let loaded = load(&package).unwrap().unwrap();
        assert_eq!(loaded, expected);
        let text = String::from_utf8(loaded.revision_information.to_xml().unwrap()).unwrap();
        assert!(text.contains("rIdNeverFetched"));
        assert!(text.contains("https://example.invalid/not-opened"));
    }

    #[test]
    fn rejects_hostile_revision_grammar() {
        let wrap = |body: &str| {
            format!(
                r#"<r:revInfo xmlns:r="{P1510_TEXT}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{body}</r:revInfo>"#
            )
        };
        let cases = [
            wrap("<r:client id=\"x\" dt=\"2020-01-01T00:00:00\"/>"),
            wrap("<r:revLst><r:client id=\"x\"/></r:revLst>"),
            wrap("<r:revLst><r:client id=\"x\" dt=\"bad\"/></r:revLst>"),
            wrap("<r:revLst><r:client id=\"x\" dt=\"2020-01-01T00:00:00\" v=\"-1\"/></r:revLst>"),
            wrap("<p:extLst/><r:revLst/>"),
            wrap("<p:extLst><p:ext uri=\"urn:x\"/></p:extLst>"),
            wrap("<p:extLst><p:ext uri=\"urn:x\"><p:wrong/></p:ext></p:extLst>"),
            wrap(
                "<p:extLst><p:ext uri=\"urn:x\"><x:a xmlns:x=\"urn:x\"/><x:b xmlns:x=\"urn:x\"/></p:ext></p:extLst>",
            ),
            wrap("<p:extLst><p:ext><x:a xmlns:x=\"urn:x\"/></p:ext></p:extLst>"),
            format!(r"<!DOCTYPE x>{}", wrap("")),
        ];
        for xml in cases {
            assert!(Info::parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn rejects_invalid_package_graphs_and_preserves_failed_store() {
        let mut external = base_package();
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                RELATIONSHIP_TYPE.into(),
                "https://example.invalid/revision.xml".into(),
                "rIdRevision".into(),
                true,
            );
        assert!(load(&external).is_err());

        let mut orphan = base_package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/revisionInfo.xml").unwrap(),
            CONTENT_TYPE.into(),
            Info::default().to_xml().unwrap(),
        )));
        assert!(load(&orphan).is_err());

        let mut outbound = base_package();
        store(&mut outbound, &value()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/revisionInfo.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.xml".into(),
                "rId1".into(),
                false,
            );
        assert!(load(&outbound).is_err());

        let mut invalid_value = value();
        invalid_value.revision_information.clients[0].date_time = "not-a-date".into();
        let mut package = base_package();
        let before = package.part_count();
        assert!(store(&mut package, &invalid_value).is_err());
        assert_eq!(package.part_count(), before);
        assert!(load(&package).unwrap().is_none());
    }
}
