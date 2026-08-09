//! Namespace-aware XML reader and deterministic writer for `model3d`.

use std::io::Write;

use litchi_ooxml_common::{relationships, xml::unqualified_attribute_value};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Prefix, QName, ResolveResult},
    reader::{NsReader, Reader},
};

use crate::{Error, Result};

use super::super::{
    Blip, Child, Id, Inert, MAX_CHILDREN, MAX_DEPTH, MAX_FRAGMENT_BYTES, MAX_NODES, Metadata,
    NAMESPACE, Namespace, RELATIONSHIP_NAMESPACE, RELATIONSHIP_NAMESPACE_STRICT, Raster,
    RasterChild, Reference,
};

/// Read one complete `m3d:model3d` element.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8]) -> Result<Metadata> {
    if xml.len() > super::super::MAX_XML_BYTES {
        return Err(limit("model3d XML bytes", super::super::MAX_XML_BYTES));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut metadata = None;
    let mut root_depth = 0usize;
    let mut root_closed = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d XML offset exceeds usize"))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolved_namespace(namespace)?;
        let event = event.into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d XML offset exceeds usize"))?;

        match event {
            Event::Start(element) if metadata.is_none() && root_depth == 0 => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                require_root(&local, &namespace)?;
                let namespaces = declarations(&element, reader.decoder())?;
                let reference = reference_from_namespace(&element, &reader)?;
                metadata = Some(Metadata {
                    reference,
                    children: Vec::new(),
                    namespaces,
                });
                root_depth = 1;
            },
            Event::Empty(element) if metadata.is_none() && root_depth == 0 => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                require_root(&local, &namespace)?;
                let namespaces = declarations(&element, reader.decoder())?;
                let reference = reference_from_namespace(&element, &reader)?;
                metadata = Some(Metadata {
                    reference,
                    children: Vec::new(),
                    namespaces,
                });
                root_closed = true;
            },
            Event::Start(element) if metadata.is_some() && root_depth == 1 => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                let child_start = event_start;
                let child_end = capture_namespace_element(&mut reader, &mut buffer)?;
                let raw = xml
                    .get(child_start..child_end)
                    .ok_or_else(|| invalid("model3d child range is outside the input"))?;
                let metadata = metadata
                    .as_mut()
                    .ok_or_else(|| invalid("model3d root metadata is missing"))?;
                push_child(metadata, local, namespace, raw.to_vec())?;
            },
            Event::Empty(element) if metadata.is_some() && root_depth == 1 => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                let raw = xml
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("model3d child range is outside the input"))?;
                let metadata = metadata
                    .as_mut()
                    .ok_or_else(|| invalid("model3d root metadata is missing"))?;
                push_child(metadata, local, namespace, raw.to_vec())?;
            },
            Event::End(_) if metadata.is_some() && root_depth == 1 => {
                root_depth = 0;
                root_closed = true;
            },
            Event::Start(_) | Event::Empty(_) if metadata.is_some() && root_depth == 0 => {
                return Err(invalid("model3d XML has more than one root element"));
            },
            Event::Text(text)
                if root_depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("model3d XML contains text outside its root"));
            },
            Event::DocType(_) => return Err(invalid("DTD is forbidden in model3d XML")),
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    let metadata = metadata.ok_or_else(|| invalid("model3d XML has no root element"))?;
    if root_depth != 0 || !root_closed {
        return Err(invalid("model3d XML root is not closed"));
    }
    crate::model3d::validation::validate(&metadata).map_err(validation_error)?;
    Ok(metadata)
}

/// Construct one inert child from a self-contained XML element.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn opaque(xml: &[u8]) -> Result<Inert> {
    if xml.len() > MAX_FRAGMENT_BYTES {
        return Err(limit("model3d inert fragment bytes", MAX_FRAGMENT_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut root_start = 0usize;
    let mut root_end = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d fragment offset exceeds usize"))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolved_namespace(namespace)?;
        let event = event.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d fragment offset exceeds usize"))?;
        match event {
            Event::Start(element) if root.is_none() => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                root = Some((local, namespace));
                root_start = start;
                root_end = capture_namespace_element(&mut reader, &mut buffer)?;
            },
            Event::Empty(element) if root.is_none() => {
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                root = Some((local, namespace));
                root_start = start;
                root_end = end;
            },
            Event::Start(_) | Event::Empty(_) if root.is_some() => {
                return Err(invalid(
                    "model3d inert fragment has more than one root element",
                ));
            },
            Event::Text(text)
                if root.is_none() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("model3d inert fragment has text outside its root"));
            },
            Event::Text(text)
                if root.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("model3d inert fragment has text outside its root"));
            },
            Event::DocType(_) | Event::Decl(_) => {
                return Err(invalid("model3d inert fragment contains document markup"));
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    let (local, namespace) = root.ok_or_else(|| invalid("model3d inert fragment has no root"))?;
    let raw = xml
        .get(root_start..root_end)
        .ok_or_else(|| invalid("model3d inert fragment range is outside the input"))?;
    let value = Inert::from_wire(raw.to_vec(), local, namespace);
    crate::model3d::validation::validate_inert_fragment(&value).map_err(validation_error)?;
    Ok(value)
}

/// Serialize one model3d fragment with stable prefixes and child order.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write(metadata: &Metadata) -> Result<Vec<u8>> {
    crate::model3d::validation::validate(metadata).map_err(validation_error)?;
    let mut output = Vec::new();
    write_inner(&mut output, metadata);
    if output.len() > super::super::MAX_XML_BYTES {
        return Err(limit("model3d output bytes", super::super::MAX_XML_BYTES));
    }
    Ok(output)
}

/// Serialize one model3d fragment to a caller-owned sink.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_to<W: Write>(writer: &mut W, metadata: &Metadata) -> Result<()> {
    let output = write(metadata)?;
    writer.write_all(&output)?;
    Ok(())
}

fn push_child(
    metadata: &mut Metadata,
    local: String,
    namespace: String,
    raw: Vec<u8>,
) -> Result<()> {
    if metadata.children.len() >= MAX_CHILDREN {
        return Err(limit("model3d children", MAX_CHILDREN));
    }
    if local == "raster" && namespace == NAMESPACE {
        metadata
            .children
            .push(Child::Raster(parse_raster(&raw, &metadata.namespaces)?));
    } else {
        let value = Inert::from_wire(raw, local, namespace);
        crate::model3d::validation::validate_inert_fragment(&value).map_err(validation_error)?;
        metadata.children.push(Child::Opaque(value));
    }
    Ok(())
}

fn parse_raster(xml: &[u8], inherited: &[Namespace]) -> Result<Raster> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (root, empty, root_start, root_end) = next_element(&mut reader, &mut buffer)?;
    if root.local_name().as_ref() != b"raster" {
        return Err(invalid("model3d raster fragment has the wrong root"));
    }
    let local_namespaces = declarations_plain(&root, reader.decoder())?;
    let namespaces = merge_namespaces(inherited, &local_namespaces);
    let renderer_name = required_attr(&root, b"rName", reader.decoder(), "raster rName")?;
    let renderer_version = required_attr(&root, b"rVer", reader.decoder(), "raster rVer")?;
    if empty {
        return Raster::from_wire(
            renderer_name,
            renderer_version,
            Vec::new(),
            local_namespaces,
        )
        .map_err(value_error);
    }

    let mut children = Vec::new();
    let mut depth = 1usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d raster offset exceeds usize"))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d raster offset exceeds usize"))?;
        match event {
            Event::Start(element) if depth == 1 => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let child_end = capture_plain_element(&mut reader, &mut buffer)?;
                let raw = xml
                    .get(start..child_end)
                    .ok_or_else(|| invalid("model3d raster child range is outside the input"))?;
                push_raster_child(&mut children, local, namespace, raw.to_vec(), &namespaces)?;
            },
            Event::Empty(element) if depth == 1 => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("model3d raster child range is outside the input"))?;
                push_raster_child(&mut children, local, namespace, raw.to_vec(), &namespaces)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("model3d raster nesting underflow"))?;
                if depth == 0 {
                    break;
                }
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(invalid("model3d raster contains unexpected text"));
            },
            Event::DocType(_) | Event::Decl(_) => {
                return Err(invalid("model3d raster contains document markup"));
            },
            Event::Eof => return Err(invalid("model3d raster is unterminated")),
            Event::Start(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    let _ = (root_start, root_end);
    if children.len() > super::super::MAX_RASTER_CHILDREN {
        return Err(limit(
            "model3d raster children",
            super::super::MAX_RASTER_CHILDREN,
        ));
    }
    Raster::from_wire(renderer_name, renderer_version, children, local_namespaces)
        .map_err(value_error)
}

fn push_raster_child(
    children: &mut Vec<RasterChild>,
    local: String,
    namespace: String,
    raw: Vec<u8>,
    namespaces: &[Namespace],
) -> Result<()> {
    if children.len() >= super::super::MAX_RASTER_CHILDREN {
        return Err(limit(
            "model3d raster children",
            super::super::MAX_RASTER_CHILDREN,
        ));
    }
    if local == "blip" && namespace == NAMESPACE {
        children.push(RasterChild::Blip(parse_blip(&raw, namespaces)?));
    } else {
        let value = Inert::from_wire(raw, local, namespace);
        crate::model3d::validation::validate_inert_fragment(&value).map_err(validation_error)?;
        children.push(RasterChild::Opaque(value));
    }
    Ok(())
}

fn parse_blip(xml: &[u8], inherited: &[Namespace]) -> Result<Blip> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (root, empty, _, _) = next_element(&mut reader, &mut buffer)?;
    if root.local_name().as_ref() != b"blip" {
        return Err(invalid("model3d blip fragment has the wrong root"));
    }
    let local_namespaces = declarations_plain(&root, reader.decoder())?;
    let namespaces = merge_namespaces(inherited, &local_namespaces);
    let reference = reference_plain(&root, &namespaces, reader.decoder())?;
    if empty {
        return Ok(Blip::from_wire(reference, Vec::new(), local_namespaces));
    }

    let mut children = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d blip offset exceeds usize"))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d blip offset exceeds usize"))?;
        match event {
            Event::Start(element) => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let child_end = capture_plain_element(&mut reader, &mut buffer)?;
                let raw = xml
                    .get(start..child_end)
                    .ok_or_else(|| invalid("model3d blip child range is outside the input"))?;
                children.push(Inert::from_wire(raw.to_vec(), local, namespace));
            },
            Event::Empty(element) => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("model3d blip child range is outside the input"))?;
                children.push(Inert::from_wire(raw.to_vec(), local, namespace));
            },
            Event::End(_) => break,
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(invalid("model3d blip contains unexpected text"));
            },
            Event::DocType(_) | Event::Decl(_) => {
                return Err(invalid("model3d blip contains document markup"));
            },
            Event::Eof => return Err(invalid("model3d blip is unterminated")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    Ok(Blip::from_wire(reference, children, local_namespaces))
}

fn write_inner(output: &mut Vec<u8>, metadata: &Metadata) {
    output.extend_from_slice(b"<m3d:model3d");
    let mut has_model = false;
    let mut has_rel = false;
    for declaration in &metadata.namespaces {
        has_model |= declaration.uri() == NAMESPACE;
        has_rel |= declaration.uri() == RELATIONSHIP_NAMESPACE
            || declaration.uri() == RELATIONSHIP_NAMESPACE_STRICT;
        write_namespace(output, declaration);
    }
    if !has_model {
        write_namespace_value(output, Some("m3d"), NAMESPACE);
    }
    if !has_rel {
        write_namespace_value(output, Some("r"), RELATIONSHIP_NAMESPACE);
    }
    write_reference(output, &metadata.reference);
    output.push(b'>');
    for child in &metadata.children {
        match child {
            Child::Opaque(value) => output.extend_from_slice(value.as_bytes()),
            Child::Raster(value) => write_raster(output, value),
        }
    }
    output.extend_from_slice(b"</m3d:model3d>");
}

fn write_raster(output: &mut Vec<u8>, raster: &Raster) {
    output.extend_from_slice(b"<m3d:raster rName=\"");
    push_escaped(output, &raster.renderer_name);
    output.extend_from_slice(b"\" rVer=\"");
    push_escaped(output, &raster.renderer_version);
    output.push(b'"');
    for declaration in &raster.namespaces {
        write_namespace(output, declaration);
    }
    output.push(b'>');
    for child in &raster.children {
        match child {
            RasterChild::Blip(value) => write_blip(output, value),
            RasterChild::Opaque(value) => output.extend_from_slice(value.as_bytes()),
        }
    }
    output.extend_from_slice(b"</m3d:raster>");
}

fn write_blip(output: &mut Vec<u8>, blip: &Blip) {
    output.extend_from_slice(b"<m3d:blip");
    for declaration in &blip.namespaces {
        write_namespace(output, declaration);
    }
    write_reference(output, &blip.reference);
    if blip.children.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        for child in &blip.children {
            output.extend_from_slice(child.as_bytes());
        }
        output.extend_from_slice(b"</m3d:blip>");
    }
}

fn write_reference(output: &mut Vec<u8>, reference: &Reference) {
    if let Some(id) = &reference.embedded {
        output.extend_from_slice(b" r:embed=\"");
        push_escaped(output, id.as_str());
        output.push(b'"');
    }
    if let Some(id) = &reference.linked {
        output.extend_from_slice(b" r:link=\"");
        push_escaped(output, id.as_str());
        output.push(b'"');
    }
}

fn write_namespace(output: &mut Vec<u8>, declaration: &Namespace) {
    write_namespace_value(output, declaration.prefix(), declaration.uri());
}

fn write_namespace_value(output: &mut Vec<u8>, prefix: Option<&str>, uri: &str) {
    output.extend_from_slice(b" xmlns");
    if let Some(prefix) = prefix {
        output.push(b':');
        output.extend_from_slice(prefix.as_bytes());
    }
    output.extend_from_slice(b"=\"");
    push_escaped(output, uri);
    output.push(b'"');
}

fn push_escaped(output: &mut Vec<u8>, value: &str) {
    let escaped = quick_xml::escape::escape(value);
    output.extend_from_slice(escaped.as_bytes());
}

fn reference_from_namespace<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<Reference> {
    let embedded =
        relationships::attribute_value(element, b"embed", reader.decoder(), reader.resolver())?;
    let linked =
        relationships::attribute_value(element, b"link", reader.decoder(), reader.resolver())?;
    reference_from_values(embedded, linked)
}

fn reference_plain(
    element: &BytesStart<'_>,
    namespaces: &[Namespace],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Reference> {
    let mut embedded = None;
    let mut linked = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let local = attribute.key.local_name();
        let is_relationship = attribute.key.prefix().is_some_and(|prefix| {
            std::str::from_utf8(prefix.as_ref())
                .ok()
                .is_some_and(|prefix| is_relationship_prefix(prefix, namespaces))
        });
        if !is_relationship || (local.as_ref() != b"embed" && local.as_ref() != b"link") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        let slot = if local.as_ref() == b"embed" {
            &mut embedded
        } else {
            &mut linked
        };
        if slot.is_some() {
            return Err(invalid(
                "model3d blip has duplicate relationship attributes",
            ));
        }
        *slot = Some(value);
    }
    reference_from_values(embedded, linked)
}

fn reference_from_values(embedded: Option<String>, linked: Option<String>) -> Result<Reference> {
    let embedded = embedded
        .map(|value| Id::new(value).map_err(value_error))
        .transpose()?;
    let linked = linked
        .map(|value| Id::new(value).map_err(value_error))
        .transpose()?;
    Ok(Reference { embedded, linked })
}

fn declarations(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Namespace>> {
    declarations_inner(element, decoder)
}

fn declarations_plain(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Namespace>> {
    declarations_inner(element, decoder)
}

fn declarations_inner(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Namespace>> {
    let mut declarations = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        let prefix = if raw == b"xmlns" {
            None
        } else if let Some(prefix) = raw.strip_prefix(b"xmlns:") {
            Some(std::str::from_utf8(prefix).map_err(xml_error)?)
        } else {
            continue;
        };
        if declarations
            .iter()
            .any(|value: &Namespace| value.prefix() == prefix)
        {
            return Err(invalid(
                "model3d element has duplicate namespace declarations",
            ));
        }
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        declarations.push(Namespace::new(prefix, uri).map_err(value_error)?);
    }
    Ok(declarations)
}

fn merge_namespaces(inherited: &[Namespace], local: &[Namespace]) -> Vec<Namespace> {
    let mut merged = inherited.to_vec();
    for declaration in local {
        if let Some(existing) = merged
            .iter_mut()
            .find(|value| value.prefix() == declaration.prefix())
        {
            *existing = declaration.clone();
        } else {
            merged.push(declaration.clone());
        }
    }
    merged
}

fn prefix_namespace(prefix: Option<Prefix<'_>>, namespaces: &[Namespace]) -> String {
    let prefix = prefix.and_then(|value| String::from_utf8(value.as_ref().to_vec()).ok());
    namespaces
        .iter()
        .find(|value| value.prefix() == prefix.as_deref())
        .map_or_else(String::new, |value| value.uri().to_owned())
}

fn is_relationship_prefix(prefix: &str, namespaces: &[Namespace]) -> bool {
    if prefix == "r" {
        return true;
    }
    namespaces.iter().any(|value| {
        value.prefix() == Some(prefix)
            && (value.uri() == RELATIONSHIP_NAMESPACE
                || value.uri() == RELATIONSHIP_NAMESPACE_STRICT)
    })
}

fn next_element(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
) -> Result<(BytesStart<'static>, bool, usize, usize)> {
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d fragment offset exceeds usize"))?;
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d fragment offset exceeds usize"))?;
        match event {
            Event::Start(element) => {
                buffer.clear();
                return Ok((element, false, start, end));
            },
            Event::Empty(element) => {
                buffer.clear();
                return Ok((element, true, start, end));
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(invalid("model3d fragment contains text outside its root"));
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid("model3d fragment contains document markup"));
            },
            Event::Eof => return Err(invalid("model3d fragment has no root")),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn capture_namespace_element<R: std::io::BufRead>(
    reader: &mut NsReader<R>,
    buffer: &mut Vec<u8>,
) -> Result<usize> {
    let mut depth = 1usize;
    let mut nodes = 0usize;
    loop {
        buffer.clear();
        let (_, event) = reader.read_resolved_event_into(buffer).map_err(xml_error)?;
        let event = event.into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d XML offset exceeds usize"))?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("model3d XML nesting is too deep"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("model3d XML depth", MAX_DEPTH));
                }
                nodes = nodes.saturating_add(1);
                if nodes > MAX_NODES {
                    return Err(limit("model3d XML nodes", MAX_NODES));
                }
            },
            Event::Empty(_) => {
                nodes = nodes.saturating_add(1);
                if nodes > MAX_NODES {
                    return Err(limit("model3d XML nodes", MAX_NODES));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("model3d XML nesting underflow"))?;
                if depth == 0 {
                    return Ok(event_end);
                }
            },
            Event::DocType(_) => return Err(invalid("DTD is forbidden in model3d XML")),
            Event::Eof => return Err(invalid("model3d element is unterminated")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn capture_plain_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    buffer: &mut Vec<u8>,
) -> Result<usize> {
    let mut depth = 1usize;
    loop {
        buffer.clear();
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("model3d XML offset exceeds usize"))?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("model3d XML nesting is too deep"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("model3d XML depth", MAX_DEPTH));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("model3d XML nesting underflow"))?;
                if depth == 0 {
                    return Ok(event_end);
                }
            },
            Event::DocType(_) => return Err(invalid("DTD is forbidden in model3d XML")),
            Event::Eof => return Err(invalid("model3d element is unterminated")),
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn required_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{description} is required")))
}

fn local_name(name: &QName<'_>) -> Result<String> {
    std::str::from_utf8(name.local_name().as_ref())
        .map(str::to_owned)
        .map_err(xml_error)
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "quick_xml yields namespace resolution tokens by value at the event boundary"
)]
fn resolved_namespace(namespace: ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .map(str::to_owned)
            .map_err(xml_error),
        ResolveResult::Unknown(_) | ResolveResult::Unbound => Ok(String::new()),
    }
}

fn resolved_name(namespace: &str, name: &QName<'_>) -> Result<(String, String)> {
    let local = local_name(name)?;
    Ok((local, namespace.to_owned()))
}

fn require_root(local: &str, namespace: &str) -> Result<()> {
    if local != "model3d" || namespace != NAMESPACE {
        return Err(invalid("model3d root name or namespace is invalid"));
    }
    Ok(())
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "map_err transfers the typed validation failure into the public error"
)]
fn validation_error(error: crate::model3d::Error) -> Error {
    Error::Invalid(error.to_string())
}

fn value_error(error: impl std::fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
