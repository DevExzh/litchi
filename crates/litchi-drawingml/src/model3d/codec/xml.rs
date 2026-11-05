//! Namespace-aware XML reader and deterministic writer for `model3d`.

use std::{borrow::Cow, io::Write};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::{relationships, xml::unqualified_attribute_value};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Prefix, QName, ResolveResult},
    reader::{NsReader, Reader},
};

use crate::{Error, LimitKind, Result};

use super::super::{
    Blip, Child, Id, Inert, MAX_NAMESPACE_TEXT_BYTES, Metadata, NAMESPACE, Namespace,
    RELATIONSHIP_NAMESPACE, RELATIONSHIP_NAMESPACE_STRICT, Raster, RasterChild, Reference,
};
use super::Limits;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// Read one complete `m3d:model3d` element.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8], limits: Limits) -> Result<Metadata> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }

    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut metadata = None;
    let mut root_depth = 0usize;
    let mut root_closed = false;
    let mut nodes = 0usize;

    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d XML offset exceeds usize"))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolved_namespace(&namespace)?;
        let event = event.into_owned();
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d XML offset exceeds usize"))?;

        match event {
            Event::Start(element) if metadata.is_none() && root_depth == 0 => {
                count_node(&mut nodes, limits)?;
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
                count_node(&mut nodes, limits)?;
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
                count_node(&mut nodes, limits)?;
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                let child_start = event_start;
                let child_end = capture_namespace_element(
                    &mut reader,
                    origin,
                    &mut buffer,
                    limits,
                    &mut nodes,
                )?;
                let raw = xml
                    .get(child_start..child_end)
                    .ok_or_else(|| invalid("model3d child range is outside the input"))?;
                let metadata = metadata
                    .as_mut()
                    .ok_or_else(|| invalid("model3d root metadata is missing"))?;
                push_child(metadata, local, namespace, raw, limits)?;
            },
            Event::Empty(element) if metadata.is_some() && root_depth == 1 => {
                count_node(&mut nodes, limits)?;
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                let raw = xml
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("model3d child range is outside the input"))?;
                let metadata = metadata
                    .as_mut()
                    .ok_or_else(|| invalid("model3d root metadata is missing"))?;
                push_child(metadata, local, namespace, raw, limits)?;
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
    crate::model3d::validation::validate(&metadata).map_err(|error| validation_error(&error))?;
    Ok(metadata)
}

/// Construct one inert child from a self-contained XML element.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn opaque(xml: &[u8], limits: Limits) -> Result<Inert> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_fragment_bytes {
        return Err(limit(
            LimitKind::InertFragmentBytes,
            xml.len(),
            limits.max_fragment_bytes,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut root_start = 0usize;
    let mut root_end = 0usize;
    let mut nodes = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d fragment offset exceeds usize"))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolved_namespace(&namespace)?;
        let event = event.into_owned();
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d fragment offset exceeds usize"))?;
        match event {
            Event::Start(element) if root.is_none() => {
                count_node(&mut nodes, limits)?;
                let (local, namespace) = resolved_name(&namespace, &element.name())?;
                root = Some((local, namespace));
                root_start = start;
                root_end = capture_namespace_element(
                    &mut reader,
                    origin,
                    &mut buffer,
                    limits,
                    &mut nodes,
                )?;
            },
            Event::Empty(element) if root.is_none() => {
                count_node(&mut nodes, limits)?;
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
    if raw.len() > limits.max_fragment_bytes {
        return Err(limit(
            LimitKind::InertFragmentBytes,
            raw.len(),
            limits.max_fragment_bytes,
        ));
    }
    let value = Inert::from_wire(raw.to_vec(), local, namespace);
    crate::model3d::validation::validate_inert_fragment(&value)
        .map_err(|error| validation_error(&error))?;
    Ok(value)
}

/// Serialize one model3d fragment with stable prefixes and child order.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write(metadata: &Metadata, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    crate::model3d::validation::validate(metadata).map_err(|error| validation_error(&error))?;
    let prefixes = Prefixes::for_metadata(metadata)?;
    let size = output_size(metadata, &prefixes)?;
    if size > limits.max_output_bytes {
        return Err(limit(LimitKind::OutputBytes, size, limits.max_output_bytes));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|error| invalid(format!("model3d output allocation failed: {error}")))?;
    write_inner(&mut output, metadata, &prefixes).map_err(Error::Io)?;
    debug_assert_eq!(output.len(), size);
    Ok(output)
}

/// Serialize one model3d fragment to a caller-owned sink.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_to<W: Write>(writer: &mut W, metadata: &Metadata, limits: Limits) -> Result<()> {
    let limits = limits.validate()?;
    crate::model3d::validation::validate(metadata).map_err(|error| validation_error(&error))?;
    let prefixes = Prefixes::for_metadata(metadata)?;
    let size = output_size(metadata, &prefixes)?;
    if size > limits.max_output_bytes {
        return Err(limit(LimitKind::OutputBytes, size, limits.max_output_bytes));
    }
    write_inner(writer, metadata, &prefixes).map_err(Error::Io)?;
    Ok(())
}

fn push_child(
    metadata: &mut Metadata,
    local: String,
    namespace: String,
    raw: &[u8],
    limits: Limits,
) -> Result<()> {
    if metadata.children.len() >= limits.max_children {
        return Err(limit(
            LimitKind::ModelChildren,
            metadata.children.len() + 1,
            limits.max_children,
        ));
    }
    if raw.len() > limits.max_fragment_bytes {
        return Err(limit(
            LimitKind::InertFragmentBytes,
            raw.len(),
            limits.max_fragment_bytes,
        ));
    }
    if local == "raster" && namespace == NAMESPACE {
        metadata.children.push(Child::Raster(parse_raster(
            raw,
            &metadata.namespaces,
            limits,
        )?));
    } else {
        let value = Inert::from_wire(raw.to_vec(), local, namespace);
        crate::model3d::validation::validate_inert_fragment(&value)
            .map_err(|error| validation_error(&error))?;
        metadata.children.push(Child::Opaque(value));
    }
    Ok(())
}

fn parse_raster(xml: &[u8], inherited: &[Namespace], limits: Limits) -> Result<Raster> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (root, empty, root_start, root_end) = next_element(&mut reader, origin, &mut buffer)?;
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
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d raster offset exceeds usize"))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d raster offset exceeds usize"))?;
        match event {
            Event::Start(element) if depth == 1 => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let child_end = capture_plain_element(&mut reader, origin, &mut buffer, limits)?;
                let raw = xml
                    .get(start..child_end)
                    .ok_or_else(|| invalid("model3d raster child range is outside the input"))?;
                push_raster_child(&mut children, local, namespace, raw, &namespaces, limits)?;
            },
            Event::Empty(element) if depth == 1 => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("model3d raster child range is outside the input"))?;
                push_raster_child(&mut children, local, namespace, raw, &namespaces, limits)?;
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
    if children.len() > limits.max_raster_children {
        return Err(limit(
            LimitKind::RasterChildren,
            children.len(),
            limits.max_raster_children,
        ));
    }
    Raster::from_wire(renderer_name, renderer_version, children, local_namespaces)
        .map_err(value_error)
}

fn push_raster_child(
    children: &mut Vec<RasterChild>,
    local: String,
    namespace: String,
    raw: &[u8],
    namespaces: &[Namespace],
    limits: Limits,
) -> Result<()> {
    if children.len() >= limits.max_raster_children {
        return Err(limit(
            LimitKind::RasterChildren,
            children.len() + 1,
            limits.max_raster_children,
        ));
    }
    if raw.len() > limits.max_fragment_bytes {
        return Err(limit(
            LimitKind::InertFragmentBytes,
            raw.len(),
            limits.max_fragment_bytes,
        ));
    }
    if local == "blip" && namespace == NAMESPACE {
        children.push(RasterChild::Blip(parse_blip(raw, namespaces, limits)?));
    } else {
        let value = Inert::from_wire(raw.to_vec(), local, namespace);
        crate::model3d::validation::validate_inert_fragment(&value)
            .map_err(|error| validation_error(&error))?;
        children.push(RasterChild::Opaque(value));
    }
    Ok(())
}

fn parse_blip(xml: &[u8], inherited: &[Namespace], limits: Limits) -> Result<Blip> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (root, empty, _, _) = next_element(&mut reader, origin, &mut buffer)?;
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
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d blip offset exceeds usize"))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d blip offset exceeds usize"))?;
        match event {
            Event::Start(element) => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let child_end = capture_plain_element(&mut reader, origin, &mut buffer, limits)?;
                let raw = xml
                    .get(start..child_end)
                    .ok_or_else(|| invalid("model3d blip child range is outside the input"))?;
                if children.len() >= limits.max_children {
                    return Err(limit(
                        LimitKind::ModelChildren,
                        children.len() + 1,
                        limits.max_children,
                    ));
                }
                if raw.len() > limits.max_fragment_bytes {
                    return Err(limit(
                        LimitKind::InertFragmentBytes,
                        raw.len(),
                        limits.max_fragment_bytes,
                    ));
                }
                children.push(Inert::from_wire(raw.to_vec(), local, namespace));
            },
            Event::Empty(element) => {
                let local = local_name(&element.name())?;
                let namespace = prefix_namespace(element.name().prefix(), &namespaces);
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("model3d blip child range is outside the input"))?;
                if children.len() >= limits.max_children {
                    return Err(limit(
                        LimitKind::ModelChildren,
                        children.len() + 1,
                        limits.max_children,
                    ));
                }
                if raw.len() > limits.max_fragment_bytes {
                    return Err(limit(
                        LimitKind::InertFragmentBytes,
                        raw.len(),
                        limits.max_fragment_bytes,
                    ));
                }
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

struct SelectedPrefix<'a> {
    value: Cow<'a, str>,
    injected: bool,
}

struct Prefixes<'a> {
    model: SelectedPrefix<'a>,
    relationship: SelectedPrefix<'a>,
}

impl<'a> Prefixes<'a> {
    fn for_metadata(metadata: &'a Metadata) -> Result<Self> {
        let model = select_prefix(metadata, is_model_namespace, "m3d", None)?;
        let relationship = select_prefix(
            metadata,
            is_relationship_namespace,
            "r",
            Some(model.value.as_ref()),
        )?;
        Ok(Self {
            model,
            relationship,
        })
    }
}

fn select_prefix<'a>(
    metadata: &'a Metadata,
    namespace: fn(&str) -> bool,
    stem: &str,
    reserved: Option<&str>,
) -> Result<SelectedPrefix<'a>> {
    if let Some(prefix) = metadata.namespaces.iter().find_map(|declaration| {
        let prefix = declaration.prefix()?;
        (prefix_is_safe(metadata, prefix, namespace) && namespace(declaration.uri()))
            .then_some(prefix)
    }) {
        return Ok(SelectedPrefix {
            value: Cow::Borrowed(prefix),
            injected: false,
        });
    }

    let mut suffix = 0usize;
    loop {
        let candidate = if suffix == 0 {
            stem.to_owned()
        } else {
            format!("{stem}{suffix}")
        };
        if candidate.len() > MAX_NAMESPACE_TEXT_BYTES {
            return Err(limit(
                LimitKind::NamespacePrefixBytes,
                candidate.len(),
                MAX_NAMESPACE_TEXT_BYTES,
            ));
        }
        if !prefix_is_in_use(metadata, candidate.as_str()) && reserved != Some(candidate.as_str()) {
            return Ok(SelectedPrefix {
                value: Cow::Owned(candidate),
                injected: true,
            });
        }
        suffix = suffix.checked_add(1).ok_or_else(|| {
            limit(
                LimitKind::NamespacePrefixBytes,
                usize::MAX,
                MAX_NAMESPACE_TEXT_BYTES,
            )
        })?;
    }
}

fn is_model_namespace(uri: &str) -> bool {
    uri == NAMESPACE
}

fn is_relationship_namespace(uri: &str) -> bool {
    uri == RELATIONSHIP_NAMESPACE || uri == RELATIONSHIP_NAMESPACE_STRICT
}

fn prefix_is_safe(metadata: &Metadata, prefix: &str, namespace: fn(&str) -> bool) -> bool {
    let mut safe = true;
    visit_namespaces(metadata, |declaration| {
        if declaration.prefix() == Some(prefix) && !namespace(declaration.uri()) {
            safe = false;
        }
    });
    safe
}

fn prefix_is_in_use(metadata: &Metadata, prefix: &str) -> bool {
    let mut in_use = false;
    visit_namespaces(metadata, |declaration| {
        in_use |= declaration.prefix() == Some(prefix);
    });
    in_use
}

fn visit_namespaces(metadata: &Metadata, mut visit: impl FnMut(&Namespace)) {
    for declaration in &metadata.namespaces {
        visit(declaration);
    }
    for child in &metadata.children {
        let Child::Raster(raster) = child else {
            continue;
        };
        for declaration in &raster.namespaces {
            visit(declaration);
        }
        for raster_child in &raster.children {
            let RasterChild::Blip(blip) = raster_child else {
                continue;
            };
            for declaration in &blip.namespaces {
                visit(declaration);
            }
        }
    }
}

fn write_inner<W: Write>(
    output: &mut W,
    metadata: &Metadata,
    prefixes: &Prefixes<'_>,
) -> std::io::Result<()> {
    // Typed names and relationship attributes use one selected prefix each.
    // A valid alternate binding is reused; otherwise the selected fresh name
    // is injected only after all retained declarations, avoiding duplicate
    // namespace attributes and preserving opaque child bindings.  The same
    // plan is used by `output_size`, so admission counts every emitted byte
    // without cloning the wire.
    output.write_all(b"<")?;
    output.write_all(prefixes.model.value.as_bytes())?;
    output.write_all(b":model3d")?;
    for declaration in &metadata.namespaces {
        write_namespace(output, declaration)?;
    }
    if prefixes.model.injected {
        write_namespace_value(output, Some(prefixes.model.value.as_ref()), NAMESPACE)?;
    }
    if prefixes.relationship.injected {
        write_namespace_value(
            output,
            Some(prefixes.relationship.value.as_ref()),
            RELATIONSHIP_NAMESPACE,
        )?;
    }
    write_reference(
        output,
        &metadata.reference,
        prefixes.relationship.value.as_ref(),
    )?;
    output.write_all(b">")?;
    for child in &metadata.children {
        match child {
            Child::Opaque(value) => output.write_all(value.as_bytes())?,
            Child::Raster(value) => write_raster(output, value, prefixes)?,
        }
    }
    output.write_all(b"</")?;
    output.write_all(prefixes.model.value.as_bytes())?;
    output.write_all(b":model3d>")?;
    Ok(())
}

fn write_raster<W: Write>(
    output: &mut W,
    raster: &Raster,
    prefixes: &Prefixes<'_>,
) -> std::io::Result<()> {
    output.write_all(b"<")?;
    output.write_all(prefixes.model.value.as_bytes())?;
    output.write_all(b":raster rName=\"")?;
    push_escaped(output, &raster.renderer_name)?;
    output.write_all(b"\" rVer=\"")?;
    push_escaped(output, &raster.renderer_version)?;
    output.write_all(b"\"")?;
    for declaration in &raster.namespaces {
        write_namespace(output, declaration)?;
    }
    output.write_all(b">")?;
    for child in &raster.children {
        match child {
            RasterChild::Blip(value) => write_blip(output, value, prefixes)?,
            RasterChild::Opaque(value) => output.write_all(value.as_bytes())?,
        }
    }
    output.write_all(b"</")?;
    output.write_all(prefixes.model.value.as_bytes())?;
    output.write_all(b":raster>")?;
    Ok(())
}

fn write_blip<W: Write>(
    output: &mut W,
    blip: &Blip,
    prefixes: &Prefixes<'_>,
) -> std::io::Result<()> {
    output.write_all(b"<")?;
    output.write_all(prefixes.model.value.as_bytes())?;
    output.write_all(b":blip")?;
    for declaration in &blip.namespaces {
        write_namespace(output, declaration)?;
    }
    write_reference(
        output,
        &blip.reference,
        prefixes.relationship.value.as_ref(),
    )?;
    if blip.children.is_empty() {
        output.write_all(b"/>")?;
    } else {
        output.write_all(b">")?;
        for child in &blip.children {
            output.write_all(child.as_bytes())?;
        }
        output.write_all(b"</")?;
        output.write_all(prefixes.model.value.as_bytes())?;
        output.write_all(b":blip>")?;
    }
    Ok(())
}

fn write_reference<W: Write>(
    output: &mut W,
    reference: &Reference,
    relationship_prefix: &str,
) -> std::io::Result<()> {
    if let Some(id) = &reference.embedded {
        output.write_all(b" ")?;
        output.write_all(relationship_prefix.as_bytes())?;
        output.write_all(b":embed=\"")?;
        push_escaped(output, id.as_str())?;
        output.write_all(b"\"")?;
    }
    if let Some(id) = &reference.linked {
        output.write_all(b" ")?;
        output.write_all(relationship_prefix.as_bytes())?;
        output.write_all(b":link=\"")?;
        push_escaped(output, id.as_str())?;
        output.write_all(b"\"")?;
    }
    Ok(())
}

fn write_namespace<W: Write>(output: &mut W, declaration: &Namespace) -> std::io::Result<()> {
    write_namespace_value(output, declaration.prefix(), declaration.uri())
}

fn write_namespace_value<W: Write>(
    output: &mut W,
    prefix: Option<&str>,
    uri: &str,
) -> std::io::Result<()> {
    output.write_all(b" xmlns")?;
    if let Some(prefix) = prefix {
        output.write_all(b":")?;
        output.write_all(prefix.as_bytes())?;
    }
    output.write_all(b"=\"")?;
    push_escaped(output, uri)?;
    output.write_all(b"\"")?;
    Ok(())
}

fn push_escaped<W: Write>(output: &mut W, value: &str) -> std::io::Result<()> {
    let escaped = quick_xml::escape::escape(value);
    output.write_all(escaped.as_bytes())
}

fn output_size(metadata: &Metadata, prefixes: &Prefixes<'_>) -> Result<usize> {
    let mut counter = CountingWriter::default();
    write_inner(&mut counter, metadata, prefixes).map_err(Error::Io)?;
    Ok(counter.bytes)
}

#[derive(Default)]
struct CountingWriter {
    bytes: usize,
}

impl Write for CountingWriter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.bytes = self
            .bytes
            .checked_add(bytes.len())
            .ok_or_else(|| std::io::Error::other("model3d output size overflow"))?;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
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
    for attribute in element.checked_attributes() {
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
    for attribute in element.checked_attributes() {
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
    origin: ReaderOrigin,
    buffer: &mut Vec<u8>,
) -> Result<(BytesStart<'static>, bool, usize, usize)> {
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d fragment offset exceeds usize"))?;
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d fragment offset exceeds usize"))?;
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
    origin: ReaderOrigin,
    buffer: &mut Vec<u8>,
    limits: Limits,
    nodes: &mut usize,
) -> Result<usize> {
    let mut depth = 1usize;
    loop {
        buffer.clear();
        let (_, event) = reader.read_resolved_event_into(buffer).map_err(xml_error)?;
        let event = event.into_owned();
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d XML offset exceeds usize"))?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("model3d XML nesting is too deep"))?;
                if depth > limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, depth, limits.max_depth));
                }
                count_node(nodes, limits)?;
            },
            Event::Empty(_) => {
                count_node(nodes, limits)?;
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

fn count_node(nodes: &mut usize, limits: Limits) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit(LimitKind::XmlNodes, usize::MAX, limits.max_nodes))?;
    if *nodes > limits.max_nodes {
        return Err(limit(LimitKind::XmlNodes, *nodes, limits.max_nodes));
    }
    Ok(())
}

fn capture_plain_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    origin: ReaderOrigin,
    buffer: &mut Vec<u8>,
    limits: Limits,
) -> Result<usize> {
    let mut depth = 1usize;
    loop {
        buffer.clear();
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("model3d XML offset exceeds usize"))?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("model3d XML nesting is too deep"))?;
                if depth > limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, depth, limits.max_depth));
                }
            },
            Event::Empty(_) => {},
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

fn resolved_namespace(namespace: &ResolveResult<'_>) -> Result<String> {
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

fn validation_error(error: &crate::model3d::Error) -> Error {
    match error {
        crate::model3d::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(*kind, *observed, *maximum),
        _ => Error::Invalid(error.to_string()),
    }
}

fn value_error(error: super::super::ValueError) -> Error {
    match error {
        super::super::ValueError::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(kind, observed, maximum),
        error => Error::Invalid(error.to_string()),
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
