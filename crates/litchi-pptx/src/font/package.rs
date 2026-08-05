//! Atomic OPC graph services for PresentationML embedded fonts.

use super::codec::*;
use super::model::*;
use super::{
    MAX_DEPTH, MAX_FONT_BYTES, MAX_MCE_MARKED_BYTES, MAX_NODES, MAX_TOTAL_FONT_BYTES,
    MAX_XML_BYTES, MCE_NS, PML, STRICT_PML, invalid, limit,
};
use crate::error::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, OffsetLimits, active_offsets};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

/// Loads embedded-font metadata and validates every referenced inert font part.
pub(super) fn load_raw(package: &OpcPackage) -> Result<Option<RawFonts>> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation)?;
    let presentation_name = presentation.partname().to_string();
    let parsed = parse_presentation(presentation.blob())?;
    let conformance = parsed.conformance;
    validate_font_relationship_sources(package, &presentation_name)?;
    let Some(mut value) = parsed.value else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let mut targets = HashSet::new();
    let mut references = HashSet::new();
    let mut resources = HashMap::<String, RawResource>::new();
    let mut total_bytes = 0usize;
    for font in &mut value.fonts {
        for face in &mut font.faces {
            references.insert(face.relationship_id.clone());
            let relationship = presentation
                .rels()
                .get(&face.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "missing embedded-font relationship '{}'",
                        face.relationship_id
                    ))
                })?;
            if relationship.reltype() != conformance.font_rel() {
                return Err(invalid(format!(
                    "relationship '{}' does not match the presentation conformance",
                    face.relationship_id,
                )));
            }
            if relationship.is_external() {
                return Err(invalid("embedded-font relationship must be internal"));
            }
            let target = relationship.target_partname()?;
            let target_name = target.to_string();
            targets.insert(target_name.clone());
            if let Some(resource) = resources.get(&target_name) {
                face.resource = Some(resource.clone());
                continue;
            }
            let part = package.get_part(&target)?;
            if !is_font_content_type(part.content_type()) {
                return Err(invalid(format!(
                    "font part '{target}' has invalid content type '{}'",
                    part.content_type()
                )));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{target}' has outbound relationships"
                )));
            }
            if part.blob().len() > MAX_FONT_BYTES {
                return Err(limit("individual font bytes"));
            }
            total_bytes = total_bytes
                .checked_add(part.blob().len())
                .ok_or_else(|| limit("total font bytes"))?;
            if total_bytes > MAX_TOTAL_FONT_BYTES {
                return Err(limit("total font bytes"));
            }
            let resource = RawResource {
                part_name: target_name.clone(),
                content_type: part.content_type().to_owned(),
                data: part.blob_arc(),
            };
            resources.insert(target_name, resource.clone());
            face.resource = Some(resource);
        }
    }
    validate_inbound_font_graph(
        package,
        &presentation_name,
        presentation,
        &references,
        &targets,
    )?;
    reject_orphan_font_parts(package, &targets)?;
    Ok(Some(value))
}

/// Atomically stores the complete embedded-font graph.
///
/// Existing font relationships are replaced. RawFont parts still referenced by
/// another relationship are retained, and unrelated presentation XML is copied
/// byte-for-byte.
pub(super) fn put_raw(
    package: &mut OpcPackage,
    value: &RawFonts,
    conformance: Conformance,
) -> Result<bool> {
    validate_value(value, true)?;
    let old = load_raw(package)?;
    let presentation = package.main_document_part()?;
    let presentation_name = presentation.partname().clone();
    let parsed = parse_presentation(presentation.blob())?;
    if parsed.conformance != conformance {
        return Err(invalid(
            "requested conformance does not match the presentation namespace",
        ));
    }
    let enabled = !value.fonts.is_empty();
    if enabled && old.as_ref() == Some(value) {
        return Ok(false);
    }
    if !enabled && old.is_none() && !embedding_enabled(presentation.blob())? {
        return Ok(false);
    }
    let fragment = if value.fonts.is_empty() {
        Vec::new()
    } else {
        write_raw(value, conformance)?
    };
    let updated_xml = patch_embedding_flag(
        &patch_font_list(presentation.blob(), &fragment, conformance)?,
        enabled,
    )?;
    let staged = parse_presentation(&updated_xml)?;
    let expected = enabled.then(|| metadata_only(value));
    if staged.conformance != conformance || staged.value != expected {
        return Err(invalid("staged embedded-font XML did not round-trip"));
    }
    let old_relationship_ids = old
        .iter()
        .flat_map(|value| &value.fonts)
        .flat_map(|font| font.faces.iter().map(|face| face.relationship_id.clone()))
        .collect::<HashSet<_>>();
    let old_part_names = old
        .iter()
        .flat_map(|value| &value.fonts)
        .flat_map(|font| font.faces.iter())
        .filter_map(|face| {
            face.resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
        .collect::<HashSet<_>>();
    let mut relationship_ids = HashMap::<String, PackURI>::new();
    let mut resources = HashMap::<String, (String, Arc<Vec<u8>>)>::new();
    let mut relationships = Vec::new();
    for font in &value.fonts {
        for face in &font.faces {
            let resource = face
                .resource
                .as_ref()
                .ok_or_else(|| invalid("embedded-font resource is required for package storage"))?;
            let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
            if let Some(existing) = relationship_ids.get(&face.relationship_id) {
                if existing != &uri {
                    return Err(invalid(format!(
                        "relationship ID '{}' resolves to conflicting font parts",
                        face.relationship_id
                    )));
                }
            } else {
                if presentation.rels().get(&face.relationship_id).is_some()
                    && !old_relationship_ids.contains(&face.relationship_id)
                {
                    return Err(invalid(format!(
                        "relationship ID '{}' already exists",
                        face.relationship_id
                    )));
                }
                relationship_ids.insert(face.relationship_id.clone(), uri.clone());
                relationships.push((uri.clone(), face.relationship_id.clone()));
            }
            if let Some((content_type, data)) = resources.get(uri.as_str()) {
                if content_type != &resource.content_type
                    || (!Arc::ptr_eq(data, &resource.data)
                        && data.as_slice() != resource.data.as_slice())
                {
                    return Err(invalid(format!(
                        "shared font part '{uri}' has conflicting resources"
                    )));
                }
            } else {
                resources.insert(
                    uri.to_string(),
                    (resource.content_type.clone(), resource.data.clone()),
                );
            }
        }
    }

    for (part_name, (content_type, data)) in &resources {
        let uri = PackURI::new(part_name).map_err(Error::Invalid)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            let stored = part.blob_arc();
            let same_data = Arc::ptr_eq(&stored, data) || stored.as_slice() == data.as_slice();
            if !same_data && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if !same_data
                && has_inbound_outside_relationships(
                    package,
                    &uri,
                    &presentation_name,
                    &old_relationship_ids,
                )?
            {
                return Err(invalid(format!(
                    "shared font part '{uri}' cannot be overwritten"
                )));
            }
        }
    }

    let mut candidate = package.clone();
    candidate.unsign();
    let existing_font_relationships = candidate
        .get_part(&presentation_name)?
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    for relationship_id in existing_font_relationships {
        candidate
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .remove(&relationship_id);
    }
    for (uri, relationship_id) in &relationships {
        candidate
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .add_relationship(
                conformance.font_rel().into(),
                uri.relative_ref(presentation_name.base_uri()),
                relationship_id.clone(),
                false,
            );
    }
    for (part_name, (content_type, data)) in resources {
        let uri = PackURI::new(&part_name).map_err(Error::Invalid)?;
        if let Ok(part) = candidate.get_part_mut(&uri) {
            part.set_blob_shared(data);
        } else {
            candidate.add_part(Box::new(BlobPart::new_shared(uri, content_type, data)));
        }
    }
    candidate
        .get_part_mut(&presentation_name)?
        .set_blob(updated_xml);
    let retained = relationships
        .iter()
        .map(|(uri, _)| uri.to_string())
        .collect::<HashSet<_>>();
    for old_part in old_part_names {
        if !retained.contains(&old_part) {
            let uri = PackURI::new(&old_part).map_err(Error::Invalid)?;
            if !part_is_referenced(&candidate, &uri)? {
                candidate.remove_part(&uri);
            }
        }
    }
    *package = candidate;
    Ok(true)
}

/// Load the complete semantic embedded-font collection.
///
/// Relationship IDs and part names remain private provenance. Font programs
/// that share one physical part share the same allocation in memory.
pub fn load(package: &OpcPackage) -> Result<Option<Fonts>> {
    load_raw(package)?.map(fonts_from_raw).transpose()
}

/// Atomically publish a complete collection, consuming its owned values.
///
/// Returns `false` for an exact semantic and physical no-op, preserving any
/// valid package signatures. A real mutation invalidates signatures only after
/// every bounded validation and staging operation succeeds.
pub fn put(package: &mut OpcPackage, fonts: Fonts) -> Result<bool> {
    validate_fonts(&fonts, true)?;
    let current = load_raw(package)?;
    let presentation = package.main_document_part()?;
    let conformance = parse_presentation(presentation.blob())?.conformance;
    let empty = RawFonts::default();
    let raw = fonts_into_raw(package, fonts, current.as_ref().unwrap_or(&empty))?;
    if current.as_ref() == Some(&raw) {
        return Ok(false);
    }
    put_raw(package, &raw, conformance)
}

/// Remove the complete embedded-font graph and return its previous semantic
/// value. Absence is an exact no-op.
pub fn remove(package: &mut OpcPackage) -> Result<Option<Fonts>> {
    let Some(current) = load_raw(package)? else {
        return Ok(None);
    };
    let value = fonts_from_raw(current.clone())?;
    let presentation = package.main_document_part()?;
    let conformance = parse_presentation(presentation.blob())?.conformance;
    let changed = put_raw(package, &RawFonts::default(), conformance)?;
    if changed { Ok(Some(value)) } else { Ok(None) }
}

/// Detect the presentation namespace profile without exposing XML details.
pub fn conformance(package: &OpcPackage) -> Result<Conformance> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation)?;
    Ok(parse_presentation(presentation.blob())?.conformance)
}

pub(super) fn fonts_from_raw(raw: RawFonts) -> Result<Fonts> {
    let mut fonts = Fonts::new();
    fonts
        .fonts
        .try_reserve(raw.fonts.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-font collection",
            source,
        })?;
    for raw_font in raw.fonts {
        let mut faces = Vec::new();
        faces
            .try_reserve(raw_font.faces.len())
            .map_err(|source| Error::Allocation {
                resource: "embedded-font faces",
                source,
            })?;
        for raw_face in raw_font.faces {
            let resource = raw_face
                .resource
                .ok_or_else(|| invalid("loaded embedded-font face has no resource"))?;
            let format = Format::parse(&resource.content_type)?;
            let data = Data::preserve(resource.data, format)?;
            faces.push(Face {
                style: raw_face.style,
                data,
                source: Some(Source {
                    relationship_id: raw_face.relationship_id,
                    part_name: resource.part_name,
                }),
            });
        }
        validate_typeface(&raw_font.typeface)?;
        fonts.fonts.push(Font {
            key: name_key(&raw_font.typeface),
            typeface: raw_font.typeface,
            panose: raw_font.panose,
            pitch_family: raw_font.pitch_family,
            charset: raw_font.charset,
            faces,
        });
    }
    fonts.reindex()?;
    validate_fonts(&fonts, false)?;
    Ok(fonts)
}

pub(super) fn fonts_into_raw(
    package: &OpcPackage,
    fonts: Fonts,
    current: &RawFonts,
) -> Result<RawFonts> {
    let presentation = package.main_document_part()?;
    let mut relationship_ids = presentation
        .rels()
        .iter()
        .filter(|relationship| !is_font_relationship(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    let current_sources = current
        .fonts
        .iter()
        .flat_map(|font| &font.faces)
        .filter_map(|face| {
            face.resource.as_ref().map(|resource| {
                (
                    (face.relationship_id.clone(), resource.part_name.clone()),
                    (resource.content_type.clone(), Arc::clone(&resource.data)),
                )
            })
        })
        .collect::<HashMap<_, _>>();
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    let mut shared_parts = HashMap::<(usize, Format), String>::new();
    let mut claimed_parts = HashMap::<String, (usize, Format)>::new();
    let mut claimed_relationships = HashMap::<String, String>::new();
    let mut raw_fonts = Vec::new();
    raw_fonts
        .try_reserve(fonts.fonts.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-font staging",
            source,
        })?;
    for font in fonts.fonts {
        let mut raw_faces = Vec::new();
        raw_faces
            .try_reserve(font.faces.len())
            .map_err(|source| Error::Allocation {
                resource: "embedded-font face staging",
                source,
            })?;
        for face in font.faces {
            let data_id = (Arc::as_ptr(&face.data.bytes) as usize, face.data.format);
            let valid_source = face.source.as_ref().filter(|source| {
                current_sources
                    .get(&(source.relationship_id.clone(), source.part_name.clone()))
                    .is_some_and(|(content_type, data)| {
                        content_type == face.data.format.content_type()
                            && Arc::ptr_eq(data, &face.data.bytes)
                    })
            });
            let part_name = if let Some(part_name) = shared_parts.get(&data_id) {
                part_name.clone()
            } else {
                let reusable =
                    valid_source
                        .map(|source| source.part_name.clone())
                        .filter(|part_name| {
                            claimed_parts
                                .get(part_name)
                                .is_none_or(|claimed| *claimed == data_id)
                        });
                let part_name = match reusable {
                    Some(part_name) => part_name,
                    None => next_font_part_name(&part_names, face.data.format)?,
                };
                part_names.insert(part_name.clone());
                claimed_parts.insert(part_name.clone(), data_id);
                shared_parts.insert(data_id, part_name.clone());
                part_name
            };
            let relationship_id = if let Some(source) = valid_source {
                match claimed_relationships.get(&source.relationship_id) {
                    Some(existing) if existing == &part_name => source.relationship_id.clone(),
                    None if !relationship_ids.contains(&source.relationship_id) => {
                        source.relationship_id.clone()
                    },
                    _ => next_font_relationship_id(&relationship_ids)?,
                }
            } else {
                next_font_relationship_id(&relationship_ids)?
            };
            relationship_ids.insert(relationship_id.clone());
            claimed_relationships
                .entry(relationship_id.clone())
                .or_insert_with(|| part_name.clone());
            raw_faces.push(RawFace {
                style: face.style,
                relationship_id,
                resource: Some(RawResource {
                    part_name,
                    content_type: face.data.format.content_type().into(),
                    data: face.data.bytes,
                }),
            });
        }
        raw_fonts.push(RawFont {
            has_descriptor: true,
            typeface: font.typeface,
            panose: font.panose,
            pitch_family: font.pitch_family,
            charset: font.charset,
            faces: raw_faces,
        });
    }
    let raw = RawFonts { fonts: raw_fonts };
    validate_value(&raw, true)?;
    Ok(raw)
}

pub(super) fn next_font_relationship_id(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdFont{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("relationship IDs"))
}

pub(super) fn next_font_part_name(used: &HashSet<String>, format: Format) -> Result<String> {
    let extension = match format {
        Format::PowerPoint => "fntdata",
        Format::Standard => "ttf",
    };
    for index in 1..=u32::MAX {
        let candidate = format!("/ppt/fonts/font{index}.{extension}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("part names"))
}

pub(super) fn embedding_enabled(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentation"
                    || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes())
                {
                    return Err(invalid("expected a PresentationML presentation root"));
                }
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() == b"embedTrueTypeFonts" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(xml_error)?;
                        return match value.as_ref() {
                            "1" | "true" => Ok(true),
                            "0" | "false" => Ok(false),
                            _ => Err(invalid(format!(
                                "invalid embedTrueTypeFonts boolean '{value}'"
                            ))),
                        };
                    }
                }
                return Ok(false);
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => return Err(invalid("missing presentation root")),
            _ => return Err(invalid("unexpected content before presentation root")),
        }
    }
}

pub(super) fn patch_embedding_flag(xml: &[u8], enabled: bool) -> Result<Vec<u8>> {
    let (start, end) = presentation_start_tag(xml)?;
    let tag = xml
        .get(start..end)
        .ok_or_else(|| invalid("presentation start-tag range is invalid"))?;
    let value_range = find_unqualified_attribute_value(tag, b"embedTrueTypeFonts")?;
    let replacement = if enabled {
        b"1".as_slice()
    } else {
        b"0".as_slice()
    };
    if let Some((value_start, value_end)) = value_range {
        if tag.get(value_start..value_end) == Some(replacement) {
            return Ok(xml.to_vec());
        }
        let absolute_start = start
            .checked_add(value_start)
            .ok_or_else(|| limit("updated presentation XML bytes"))?;
        let absolute_end = start
            .checked_add(value_end)
            .ok_or_else(|| limit("updated presentation XML bytes"))?;
        return replace_bytes(xml, absolute_start, absolute_end, replacement);
    }
    if !enabled {
        return Ok(xml.to_vec());
    }
    let mut insertion = end
        .checked_sub(1)
        .ok_or_else(|| invalid("presentation start tag is empty"))?;
    if xml.get(insertion.wrapping_sub(1)) == Some(&b'/') {
        insertion = insertion
            .checked_sub(1)
            .ok_or_else(|| invalid("presentation start tag is empty"))?;
    }
    replace_bytes(xml, insertion, insertion, b" embedTrueTypeFonts=\"1\"")
}

pub(super) fn presentation_start_tag(xml: &[u8]) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentation"
                    || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes())
                {
                    return Err(invalid("expected a PresentationML presentation root"));
                }
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("presentation XML offset overflow"))?;
                return Ok((start, end));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => return Err(invalid("missing presentation root")),
            _ => return Err(invalid("unexpected content before presentation root")),
        }
    }
}

pub(super) fn find_unqualified_attribute_value(
    tag: &[u8],
    wanted: &[u8],
) -> Result<Option<(usize, usize)>> {
    let mut offset = tag
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or_else(|| invalid("presentation start tag has no opening delimiter"))?
        + 1;
    while tag
        .get(offset)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'>' | b'/'))
    {
        offset += 1;
    }
    loop {
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        if tag
            .get(offset)
            .is_none_or(|byte| matches!(byte, b'>' | b'/'))
        {
            return Ok(None);
        }
        let name_start = offset;
        while tag
            .get(offset)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            offset += 1;
        }
        let name_end = offset;
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        if tag.get(offset) != Some(&b'=') {
            return Err(invalid("presentation root contains a malformed attribute"));
        }
        offset += 1;
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        let quote = *tag
            .get(offset)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("presentation root attribute is not quoted"))?;
        offset += 1;
        let value_start = offset;
        while tag.get(offset).is_some_and(|byte| *byte != quote) {
            offset += 1;
        }
        let value_end = offset;
        if tag.get(offset) != Some(&quote) {
            return Err(invalid("presentation root attribute is unterminated"));
        }
        offset += 1;
        if tag.get(name_start..name_end) == Some(wanted) {
            return Ok(Some((value_start, value_end)));
        }
    }
}

pub(super) fn replace_bytes(
    xml: &[u8],
    start: usize,
    end: usize,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if start > end || end > xml.len() {
        return Err(invalid("presentation XML replacement range is invalid"));
    }
    let length = xml
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| limit("updated presentation XML bytes"))?;
    if length > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "embedded-font presentation patch",
            source,
        })?;
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

pub(super) fn patch_font_list(
    xml: &[u8],
    fragment: &[u8],
    conformance: Conformance,
) -> Result<Vec<u8>> {
    let scan = active_direct_elements(xml, conformance)?;
    let mut lists = scan
        .elements
        .iter()
        .filter(|element| element.rank == presentation_child_rank("embeddedFontLst"));
    let list = lists.next();
    if lists.next().is_some() {
        return Err(invalid(
            "presentation has multiple active direct embeddedFontLst elements",
        ));
    }
    if let Some(list) = list {
        return replace_bytes(xml, list.start, list.end, fragment);
    }
    if fragment.is_empty() {
        Ok(xml.to_vec())
    } else {
        insert_font_list(xml, fragment, &scan)
    }
}

#[derive(Clone, Copy)]
pub(super) struct DirectElement {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) rank: Option<usize>,
}

pub(super) struct DirectScan {
    pub(super) elements: Vec<DirectElement>,
    pub(super) root_close: usize,
}

#[derive(Clone, Copy)]
pub(super) struct DirectFrame {
    pub(super) mce_wrapper: bool,
    pub(super) element: Option<usize>,
}

pub(super) fn active_direct_elements(xml: &[u8], conformance: Conformance) -> Result<DirectScan> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("presentation XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut frames = Vec::<DirectFrame>::new();
    let mut elements = Vec::<DirectElement>::new();
    let mut offsets = Vec::<u32>::new();
    let mut root_seen = false;
    let mut root_close = None;
    let mut nodes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if frames.len() >= MAX_DEPTH {
                    return Err(limit("presentation XML depth"));
                }
                let is_pml = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if value == conformance.pml().as_bytes()
                );
                let is_mce = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value)) if value == MCE_NS.as_bytes()
                );
                let local = element.local_name();
                if frames.is_empty() {
                    if root_seen || !is_pml || local.as_ref() != b"presentation" {
                        return Err(invalid(
                            "presentation root does not match requested conformance",
                        ));
                    }
                    root_seen = true;
                }
                let effective_direct =
                    !frames.is_empty() && frames.iter().skip(1).all(|frame| frame.mce_wrapper);
                let rank = effective_direct
                    .then(|| std::str::from_utf8(local.as_ref()).ok())
                    .flatten()
                    .and_then(presentation_child_rank);
                let direct = if is_pml && rank.is_some() {
                    let index = elements.len();
                    elements.push(DirectElement {
                        start,
                        end: 0,
                        rank,
                    });
                    offsets.push(
                        u32::try_from(start)
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                    );
                    Some(index)
                } else {
                    None
                };
                let mce_wrapper = is_mce
                    && matches!(
                        local.as_ref(),
                        b"AlternateContent" | b"Choice" | b"Fallback"
                    );
                frames.push(DirectFrame {
                    mce_wrapper,
                    element: direct,
                });
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if frames.is_empty() {
                    return Err(invalid("presentation root cannot be empty"));
                }
                let is_pml = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if value == conformance.pml().as_bytes()
                );
                let effective_direct = frames.iter().skip(1).all(|frame| frame.mce_wrapper);
                let local = element.local_name();
                let rank = effective_direct
                    .then(|| std::str::from_utf8(local.as_ref()).ok())
                    .flatten()
                    .and_then(presentation_child_rank);
                if is_pml && rank.is_some() {
                    elements.push(DirectElement {
                        start,
                        end: usize::try_from(reader.buffer_position())
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                        rank,
                    });
                    offsets.push(
                        u32::try_from(start)
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                    );
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("unexpected presentation closing element"))?;
                if let Some(index) = frame.element {
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("presentation XML offset overflow"))?;
                    let len = elements.len();
                    elements
                        .get_mut(index)
                        .ok_or(Error::FontIndexOutOfBounds { index, len })?
                        .end = end;
                }
                if frames.is_empty() {
                    root_close = Some(start);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !frames.is_empty() {
        return Err(invalid("invalid presentation XML"));
    }
    let defaults = OffsetLimits::default();
    let limits = OffsetLimits {
        max_source_bytes: MAX_XML_BYTES,
        max_offsets: MAX_NODES,
        max_marked_bytes: MAX_MCE_MARKED_BYTES,
        processing: Limits {
            max_input_bytes: MAX_MCE_MARKED_BYTES,
            max_output_bytes: MAX_MCE_MARKED_BYTES,
            max_depth: MAX_DEPTH,
            ..defaults.processing
        },
    };
    let active = active_offsets(xml, &offsets, &Capabilities::default(), &limits)?;
    let mut active = active.into_iter().peekable();
    elements.retain(|element| {
        let Ok(start) = u32::try_from(element.start) else {
            return false;
        };
        while active.peek().is_some_and(|offset| *offset < start) {
            active.next();
        }
        if active.peek() == Some(&start) {
            active.next();
            true
        } else {
            false
        }
    });
    Ok(DirectScan {
        elements,
        root_close: root_close.ok_or_else(|| invalid("missing presentation closing element"))?,
    })
}

pub(super) fn insert_font_list(xml: &[u8], fragment: &[u8], scan: &DirectScan) -> Result<Vec<u8>> {
    let font_rank = presentation_child_rank("embeddedFontLst")
        .ok_or_else(|| invalid("missing embeddedFontLst schema rank"))?;
    let position = scan
        .elements
        .iter()
        .filter(|element| element.rank.is_some_and(|rank| rank > font_rank))
        .map(|element| element.start)
        .min()
        .unwrap_or(scan.root_close);
    let length = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated presentation XML bytes"))?;
    if length > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "embedded-font presentation patch",
            source,
        })?;
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

pub(super) fn validate_font_relationship_sources(
    package: &OpcPackage,
    presentation: &str,
) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_relationship(relationship.reltype()))
    {
        return Err(invalid("package root cannot source font relationships"));
    }
    for part in package.iter_parts() {
        if part.partname().as_str() != presentation
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_relationship(relationship.reltype()))
        {
            return Err(invalid(format!(
                "font relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

pub(super) fn validate_inbound_font_graph(
    package: &OpcPackage,
    presentation_name: &str,
    presentation: &dyn Part,
    references: &HashSet<String>,
    targets: &HashSet<String>,
) -> Result<()> {
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
    {
        if !references.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced font relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        let target = relationship.target_partname()?;
        if targets.contains(target.as_str()) {
            return Err(invalid(format!(
                "font part '{target}' has an invalid package-root relationship"
            )));
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = relationship.target_partname()?;
            if targets.contains(target.as_str())
                && (part.partname().as_str() != presentation_name
                    || !is_font_relationship(relationship.reltype())
                    || !references.contains(relationship.r_id()))
            {
                return Err(invalid(format!(
                    "font part '{target}' has an invalid inbound relationship"
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn reject_orphan_font_parts(
    package: &OpcPackage,
    targets: &HashSet<String>,
) -> Result<()> {
    for part in package.iter_parts() {
        if is_font_content_type(part.content_type())
            && !targets.contains(part.partname().as_str())
            && !part_is_referenced(package, part.partname())?
        {
            return Err(invalid(format!("orphan font part '{}'", part.partname())));
        }
    }
    Ok(())
}

pub(super) fn require_presentation(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has unsupported presentation content type '{}'",
            part.content_type()
        )))
    }
}

pub(super) fn metadata_only(value: &RawFonts) -> RawFonts {
    let mut value = value.clone();
    for font in &mut value.fonts {
        for face in &mut font.faces {
            face.resource = None;
        }
    }
    value
}

pub(super) fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

pub(super) fn has_inbound_outside_relationships(
    package: &OpcPackage,
    target: &PackURI,
    presentation: &PackURI,
    replaced_relationships: &HashSet<String>,
) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target
                && (part.partname() != presentation
                    || !replaced_relationships.contains(relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}
