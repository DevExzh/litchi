#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! XLSB OPC ownership for timeline cache and worksheet timeline XML parts.

use super::codec::{
    CACHE_CONTENT_TYPE, CACHE_RELATIONSHIP_TYPE, VIEWS_CONTENT_TYPE, VIEWS_RELATIONSHIP_TYPE,
    parse_cache, parse_views, write_cache, write_views,
};
use super::model::{Cache, MAX_CACHES, MAX_VIEWS, Views, validate_cache, validate_views};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Records, Writer, kind};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::char;
use std::collections::HashSet;

/// A timeline cache resolved from a workbook relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachePart {
    /// Workbook relationship identifier.
    pub relationship_id: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed timeline cache snapshot.
    pub cache: Cache,
}

/// Worksheet timeline views resolved from a worksheet relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ViewPart {
    /// Worksheet relationship identifier.
    pub relationship_id: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed timeline view snapshot.
    pub views: Views,
}

const MAX_REL_ID_UNITS: usize = 255;

fn invalid(typ: &str, value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.to_string(),
        val: value.into(),
    }
}

fn parse_rel_id(payload: &[u8], typ: &str) -> Result<String> {
    let mut cursor = Cursor::new(payload, "FRTHeader.RelID");
    let flags = cursor.read_u32()?;
    if flags != 0x08 {
        return Err(invalid(
            typ,
            format!("invalid FRTHeader flags 0x{flags:08X}"),
        ));
    }
    let units = usize::from(cursor.read_u16()?);
    if units == 0 || units > MAX_REL_ID_UNITS {
        return Err(invalid(typ, "invalid relationship ID length"));
    }
    let bytes = cursor.read_bytes(units * 2)?;
    let value: String = char::decode_utf16(
        bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
    )
    .collect::<std::result::Result<_, _>>()
    .map_err(|_| invalid(typ, "invalid UTF-16 relationship ID"))?;
    if value.contains('\0') {
        return Err(invalid(typ, "relationship ID contains NUL"));
    }
    cursor.finish()?;
    Ok(value)
}

fn write_rel_id(value: &str) -> Result<Vec<u8>> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_REL_ID_UNITS || value.contains('\0') {
        return Err(invalid("FRTHeader.RelID", "invalid relationship ID"));
    }
    let mut output = Vec::with_capacity(6 + units * 2);
    let mut writer = Writer::new(&mut output);
    writer.write_u32(0x08)?;
    writer.write_u16(units as u16)?;
    for unit in value.encode_utf16() {
        writer.write_u16(unit)?;
    }
    Ok(output)
}

fn record_span(
    data: &[u8],
    begin: crate::raw::Kind,
    end: crate::raw::Kind,
) -> Result<Option<(usize, usize)>> {
    let mut iterator = Records::new(data);
    let mut start = None;
    let mut result = None;
    while let Some(record) = iterator.next() {
        let record = record?;
        if record.kind() == begin {
            if start.is_some() || result.is_some() {
                return Err(invalid(
                    "timeline BIFF12 integration",
                    "duplicate collection",
                ));
            }
            start = Some(record.offset());
        } else if record.kind() == end {
            let begin_offset = start.take().ok_or_else(|| {
                invalid("timeline BIFF12 integration", "end has no matching begin")
            })?;
            result = Some((begin_offset, iterator.offset()));
        }
    }
    if start.is_some() {
        return Err(Error::UnexpectedEndOfStream(
            "timeline BIFF12 integration".to_string(),
        ));
    }
    Ok(result)
}

fn parse_collection_refs(data: &[u8], typ: &str) -> Result<Vec<String>> {
    let mut iterator = Records::new(data);
    let mut inside = false;
    let mut item_open = false;
    let mut refs = Vec::new();
    while let Some(record) = iterator.next() {
        let record = record?;
        match record.kind() {
            kind::BEGIN_TIMELINE_CACHE_IDS => {
                if inside || !record.payload().is_empty() {
                    return Err(invalid(typ, "invalid collection begin"));
                }
                inside = true;
            },
            kind::BEGIN_TIMELINE_CACHE_ID => {
                if !inside || item_open {
                    return Err(invalid(typ, "invalid cache reference begin"));
                }
                refs.push(parse_rel_id(record.payload(), typ)?);
                item_open = true;
            },
            kind::END_TIMELINE_CACHE_ID => {
                if !item_open || !record.payload().is_empty() {
                    return Err(invalid(typ, "invalid cache reference end"));
                }
                item_open = false;
            },
            kind::END_TIMELINE_CACHE_IDS => {
                if !inside || item_open || !record.payload().is_empty() {
                    return Err(invalid(typ, "invalid collection end"));
                }
                inside = false;
            },
            _ if inside => return Err(invalid(typ, "unexpected record in collection")),
            _ => {},
        }
    }
    if inside || item_open {
        return Err(Error::UnexpectedEndOfStream(typ.to_string()));
    }
    let mut unique = HashSet::with_capacity(refs.len());
    if refs
        .iter()
        .any(|value| !unique.insert(value.to_ascii_lowercase()))
    {
        return Err(invalid(typ, "duplicate relationship ID"));
    }
    Ok(refs)
}

fn parse_single_ref(data: &[u8], typ: &str) -> Result<Option<String>> {
    let mut iterator = Records::new(data);
    let mut value = None;
    let mut open = false;
    while let Some(record) = iterator.next() {
        let record = record?;
        match record.kind() {
            kind::BEGIN_TIMELINE_EX => {
                if open || value.is_some() {
                    return Err(invalid(typ, "duplicate timeline reference"));
                }
                value = Some(parse_rel_id(record.payload(), typ)?);
                open = true;
            },
            kind::END_TIMELINE_EX => {
                if !open || !record.payload().is_empty() {
                    return Err(invalid(typ, "unbalanced timeline reference"));
                }
                open = false;
            },
            _ => {},
        }
    }
    if open {
        return Err(Error::UnexpectedEndOfStream(typ.to_string()));
    }
    Ok(value)
}

fn write_collection_refs(refs: &[String]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_TIMELINE_CACHE_IDS, &[])?;
    for value in refs {
        writer.write_record(kind::BEGIN_TIMELINE_CACHE_ID, &write_rel_id(value)?)?;
        writer.write_record(kind::END_TIMELINE_CACHE_ID, &[])?;
    }
    writer.write_record(kind::END_TIMELINE_CACHE_IDS, &[])?;
    Ok(output)
}

fn write_single_ref(value: &str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_TIMELINE_EX, &write_rel_id(value)?)?;
    writer.write_record(kind::END_TIMELINE_EX, &[])?;
    Ok(output)
}

fn rewrite_block(
    data: &[u8],
    existing: Option<(usize, usize)>,
    replacement: Option<&[u8]>,
    boundary: crate::raw::Kind,
) -> Result<Vec<u8>> {
    let (start, end) = if let Some(span) = existing {
        span
    } else if replacement.is_some() {
        let mut iterator = Records::new(data);
        let mut offset = None;
        while let Some(record) = iterator.next() {
            let record = record?;
            if record.kind() == boundary {
                if offset.replace(record.offset()).is_some() {
                    return Err(invalid(
                        "timeline BIFF12 integration",
                        "duplicate insertion boundary",
                    ));
                }
            }
        }
        let offset = offset
            .ok_or_else(|| invalid("timeline BIFF12 integration", "missing insertion boundary"))?;
        (offset, offset)
    } else {
        return Ok(data.to_vec());
    };
    let replacement_len = replacement.map_or(0, <[u8]>::len);
    let capacity = data
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or(Error::InvalidLength {
            expected: usize::MAX,
            found: data.len(),
        })?;
    let mut output = Vec::with_capacity(capacity);
    output.extend_from_slice(&data[..start]);
    if let Some(replacement) = replacement {
        output.extend_from_slice(replacement);
    }
    output.extend_from_slice(&data[end..]);
    Ok(output)
}

fn next_rel_id(part: &dyn Part) -> String {
    let mut index = 1u32;
    loop {
        let value = format!("rId{index}");
        if part.rels().get(&value).is_none() {
            return value;
        }
        index = index.saturating_add(1);
    }
}

fn next_part(
    package: &OpcPackage,
    directory: &str,
    stem: &str,
    extension: &str,
) -> Result<PackURI> {
    for index in 1..=65_536u32 {
        let uri = PackURI::new(format!("/{directory}/{stem}{index}.{extension}"))?;
        if package.validate_new_part_name(&uri).is_ok() {
            return Ok(uri);
        }
    }
    Err(invalid(
        "timeline part allocation",
        "part-name limit exceeded",
    ))
}

fn validate_cache_set(caches: &[Cache]) -> Result<()> {
    if caches.len() > MAX_CACHES {
        return Err(Error::InvalidLength {
            expected: MAX_CACHES,
            found: caches.len(),
        });
    }
    let mut names = HashSet::with_capacity(caches.len());
    for cache in caches {
        validate_cache(cache)?;
        if !names.insert(cache.name.to_ascii_lowercase()) {
            return Err(invalid("timeline cache collection", "duplicate cache name"));
        }
    }
    Ok(())
}

/// Load all workbook timeline cache parts.
pub fn load_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<CachePart>> {
    let workbook_part = package.get_part(workbook)?;
    let refs = parse_collection_refs(workbook_part.blob(), "timeline cache references")?;
    let mut output = Vec::with_capacity(refs.len());
    for relationship_id in &refs {
        let relationship = workbook_part.rels().get(relationship_id).ok_or_else(|| {
            invalid(
                "timeline cache references",
                format!("missing relationship {relationship_id}"),
            )
        })?;
        if relationship.is_external() || relationship.reltype() != CACHE_RELATIONSHIP_TYPE {
            return Err(invalid(
                "timeline cache relationship",
                "wrong type or external target",
            ));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != CACHE_CONTENT_TYPE || !part.rels().is_empty() {
            return Err(invalid(
                "timeline cache part",
                "wrong content type or outbound relationships",
            ));
        }
        output.push(CachePart {
            relationship_id: relationship_id.clone(),
            part_name: target.as_str().to_string(),
            cache: parse_cache(part.blob())?,
        });
    }
    let referenced: HashSet<_> = output.iter().map(|part| part.part_name.as_str()).collect();
    for part in package.iter_parts() {
        if part.content_type() == CACHE_CONTENT_TYPE
            && !referenced.contains(part.partname().as_str())
        {
            return Err(invalid("timeline cache package graph", "orphan cache part"));
        }
    }
    Ok(output)
}

fn remove_caches(package: &mut OpcPackage, workbook: &PackURI) -> Result<()> {
    let refs = parse_collection_refs(
        package.get_part(workbook)?.blob(),
        "timeline cache references",
    )?;
    let mut targets = Vec::with_capacity(refs.len());
    {
        let part = package.get_part(workbook)?;
        for relationship_id in &refs {
            let relationship = part.rels().get(relationship_id).ok_or_else(|| {
                invalid(
                    "timeline cache references",
                    format!("missing relationship {relationship_id}"),
                )
            })?;
            if relationship.is_external() || relationship.reltype() != CACHE_RELATIONSHIP_TYPE {
                return Err(invalid(
                    "timeline cache relationship",
                    "wrong type or external target",
                ));
            }
            targets.push((relationship_id.clone(), relationship.target_partname()?));
        }
    }
    crate::package::owner_transaction::require_exclusive_inbound(
        package,
        workbook,
        &targets,
        "timeline cache",
    )?;
    let part = package.get_part_mut(workbook)?;
    for (relationship_id, _) in &targets {
        part.rels_mut().remove(relationship_id);
    }
    for (_, target) in targets {
        package.remove_part(&target);
    }
    Ok(())
}

/// Replace workbook timeline caches and their BIFF12 references.
pub fn store_caches(package: &mut OpcPackage, workbook: &PackURI, caches: &[Cache]) -> Result<()> {
    validate_cache_set(caches)?;
    let _ = load_caches(package, workbook)?;
    let old_block = record_span(
        package.get_part(workbook)?.blob(),
        kind::BEGIN_TIMELINE_CACHE_IDS,
        kind::END_TIMELINE_CACHE_IDS,
    )?;
    let mut workbook_blob = package.get_part(workbook)?.blob().to_vec();
    remove_caches(package, workbook)?;
    if caches.is_empty() {
        let updated = rewrite_block(&workbook_blob, old_block, None, kind::END_BOOK)?;
        package.get_part_mut(workbook)?.set_blob(updated);
        package.unsign();
        return Ok(());
    }
    let mut used: HashSet<String> = package
        .get_part(workbook)?
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    let mut planned = Vec::with_capacity(caches.len());
    for cache in caches {
        let relationship_id = loop {
            let candidate = format!("rId{}", used.len() + 1);
            if used.insert(candidate.clone()) {
                break candidate;
            }
        };
        let uri = next_part(package, "xl/timelineCaches", "timelineCache", "xml")?;
        let bytes = write_cache(cache)?;
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            CACHE_CONTENT_TYPE.to_string(),
            bytes.clone(),
        )))?;
        planned.push((relationship_id, uri));
    }
    let refs: Vec<String> = planned.iter().map(|(id, _)| id.clone()).collect();
    let replacement = write_collection_refs(&refs)?;
    workbook_blob = rewrite_block(
        &workbook_blob,
        old_block,
        Some(&replacement),
        kind::END_BOOK,
    )?;
    let part = package.get_part_mut(workbook)?;
    for (relationship_id, uri) in &planned {
        part.rels_mut().add_relationship(
            CACHE_RELATIONSHIP_TYPE.to_string(),
            uri.relative_ref(workbook.base_uri()),
            relationship_id.clone(),
            false,
        );
    }
    part.set_blob(workbook_blob);
    package.unsign();
    Ok(())
}

/// Load one worksheet timeline view part, if attached.
pub fn load_views(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<ViewPart>> {
    let worksheet_part = package.get_part(worksheet)?;
    let reference = parse_single_ref(worksheet_part.blob(), "timeline views")?;
    let relationships: Vec<_> = worksheet_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == VIEWS_RELATIONSHIP_TYPE)
        .collect();
    if relationships.len() > 1 {
        return Err(invalid(
            "timeline views",
            "multiple worksheet relationships",
        ));
    }
    let Some(relationship) = relationships.into_iter().next() else {
        if reference.is_some() {
            return Err(invalid(
                "timeline views",
                "BIFF12 reference has no relationship",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() || reference.as_deref() != Some(relationship.r_id()) {
        return Err(invalid(
            "timeline views",
            "BIFF12 reference and relationship disagree",
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != VIEWS_CONTENT_TYPE || !part.rels().is_empty() {
        return Err(invalid(
            "timeline views",
            "wrong content type or outbound relationships",
        ));
    }
    Ok(Some(ViewPart {
        relationship_id: relationship.r_id().to_string(),
        part_name: target.as_str().to_string(),
        views: parse_views(part.blob())?,
    }))
}

/// Replace one worksheet's timeline view part and BIFF12 reference.
pub fn store_views(package: &mut OpcPackage, worksheet: &PackURI, views: &Views) -> Result<()> {
    validate_views(views)?;
    if let Some(existing) = load_views(package, worksheet)? {
        let target = PackURI::new(&existing.part_name)?;
        crate::package::owner_transaction::require_exclusive_inbound(
            package,
            worksheet,
            &[(existing.relationship_id, target)],
            "timeline view",
        )?;
    }
    if views.items.len() > MAX_VIEWS {
        return Err(Error::InvalidLength {
            expected: MAX_VIEWS,
            found: views.items.len(),
        });
    }
    let old_block = record_span(
        package.get_part(worksheet)?.blob(),
        kind::BEGIN_TIMELINE_EX,
        kind::END_TIMELINE_EX,
    )?;
    let mut worksheet_blob = package.get_part(worksheet)?.blob().to_vec();
    if let Some(old) = load_views(package, worksheet)? {
        let part = package.get_part_mut(worksheet)?;
        part.rels_mut().remove(&old.relationship_id);
        package.remove_part(&PackURI::new(&old.part_name)?);
    }
    if views.items.is_empty() {
        let updated = rewrite_block(&worksheet_blob, old_block, None, kind::END_SHEET)?;
        package.get_part_mut(worksheet)?.set_blob(updated);
        package.unsign();
        return Ok(());
    }
    let uri = next_part(package, "xl/timelines", "timeline", "xml")?;
    let relationship_id = next_rel_id(package.get_part(worksheet)?);
    package.try_add_part(Box::new(BlobPart::new(
        uri.clone(),
        VIEWS_CONTENT_TYPE.to_string(),
        write_views(views)?,
    )))?;
    let replacement = write_single_ref(&relationship_id)?;
    worksheet_blob = rewrite_block(
        &worksheet_blob,
        old_block,
        Some(&replacement),
        kind::END_SHEET,
    )?;
    let part = package.get_part_mut(worksheet)?;
    part.rels_mut().add_relationship(
        VIEWS_RELATIONSHIP_TYPE.to_string(),
        uri.relative_ref(worksheet.base_uri()),
        relationship_id,
        false,
    );
    part.set_blob(worksheet_blob);
    package.unsign();
    Ok(())
}
