#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! XLSB OPC ownership for slicer cache and worksheet slicer parts.
//!
//! The owner publishes only validated, inert snapshots. It edits workbook and
//! worksheet BIFF12 relationship references transactionally at the workbook
//! facade; it never evaluates a PivotCache or applies a selection.

use super::codec::{parse_cache, parse_views, write_cache, write_views};
use super::model::{Cache, MAX_CACHES, MAX_VIEWS, Views};
use super::validation::{cache as validate_cache, views as validate_views};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Records, Writer, kind};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::char;
use std::collections::HashSet;

/// XLSB slicer cache content type.
pub const CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicerCache";
/// XLSB slicer cache relationship type.
pub const CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
/// XLSB worksheet slicers content type.
pub const VIEWS_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicer";
/// XLSB worksheet slicers relationship type.
pub const VIEWS_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicer";

const MAX_REL_ID_UNITS: usize = 255;

/// A cache definition resolved from one workbook relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachePart {
    /// Workbook relationship identifier.
    pub relationship_id: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed cache snapshot.
    pub cache: Cache,
}

/// A worksheet slicer view part resolved from one worksheet relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ViewPart {
    /// Worksheet relationship identifier.
    pub relationship_id: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed view snapshot.
    pub views: Views,
}

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
        return Err(invalid(
            typ,
            format!("invalid relationship ID length {units}"),
        ));
    }
    let bytes = cursor.read_bytes(units.checked_mul(2).ok_or(Error::InvalidLength {
        expected: usize::MAX,
        found: units,
    })?)?;
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
    let mut payload = Vec::with_capacity(6 + units * 2);
    let mut writer = Writer::new(&mut payload);
    writer.write_u32(0x08)?;
    writer.write_u16(units as u16)?;
    for unit in value.encode_utf16() {
        writer.write_u16(unit)?;
    }
    Ok(payload)
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
                    "BIFF12 integration",
                    "duplicate reference collection",
                ));
            }
            start = Some(record.offset());
        } else if record.kind() == end {
            let begin_offset = start
                .take()
                .ok_or_else(|| invalid("BIFF12 integration", "end record has no matching begin"))?;
            result = Some((begin_offset, iterator.offset()));
        }
    }
    if start.is_some() {
        return Err(Error::UnexpectedEndOfStream(format!("end record {end}")));
    }
    Ok(result)
}

fn parse_collection_refs(
    data: &[u8],
    collection_begin: crate::raw::Kind,
    collection_end: crate::raw::Kind,
    item_begin: crate::raw::Kind,
    item_end: crate::raw::Kind,
    typ: &str,
) -> Result<Vec<String>> {
    let mut iterator = Records::new(data);
    let mut inside = false;
    let mut item_open = false;
    let mut refs = Vec::new();
    while let Some(record) = iterator.next() {
        let record = record?;
        match record.kind() {
            kind if kind == collection_begin => {
                if inside {
                    return Err(invalid(typ, "nested collection"));
                }
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                inside = true;
            },
            kind if kind == item_begin => {
                if !inside || item_open {
                    return Err(invalid(typ, "unbalanced item begin"));
                }
                refs.push(parse_rel_id(record.payload(), typ)?);
                item_open = true;
            },
            kind if kind == item_end => {
                if !item_open || !record.payload().is_empty() {
                    return Err(invalid(typ, "unbalanced item end"));
                }
                item_open = false;
            },
            kind if kind == collection_end => {
                if !inside || item_open || !record.payload().is_empty() {
                    return Err(invalid(typ, "unbalanced collection end"));
                }
                inside = false;
            },
            _ if inside => return Err(invalid(typ, "unexpected record in reference collection")),
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

fn parse_single_ref(
    data: &[u8],
    begin: crate::raw::Kind,
    end: crate::raw::Kind,
    typ: &str,
) -> Result<Option<String>> {
    let mut iterator = Records::new(data);
    let mut value = None;
    let mut open = false;
    while let Some(record) = iterator.next() {
        let record = record?;
        match record.kind() {
            kind if kind == begin => {
                if open || value.is_some() {
                    return Err(invalid(typ, "duplicate reference record"));
                }
                value = Some(parse_rel_id(record.payload(), typ)?);
                open = true;
            },
            kind if kind == end => {
                if !open || !record.payload().is_empty() {
                    return Err(invalid(typ, "unbalanced reference record"));
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

fn write_collection_refs(
    refs: &[String],
    collection_begin: crate::raw::Kind,
    collection_end: crate::raw::Kind,
    item_begin: crate::raw::Kind,
    item_end: crate::raw::Kind,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(collection_begin, &[])?;
    for value in refs {
        writer.write_record(item_begin, &write_rel_id(value)?)?;
        writer.write_record(item_end, &[])?;
    }
    writer.write_record(collection_end, &[])?;
    Ok(output)
}

fn write_single_ref(
    value: &str,
    begin: crate::raw::Kind,
    end: crate::raw::Kind,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(begin, &write_rel_id(value)?)?;
    writer.write_record(end, &[])?;
    Ok(output)
}

fn rewrite_block(
    data: &[u8],
    existing: Option<(usize, usize)>,
    replacement: Option<&[u8]>,
    insert_before: crate::raw::Kind,
) -> Result<Vec<u8>> {
    let (start, end) = if let Some(span) = existing {
        span
    } else if replacement.is_some() {
        let mut iterator = Records::new(data);
        let mut offset = None;
        while let Some(record) = iterator.next() {
            let record = record?;
            if record.kind() == insert_before {
                if offset.replace(record.offset()).is_some() {
                    return Err(invalid(
                        "BIFF12 integration",
                        "duplicate insertion boundary",
                    ));
                }
            }
        }
        let offset =
            offset.ok_or_else(|| invalid("BIFF12 integration", "missing insertion boundary"))?;
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
        let candidate = format!("rId{index}");
        if part.rels().get(&candidate).is_none() {
            return candidate;
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
    Err(invalid("XLSB part allocation", "part-name limit exceeded"))
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
            return Err(invalid("slicer cache collection", "duplicate cache name"));
        }
    }
    Ok(())
}

fn load_cache_refs(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<String>> {
    let part = package.get_part(workbook)?;
    parse_collection_refs(
        part.blob(),
        kind::BEGIN_SLICER_CACHE_IDS,
        kind::END_SLICER_CACHE_IDS,
        kind::BEGIN_SLICER_CACHE_ID,
        kind::END_SLICER_CACHE_ID,
        "slicer cache references",
    )
}

/// Load all workbook slicer cache parts.
pub fn load_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<CachePart>> {
    let refs = load_cache_refs(package, workbook)?;
    let workbook_part = package.get_part(workbook)?;
    let mut output = Vec::with_capacity(refs.len());
    for relationship_id in &refs {
        let relationship = workbook_part.rels().get(relationship_id).ok_or_else(|| {
            invalid(
                "slicer cache references",
                format!("missing relationship {relationship_id}"),
            )
        })?;
        if relationship.is_external() || relationship.reltype() != CACHE_RELATIONSHIP_TYPE {
            return Err(invalid(
                "slicer cache relationship",
                "wrong type or external target",
            ));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != CACHE_CONTENT_TYPE || !part.rels().is_empty() {
            return Err(invalid(
                "slicer cache part",
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
            return Err(invalid("slicer cache package graph", "orphan cache part"));
        }
    }
    Ok(output)
}

fn remove_cache_parts(package: &mut OpcPackage, workbook: &PackURI) -> Result<Vec<String>> {
    let refs = load_cache_refs(package, workbook)?;
    let mut targets = Vec::with_capacity(refs.len());
    {
        let part = package.get_part(workbook)?;
        for relationship_id in &refs {
            let relationship = part.rels().get(relationship_id).ok_or_else(|| {
                invalid(
                    "slicer cache references",
                    format!("missing relationship {relationship_id}"),
                )
            })?;
            if relationship.is_external() || relationship.reltype() != CACHE_RELATIONSHIP_TYPE {
                return Err(invalid(
                    "slicer cache relationship",
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
        "slicer cache",
    )?;
    let part = package.get_part_mut(workbook)?;
    for (relationship_id, _) in &targets {
        part.rels_mut().remove(relationship_id);
    }
    for (_, target) in &targets {
        package.remove_part(target);
    }
    Ok(refs)
}

/// Replace workbook slicer caches and their BIFF12 relationship collection.
pub fn store_caches(package: &mut OpcPackage, workbook: &PackURI, caches: &[Cache]) -> Result<()> {
    validate_cache_set(caches)?;
    let _ = load_caches(package, workbook)?;
    let old_refs = load_cache_refs(package, workbook)?;
    let old_block = record_span(
        package.get_part(workbook)?.blob(),
        kind::BEGIN_SLICER_CACHE_IDS,
        kind::END_SLICER_CACHE_IDS,
    )?;
    let mut workbook_blob = package.get_part(workbook)?.blob().to_vec();
    remove_cache_parts(package, workbook)?;
    let _ = old_refs;
    if caches.is_empty() {
        workbook_blob = rewrite_block(&workbook_blob, old_block, None, kind::END_BOOK)?;
        package.get_part_mut(workbook)?.set_blob(workbook_blob);
        package.unsign();
        return Ok(());
    }

    let mut planned = Vec::with_capacity(caches.len());
    let mut used_relationship_ids: HashSet<String> = package
        .get_part(workbook)?
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    for cache in caches {
        let uri = next_part(package, "xl/slicerCaches", "slicerCache", "bin")?;
        let relationship_id = loop {
            let candidate = format!("rId{}", used_relationship_ids.len() + 1);
            if used_relationship_ids.insert(candidate.clone()) {
                break candidate;
            }
        };
        let bytes = write_cache(cache)?;
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            CACHE_CONTENT_TYPE.to_string(),
            bytes.clone(),
        )))?;
        planned.push((relationship_id, uri, bytes));
    }
    let refs: Vec<String> = planned.iter().map(|(id, _, _)| id.clone()).collect();
    let replacement = write_collection_refs(
        &refs,
        kind::BEGIN_SLICER_CACHE_IDS,
        kind::END_SLICER_CACHE_IDS,
        kind::BEGIN_SLICER_CACHE_ID,
        kind::END_SLICER_CACHE_ID,
    )?;
    workbook_blob = rewrite_block(
        &workbook_blob,
        old_block,
        Some(&replacement),
        kind::END_BOOK,
    )?;
    {
        let part = package.get_part_mut(workbook)?;
        for (relationship_id, uri, _) in &planned {
            part.rels_mut().add_relationship(
                CACHE_RELATIONSHIP_TYPE.to_string(),
                uri.relative_ref(workbook.base_uri()),
                relationship_id.clone(),
                false,
            );
        }
        part.set_blob(workbook_blob);
    }
    package.unsign();
    Ok(())
}

fn load_view_ref(
    package: &OpcPackage,
    worksheet: &PackURI,
    relation_type: &str,
    begin: crate::raw::Kind,
    end: crate::raw::Kind,
    typ: &str,
) -> Result<Option<ViewPart>> {
    let worksheet_part = package.get_part(worksheet)?;
    let reference = parse_single_ref(worksheet_part.blob(), begin, end, typ)?;
    let relationships: Vec<_> = worksheet_part
        .rels()
        .iter()
        .filter(|rel| rel.reltype() == relation_type)
        .collect();
    if relationships.len() > 1 {
        return Err(invalid(typ, "multiple worksheet relationships"));
    }
    let Some(relationship) = relationships.into_iter().next() else {
        if reference.is_some() {
            return Err(invalid(typ, "BIFF12 reference has no relationship"));
        }
        return Ok(None);
    };
    if relationship.is_external() || reference.as_deref() != Some(relationship.r_id()) {
        return Err(invalid(typ, "BIFF12 reference and relationship disagree"));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != VIEWS_CONTENT_TYPE || !part.rels().is_empty() {
        return Err(invalid(typ, "wrong content type or outbound relationships"));
    }
    Ok(Some(ViewPart {
        relationship_id: relationship.r_id().to_string(),
        part_name: target.as_str().to_string(),
        views: parse_views(part.blob())?,
    }))
}

/// Load worksheet slicer views, if a slicers part is attached.
pub fn load_views(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<ViewPart>> {
    load_view_ref(
        package,
        worksheet,
        VIEWS_RELATIONSHIP_TYPE,
        kind::BEGIN_SLICER_EX,
        kind::END_SLICER_EX,
        "slicer views",
    )
}

/// Replace a worksheet's slicer views and relationship references.
pub fn store_views(package: &mut OpcPackage, worksheet: &PackURI, views: &Views) -> Result<()> {
    validate_views(views)?;
    if let Some(existing) = load_views(package, worksheet)? {
        let target = PackURI::new(&existing.part_name)?;
        crate::package::owner_transaction::require_exclusive_inbound(
            package,
            worksheet,
            &[(existing.relationship_id, target)],
            "slicer view",
        )?;
    }
    if views.items.len() > MAX_VIEWS {
        return Err(Error::InvalidLength {
            expected: MAX_VIEWS,
            found: views.items.len(),
        });
    }
    let old = load_views(package, worksheet)?;
    let old_block = record_span(
        package.get_part(worksheet)?.blob(),
        kind::BEGIN_SLICER_EX,
        kind::END_SLICER_EX,
    )?;
    let mut worksheet_blob = package.get_part(worksheet)?.blob().to_vec();
    if let Some(old) = old {
        let part = package.get_part_mut(worksheet)?;
        part.rels_mut().remove(&old.relationship_id);
        package.remove_part(&PackURI::new(&old.part_name)?);
    }
    if views.items.is_empty() {
        worksheet_blob = rewrite_block(&worksheet_blob, old_block, None, kind::END_SHEET)?;
        package.get_part_mut(worksheet)?.set_blob(worksheet_blob);
        package.unsign();
        return Ok(());
    }
    let uri = next_part(package, "xl/slicers", "slicer", "bin")?;
    let relationship_id = next_rel_id(package.get_part(worksheet)?);
    let bytes = write_views(views)?;
    package.try_add_part(Box::new(BlobPart::new(
        uri.clone(),
        VIEWS_CONTENT_TYPE.to_string(),
        bytes,
    )))?;
    let replacement =
        write_single_ref(&relationship_id, kind::BEGIN_SLICER_EX, kind::END_SLICER_EX)?;
    worksheet_blob = rewrite_block(
        &worksheet_blob,
        old_block,
        Some(&replacement),
        kind::END_SHEET,
    )?;
    {
        let part = package.get_part_mut(worksheet)?;
        part.rels_mut().add_relationship(
            VIEWS_RELATIONSHIP_TYPE.to_string(),
            uri.relative_ref(worksheet.base_uri()),
            relationship_id,
            false,
        );
        part.set_blob(worksheet_blob);
    }
    package.unsign();
    Ok(())
}
