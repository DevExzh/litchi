//! Granular package CRUD for slicers, slicer caches, timelines, and timeline caches.

use std::collections::{HashMap, HashSet};

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{OoxmlError, Result};
use crate::xlsx::parsers::workbook_parser;
use crate::xlsx::pivot::read_pivot_tables;
use crate::xlsx::slicer_cache::{
    SlicerCacheDefinition, WorkbookSlicerCache, load_slicer_caches, store_slicer_cache,
    write_slicer_cache_definition,
};
use crate::xlsx::slicers::{
    SLICERS_CONTENT_TYPE, Slicer, Slicers, WorksheetSlicers, load_worksheet_slicers,
    store_worksheet_slicers, write_slicers,
};
use crate::xlsx::timelines::{
    TIMELINE_CACHE_CONTENT_TYPE, TIMELINE_CACHE_EXTENSION_URI,
    TIMELINE_CACHE_RELATIONSHIP_TYPE, TIMELINES_EXTENSION_URI, Timeline,
    TimelineCacheDefinition, Timelines,
    WorkbookTimelineCache, WorksheetTimelines, load_timeline_caches, load_timelines,
    store_timeline_caches, store_worksheet_timelines, write_timeline_cache_definition,
    write_timelines,
};

const SLICER_CACHE_EXTENSION_URI: &str = "{BBE1A952-AA13-448E-AADC-164F8A28A991}";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;

pub fn find_slicer(
    package: &OpcPackage,
    worksheet_name: &PackURI,
    slicer_name: &str,
) -> Result<Option<Slicer>> {
    Ok(load_worksheet_slicers(package, worksheet_name)?
        .into_iter()
        .flat_map(|part| part.slicers.slicers)
        .find(|slicer| slicer.name.eq_ignore_ascii_case(slicer_name)))
}

pub fn add_slicer(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    slicer: Slicer,
) -> Result<WorksheetSlicers> {
    let caches = load_slicer_caches(package)?;
    validate_slicer_cache_reference(&slicer, &caches)?;
    let mut parts = load_worksheet_slicers(package, worksheet_name)?;
    if all_slicer_names(package)?.contains(&slicer.name.to_ascii_lowercase()) {
        return Err(invalid(format!("duplicate slicer name '{}'", slicer.name)));
    }
    if let Some(part) = parts.first_mut() {
        part.slicers.slicers.push(slicer);
        let xml = write_slicers(&part.slicers)?;
        let staged = part.clone();
        validate_slicer_views(package, Some((&staged.part_name, &staged.slicers)), &caches)?;
        package
            .get_part_mut(&PackURI::new(&staged.part_name).map_err(OoxmlError::InvalidUri)?)?
            .set_blob(xml);
        let _ = package.clear_digital_signatures();
        return Ok(staged);
    }
    let part_name = next_part_name(package, "/xl/slicers/slicer%d.xml")?;
    let relationship_id = next_relationship_id(package.get_part(worksheet_name)?, "rIdSlicer")?;
    let value = WorksheetSlicers {
        relationship_id,
        part_name: part_name.to_string(),
        slicers: Slicers::new(vec![slicer]),
    };
    store_worksheet_slicers(package, worksheet_name, &value)?;
    let _ = package.clear_digital_signatures();
    Ok(value)
}

pub fn update_slicer<F>(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    slicer_name: &str,
    update: F,
) -> Result<bool>
where
    F: FnOnce(&mut Slicer),
{
    let caches = load_slicer_caches(package)?;
    let mut parts = load_worksheet_slicers(package, worksheet_name)?;
    let mut update = Some(update);
    let mut changed = None;
    for (part_index, part) in parts.iter_mut().enumerate() {
        if let Some(slicer) = part
            .slicers
            .slicers
            .iter_mut()
            .find(|slicer| slicer.name.eq_ignore_ascii_case(slicer_name))
        {
            update.take().expect("used once")(slicer);
            if !slicer.name.eq_ignore_ascii_case(slicer_name) {
                return Err(invalid("slicer update cannot change its name"));
            }
            changed = Some(part_index);
            break;
        }
    }
    let Some(index) = changed else { return Ok(false); };
    validate_slicer_views(package, Some((&parts[index].part_name, &parts[index].slicers)), &caches)?;
    let xml = write_slicers(&parts[index].slicers)?;
    let part_name = PackURI::new(&parts[index].part_name).map_err(OoxmlError::InvalidUri)?;
    package.get_part_mut(&part_name)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn replace_slicer(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    slicer_name: &str,
    replacement: Slicer,
) -> Result<bool> {
    if !replacement.name.eq_ignore_ascii_case(slicer_name) {
        return Err(invalid("replacement slicer name must match"));
    }
    update_slicer(package, worksheet_name, slicer_name, move |slicer| *slicer = replacement)
}

pub fn remove_slicer(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    slicer_name: &str,
) -> Result<bool> {
    let mut parts = load_worksheet_slicers(package, worksheet_name)?;
    let Some(part_index) = parts.iter().position(|part| {
        part.slicers.slicers.iter().any(|slicer| slicer.name.eq_ignore_ascii_case(slicer_name))
    }) else { return Ok(false); };
    let item_index = parts[part_index]
        .slicers
        .slicers
        .iter()
        .position(|slicer| slicer.name.eq_ignore_ascii_case(slicer_name))
        .expect("located");
    parts[part_index].slicers.slicers.remove(item_index);
    if !parts[part_index].slicers.slicers.is_empty() {
        let xml = write_slicers(&parts[part_index].slicers)?;
        let uri = PackURI::new(&parts[part_index].part_name).map_err(OoxmlError::InvalidUri)?;
        package.get_part_mut(&uri)?.set_blob(xml);
    } else {
        let removed = parts.remove(part_index);
        let uri = PackURI::new(&removed.part_name).map_err(OoxmlError::InvalidUri)?;
        package
            .get_part_mut(worksheet_name)?
            .rels_mut()
            .remove(&removed.relationship_id);
        if !part_is_referenced(package, &uri) {
            package.remove_part(&uri);
        }
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn reorder_slicers(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    ordered_names: &[String],
) -> Result<Vec<Slicer>> {
    let caches = load_slicer_caches(package)?;
    let mut parts = load_worksheet_slicers(package, worksheet_name)?;
    let counts: Vec<usize> = parts.iter().map(|part| part.slicers.slicers.len()).collect();
    let flattened: Vec<Slicer> = parts
        .iter_mut()
        .flat_map(|part| std::mem::take(&mut part.slicers.slicers))
        .collect();
    let ordered = reorder_by_key(flattened, ordered_names, |slicer| &slicer.name, "slicer")?;
    let mut offset = 0usize;
    for (part, count) in parts.iter_mut().zip(counts) {
        part.slicers.slicers = ordered[offset..offset + count].to_vec();
        offset += count;
    }
    validate_slicer_views(package, None, &caches)?;
    let plans: Vec<(PackURI, Vec<u8>)> = parts
        .iter()
        .map(|part| {
            Ok((
                PackURI::new(&part.part_name).map_err(OoxmlError::InvalidUri)?,
                write_slicers(&part.slicers)?,
            ))
        })
        .collect::<Result<_>>()?;
    for (uri, xml) in plans { package.get_part_mut(&uri)?.set_blob(xml); }
    let _ = package.clear_digital_signatures();
    Ok(ordered)
}

pub fn find_slicer_cache(package: &OpcPackage, name: &str) -> Result<Option<WorkbookSlicerCache>> {
    Ok(load_slicer_caches(package)?
        .into_iter()
        .find(|cache| cache.definition.name.eq_ignore_ascii_case(name)))
}

pub fn add_slicer_cache(
    package: &mut OpcPackage,
    definition: SlicerCacheDefinition,
) -> Result<WorkbookSlicerCache> {
    let workbook_name = package.main_document_part()?.partname().clone();
    let value = WorkbookSlicerCache {
        relationship_id: next_relationship_id(package.get_part(&workbook_name)?, "rIdSlicerCache")?,
        part_name: next_part_name(package, "/xl/slicerCaches/slicerCache%d.xml")?.to_string(),
        definition,
    };
    validate_slicer_cache_pivot_links(package, std::slice::from_ref(&value))?;
    store_slicer_cache(package, &value)?;
    let _ = package.clear_digital_signatures();
    Ok(value)
}

pub fn update_slicer_cache<F>(package: &mut OpcPackage, name: &str, update: F) -> Result<bool>
where
    F: FnOnce(&mut SlicerCacheDefinition),
{
    let mut caches = load_slicer_caches(package)?;
    let Some(index) = caches.iter().position(|cache| cache.definition.name.eq_ignore_ascii_case(name)) else {
        return Ok(false);
    };
    update(&mut caches[index].definition);
    if !caches[index].definition.name.eq_ignore_ascii_case(name) {
        return Err(invalid("Slicer Cache update cannot change its name"));
    }
    validate_slicer_cache_set(package, &caches)?;
    let xml = write_slicer_cache_definition(&caches[index].definition)?;
    let uri = PackURI::new(&caches[index].part_name).map_err(OoxmlError::InvalidUri)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn replace_slicer_cache(
    package: &mut OpcPackage,
    name: &str,
    replacement: SlicerCacheDefinition,
) -> Result<bool> {
    if !replacement.name.eq_ignore_ascii_case(name) {
        return Err(invalid("replacement Slicer Cache name must match"));
    }
    update_slicer_cache(package, name, move |definition| *definition = replacement)
}

pub fn remove_slicer_cache(package: &mut OpcPackage, name: &str) -> Result<bool> {
    let mut caches = load_slicer_caches(package)?;
    let Some(index) = caches.iter().position(|cache| cache.definition.name.eq_ignore_ascii_case(name)) else {
        return Ok(false);
    };
    if all_slicers(package)?.iter().any(|slicer| slicer.cache.eq_ignore_ascii_case(name)) {
        return Err(invalid(format!("Slicer Cache '{name}' is still referenced")));
    }
    let removed = caches.remove(index);
    validate_slicer_cache_set(package, &caches)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let ids: Vec<String> = caches.iter().map(|cache| cache.relationship_id.clone()).collect();
    let updated = rewrite_integration_refs(
        package.get_part(&workbook_name)?.blob(),
        SLICER_CACHE_EXTENSION_URI,
        X14,
        "slicerCaches",
        "slicerCache",
        &ids,
    )?;
    package.get_part_mut(&workbook_name)?.set_blob(updated);
    package.get_part_mut(&workbook_name)?.rels_mut().remove(&removed.relationship_id);
    let uri = PackURI::new(&removed.part_name).map_err(OoxmlError::InvalidUri)?;
    if !part_is_referenced(package, &uri) { package.remove_part(&uri); }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn reorder_slicer_caches(
    package: &mut OpcPackage,
    ordered_names: &[String],
) -> Result<Vec<WorkbookSlicerCache>> {
    let caches = reorder_by_key(load_slicer_caches(package)?, ordered_names, |cache| &cache.definition.name, "Slicer Cache")?;
    validate_slicer_cache_set(package, &caches)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let ids: Vec<String> = caches.iter().map(|cache| cache.relationship_id.clone()).collect();
    let updated = rewrite_integration_refs(
        package.get_part(&workbook_name)?.blob(),
        SLICER_CACHE_EXTENSION_URI,
        X14,
        "slicerCaches",
        "slicerCache",
        &ids,
    )?;
    package.get_part_mut(&workbook_name)?.set_blob(updated);
    let _ = package.clear_digital_signatures();
    Ok(caches)
}

pub fn find_timeline_cache(package: &OpcPackage, name: &str) -> Result<Option<WorkbookTimelineCache>> {
    let workbook = package.main_document_part()?.partname().clone();
    Ok(load_timeline_caches(package, &workbook)?
        .into_iter()
        .find(|cache| cache.definition.name.eq_ignore_ascii_case(name)))
}

pub fn add_timeline_cache(
    package: &mut OpcPackage,
    definition: TimelineCacheDefinition,
) -> Result<WorkbookTimelineCache> {
    let workbook = package.main_document_part()?.partname().clone();
    let mut caches = load_timeline_caches(package, &workbook)?;
    let value = WorkbookTimelineCache {
        relationship_id: next_relationship_id(package.get_part(&workbook)?, "rIdTimelineCache")?,
        part_name: next_part_name(package, "/xl/timelineCaches/timelineCache%d.xml")?.to_string(),
        definition,
    };
    validate_timeline_cache_set(package, std::slice::from_ref(&value))?;
    if caches.is_empty() {
        store_timeline_caches(package, &workbook, std::slice::from_ref(&value))?;
        let _ = package.clear_digital_signatures();
        return Ok(value);
    }
    caches.push(value.clone());
    validate_timeline_cache_set(package, &caches)?;
    let ids: Vec<String> = caches.iter().map(|cache| cache.relationship_id.clone()).collect();
    let updated = rewrite_integration_refs(
        package.get_part(&workbook)?.blob(),
        TIMELINE_CACHE_EXTENSION_URI,
        X15,
        "timelineCacheRefs",
        "timelineCacheRef",
        &ids,
    )?;
    let uri = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    let xml = write_timeline_cache_definition(&value.definition)?;
    package.try_add_part(Box::new(BlobPart::new(uri.clone(), TIMELINE_CACHE_CONTENT_TYPE.into(), xml)))?;
    package.get_part_mut(&workbook)?.rels_mut().add_relationship(
        TIMELINE_CACHE_RELATIONSHIP_TYPE.into(),
        uri.relative_ref(workbook.base_uri()),
        value.relationship_id.clone(),
        false,
    );
    package.get_part_mut(&workbook)?.set_blob(updated);
    let _ = package.clear_digital_signatures();
    Ok(value)
}

pub fn update_timeline_cache<F>(package: &mut OpcPackage, name: &str, update: F) -> Result<bool>
where
    F: FnOnce(&mut TimelineCacheDefinition),
{
    let workbook = package.main_document_part()?.partname().clone();
    let mut caches = load_timeline_caches(package, &workbook)?;
    let Some(index) = caches.iter().position(|cache| cache.definition.name.eq_ignore_ascii_case(name)) else { return Ok(false); };
    update(&mut caches[index].definition);
    if !caches[index].definition.name.eq_ignore_ascii_case(name) {
        return Err(invalid("Timeline Cache update cannot change its name"));
    }
    validate_timeline_cache_set(package, &caches)?;
    let xml = write_timeline_cache_definition(&caches[index].definition)?;
    let uri = PackURI::new(&caches[index].part_name).map_err(OoxmlError::InvalidUri)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn replace_timeline_cache(package: &mut OpcPackage, name: &str, replacement: TimelineCacheDefinition) -> Result<bool> {
    if !replacement.name.eq_ignore_ascii_case(name) {
        return Err(invalid("replacement Timeline Cache name must match"));
    }
    update_timeline_cache(package, name, move |definition| *definition = replacement)
}

pub fn remove_timeline_cache(package: &mut OpcPackage, name: &str) -> Result<bool> {
    let workbook = package.main_document_part()?.partname().clone();
    let mut caches = load_timeline_caches(package, &workbook)?;
    let Some(index) = caches.iter().position(|cache| cache.definition.name.eq_ignore_ascii_case(name)) else { return Ok(false); };
    if load_timelines(package, &workbook)?.iter().any(|sheet| {
        sheet.timelines.timelines.iter().any(|timeline| timeline.cache.eq_ignore_ascii_case(name))
    }) {
        return Err(invalid(format!("Timeline Cache '{name}' is still referenced")));
    }
    let removed = caches.remove(index);
    validate_timeline_cache_set(package, &caches)?;
    let ids: Vec<String> = caches.iter().map(|cache| cache.relationship_id.clone()).collect();
    let updated = rewrite_integration_refs(
        package.get_part(&workbook)?.blob(),
        TIMELINE_CACHE_EXTENSION_URI,
        X15,
        "timelineCacheRefs",
        "timelineCacheRef",
        &ids,
    )?;
    package.get_part_mut(&workbook)?.set_blob(updated);
    package.get_part_mut(&workbook)?.rels_mut().remove(&removed.relationship_id);
    let uri = PackURI::new(&removed.part_name).map_err(OoxmlError::InvalidUri)?;
    if !part_is_referenced(package, &uri) { package.remove_part(&uri); }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn reorder_timeline_caches(package: &mut OpcPackage, ordered_names: &[String]) -> Result<Vec<WorkbookTimelineCache>> {
    let workbook = package.main_document_part()?.partname().clone();
    let caches = reorder_by_key(load_timeline_caches(package, &workbook)?, ordered_names, |cache| &cache.definition.name, "Timeline Cache")?;
    validate_timeline_cache_set(package, &caches)?;
    let ids: Vec<String> = caches.iter().map(|cache| cache.relationship_id.clone()).collect();
    let updated = rewrite_integration_refs(package.get_part(&workbook)?.blob(), TIMELINE_CACHE_EXTENSION_URI, X15, "timelineCacheRefs", "timelineCacheRef", &ids)?;
    package.get_part_mut(&workbook)?.set_blob(updated);
    let _ = package.clear_digital_signatures();
    Ok(caches)
}

pub fn find_timeline(package: &OpcPackage, worksheet: &PackURI, name: &str) -> Result<Option<Timeline>> {
    let workbook = package.main_document_part()?.partname().clone();
    Ok(load_timelines(package, &workbook)?
        .into_iter()
        .find(|sheet| sheet.worksheet_part_name == worksheet.to_string())
        .and_then(|sheet| sheet.timelines.timelines.into_iter().find(|timeline| timeline.name.eq_ignore_ascii_case(name))))
}

pub fn add_timeline(package: &mut OpcPackage, worksheet: &PackURI, timeline: Timeline) -> Result<WorksheetTimelines> {
    let workbook = package.main_document_part()?.partname().clone();
    let caches = load_timeline_caches(package, &workbook)?;
    validate_timeline_cache_reference(&timeline, &caches)?;
    let mut sheets = load_timelines(package, &workbook)?;
    if sheets.iter().flat_map(|sheet| &sheet.timelines.timelines).any(|candidate| candidate.name.eq_ignore_ascii_case(&timeline.name)) {
        return Err(invalid(format!("duplicate timeline name '{}'", timeline.name)));
    }
    if let Some(index) = sheets.iter().position(|sheet| sheet.worksheet_part_name == worksheet.to_string()) {
        sheets[index].timelines.timelines.push(timeline);
        validate_timeline_views(&sheets, &caches)?;
        let xml = write_timelines(&sheets[index].timelines)?;
        let uri = PackURI::new(&sheets[index].part_name).map_err(OoxmlError::InvalidUri)?;
        package.get_part_mut(&uri)?.set_blob(xml);
        let _ = package.clear_digital_signatures();
        return Ok(sheets[index].clone());
    }
    let value = WorksheetTimelines {
        worksheet_part_name: worksheet.to_string(),
        relationship_id: next_relationship_id(package.get_part(worksheet)?, "rIdTimeline")?,
        part_name: next_part_name(package, "/xl/timelines/timeline%d.xml")?.to_string(),
        timelines: Timelines { timelines: vec![timeline] },
    };
    store_worksheet_timelines(package, &workbook, &value)?;
    let _ = package.clear_digital_signatures();
    Ok(value)
}

pub fn update_timeline<F>(package: &mut OpcPackage, worksheet: &PackURI, name: &str, update: F) -> Result<bool>
where F: FnOnce(&mut Timeline) {
    let workbook = package.main_document_part()?.partname().clone();
    let caches = load_timeline_caches(package, &workbook)?;
    let mut sheets = load_timelines(package, &workbook)?;
    let Some(sheet_index) = sheets.iter().position(|sheet| sheet.worksheet_part_name == worksheet.to_string()) else { return Ok(false); };
    let Some(index) = sheets[sheet_index].timelines.timelines.iter().position(|timeline| timeline.name.eq_ignore_ascii_case(name)) else { return Ok(false); };
    update(&mut sheets[sheet_index].timelines.timelines[index]);
    if !sheets[sheet_index].timelines.timelines[index].name.eq_ignore_ascii_case(name) {
        return Err(invalid("timeline update cannot change its name"));
    }
    validate_timeline_views(&sheets, &caches)?;
    let xml = write_timelines(&sheets[sheet_index].timelines)?;
    let uri = PackURI::new(&sheets[sheet_index].part_name).map_err(OoxmlError::InvalidUri)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn replace_timeline(package: &mut OpcPackage, worksheet: &PackURI, name: &str, replacement: Timeline) -> Result<bool> {
    if !replacement.name.eq_ignore_ascii_case(name) { return Err(invalid("replacement timeline name must match")); }
    update_timeline(package, worksheet, name, move |timeline| *timeline = replacement)
}

pub fn remove_timeline(package: &mut OpcPackage, worksheet: &PackURI, name: &str) -> Result<bool> {
    let workbook = package.main_document_part()?.partname().clone();
    let caches = load_timeline_caches(package, &workbook)?;
    let mut sheets = load_timelines(package, &workbook)?;
    let Some(sheet_index) = sheets.iter().position(|sheet| sheet.worksheet_part_name == worksheet.to_string()) else { return Ok(false); };
    let Some(index) = sheets[sheet_index].timelines.timelines.iter().position(|timeline| timeline.name.eq_ignore_ascii_case(name)) else { return Ok(false); };
    sheets[sheet_index].timelines.timelines.remove(index);
    if !sheets[sheet_index].timelines.timelines.is_empty() {
        validate_timeline_views(&sheets, &caches)?;
        let xml = write_timelines(&sheets[sheet_index].timelines)?;
        let uri = PackURI::new(&sheets[sheet_index].part_name).map_err(OoxmlError::InvalidUri)?;
        package.get_part_mut(&uri)?.set_blob(xml);
    } else {
        let removed = sheets.remove(sheet_index);
        validate_timeline_views(&sheets, &caches)?;
        let updated = rewrite_integration_refs(package.get_part(worksheet)?.blob(), TIMELINES_EXTENSION_URI, X15, "timelineRefs", "timelineRef", &[])?;
        package.get_part_mut(worksheet)?.set_blob(updated);
        package.get_part_mut(worksheet)?.rels_mut().remove(&removed.relationship_id);
        let uri = PackURI::new(&removed.part_name).map_err(OoxmlError::InvalidUri)?;
        if !part_is_referenced(package, &uri) { package.remove_part(&uri); }
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub fn reorder_timelines(package: &mut OpcPackage, worksheet: &PackURI, ordered_names: &[String]) -> Result<Vec<Timeline>> {
    let workbook = package.main_document_part()?.partname().clone();
    let caches = load_timeline_caches(package, &workbook)?;
    let mut sheets = load_timelines(package, &workbook)?;
    let Some(sheet_index) = sheets.iter().position(|sheet| sheet.worksheet_part_name == worksheet.to_string()) else {
        if ordered_names.is_empty() { return Ok(Vec::new()); }
        return Err(invalid("worksheet has no Timelines part"));
    };
    let ordered = reorder_by_key(std::mem::take(&mut sheets[sheet_index].timelines.timelines), ordered_names, |timeline| &timeline.name, "timeline")?;
    sheets[sheet_index].timelines.timelines = ordered.clone();
    validate_timeline_views(&sheets, &caches)?;
    let xml = write_timelines(&sheets[sheet_index].timelines)?;
    let uri = PackURI::new(&sheets[sheet_index].part_name).map_err(OoxmlError::InvalidUri)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(ordered)
}

fn validate_slicer_cache_set(package: &OpcPackage, caches: &[WorkbookSlicerCache]) -> Result<()> {
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    let any_uid = caches.iter().any(|cache| cache.definition.uid.is_some());
    for cache in caches {
        write_slicer_cache_definition(&cache.definition)?;
        if !names.insert(cache.definition.name.to_ascii_lowercase()) { return Err(invalid("duplicate Slicer Cache name")); }
        match &cache.definition.uid {
            Some(uid) if !uids.insert(uid.to_ascii_lowercase()) => return Err(invalid("duplicate Slicer Cache uid")),
            None if any_uid => return Err(invalid("Slicer Cache uid must be present on every cache or none")),
            _ => {},
        }
    }
    validate_slicer_views(package, None, caches)?;
    validate_slicer_cache_pivot_links(package, caches)
}

fn validate_slicer_cache_pivot_links(package: &OpcPackage, caches: &[WorkbookSlicerCache]) -> Result<()> {
    let workbook = package.main_document_part()?;
    let details = workbook_parser::parse_workbook_details(std::str::from_utf8(workbook.blob()).map_err(|e| invalid(e.to_string()))?)
        .map_err(|e| invalid(e.to_string()))?;
    if details.pivot_caches.is_empty() { return Ok(()); }
    let sheet_ids: HashMap<String, u32> = details.sheets.into_iter().map(|sheet| (sheet.name, sheet.sheet_id)).collect();
    let pivots = read_pivot_tables(package).map_err(|e| invalid(e.to_string()))?;
    let bindings: HashSet<(u32, String)> = pivots.into_iter().filter_map(|pivot| {
        sheet_ids.get(&pivot.sheet_name).map(|id| (*id, pivot.name.to_ascii_lowercase()))
    }).collect();
    for cache in caches {
        for pivot in &cache.definition.pivot_tables {
            if !bindings.contains(&(pivot.tab_id, pivot.name.to_ascii_lowercase())) {
                return Err(invalid(format!("Slicer Cache '{}' references missing PivotTable '{}' on sheet ID {}", cache.definition.name, pivot.name, pivot.tab_id)));
            }
        }
    }
    Ok(())
}

fn validate_timeline_cache_set(package: &OpcPackage, caches: &[WorkbookTimelineCache]) -> Result<()> {
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    let any_uid = caches.iter().any(|cache| cache.definition.uid.is_some());
    let mut filter_ids = HashSet::new();
    for cache in caches {
        write_timeline_cache_definition(&cache.definition)?;
        if !names.insert(cache.definition.name.to_ascii_lowercase()) { return Err(invalid("duplicate Timeline Cache name")); }
        match &cache.definition.uid {
            Some(uid) if !uids.insert(uid.to_ascii_lowercase()) => return Err(invalid("duplicate Timeline Cache uid")),
            None if any_uid => return Err(invalid("Timeline Cache uid must be present on every cache or none")),
            _ => {},
        }
        if let Some(filter) = &cache.definition.timeline_pivot_filter
            && !filter_ids.insert(filter.id)
        {
            return Err(invalid(format!("duplicate timeline pivot filter ID {}", filter.id)));
        }
    }
    let workbook = package.main_document_part()?;
    let details = workbook_parser::parse_workbook_details(std::str::from_utf8(workbook.blob()).map_err(|e| invalid(e.to_string()))?)
        .map_err(|e| invalid(e.to_string()))?;
    if !details.pivot_caches.is_empty() {
        let ids: HashSet<u32> = details.pivot_caches.into_iter().map(|cache| cache.cache_id).collect();
        for cache in caches {
            if !ids.contains(&cache.definition.state.pivot_cache_id) {
                return Err(invalid(format!("Timeline Cache '{}' references missing pivot cache ID {}", cache.definition.name, cache.definition.state.pivot_cache_id)));
            }
        }
    }
    validate_timeline_views(&load_timelines(package, workbook.partname())?, caches)
}

fn validate_slicer_views(package: &OpcPackage, replacement: Option<(&str, &Slicers)>, caches: &[WorkbookSlicerCache]) -> Result<()> {
    let cache_names: HashSet<String> = caches.iter().map(|cache| cache.definition.name.to_ascii_lowercase()).collect();
    let mut names = HashSet::new();
    for part in package.iter_parts().filter(|part| part.content_type() == SLICERS_CONTENT_TYPE) {
        let value = if replacement.is_some_and(|(name, _)| name == part.partname().as_str()) {
            replacement.expect("checked").1.clone()
        } else {
            crate::xlsx::slicers::parse_slicers(part.blob())?
        };
        for slicer in value.slicers {
            if !names.insert(slicer.name.to_ascii_lowercase()) { return Err(invalid("duplicate workbook slicer name")); }
            if !cache_names.contains(&slicer.cache.to_ascii_lowercase()) { return Err(invalid(format!("slicer '{}' references missing cache '{}'", slicer.name, slicer.cache))); }
        }
    }
    Ok(())
}

fn validate_timeline_views(sheets: &[WorksheetTimelines], caches: &[WorkbookTimelineCache]) -> Result<()> {
    let cache_names: HashSet<String> = caches.iter().map(|cache| cache.definition.name.to_ascii_lowercase()).collect();
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    for sheet in sheets {
        write_timelines(&sheet.timelines)?;
        for timeline in &sheet.timelines.timelines {
            if !names.insert(timeline.name.to_ascii_lowercase()) { return Err(invalid("duplicate workbook timeline name")); }
            if let Some(uid) = &timeline.uid
                && !uids.insert(uid.to_ascii_lowercase())
            { return Err(invalid("duplicate workbook timeline uid")); }
            if !cache_names.contains(&timeline.cache.to_ascii_lowercase()) { return Err(invalid(format!("timeline '{}' references missing cache '{}'", timeline.name, timeline.cache))); }
        }
    }
    Ok(())
}

fn validate_slicer_cache_reference(slicer: &Slicer, caches: &[WorkbookSlicerCache]) -> Result<()> {
    if caches.iter().any(|cache| cache.definition.name.eq_ignore_ascii_case(&slicer.cache)) { Ok(()) }
    else { Err(invalid(format!("slicer '{}' references missing cache '{}'", slicer.name, slicer.cache))) }
}

fn validate_timeline_cache_reference(timeline: &Timeline, caches: &[WorkbookTimelineCache]) -> Result<()> {
    if caches.iter().any(|cache| cache.definition.name.eq_ignore_ascii_case(&timeline.cache)) { Ok(()) }
    else { Err(invalid(format!("timeline '{}' references missing cache '{}'", timeline.name, timeline.cache))) }
}

fn all_slicers(package: &OpcPackage) -> Result<Vec<Slicer>> {
    let mut output = Vec::new();
    for part in package.iter_parts().filter(|part| part.content_type() == SLICERS_CONTENT_TYPE) {
        output.extend(crate::xlsx::slicers::parse_slicers(part.blob())?.slicers);
    }
    Ok(output)
}

fn all_slicer_names(package: &OpcPackage) -> Result<HashSet<String>> {
    Ok(all_slicers(package)?.into_iter().map(|slicer| slicer.name.to_ascii_lowercase()).collect())
}

fn reorder_by_key<T, F>(values: Vec<T>, order: &[String], key: F, description: &str) -> Result<Vec<T>>
where F: Fn(&T) -> &str {
    if values.len() != order.len() { return Err(invalid(format!("{description} reorder must contain every item"))); }
    let mut remaining = HashMap::new();
    for value in values {
        let folded = key(&value).to_ascii_lowercase();
        if remaining.insert(folded, value).is_some() { return Err(invalid(format!("duplicate {description} name"))); }
    }
    let mut output = Vec::with_capacity(order.len());
    for name in order {
        output.push(remaining.remove(&name.to_ascii_lowercase()).ok_or_else(|| invalid(format!("unknown or duplicate {description} '{name}'")))?);
    }
    if !remaining.is_empty() { return Err(invalid(format!("{description} reorder must contain every item"))); }
    Ok(output)
}

fn rewrite_integration_refs(xml: &[u8], uri: &str, family_ns: &str, list: &str, item: &str, ids: &[String]) -> Result<Vec<u8>> {
    let (core, rel) = root_namespaces(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut target: Option<(usize, usize)> = None;
    let mut open: Option<(usize, usize)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))?;
        let event = reader.read_event().map_err(|e| invalid(e.to_string()))?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 2 && is_core_namespace(&namespace, core) && element.local_name().as_ref() == b"ext" && attribute_value(&element, "uri", reader.decoder())?.as_deref() == Some(uri) {
                    if open.is_some() || target.is_some() { return Err(invalid("duplicate integration extension")); }
                    open = Some((start, depth));
                }
                depth += 1;
                if depth > MAX_DEPTH { return Err(invalid("integration XML depth exceeds limit")); }
            },
            Event::Empty(element) => {
                if depth == 2 && is_core_namespace(&namespace, core) && element.local_name().as_ref() == b"ext" && attribute_value(&element, "uri", reader.decoder())?.as_deref() == Some(uri) {
                    if target.is_some() || open.is_some() { return Err(invalid("duplicate integration extension")); }
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))?;
                    target = Some((start, end));
                }
            },
            Event::End(_) => {
                if depth == 0 { return Err(invalid("unexpected XML close")); }
                depth -= 1;
                if open.is_some_and(|(_, open_depth)| open_depth == depth) {
                    let (begin, _) = open.take().expect("checked");
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))?;
                    target = Some((begin, end));
                }
            },
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")),
            Event::Eof => break,
            _ => {},
        }
    }
    let (start, end) = target.ok_or_else(|| invalid(format!("integration extension '{uri}' is missing")))?;
    let replacement = if ids.is_empty() { Vec::new() } else { integration_fragment(core, rel, uri, family_ns, list, item, ids) };
    let size = xml.len().checked_sub(end - start).and_then(|value| value.checked_add(replacement.len())).ok_or_else(|| invalid("rewrite size overflow"))?;
    if size > MAX_REWRITE_BYTES { return Err(invalid("rewritten XML exceeds limit")); }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

fn root_namespaces(xml: &[u8]) -> Result<(&'static str, &'static str)> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|e| invalid(e.to_string()))?;
        match event {
            Event::Start(_) => return match namespace {
                ResolveResult::Bound(Namespace(value)) if value.as_ref() == SML.as_bytes() => Ok((SML, REL)),
                ResolveResult::Bound(Namespace(value)) if value.as_ref() == STRICT_SML.as_bytes() => Ok((STRICT_SML, STRICT_REL)),
                _ => Err(invalid("unsupported SpreadsheetML root namespace")),
            },
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")),
            Event::Eof => return Err(invalid("missing XML root")),
            _ => {},
        }
    }
}

fn integration_fragment(core: &str, rel: &str, uri: &str, family_ns: &str, list: &str, item: &str, ids: &[String]) -> Vec<u8> {
    let mut output = format!("<ext xmlns=\"{}\" uri=\"{}\"><f:{} xmlns:f=\"{}\" xmlns:r=\"{}\">", core, uri, list, family_ns, rel);
    for id in ids { output.push_str(&format!("<f:{item} r:id=\"{}\"/>", xml_escape(id))); }
    output.push_str(&format!("</f:{list}></ext>"));
    output.into_bytes()
}

fn attribute_value(element: &quick_xml::events::BytesStart<'_>, name: &str, decoder: quick_xml::encoding::Decoder) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|e| invalid(e.to_string()))?;
        if attribute.key.as_ref() == name.as_bytes() {
            return Ok(Some(attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder).map_err(|e| invalid(e.to_string()))?.into_owned()));
        }
    }
    Ok(None)
}

fn is_core_namespace(namespace: &ResolveResult<'_>, core: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == core.as_bytes())
}

fn next_part_name(package: &OpcPackage, template: &str) -> Result<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(&template.replace("%d", &suffix.to_string())).map_err(OoxmlError::InvalidUri)?;
        if package.get_part(&candidate).is_err() { return Ok(candidate); }
    }
    Err(invalid("no free package part name"))
}

fn next_relationship_id(owner: &dyn Part, prefix: &str) -> Result<String> {
    for suffix in 1..=65_537u32 {
        let candidate = format!("{prefix}{suffix}");
        if owner.rels().get(&candidate).is_none() { return Ok(candidate); }
    }
    Err(invalid("no free relationship ID"))
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| part.rels().iter().any(|relationship| !relationship.is_external() && relationship.target_partname().is_ok_and(|name| name == *target)))
        || package.rels().iter().any(|relationship| !relationship.is_external() && relationship.target_partname().is_ok_and(|name| name == *target))
}

fn xml_escape(value: &str) -> String {
    value.replace('&', "&amp;").replace('<', "&lt;").replace('"', "&quot;")
}

fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
