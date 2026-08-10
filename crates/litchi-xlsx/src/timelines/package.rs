//! OPC part lifecycle, relationship graph validation, and atomic timeline CRUD.

use super::codec::{
    Node, attr, empty, escape, no_attributes, optional, parse_document,
    parse_timeline_cache_definition, parse_timelines, require, required, whitespace,
    write_timeline_cache_definition, write_timelines,
};
use super::model::{
    Cache, Part, validate_cache_set, validate_global_views, validate_relationship_id,
    validate_views_local,
};
use super::{
    MAX_CACHES, MAX_DEPTH, MAX_REWRITE_BYTES, REL, SML, STRICT_REL, STRICT_SML,
    TIMELINE_CACHE_CONTENT_TYPE, TIMELINE_CACHE_EXTENSION_URI, TIMELINE_CACHE_RELATIONSHIP_TYPE,
    TIMELINES_CONTENT_TYPE, TIMELINES_EXTENSION_URI, TIMELINES_RELATIONSHIP_TYPE, X15, invalid,
    limit, xml_error,
};
use crate::error::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub fn load_timeline_caches(package: &OpcPackage, workbook_name: &PackURI) -> Result<Vec<Cache>> {
    reject_root_relationships(package, TIMELINE_CACHE_RELATIONSHIP_TYPE, "View Cache")?;
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = parse_document(workbook.blob())?;
    let (core, rel) = source_namespaces(&workbook_root, "workbook")?;
    let refs = integration_refs(
        &workbook_root,
        core,
        rel,
        TIMELINE_CACHE_EXTENSION_URI,
        "timelineCacheRefs",
        "timelineCacheRef",
    )?
    .unwrap_or_default();
    if refs.len() > MAX_CACHES {
        return Err(limit("cache reference count"));
    }
    for part in package.iter_parts() {
        if part.partname().as_str() != workbook_name.as_str()
            && part
                .rels()
                .iter()
                .any(|relationship| relationship.reltype() == TIMELINE_CACHE_RELATIONSHIP_TYPE)
        {
            return Err(invalid(format!(
                "non-workbook part '{}' sources a View Cache relationship",
                part.partname()
            )));
        }
    }
    let mut ids = HashSet::new();
    let mut targets = HashSet::new();
    let mut output = Vec::with_capacity(refs.len());
    for id in refs {
        validate_relationship_id(&id)?;
        if !ids.insert(id.clone()) {
            return Err(invalid(format!("duplicate View Cache reference '{id}'")));
        }
        let relationship = workbook
            .rels()
            .get(&id)
            .ok_or_else(|| invalid(format!("missing View Cache relationship '{id}'")))?;
        if relationship.reltype() != TIMELINE_CACHE_RELATIONSHIP_TYPE || relationship.is_external()
        {
            return Err(invalid(format!(
                "View Cache reference '{id}' must target an internal View Cache relationship"
            )));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple View Cache references target '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != TIMELINE_CACHE_CONTENT_TYPE {
            return Err(invalid(format!(
                "View Cache part '{target}' has content type '{}'",
                part.content_type()
            )));
        }
        if !part.rels().is_empty() {
            return Err(invalid(format!(
                "View Cache part '{target}' has forbidden outbound relationships"
            )));
        }
        output.push(Cache {
            relationship_id: id,
            part_name: target.to_string(),
            definition: parse_timeline_cache_definition(part.blob())?,
        });
    }
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TIMELINE_CACHE_RELATIONSHIP_TYPE)
    {
        if !ids.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced View Cache relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == TIMELINE_CACHE_CONTENT_TYPE)
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan View Cache part '{}'",
                part.partname()
            )));
        }
    }
    validate_cache_set(&output)?;
    Ok(output)
}

/// Store a complete workbook View Cache set and its single integration extension.
pub fn store_timeline_caches(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    caches: &[Cache],
) -> Result<()> {
    if caches.is_empty() {
        return Err(invalid("at least one View Cache is required for storage"));
    }
    validate_cache_set(caches)?;
    if !load_timeline_caches(package, workbook_name)?.is_empty() {
        return Err(invalid("workbook already contains View Caches"));
    }
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    let (core, rel) = source_namespaces(&root, "workbook")?;
    let mut targets = HashSet::new();
    let mut ids = HashSet::new();
    let mut plans = Vec::with_capacity(caches.len());
    for cache in caches {
        validate_relationship_id(&cache.relationship_id)?;
        if !ids.insert(cache.relationship_id.clone()) {
            return Err(invalid("duplicate View Cache relationship ID"));
        }
        if workbook.rels().get(&cache.relationship_id).is_some() {
            return Err(invalid(format!(
                "workbook relationship ID '{}' already exists",
                cache.relationship_id
            )));
        }
        let uri = PackURI::new(&cache.part_name)
            .map_err(|error| Error::Invalid(format!("invalid View Cache part URI: {error}")))?;
        if !uri.as_str().starts_with("/xl/timelineCaches/") || !uri.as_str().ends_with(".xml") {
            return Err(invalid(format!(
                "View Cache part '{uri}' must be under /xl/timelineCaches and end in .xml"
            )));
        }
        if !targets.insert(uri.to_string())
            || package
                .iter_parts()
                .any(|part| part.partname().as_str() == uri.as_str())
        {
            return Err(invalid(format!("View Cache target '{uri}' already exists")));
        }
        plans.push((
            cache.relationship_id.clone(),
            uri,
            write_timeline_cache_definition(&cache.definition)?,
        ));
    }
    let refs: Vec<String> = plans.iter().map(|plan| plan.0.clone()).collect();
    let fragment = integration_extension(
        core,
        rel,
        TIMELINE_CACHE_EXTENSION_URI,
        "timelineCacheRefs",
        "timelineCacheRef",
        &refs,
    );
    let updated = insert_extension(
        workbook.blob(),
        &root,
        core,
        TIMELINE_CACHE_EXTENSION_URI,
        &fragment,
    )?;
    for (_, uri, xml) in &plans {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            TIMELINE_CACHE_CONTENT_TYPE.into(),
            xml.clone(),
        )));
    }
    for (id, uri, _) in &plans {
        package
            .get_part_mut(workbook_name)?
            .rels_mut()
            .add_relationship(
                TIMELINE_CACHE_RELATIONSHIP_TYPE.into(),
                uri.relative_ref(workbook_name.base_uri()),
                id.clone(),
                false,
            );
    }
    package.get_part_mut(workbook_name)?.set_blob(updated);
    Ok(())
}

/// Load every worksheet Views part and cross-validate views against View Caches.
pub fn load_timelines(package: &OpcPackage, workbook_name: &PackURI) -> Result<Vec<Part>> {
    reject_root_relationships(package, TIMELINES_RELATIONSHIP_TYPE, "Views")?;
    let caches = load_timeline_caches(package, workbook_name)?;
    let cache_names: HashSet<String> = caches
        .iter()
        .map(|cache| cache.definition.name.to_lowercase())
        .collect();
    let mut targets = HashSet::new();
    let mut output = Vec::new();
    for part in package.iter_parts() {
        let timeline_relationships: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == TIMELINES_RELATIONSHIP_TYPE)
            .collect();
        if part.content_type() != ct::SML_WORKSHEET {
            if !timeline_relationships.is_empty() {
                return Err(invalid(format!(
                    "non-worksheet part '{}' sources a Views relationship",
                    part.partname()
                )));
            }
            continue;
        }
        let root = parse_document(part.blob())?;
        let (core, rel) = source_namespaces(&root, "worksheet")?;
        let refs = integration_refs(
            &root,
            core,
            rel,
            TIMELINES_EXTENSION_URI,
            "timelineRefs",
            "timelineRef",
        )?
        .unwrap_or_default();
        if refs.is_empty() {
            if !timeline_relationships.is_empty() {
                return Err(invalid(format!(
                    "worksheet '{}' has a Views relationship without timelineRefs",
                    part.partname()
                )));
            }
            continue;
        }
        if refs.len() != 1 {
            return Err(invalid(
                "worksheet timelineRefs must contain exactly one timelineRef",
            ));
        }
        let id = refs[0].clone();
        validate_relationship_id(&id)?;
        let relationship = part
            .rels()
            .get(&id)
            .ok_or_else(|| invalid(format!("missing Views relationship '{id}'")))?;
        if relationship.reltype() != TIMELINES_RELATIONSHIP_TYPE || relationship.is_external() {
            return Err(invalid(format!(
                "Views reference '{id}' must target an internal Views relationship"
            )));
        }
        if timeline_relationships.len() != 1 {
            return Err(invalid(format!(
                "worksheet '{}' has unreferenced or duplicate Views relationships",
                part.partname()
            )));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple worksheets target Views part '{target}'"
            )));
        }
        let timeline_part = package.get_part(&target)?;
        if timeline_part.content_type() != TIMELINES_CONTENT_TYPE {
            return Err(invalid(format!(
                "Views part '{target}' has content type '{}'",
                timeline_part.content_type()
            )));
        }
        if !timeline_part.rels().is_empty() {
            return Err(invalid(format!(
                "Views part '{target}' has forbidden outbound relationships"
            )));
        }
        let timelines = parse_timelines(timeline_part.blob())?;
        for view in &timelines.timelines {
            if !cache_names.contains(&view.cache.to_lowercase()) {
                return Err(invalid(format!(
                    "timeline '{}' references unknown cache '{}'",
                    view.name, view.cache
                )));
            }
        }
        output.push(Part {
            worksheet_part_name: part.partname().to_string(),
            relationship_id: id,
            part_name: target.to_string(),
            timelines,
        });
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == TIMELINES_CONTENT_TYPE)
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!("orphan Views part '{}'", part.partname())));
        }
    }
    validate_global_views(&output)?;
    Ok(output)
}

/// Store one worksheet's Views part, then validate it against all workbook caches and views.
pub fn store_worksheet_timelines(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Part,
) -> Result<()> {
    validate_views_local(&value.timelines)?;
    let caches = load_timeline_caches(package, workbook_name)?;
    let cache_names: HashSet<String> = caches
        .iter()
        .map(|cache| cache.definition.name.to_lowercase())
        .collect();
    for view in &value.timelines.timelines {
        if !cache_names.contains(&view.cache.to_lowercase()) {
            return Err(invalid(format!(
                "timeline '{}' references unknown cache '{}'",
                view.name, view.cache
            )));
        }
    }
    let existing = load_timelines(package, workbook_name)?;
    if existing
        .iter()
        .any(|sheet| sheet.worksheet_part_name == value.worksheet_part_name)
    {
        return Err(invalid("worksheet already contains a Views part"));
    }
    let mut combined = existing;
    combined.push(value.clone());
    validate_global_views(&combined)?;
    validate_relationship_id(&value.relationship_id)?;
    let worksheet_uri = PackURI::new(&value.worksheet_part_name)
        .map_err(|error| Error::Invalid(format!("invalid Views worksheet part URI: {error}")))?;
    let target_uri = PackURI::new(&value.part_name)
        .map_err(|error| Error::Invalid(format!("invalid Views part URI: {error}")))?;
    if !target_uri.as_str().starts_with("/xl/timelines/") || !target_uri.as_str().ends_with(".xml")
    {
        return Err(invalid(
            "Views part must be under /xl/timelines and end in .xml",
        ));
    }
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == target_uri.as_str())
    {
        return Err(invalid(format!("Views part '{target_uri}' already exists")));
    }
    let worksheet = package.get_part(&worksheet_uri)?;
    if worksheet.content_type() != ct::SML_WORKSHEET {
        return Err(invalid(format!(
            "part '{worksheet_uri}' is not a worksheet"
        )));
    }
    if worksheet.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(format!(
            "worksheet relationship ID '{}' already exists",
            value.relationship_id
        )));
    }
    let root = parse_document(worksheet.blob())?;
    let (core, rel) = source_namespaces(&root, "worksheet")?;
    let fragment = integration_extension(
        core,
        rel,
        TIMELINES_EXTENSION_URI,
        "timelineRefs",
        "timelineRef",
        std::slice::from_ref(&value.relationship_id),
    );
    let updated = insert_extension(
        worksheet.blob(),
        &root,
        core,
        TIMELINES_EXTENSION_URI,
        &fragment,
    )?;
    let xml = write_timelines(&value.timelines)?;
    package.add_part(Box::new(BlobPart::new(
        target_uri.clone(),
        TIMELINES_CONTENT_TYPE.into(),
        xml,
    )));
    package
        .get_part_mut(&worksheet_uri)?
        .rels_mut()
        .add_relationship(
            TIMELINES_RELATIONSHIP_TYPE.into(),
            target_uri.relative_ref(worksheet_uri.base_uri()),
            value.relationship_id.clone(),
            false,
        );
    package.get_part_mut(&worksheet_uri)?.set_blob(updated);
    Ok(())
}

fn integration_refs(
    root: &Node,
    core: &str,
    rel: &str,
    uri: &str,
    list_name: &str,
    ref_name: &str,
) -> Result<Option<Vec<String>>> {
    let mut found = None;
    for ext_lst in root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
    {
        whitespace(ext_lst)?;
        for ext in ext_lst.children.iter().filter(|child| {
            child.namespace == core
                && child.name == "ext"
                && optional(child, "", "uri") == Some(uri)
        }) {
            if found.is_some() {
                return Err(invalid(format!(
                    "duplicate integration extension URI '{uri}'"
                )));
            }
            no_attributes(ext, &[("", "uri")])?;
            whitespace(ext)?;
            if ext.children.len() != 1 {
                return Err(invalid(format!(
                    "integration extension '{uri}' must contain exactly one child"
                )));
            }
            let list = &ext.children[0];
            require(list, X15, list_name)?;
            no_attributes(list, &[])?;
            whitespace(list)?;
            if list.children.is_empty() {
                return Err(invalid(format!("{list_name} must not be empty")));
            }
            let mut ids = Vec::with_capacity(list.children.len());
            for item in &list.children {
                require(item, X15, ref_name)?;
                no_attributes(item, &[(rel, "id")])?;
                empty(item)?;
                ids.push(required(item, rel, "id")?.to_owned());
            }
            found = Some(ids);
        }
    }
    Ok(found)
}

fn integration_extension(
    core: &str,
    rel: &str,
    uri: &str,
    list_name: &str,
    ref_name: &str,
    ids: &[String],
) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<ext xmlns=\"");
    escape(&mut output, core);
    output.extend_from_slice(b"\" uri=\"");
    escape(&mut output, uri);
    output.extend_from_slice(b"\"><x15:");
    output.extend_from_slice(list_name.as_bytes());
    output.extend_from_slice(b" xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, rel);
    output.extend_from_slice(b"\">");
    for id in ids {
        output.extend_from_slice(b"<x15:");
        output.extend_from_slice(ref_name.as_bytes());
        attr(&mut output, "r:id", id);
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</x15:");
    output.extend_from_slice(list_name.as_bytes());
    output.extend_from_slice(b"></ext>");
    output
}

fn insert_extension(
    xml: &[u8],
    root: &Node,
    core: &str,
    uri: &str,
    fragment: &[u8],
) -> Result<Vec<u8>> {
    if root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
        .flat_map(|list| &list.children)
        .any(|child| {
            child.namespace == core
                && child.name == "ext"
                && optional(child, "", "uri") == Some(uri)
        })
    {
        return Err(invalid(format!(
            "integration extension '{uri}' already exists"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut open_ext = None;
    let mut empty_ext = None;
    let mut root_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let mut empty_candidate = None;
        match event {
            Event::Start(element) => {
                let is_core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == core.as_bytes());
                if depth == 1 && is_core && element.local_name().as_ref() == b"extLst" {
                    open_ext = Some(2usize);
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
            },
            Event::Empty(element) => {
                let is_core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == core.as_bytes());
                if depth == 1 && is_core && element.local_name().as_ref() == b"extLst" {
                    empty_candidate = Some((start, element.name().as_ref().to_vec()));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected closing element"));
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                if open_ext == Some(depth) && element.local_name().as_ref() == b"extLst" {
                    let size = xml
                        .len()
                        .checked_add(fragment.len())
                        .ok_or_else(|| limit("rewrite bytes"))?;
                    if size > MAX_REWRITE_BYTES {
                        return Err(limit("rewrite bytes"));
                    }
                    let mut output = Vec::with_capacity(size);
                    output.extend_from_slice(&xml[..start]);
                    output.extend_from_slice(fragment);
                    output.extend_from_slice(&xml[start..]);
                    return Ok(output);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
        if let Some((start, qname)) = empty_candidate {
            let end = usize::try_from(reader.buffer_position())
                .map_err(|_source| invalid("XML offset overflow"))?;
            empty_ext = Some((start, end, qname));
        }
    }
    if let Some((start, end, qname)) = empty_ext {
        let size = xml
            .len()
            .checked_add(fragment.len() + qname.len() + 2)
            .ok_or_else(|| limit("rewrite bytes"))?;
        if size > MAX_REWRITE_BYTES {
            return Err(limit("rewrite bytes"));
        }
        let raw = &xml[start..end];
        let close = raw
            .windows(2)
            .rposition(|window| window == b"/>")
            .ok_or_else(|| invalid("invalid empty extLst"))?;
        let mut output = Vec::with_capacity(size);
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(&raw[..close]);
        output.push(b'>');
        output.extend_from_slice(fragment);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&qname);
        output.push(b'>');
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    let position = root_close.ok_or_else(|| invalid("missing source root closing element"))?;
    let mut wrapper = Vec::new();
    wrapper.extend_from_slice(b"<extLst xmlns=\"");
    escape(&mut wrapper, core);
    wrapper.extend_from_slice(b"\">");
    wrapper.extend_from_slice(fragment);
    wrapper.extend_from_slice(b"</extLst>");
    let size = xml
        .len()
        .checked_add(wrapper.len())
        .ok_or_else(|| limit("rewrite bytes"))?;
    if size > MAX_REWRITE_BYTES {
        return Err(limit("rewrite bytes"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(&wrapper);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn source_namespaces<'a>(root: &'a Node, name: &str) -> Result<(&'a str, &'static str)> {
    if root.name != name {
        return Err(invalid(format!("expected {name} source root")));
    }
    match root.namespace.as_str() {
        SML => Ok((SML, REL)),
        STRICT_SML => Ok((STRICT_SML, STRICT_REL)),
        _ => Err(invalid(format!("unsupported {name} namespace"))),
    }
}
fn reject_root_relationships(package: &OpcPackage, kind: &str, name: &str) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == kind)
    {
        Err(invalid(format!(
            "package root cannot source {name} relationships"
        )))
    } else {
        Ok(())
    }
}
