//! OPC part lifecycle, relationship graph validation, and atomic timeline CRUD.

use super::codec::{
    Node, empty, no_attributes, optional, parse_document, parse_timeline_cache_definition,
    parse_timelines, require, required, whitespace, write_timeline_cache_definition,
    write_timelines,
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
use litchi_core::xml::ReaderOrigin;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, SourceLimits};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

struct IntegrationXml {
    bytes: Vec<u8>,
    limit: usize,
}

impl IntegrationXml {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn append(&mut self, bytes: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| limit("integration XML bytes"))?;
        if length > self.limit {
            return Err(limit("integration XML bytes"));
        }
        self.bytes
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "Timeline integration XML",
                source,
            })?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    fn escape(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.append(b"&amp;")?,
                '<' => self.append(b"&lt;")?,
                '"' => self.append(b"&quot;")?,
                '\t' => self.append(b"&#x9;")?,
                '\n' => self.append(b"&#xA;")?,
                '\r' => self.append(b"&#xD;")?,
                _ => {
                    let mut bytes = [0; 4];
                    self.append(character.encode_utf8(&mut bytes).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

pub fn load_timeline_caches(package: &OpcPackage, workbook_name: &PackURI) -> Result<Vec<Cache>> {
    let limits = package.read_limits();
    load_timeline_caches_limited(package, workbook_name, &limits)
}

pub(crate) fn load_timeline_caches_limited(
    package: &OpcPackage,
    workbook_name: &PackURI,
    limits: &SourceLimits,
) -> Result<Vec<Cache>> {
    reject_root_relationships(package, super::is_cache_relationship, "View Cache")?;
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = super::codec::parse_document_limited(workbook.blob(), limits)?;
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
                .any(|relationship| super::is_cache_relationship(relationship.reltype()))
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
        if !super::is_cache_relationship(relationship.reltype()) || relationship.is_external() {
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
        if !super::is_cache_content(part.content_type()) {
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
            definition: super::codec::parse_timeline_cache_definition_limited(part.blob(), limits)?,
        });
    }
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| super::is_cache_relationship(relationship.reltype()))
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
        .filter(|part| super::is_cache_content(part.content_type()))
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan View Cache part '{}'",
                part.partname()
            )));
        }
    }
    let maximum = crate::raw::output_ceiling(limits);
    for cache in &output {
        super::codec::preflight_cache_definition_limited(&cache.definition, maximum)?;
    }
    validate_cache_set(&output)?;
    Ok(output)
}

/// Store a complete workbook View Cache set and its single integration extension.
///
/// Storing changes the package, so a signed package is refused up front with
/// [`crate::Error::Signed`] (it used to be changed in memory and refused only
/// when saved): remove the signature explicitly first
/// ([`OpcPackage::unsign`]).
pub(crate) fn store_timeline_caches(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    caches: &[Cache],
) -> Result<()> {
    let limits = package.read_limits();
    store_timeline_caches_limited(package, workbook_name, caches, &limits)
}

fn preflight_timeline_cache_insertions(
    package: &OpcPackage,
    workbook_name: &PackURI,
    caches: &[Cache],
    limits: &SourceLimits,
) -> Result<()> {
    if caches.is_empty() {
        return Err(invalid("at least one View Cache is required for storage"));
    }
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::RelationshipsPerPart,
        caches.len() as u64,
        limits.max_relationships_per_part() as u64,
    )?;
    crate::raw::check_new_part_count(package, caches.len(), limits)?;
    validate_cache_set(caches.iter())?;

    let mut ids = HashSet::new();
    ids.try_reserve(caches.len())
        .map_err(|source| Error::Allocation {
            resource: "timeline cache relationship IDs",
            source,
        })?;
    let mut targets = HashSet::new();
    targets
        .try_reserve(caches.len())
        .map_err(|source| Error::Allocation {
            resource: "timeline cache part names",
            source,
        })?;
    for cache in caches {
        let uri = crate::timeline::raw::validate_cache_identity_fields(
            &cache.relationship_id,
            &cache.part_name,
            workbook_name.base_uri(),
            limits,
        )?;
        if !ids.insert(cache.relationship_id.as_str()) {
            return Err(invalid("duplicate View Cache relationship ID"));
        }
        if !targets.insert(uri.to_string()) {
            return Err(invalid("duplicate View Cache part name"));
        }
        crate::raw::check_new_relationship_identity(
            package,
            workbook_name,
            &cache.relationship_id,
            &uri,
        )?;
    }
    if !load_timeline_caches_limited(package, workbook_name, limits)?.is_empty() {
        return Err(invalid("workbook already contains View Caches"));
    }
    Ok(())
}

pub(crate) fn store_timeline_caches_limited(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    caches: &[Cache],
    limits: &SourceLimits,
) -> Result<()> {
    preflight_timeline_cache_insertions(package, workbook_name, caches, limits)?;
    let additions: Vec<_> = caches
        .iter()
        .map(|cache| {
            let uri = crate::timeline::raw::validate_cache_identity_fields(
                &cache.relationship_id,
                &cache.part_name,
                workbook_name.base_uri(),
                limits,
            )?;
            Ok((
                cache.relationship_id.clone(),
                TIMELINE_CACHE_RELATIONSHIP_TYPE.to_owned(),
                uri.relative_ref(workbook_name.base_uri()),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    let budget = crate::raw::prepare_package_budget(package, workbook_name, &additions, limits)?;
    let prepared = caches
        .iter()
        .cloned()
        .map(|cache| {
            crate::timeline::raw::prepare_cache(
                crate::timeline::raw::Cache {
                    relationship_id: cache.relationship_id,
                    part_name: cache.part_name,
                    definition: cache.definition,
                },
                workbook_name.base_uri(),
                limits,
            )
        })
        .collect::<Result<Vec<_>>>()?;
    store_timeline_caches_prepared_limited_with_budget(
        package,
        workbook_name,
        &prepared,
        limits,
        &budget,
    )
}

pub(crate) fn store_timeline_caches_prepared_limited(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    prepared: &[crate::timeline::raw::PreparedCache],
    limits: &SourceLimits,
) -> Result<()> {
    if prepared.is_empty() {
        return Err(invalid("at least one View Cache is required for storage"));
    }
    let additions: Vec<_> = prepared
        .iter()
        .map(|cache| {
            (
                cache.value.relationship_id.clone(),
                TIMELINE_CACHE_RELATIONSHIP_TYPE.to_owned(),
                cache.uri.relative_ref(workbook_name.base_uri()),
            )
        })
        .collect();
    let budget = crate::raw::prepare_package_budget(package, workbook_name, &additions, limits)?;
    store_timeline_caches_prepared_limited_with_budget(
        package,
        workbook_name,
        prepared,
        limits,
        &budget,
    )
}

pub(crate) fn store_timeline_caches_prepared_limited_with_budget(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    prepared: &[crate::timeline::raw::PreparedCache],
    limits: &SourceLimits,
    budget: &crate::raw::PackageBudget,
) -> Result<()> {
    if prepared.is_empty() {
        return Err(invalid("at least one View Cache is required for storage"));
    }
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::RelationshipsPerPart,
        prepared.len() as u64,
        limits.max_relationships_per_part() as u64,
    )?;
    crate::raw::check_new_part_count(package, prepared.len(), limits)?;
    validate_cache_set(prepared.iter().map(|cache| &cache.value))?;
    if !load_timeline_caches_limited(package, workbook_name, limits)?.is_empty() {
        return Err(invalid("workbook already contains View Caches"));
    }
    let workbook = package.get_part(workbook_name)?;
    let root = super::codec::parse_document_limited(workbook.blob(), limits)?;
    let (core, rel) = source_namespaces(&root, "workbook")?;
    let mut targets = HashSet::new();
    let mut ids = HashSet::new();
    for cache in prepared {
        let value = &cache.value;
        if !ids.insert(value.relationship_id.clone()) {
            return Err(invalid("duplicate View Cache relationship ID"));
        }
        if workbook.rels().get(&value.relationship_id).is_some() {
            return Err(invalid(format!(
                "workbook relationship ID '{}' already exists",
                value.relationship_id
            )));
        }
        let uri = cache.uri.clone();
        if !targets.insert(uri.to_string()) || package.get_part(&uri).is_ok() {
            return Err(invalid(format!("View Cache target '{uri}' already exists")));
        }
    }
    let added_bytes = prepared.iter().try_fold(0usize, |total, cache| {
        total
            .checked_add(cache.xml.len())
            .ok_or_else(|| limit("aggregate part bytes"))
    })?;
    let refs: Vec<String> = prepared
        .iter()
        .map(|cache| cache.value.relationship_id.clone())
        .collect();
    let fragment = integration_extension(
        core,
        rel,
        TIMELINE_CACHE_EXTENSION_URI,
        "timelineCacheRefs",
        "timelineCacheRef",
        &refs,
        limits,
    )?;
    let updated = insert_extension(
        workbook.blob(),
        &root,
        core,
        TIMELINE_CACHE_EXTENSION_URI,
        &fragment,
        limits,
        |replacement_len| {
            crate::raw::check_part_budget_with_ledger(
                package,
                budget,
                prepared.len(),
                added_bytes,
                &[(workbook_name, replacement_len)],
                limits,
            )
        },
    )?;
    crate::signature_policy::ensure_unsigned(package)?;
    for cache in prepared {
        package.add_part(Box::new(BlobPart::new(
            cache.uri.clone(),
            TIMELINE_CACHE_CONTENT_TYPE.into(),
            cache.xml.clone(),
        )));
    }
    for cache in prepared {
        package
            .get_part_mut(workbook_name)?
            .rels_mut()
            .add_relationship(
                TIMELINE_CACHE_RELATIONSHIP_TYPE.into(),
                cache.uri.relative_ref(workbook_name.base_uri()),
                cache.value.relationship_id.clone(),
                false,
            );
    }
    package.get_part_mut(workbook_name)?.set_blob(updated);
    package.source_content_types(*limits)?;
    Ok(())
}

/// Load every worksheet Views part and cross-validate views against View Caches.
pub fn load_timelines(package: &OpcPackage, workbook_name: &PackURI) -> Result<Vec<Part>> {
    let limits = package.read_limits();
    load_timelines_limited(package, workbook_name, &limits)
}

pub(crate) fn load_timelines_limited(
    package: &OpcPackage,
    workbook_name: &PackURI,
    limits: &SourceLimits,
) -> Result<Vec<Part>> {
    reject_root_relationships(package, super::is_views_relationship, "Views")?;
    let caches = load_timeline_caches_limited(package, workbook_name, limits)?;
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
            .filter(|relationship| super::is_views_relationship(relationship.reltype()))
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
        let root = super::codec::parse_document_limited(
            package.get_part(part.partname())?.blob(),
            limits,
        )?;
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
        if !super::is_views_relationship(relationship.reltype()) || relationship.is_external() {
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
        if !super::is_views_content(timeline_part.content_type()) {
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
        let timelines = super::codec::parse_timelines_limited(timeline_part.blob(), limits)?;
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
        .filter(|part| super::is_views_content(part.content_type()))
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!("orphan Views part '{}'", part.partname())));
        }
    }
    validate_global_views(&output)?;
    Ok(output)
}

/// Store one worksheet's Views part, then validate it against all workbook caches and views.
///
/// Storing changes the package, so a signed package is refused up front with
/// [`crate::Error::Signed`] (it used to be changed in memory and refused only
/// when saved): remove the signature explicitly first
/// ([`OpcPackage::unsign`]).
pub(crate) fn store_worksheet_timelines(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Part,
) -> Result<()> {
    let limits = package.read_limits();
    store_worksheet_timelines_limited(package, workbook_name, value, &limits)
}

pub(crate) fn store_worksheet_timelines_limited(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Part,
    limits: &SourceLimits,
) -> Result<()> {
    preflight_timeline_part_insertion(package, value, limits)?;
    let (target_uri, worksheet_uri) = crate::timeline::raw::validate_part_identity_fields(
        &value.worksheet_part_name,
        &value.relationship_id,
        &value.part_name,
        limits,
    )?;
    let additions = [(
        value.relationship_id.clone(),
        TIMELINES_RELATIONSHIP_TYPE.to_owned(),
        target_uri.relative_ref(worksheet_uri.base_uri()),
    )];
    let budget = crate::raw::prepare_package_budget(package, &worksheet_uri, &additions, limits)?;
    let prepared = crate::timeline::raw::prepare_part(
        crate::timeline::raw::Part {
            worksheet_part_name: value.worksheet_part_name.clone(),
            relationship_id: value.relationship_id.clone(),
            part_name: value.part_name.clone(),
            timelines: value.timelines.clone(),
        },
        limits,
    )?;
    store_worksheet_timelines_prepared_limited_with_budget(
        package,
        workbook_name,
        &prepared,
        limits,
        &budget,
    )
}

fn preflight_timeline_part_insertion(
    package: &OpcPackage,
    value: &Part,
    limits: &SourceLimits,
) -> Result<()> {
    validate_views_local(&value.timelines)?;
    crate::raw::check_new_part_count(package, 1, limits)?;
    let (target_uri, worksheet_uri) = crate::timeline::raw::validate_part_identity_fields(
        &value.worksheet_part_name,
        &value.relationship_id,
        &value.part_name,
        limits,
    )?;
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
    crate::raw::check_new_relationship_identity(
        package,
        &worksheet_uri,
        &value.relationship_id,
        &target_uri,
    )?;
    Ok(())
}

pub(crate) fn store_worksheet_timelines_prepared_limited(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    prepared: &crate::timeline::raw::PreparedPart,
    limits: &SourceLimits,
) -> Result<()> {
    let additions = [(
        prepared.value.relationship_id.clone(),
        TIMELINES_RELATIONSHIP_TYPE.to_owned(),
        prepared.uri.relative_ref(prepared.worksheet.base_uri()),
    )];
    let budget =
        crate::raw::prepare_package_budget(package, &prepared.worksheet, &additions, limits)?;
    store_worksheet_timelines_prepared_limited_with_budget(
        package,
        workbook_name,
        prepared,
        limits,
        &budget,
    )
}

pub(crate) fn store_worksheet_timelines_prepared_limited_with_budget(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    prepared: &crate::timeline::raw::PreparedPart,
    limits: &SourceLimits,
    budget: &crate::raw::PackageBudget,
) -> Result<()> {
    let value = &prepared.value;
    validate_views_local(&value.timelines)?;
    let caches = load_timeline_caches_limited(package, workbook_name, limits)?;
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
    let existing = load_timelines_limited(package, workbook_name, limits)?;
    if existing
        .iter()
        .any(|sheet| sheet.worksheet_part_name == value.worksheet_part_name)
    {
        return Err(invalid("worksheet already contains a Views part"));
    }
    validate_global_views(existing.iter().chain(std::iter::once(value)))?;
    let worksheet_uri = prepared.worksheet.clone();
    let target_uri = prepared.uri.clone();
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
    let root = super::codec::parse_document_limited(worksheet.blob(), limits)?;
    let (core, rel) = source_namespaces(&root, "worksheet")?;
    let fragment = integration_extension(
        core,
        rel,
        TIMELINES_EXTENSION_URI,
        "timelineRefs",
        "timelineRef",
        std::slice::from_ref(&value.relationship_id),
        limits,
    )?;
    let updated = insert_extension(
        worksheet.blob(),
        &root,
        core,
        TIMELINES_EXTENSION_URI,
        &fragment,
        limits,
        |replacement_len| {
            crate::raw::check_part_budget_with_ledger(
                package,
                budget,
                1,
                prepared.xml.len(),
                &[(&worksheet_uri, replacement_len)],
                limits,
            )
        },
    )?;
    crate::signature_policy::ensure_unsigned(package)?;
    package.add_part(Box::new(BlobPart::new(
        target_uri.clone(),
        TIMELINES_CONTENT_TYPE.into(),
        prepared.xml.clone(),
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
    package.source_content_types(*limits)?;
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
    limits: &SourceLimits,
) -> Result<Vec<u8>> {
    let maximum = crate::raw::output_ceiling(limits).min(MAX_REWRITE_BYTES);
    let mut output = IntegrationXml::new(maximum);
    output.append(b"<ext xmlns=\"")?;
    output.escape(core)?;
    output.append(b"\" uri=\"")?;
    output.escape(uri)?;
    output.append(b"\"><x15:")?;
    output.append(list_name.as_bytes())?;
    output.append(b" xmlns:x15=\"")?;
    output.escape(X15)?;
    output.append(b"\" xmlns:r=\"")?;
    output.escape(rel)?;
    output.append(b"\">")?;
    for id in ids {
        output.append(b"<x15:")?;
        output.append(ref_name.as_bytes())?;
        output.append(b" r:id=\"")?;
        output.escape(id)?;
        output.append(b"\"/>")?;
    }
    output.append(b"</x15:")?;
    output.append(list_name.as_bytes())?;
    output.append(b"></ext>")?;
    Ok(output.finish())
}

fn insert_extension(
    xml: &[u8],
    root: &Node,
    core: &str,
    uri: &str,
    fragment: &[u8],
    limits: &SourceLimits,
    mut admit: impl FnMut(usize) -> Result<()>,
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
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut open_ext = None;
    let mut empty_ext = None;
    let mut root_close = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML offset overflow"))?;
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
                    if size > MAX_REWRITE_BYTES
                        || u64::try_from(size).unwrap_or(u64::MAX) > limits.max_part_bytes()
                    {
                        return Err(limit("rewrite bytes"));
                    }
                    admit(size)?;
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
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("XML offset overflow"))?;
            empty_ext = Some((start, end, qname));
        }
    }
    if let Some((start, end, qname)) = empty_ext {
        let size = xml
            .len()
            .checked_add(fragment.len() + qname.len() + 2)
            .ok_or_else(|| limit("rewrite bytes"))?;
        if size > MAX_REWRITE_BYTES
            || u64::try_from(size).unwrap_or(u64::MAX) > limits.max_part_bytes()
        {
            return Err(limit("rewrite bytes"));
        }
        admit(size)?;
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
    let maximum = crate::raw::output_ceiling(limits).min(MAX_REWRITE_BYTES);
    let mut wrapper = IntegrationXml::new(maximum);
    wrapper.append(b"<extLst xmlns=\"")?;
    wrapper.escape(core)?;
    wrapper.append(b"\">")?;
    wrapper.append(fragment)?;
    wrapper.append(b"</extLst>")?;
    let wrapper = wrapper.finish();
    let size = xml
        .len()
        .checked_add(wrapper.len())
        .ok_or_else(|| limit("rewrite bytes"))?;
    if size > MAX_REWRITE_BYTES || u64::try_from(size).unwrap_or(u64::MAX) > limits.max_part_bytes()
    {
        return Err(limit("rewrite bytes"));
    }
    admit(size)?;
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
fn reject_root_relationships(
    package: &OpcPackage,
    matches_kind: fn(&str) -> bool,
    name: &str,
) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| matches_kind(relationship.reltype()))
    {
        Err(invalid(format!(
            "package root cannot source {name} relationships"
        )))
    } else {
        Ok(())
    }
}
