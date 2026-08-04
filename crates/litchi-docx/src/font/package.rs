//! OPC relationship graph ownership for the WordprocessingML font table.

use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, XmlPart};
use quick_xml::{XmlVersion, events::Event, reader::Reader};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use super::codec::{parse, write, xml_error};
use super::model::{
    Conformance, FONT_CT, FT_CT, MAX_ALL_FONTS, MAX_FONT, MAX_NODES, MAX_XML, Resource, Table,
    bounded, invalid, is_font_relationship, is_font_table_relationship, name_key,
    validate_table_value,
};

impl Table {
    fn extract_from_part(part: &dyn Part, pkg: &OpcPackage) -> Result<Self> {
        if part.content_type() != FT_CT {
            return Err(Error::ContentType {
                expected: FT_CT.into(),
                actual: part.content_type().into(),
            });
        }
        let mut v = parse(part.blob())?;
        v.resolve(part, pkg)?;
        Ok(v)
    }
    fn resolve(&mut self, source: &dyn Part, pkg: &OpcPackage) -> Result<()> {
        validate_font_relationship_sources(pkg, source.partname())?;
        let mut used = HashSet::new();
        let mut cached = HashMap::<String, Resource>::new();
        let mut targets = HashSet::new();
        let mut total = 0usize;
        for font in &mut self.fonts {
            for embed in &mut font.embedded_fonts {
                used.insert(embed.relationship_id.clone());
                let rel = source.rels().get(&embed.relationship_id).ok_or_else(|| {
                    invalid(format!(
                        "missing embedded-font relationship '{}'",
                        embed.relationship_id
                    ))
                })?;
                if !is_font_relationship(rel.reltype()) {
                    return Err(invalid(format!(
                        "invalid embedded-font relationship type '{}'",
                        rel.reltype()
                    )));
                }
                if rel.is_external() {
                    return Err(invalid("embedded-font relationship cannot be external"));
                }
                let uri = rel.target_partname()?;
                let target_name = uri.to_string();
                targets.insert(target_name.clone());
                if let Some(v) = cached.get(&target_name) {
                    embed.resource = Some(v.clone());
                    continue;
                }
                let part = pkg.get_part(&uri)?;
                if part.content_type() != FONT_CT {
                    return Err(Error::ContentType {
                        expected: FONT_CT.into(),
                        actual: part.content_type().into(),
                    });
                }
                if part.blob().len() > MAX_FONT {
                    return Err(invalid(format!("embedded font '{uri}' is too large")));
                }
                if part.blob().len() < 32 {
                    return Err(invalid(format!("embedded font '{uri}' is too short")));
                }
                total = total
                    .checked_add(part.blob().len())
                    .ok_or_else(|| invalid("embedded-font size overflow"))?;
                if total > MAX_ALL_FONTS {
                    return Err(invalid("embedded fonts exceed total size limit"));
                }
                if part.rels().iter().next().is_some() {
                    return Err(invalid(format!(
                        "embedded font '{uri}' has nested relationships"
                    )));
                }
                let resource = Resource {
                    part_name: uri.to_string(),
                    content_type: part.content_type().into(),
                    data: part.blob_arc(),
                };
                cached.insert(target_name, resource.clone());
                embed.resource = Some(resource)
            }
        }
        for rel in source.rels().iter() {
            if is_font_relationship(rel.reltype()) && !used.contains(rel.r_id()) {
                return Err(invalid(format!(
                    "unreferenced font-table relationship '{}'",
                    rel.r_id()
                )));
            }
        }
        reject_orphan_font_parts(pkg, &targets)?;
        Ok(())
    }
}

/// Read the document font table and its bounded, inert font resources.
///
/// Embedded payload allocations are shared with the OPC package rather than
/// copied. The returned table can therefore be queried repeatedly without
/// reparsing or rediscovering the package graph.
pub fn read(package: &OpcPackage) -> Result<Option<Table>> {
    let (main_name, table_name, _) = locate_font_table(package)?;
    validate_font_table_relationship_sources(package, &main_name)?;
    let Some(table_name) = table_name else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let part = package.get_part(&table_name)?;
    Ok(Some(Table::extract_from_part(part, package)?))
}

/// Move a complete font table into the package after validating the staged XML
/// and OPC graph.
///
/// Font bytes are stored exactly as supplied. Callers that have unobfuscated
/// bytes must explicitly call [`obfuscate`] first. The API
/// operates on an already decrypted in-memory `OpcPackage` and invalidates any
/// package signatures immediately before the mutation phase. Moving a default,
/// empty [`Table`] removes the optional font-table graph and any font resources
/// that become unreferenced.
pub fn put(package: &mut OpcPackage, mut value: Table, conformance: Conformance) -> Result<bool> {
    validate_package_conformance(package, conformance)?;
    let old = read(package)?.unwrap_or_default();
    let (main_name, old_table_name, old_table_relationship_id) = locate_font_table(package)?;
    if value.fonts.is_empty()
        && value.namespaces.is_empty()
        && value.extension_attributes.is_empty()
    {
        return remove_graph(
            package,
            &old,
            &main_name,
            old_table_name.as_ref(),
            old_table_relationship_id.as_deref(),
        );
    }
    if old == value {
        return Ok(false);
    }
    allocate_font_identifiers(package, &mut value)?;
    validate_table_value(&value, true)?;
    let table_name = match old_table_name.clone() {
        Some(name) => name,
        None => next_font_table_part_name(package)?,
    };
    let table_relationship_id = match old_table_relationship_id.clone() {
        Some(id) => id,
        None => next_named_relationship_id(package.get_part(&main_name)?, "rIdTable")?,
    };
    if let Some(existing) = &old_table_name {
        let replaced = old_table_relationship_id
            .iter()
            .cloned()
            .collect::<HashSet<_>>();
        if has_inbound_outside_relationships(package, existing, &main_name, &replaced)? {
            return Err(invalid(format!(
                "shared font-table part '{existing}' cannot be overwritten"
            )));
        }
    }

    let xml = write(&value, conformance)?;
    let staged = parse(&xml)?;
    if !same_metadata(&staged, &value) {
        return Err(invalid("staged font-table XML did not round-trip"));
    }

    let old_relationship_ids = if let Some(name) = &old_table_name {
        package
            .get_part(name)?
            .rels()
            .iter()
            .filter(|relationship| is_font_relationship(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect::<HashSet<_>>()
    } else {
        HashSet::new()
    };
    let old_part_names = old
        .fonts
        .iter()
        .flat_map(|font| font.embedded_fonts.iter())
        .filter_map(|font| {
            font.resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
        .collect::<HashSet<_>>();
    let old_part_uris = old_part_names
        .iter()
        .map(|name| PackURI::new(name).map_err(Error::Uri))
        .collect::<Result<Vec<_>>>()?;

    let table_part = old_table_name
        .as_ref()
        .map(|name| package.get_part(name))
        .transpose()?;
    let mut relationships = HashMap::<String, PackURI>::new();
    let mut resources = HashMap::<String, (String, Arc<Vec<u8>>)>::new();
    for font in &value.fonts {
        for embedded in &font.embedded_fonts {
            if let Some(part) = table_part
                && part.rels().get(&embedded.relationship_id).is_some()
                && !old_relationship_ids.contains(&embedded.relationship_id)
            {
                return Err(invalid(format!(
                    "relationship ID '{}' already exists",
                    embedded.relationship_id
                )));
            }
            let resource = embedded
                .resource
                .as_ref()
                .ok_or_else(|| invalid("embedded-font resource is required for package storage"))?;
            let uri = PackURI::new(&resource.part_name).map_err(Error::Uri)?;
            if let Some(previous) = relationships.get(&embedded.relationship_id) {
                if previous != &uri {
                    return Err(invalid(format!(
                        "relationship ID '{}' has conflicting font targets",
                        embedded.relationship_id
                    )));
                }
            } else {
                relationships.insert(embedded.relationship_id.clone(), uri.clone());
            }
            if let Some((content_type, data)) = resources.get(uri.as_str()) {
                if content_type != &resource.content_type || data.as_slice() != resource.bytes() {
                    return Err(invalid(format!(
                        "shared font part '{uri}' has conflicting resources"
                    )));
                }
            } else {
                resources.insert(
                    uri.to_string(),
                    (resource.content_type.clone(), resource.share()),
                );
            }
        }
    }

    for (part_name, (content_type, data)) in &resources {
        let uri = PackURI::new(part_name).map_err(Error::Uri)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            if part.blob() != data.as_slice() && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if part.blob() != data.as_slice()
                && old_table_name.as_ref().is_some_and(|table| {
                    has_inbound_outside_relationships(package, &uri, table, &old_relationship_ids)
                        .unwrap_or(true)
                })
            {
                return Err(invalid(format!(
                    "shared font part '{uri}' cannot be overwritten"
                )));
            }
        }
    }
    validate_all_internal_relationship_targets(package)?;

    let resource_parts = resources
        .into_iter()
        .map(|(name, (content_type, data))| {
            PackURI::new(&name)
                .map(|uri| (uri, content_type, data))
                .map_err(Error::Uri)
        })
        .collect::<Result<Vec<_>>>()?;
    package.unsign();

    for (uri, content_type, data) in resource_parts {
        if let Ok(part) = package.get_part_mut(&uri) {
            part.set_blob_shared(data);
        } else {
            package.add_part(Box::new(BlobPart::new_shared(uri, content_type, data)));
        }
    }
    if let Some(existing) = &old_table_name {
        let part = package.get_part_mut(existing)?;
        let font_relationships = part
            .rels()
            .iter()
            .filter(|relationship| is_font_relationship(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect::<Vec<_>>();
        for relationship_id in font_relationships {
            part.rels_mut().remove(&relationship_id);
        }
        for (relationship_id, target) in &relationships {
            part.rels_mut().add_relationship(
                conformance.font_rel().into(),
                target.relative_ref(table_name.base_uri()),
                relationship_id.clone(),
                false,
            );
        }
        part.set_blob(xml);
    } else {
        let mut part = XmlPart::new(table_name.clone(), FT_CT.into(), xml);
        for (relationship_id, target) in &relationships {
            part.rels_mut().add_relationship(
                conformance.font_rel().into(),
                target.relative_ref(table_name.base_uri()),
                relationship_id.clone(),
                false,
            );
        }
        package.add_part(Box::new(part));
        package
            .get_part_mut(&main_name)?
            .rels_mut()
            .add_relationship(
                conformance.font_table_rel().into(),
                table_name.relative_ref(main_name.base_uri()),
                table_relationship_id,
                false,
            );
    }

    let retained = relationships
        .values()
        .map(ToString::to_string)
        .collect::<HashSet<_>>();
    for uri in old_part_uris {
        if !retained.contains(uri.as_str()) && !part_is_referenced(package, &uri)? {
            package.remove_part(&uri);
        }
    }
    Ok(true)
}

/// Remove the optional font-table graph and every font resource that becomes
/// unreferenced.
///
/// The complete relationship graph is validated before signatures, parts, or
/// relationships are mutated. Resources shared by another source are retained.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let old = read(package)?.unwrap_or_default();
    let (main_name, table_name, relationship_id) = locate_font_table(package)?;
    remove_graph(
        package,
        &old,
        &main_name,
        table_name.as_ref(),
        relationship_id.as_deref(),
    )
}

fn remove_graph(
    package: &mut OpcPackage,
    old: &Table,
    main_name: &PackURI,
    table_name: Option<&PackURI>,
    table_relationship_id: Option<&str>,
) -> Result<bool> {
    let Some(table_name) = table_name else {
        return Ok(false);
    };
    let table_relationship_id =
        table_relationship_id.ok_or_else(|| invalid("font-table relationship ID is missing"))?;
    let table_part = package.get_part(table_name)?;
    if table_part
        .rels()
        .iter()
        .any(|relationship| !is_font_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "font table with unknown outbound relationships cannot be removed safely",
        ));
    }
    let font_relationship_ids = table_part
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    let replaced_table_relationship = HashSet::from([table_relationship_id.to_owned()]);
    if has_inbound_outside_relationships(
        package,
        table_name,
        main_name,
        &replaced_table_relationship,
    )? {
        return Err(invalid(format!(
            "shared font-table part '{table_name}' cannot be removed"
        )));
    }

    let resource_names = old
        .fonts
        .iter()
        .flat_map(|font| font.embedded_fonts.iter())
        .filter_map(|embed| embed.resource.as_ref())
        .map(|resource| resource.part_name.as_str())
        .collect::<HashSet<_>>();
    let mut resources_to_remove = Vec::with_capacity(resource_names.len());
    for name in resource_names {
        let uri = PackURI::new(name).map_err(Error::Uri)?;
        if !has_inbound_outside_relationships(package, &uri, table_name, &font_relationship_ids)? {
            resources_to_remove.push(uri);
        }
    }
    validate_all_internal_relationship_targets(package)?;

    package.unsign();
    package
        .get_part_mut(main_name)?
        .rels_mut()
        .remove(table_relationship_id);
    package.remove_part(table_name);
    for uri in resources_to_remove {
        package.remove_part(&uri);
    }
    Ok(true)
}

/// Reject embedded typefaces that are not directly named by any `w:rFonts`.
/// Theme-based font resolution is intentionally not attempted.
pub fn validate_usage(package: &OpcPackage, table: &Table) -> Result<()> {
    let used = directly_used_font_names(package)?;
    let unused = table
        .fonts
        .iter()
        .filter(|font| !font.embedded_fonts.is_empty())
        .filter(|font| !used.contains(&name_key(&font.name)))
        .map(|font| font.name.clone())
        .collect::<Vec<_>>();
    if unused.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!(
            "embedded fonts are not directly used by the document: {}",
            unused.join(", ")
        )))
    }
}

fn locate_font_table(package: &OpcPackage) -> Result<(PackURI, Option<PackURI>, Option<String>)> {
    let main = package.main_document_part()?;
    let main_name = main.partname().clone();
    let mut matching = main
        .rels()
        .iter()
        .filter(|relationship| is_font_table_relationship(relationship.reltype()));
    let Some(relationship) = matching.next() else {
        return Ok((main_name, None, None));
    };
    if matching.next().is_some() {
        return Err(invalid("document has multiple font-table relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("font-table relationship cannot be external"));
    }
    Ok((
        main_name,
        Some(relationship.target_partname()?),
        Some(relationship.r_id().to_owned()),
    ))
}

fn allocate_font_identifiers(package: &OpcPackage, table: &mut Table) -> Result<()> {
    let (_, table_name, _) = locate_font_table(package)?;
    let mut relationship_ids = table_name
        .as_ref()
        .map(|name| {
            package.get_part(name).map(|part| {
                part.rels()
                    .iter()
                    .map(|relationship| relationship.r_id().to_owned())
                    .collect::<HashSet<_>>()
            })
        })
        .transpose()?
        .unwrap_or_default();
    relationship_ids.extend(table.fonts.iter().flat_map(|font| {
        font.embedded_fonts
            .iter()
            .filter(|embedded| !embedded.relationship_id.is_empty())
            .map(|embedded| embedded.relationship_id.clone())
    }));
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    part_names.extend(table.fonts.iter().flat_map(|font| {
        font.embedded_fonts.iter().filter_map(|embedded| {
            embedded
                .resource
                .as_ref()
                .filter(|resource| !resource.part_name.is_empty())
                .map(|resource| resource.part_name.clone())
        })
    }));
    let mut shared_names = HashMap::<usize, String>::new();
    for font in &table.fonts {
        for embedded in &font.embedded_fonts {
            if let Some(resource) = &embedded.resource
                && !resource.part_name.is_empty()
            {
                shared_names.insert(
                    Arc::as_ptr(&resource.data) as usize,
                    resource.part_name.clone(),
                );
            }
        }
    }
    for font in &mut table.fonts {
        for embedded in &mut font.embedded_fonts {
            if embedded.relationship_id.is_empty() {
                embedded.relationship_id = next_font_relationship_id(&relationship_ids)?;
            }
            relationship_ids.insert(embedded.relationship_id.clone());
            let resource = embedded
                .resource
                .as_mut()
                .ok_or_else(|| invalid("embedded-font resource is required"))?;
            if resource.part_name.is_empty() {
                let identity = Arc::as_ptr(&resource.data) as usize;
                resource.part_name = match shared_names.get(&identity) {
                    Some(name) => name.clone(),
                    None => {
                        let name = next_font_part_name(&part_names)?;
                        shared_names.insert(identity, name.clone());
                        name
                    },
                };
            }
            part_names.insert(resource.part_name.clone());
            if resource.content_type.is_empty() {
                resource.content_type = FONT_CT.into();
            }
        }
    }
    Ok(())
}

fn next_font_relationship_id(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdFont{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("too many font relationship IDs"))
}
fn next_font_part_name(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("/word/fonts/font{index}.odttf");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("too many font part names"))
}
fn next_font_table_part_name(package: &OpcPackage) -> Result<PackURI> {
    let used = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    if !used.contains("/word/fontTable.xml") {
        return PackURI::new("/word/fontTable.xml").map_err(Error::Uri);
    }
    for index in 1..=u32::MAX {
        let candidate = format!("/word/fontTable{index}.xml");
        if !used.contains(&candidate) {
            return PackURI::new(&candidate).map_err(Error::Uri);
        }
    }
    Err(invalid("too many font-table part names"))
}
fn next_named_relationship_id(source: &dyn Part, prefix: &str) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("{prefix}{index}");
        if source.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("too many relationship IDs"))
}

fn same_metadata(left: &Table, right: &Table) -> bool {
    left.namespaces == right.namespaces
        && left.extension_attributes == right.extension_attributes
        && left.fonts.len() == right.fonts.len()
        && left.fonts.iter().zip(&right.fonts).all(|(left, right)| {
            left.name == right.name
                && left.alternate_name == right.alternate_name
                && left.panose == right.panose
                && left.character_set == right.character_set
                && left.family == right.family
                && left.not_true_type == right.not_true_type
                && left.pitch == right.pitch
                && left.signature == right.signature
                && left.extension_attributes == right.extension_attributes
                && left.embedded_fonts.len() == right.embedded_fonts.len()
                && left
                    .embedded_fonts
                    .iter()
                    .zip(&right.embedded_fonts)
                    .all(|(left, right)| {
                        left.style == right.style
                            && left.relationship_id == right.relationship_id
                            && left.font_key == right.font_key
                            && left.subsetted == right.subsetted
                            && left.extension_attributes == right.extension_attributes
                    })
        })
}

fn validate_font_table_relationship_sources(package: &OpcPackage, main: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_table_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source a font-table relationship",
        ));
    }
    for part in package.iter_parts() {
        if part.partname() != main
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_table_relationship(relationship.reltype()))
            && part.content_type()
                != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
        {
            return Err(invalid(format!(
                "font-table relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_font_relationship_sources(package: &OpcPackage, table: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_relationship(relationship.reltype()))
    {
        return Err(invalid("package root cannot source a font relationship"));
    }
    for part in package.iter_parts() {
        if part.partname() != table
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_relationship(relationship.reltype()))
            && part.content_type() != FT_CT
        {
            return Err(invalid(format!(
                "font relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_package_conformance(package: &OpcPackage, requested: Conformance) -> Result<()> {
    const STRICT_OFFICE_DOCUMENT: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    | STRICT_OFFICE_DOCUMENT
            )
        })
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    let actual = if relationship.reltype() == STRICT_OFFICE_DOCUMENT {
        Conformance::Strict
    } else {
        Conformance::Transitional
    };
    if actual == requested {
        Ok(())
    } else {
        Err(invalid(
            "requested font-table conformance does not match the package relationship namespace",
        ))
    }
}

fn reject_orphan_font_parts(package: &OpcPackage, targets: &HashSet<String>) -> Result<()> {
    for part in package.iter_parts() {
        if (part.content_type() == FONT_CT || part.partname().as_str().starts_with("/word/fonts/"))
            && !targets.contains(part.partname().as_str())
            && !part_is_referenced(package, part.partname())?
        {
            return Err(invalid(format!("orphan font part '{}'", part.partname())));
        }
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> Result<bool> {
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

fn has_inbound_outside_relationships(
    package: &OpcPackage,
    target: &PackURI,
    table: &PackURI,
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
                && (part.partname() != table
                    || !replaced_relationships.contains(relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn validate_all_internal_relationship_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        relationship.target_partname()?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            relationship.target_partname()?;
        }
    }
    Ok(())
}

fn directly_used_font_names(package: &OpcPackage) -> Result<HashSet<String>> {
    let mut output = HashSet::new();
    for part in package.iter_parts().filter(|part| {
        part.content_type().contains("wordprocessingml")
            && part.content_type().ends_with("+xml")
            && part.content_type() != FT_CT
    }) {
        if part.blob().len() > MAX_XML {
            return Err(invalid(format!(
                "WordprocessingML part '{}' is too large for font-usage validation",
                part.partname()
            )));
        }
        let mut reader = Reader::from_reader(part.blob());
        let mut nodes = 0usize;
        loop {
            match reader.read_event().map_err(xml_error)? {
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"rFonts" =>
                {
                    nodes += 1;
                    if nodes > MAX_NODES {
                        return Err(invalid("font-usage XML node limit exceeded"));
                    }
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(xml_error)?;
                        if matches!(
                            attribute.key.local_name().as_ref(),
                            b"ascii" | b"hAnsi" | b"eastAsia" | b"cs"
                        ) {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(xml_error)?;
                            bounded(&value)?;
                            output.insert(name_key(&value));
                        }
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTDs and processing instructions are rejected"));
                },
                Event::Eof => break,
                _ => {},
            }
        }
    }
    Ok(output)
}
