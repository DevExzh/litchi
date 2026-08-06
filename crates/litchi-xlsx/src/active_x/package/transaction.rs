//! Snapshot-safe package transactions for inert ActiveX graphs.

use super::super::codec::{controls_span, descriptor_relationship_ids, replace_controls_xml};
use super::super::model::*;
use super::super::validation::{validate_control_set, validate_part_location, validate_rel_id};
use super::super::{
    BINARY_CONTENT_TYPE, BINARY_REL, CONTROL_REL, CONTROL_REL_STRICT, DESCRIPTOR_CONTENT_TYPE,
    IMAGE_REL, IMAGE_REL_STRICT, MAX_BINARY, Result, WORKSHEET_CONTENT_TYPE, invalid, limit,
    relerr,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use std::collections::HashSet;

/// Stores a complete graph after preparing all XML, relationship, URI, and
/// resource invariants. The package is mutated only after preparation passes.
pub(super) fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
) -> Result<()> {
    let prepared = prepare_graph(package, worksheet_uri, value, true)?;
    install_graph(package, worksheet_uri, prepared)
}

struct PreparedGraph {
    worksheet_xml: Vec<u8>,
    strict: bool,
    descriptors: Vec<PreparedDescriptor>,
    resources: Vec<(PackURI, String, Vec<u8>)>,
    worksheet_relationships: Vec<(String, PackURI, bool)>,
}

struct PreparedDescriptor {
    uri: PackURI,
    xml: Vec<u8>,
    relationships: Vec<(String, PackURI)>,
}

fn prepare_graph(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
    require_empty: bool,
) -> Result<PreparedGraph> {
    validate_control_set(value)?;
    let worksheet = package.get_part(worksheet_uri)?;
    if worksheet.content_type() != WORKSHEET_CONTENT_TYPE {
        return Err(super::super::content_type(
            WORKSHEET_CONTENT_TYPE,
            worksheet.content_type(),
        ));
    }
    let existing = Controls::parse(worksheet.blob())?;
    if require_empty
        && (!existing.controls.is_empty()
            || worksheet
                .rels()
                .iter()
                .any(|rel| matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT)))
    {
        return Err(invalid("worksheet already has an ActiveX control graph"));
    }
    let controls = Controls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    };
    let worksheet_xml = replace_controls_xml(worksheet.blob(), &controls)?;
    let strict = controls_span(worksheet.blob())?.strict;

    let mut occupied_ids: HashSet<String> = worksheet
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    let mut part_uris = HashSet::new();
    let mut descriptors = Vec::with_capacity(value.controls.len());
    let mut resources = Vec::new();
    let mut worksheet_relationships = Vec::new();
    for item in &value.controls {
        validate_rel_id(&item.control.relationship_id)?;
        if !occupied_ids.insert(item.control.relationship_id.clone()) {
            return Err(relerr("duplicate or occupied worksheet relationship ID"));
        }
        validate_part_location(&item.descriptor_uri, "/xl/activeX/", "ActiveX descriptor")?;
        reserve_new_part(package, &mut part_uris, &item.descriptor_uri)?;
        let descriptor_xml = item.descriptor.to_xml()?;
        let expected_ids = descriptor_relationship_ids(&item.descriptor)?;
        let actual_ids: HashSet<String> = item
            .binaries
            .iter()
            .map(|binary| binary.relationship_id.clone())
            .collect();
        if expected_ids.iter().collect::<HashSet<_>>() != actual_ids.iter().collect::<HashSet<_>>()
            || actual_ids.len() != item.binaries.len()
        {
            return Err(relerr(
                "descriptor relationship IDs must exactly match supplied binaries",
            ));
        }
        let mut descriptor_rels = Vec::with_capacity(item.binaries.len());
        for binary in &item.binaries {
            validate_rel_id(&binary.relationship_id)?;
            validate_part_location(&binary.part_uri, "/xl/activeX/", "ActiveX binary")?;
            if binary.bytes.len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            reserve_new_part(package, &mut part_uris, &binary.part_uri)?;
            descriptor_rels.push((binary.relationship_id.clone(), binary.part_uri.clone()));
            resources.push((
                binary.part_uri.clone(),
                BINARY_CONTENT_TYPE.into(),
                binary.bytes.clone(),
            ));
        }
        worksheet_relationships.push((
            item.control.relationship_id.clone(),
            item.descriptor_uri.clone(),
            false,
        ));
        match (&item.control.properties, &item.preview) {
            (Some(properties), Some(preview)) => {
                if properties.preview_relationship_id.as_deref()
                    != Some(preview.relationship_id.as_str())
                {
                    return Err(relerr(
                        "control preview relationship ID does not match supplied preview",
                    ));
                }
                validate_rel_id(&preview.relationship_id)?;
                if !occupied_ids.insert(preview.relationship_id.clone()) {
                    return Err(relerr("duplicate or occupied worksheet relationship ID"));
                }
                validate_part_location(&preview.part_uri, "/xl/media/", "ActiveX preview")?;
                if !preview.content_type.starts_with("image/") {
                    return Err(invalid("ActiveX preview content type must be image/*"));
                }
                if preview.bytes.len() > MAX_BINARY {
                    return Err(limit("ActiveX preview image bytes"));
                }
                reserve_new_part(package, &mut part_uris, &preview.part_uri)?;
                worksheet_relationships.push((
                    preview.relationship_id.clone(),
                    preview.part_uri.clone(),
                    true,
                ));
                resources.push((
                    preview.part_uri.clone(),
                    preview.content_type.clone(),
                    preview.bytes.clone(),
                ));
            },
            (Some(properties), None) if properties.preview_relationship_id.is_some() => {
                return Err(relerr("control references a preview that was not supplied"));
            },
            (_, Some(_)) => return Err(relerr("supplied preview is not referenced by controlPr")),
            _ => {},
        }
        descriptors.push(PreparedDescriptor {
            uri: item.descriptor_uri.clone(),
            xml: descriptor_xml,
            relationships: descriptor_rels,
        });
    }
    Ok(PreparedGraph {
        worksheet_xml,
        strict,
        descriptors,
        resources,
        worksheet_relationships,
    })
}

fn install_graph(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    prepared: PreparedGraph,
) -> Result<()> {
    package.unsign();
    for (uri, content_type, bytes) in prepared.resources {
        package.try_add_part(Box::new(BlobPart::new(uri, content_type, bytes)))?;
    }
    for descriptor in prepared.descriptors {
        let mut part = BlobPart::new(
            descriptor.uri.clone(),
            DESCRIPTOR_CONTENT_TYPE.into(),
            descriptor.xml,
        );
        for (id, target) in descriptor.relationships {
            part.rels_mut().try_add_relationship(
                BINARY_REL.into(),
                target.relative_ref(descriptor.uri.base_uri()),
                id,
                TargetMode::Internal,
            )?;
        }
        package.try_add_part(Box::new(part))?;
    }
    let worksheet = package.get_part_mut(worksheet_uri)?;
    for (id, target, preview) in prepared.worksheet_relationships {
        worksheet.rels_mut().try_add_relationship(
            if preview {
                if prepared.strict {
                    IMAGE_REL_STRICT
                } else {
                    IMAGE_REL
                }
            } else if prepared.strict {
                CONTROL_REL_STRICT
            } else {
                CONTROL_REL
            }
            .into(),
            target.relative_ref(worksheet_uri.base_uri()),
            id,
            TargetMode::Internal,
        )?;
    }
    worksheet.set_blob(prepared.worksheet_xml);
    Ok(())
}

fn reserve_new_part(
    package: &OpcPackage,
    reserved: &mut HashSet<PackURI>,
    uri: &PackURI,
) -> Result<()> {
    if reserved.iter().any(|other| other.is_equivalent_to(uri)) {
        return Err(invalid("ActiveX graph contains conflicting part names"));
    }
    package.validate_new_part_name(uri)?;
    reserved.insert(uri.clone());
    Ok(())
}

pub(super) fn part_has_inbound_relationship(
    package: &OpcPackage,
    target: &PackURI,
) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}
