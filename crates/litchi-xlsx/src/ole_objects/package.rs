//! OPC graph ownership for worksheet OLE objects.

use crate::error::Result;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::collections::{HashMap, HashSet};

use super::codec::{
    crate_conformance, insert_collection, parse_document, parse_ole_objects, write_ole_objects,
};
use super::invalid;
use super::model::{
    OleObjectConformance, OleObjectRelationshipKind, OleObjectResource, OleObjectTarget,
    OleObjects, add_payload, is_image_content_type, validate_id, validate_resource, validate_value,
};

/// Validate every worksheet OLE owner and reject orphan embedded payloads.
pub(super) fn validate_graph(
    package: &OpcPackage,
    worksheet_name: &PackURI,
) -> Result<Option<OleObjects>> {
    if package
        .rels()
        .iter()
        .any(|relationship| embedded_kind(relationship.reltype()).is_some())
    {
        return Err(invalid(
            "package root cannot source embedded-object relationships",
        ));
    }

    let worksheet_names = package
        .iter_parts()
        .filter(|part| part.content_type() == ct::SML_WORKSHEET)
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    let mut embedded_targets = HashSet::new();
    for worksheet in &worksheet_names {
        let _ = load_ole_objects(package, worksheet)?;
        let part = package.get_part(worksheet)?;
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| embedded_kind(relationship.reltype()).is_some())
        {
            if relationship.is_external() {
                continue;
            }
            embedded_targets.insert(relationship.target_partname()?.to_string());
        }
    }
    for part in package.iter_parts().filter(|part| {
        part.partname().as_str().starts_with("/xl/embeddings/")
            && (part.content_type() == ct::OFC_OLE_OBJECT
                || embedded_targets.contains(part.partname().as_str()))
    }) {
        if !embedded_targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "embedded payload '{}' has no owning relationship",
                part.partname()
            )));
        }
    }
    load_ole_objects(package, worksheet_name)
}

pub fn load_ole_objects(
    package: &OpcPackage,
    worksheet_name: &PackURI,
) -> Result<Option<OleObjects>> {
    if package
        .rels()
        .iter()
        .any(|relationship| embedded_kind(relationship.reltype()).is_some())
    {
        return Err(invalid(
            "package root cannot source embedded-object relationships",
        ));
    }
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    let Some(mut value) = parse_ole_objects(worksheet.blob())? else {
        if worksheet
            .rels()
            .iter()
            .any(|relationship| embedded_kind(relationship.reltype()).is_some())
        {
            return Err(invalid(
                "worksheet has embedded-object relationships without oleObjects markup",
            ));
        }
        return Ok(None);
    };
    let mut referenced = HashSet::new();
    let mut targets = HashSet::new();
    let mut total = 0usize;
    for object in &mut value.objects {
        if !referenced.insert(object.relationship_id.clone()) {
            return Err(invalid(format!(
                "duplicate object relationship reference '{}'",
                object.relationship_id
            )));
        }
        let relationship = worksheet
            .rels()
            .get(&object.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing object relationship '{}'",
                    object.relationship_id
                ))
            })?;
        let kind = embedded_kind(relationship.reltype()).ok_or_else(|| {
            invalid(format!(
                "relationship '{}' is not an embedded-object relationship",
                object.relationship_id
            ))
        })?;
        object.relationship_kind = kind;
        object.target = Some(if relationship.is_external() {
            if object.link.is_none() {
                return Err(invalid("external OLE relationship requires a link moniker"));
            }
            OleObjectTarget::External(relationship.target_ref().to_owned())
        } else {
            let target = relationship.target_partname()?;
            if !targets.insert(target.to_string()) {
                return Err(invalid(format!("multiple OLE objects target '{target}'")));
            }
            if !target.as_str().starts_with("/xl/embeddings/") {
                return Err(invalid(format!(
                    "embedded object '{target}' is outside /xl/embeddings"
                )));
            }
            let part = package.get_part(&target)?;
            validate_payload(part, kind)?;
            add_payload(&mut total, part.blob().len())?;
            OleObjectTarget::Internal(OleObjectResource {
                part_name: target.to_string(),
                content_type: part.content_type().to_owned(),
                data: part.blob().to_vec(),
            })
        });
        if let Some(properties) = object.properties.as_mut() {
            let relationship = worksheet
                .rels()
                .get(&properties.preview_relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "missing object preview relationship '{}'",
                        properties.preview_relationship_id
                    ))
                })?;
            if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
                || relationship.is_external()
            {
                return Err(invalid(
                    "object preview relationship must be an internal image",
                ));
            }
            let target = relationship.target_partname()?;
            if !target.as_str().starts_with("/xl/media/") {
                return Err(invalid(format!(
                    "object preview '{target}' is outside /xl/media"
                )));
            }
            let part = package.get_part(&target)?;
            if !is_image_content_type(part.content_type()) || !part.rels().is_empty() {
                return Err(invalid(format!(
                    "object preview '{target}' is not a relationship-free image"
                )));
            }
            add_payload(&mut total, part.blob().len())?;
            properties.preview = Some(OleObjectResource {
                part_name: target.to_string(),
                content_type: part.content_type().to_owned(),
                data: part.blob().to_vec(),
            });
        }
    }
    for relationship in worksheet
        .rels()
        .iter()
        .filter(|relationship| embedded_kind(relationship.reltype()).is_some())
    {
        if !referenced.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced embedded-object relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    Ok(Some(value))
}

/// Adds a new OLE collection and its inert relationships to one worksheet.
pub fn store_ole_objects(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    value: &OleObjects,
    conformance: OleObjectConformance,
) -> Result<()> {
    validate_value(value, true)?;
    if load_ole_objects(package, worksheet_name)?.is_some() {
        return Err(invalid("worksheet already contains OLE objects"));
    }
    let worksheet = package.get_part(worksheet_name)?;
    let root = parse_document(worksheet.blob())?;
    if crate_conformance(&root)? != conformance {
        return Err(invalid(
            "requested conformance does not match worksheet namespace",
        ));
    }
    let fragment = write_ole_objects(value, conformance)?;
    let updated = insert_collection(worksheet.blob(), &fragment, conformance)?;
    let mut relationships: HashMap<String, (String, String, bool)> = HashMap::new();
    let mut parts: HashMap<String, OleObjectResource> = HashMap::new();
    for object in &value.objects {
        let target = object
            .target
            .as_ref()
            .ok_or_else(|| invalid("OLE target is required for package storage"))?;
        let relationship_type = match object.relationship_kind {
            OleObjectRelationshipKind::OleObject => conformance.ole_rel(),
            OleObjectRelationshipKind::Package => conformance.package_rel(),
        };
        match target {
            OleObjectTarget::External(target) => add_relationship_plan(
                &mut relationships,
                &object.relationship_id,
                relationship_type,
                target,
                true,
            )?,
            OleObjectTarget::Internal(resource) => {
                let uri = resource_uri(resource, "/xl/embeddings/")?;
                add_part_plan(package, &mut parts, resource)?;
                add_relationship_plan(
                    &mut relationships,
                    &object.relationship_id,
                    relationship_type,
                    &uri.relative_ref(worksheet_name.base_uri()),
                    false,
                )?;
            },
        }
        if let Some(properties) = &object.properties {
            let preview = properties.preview.as_ref().ok_or_else(|| {
                invalid("object preview resource is required for package storage")
            })?;
            let uri = resource_uri(preview, "/xl/media/")?;
            add_part_plan(package, &mut parts, preview)?;
            add_relationship_plan(
                &mut relationships,
                &properties.preview_relationship_id,
                conformance.image_rel(),
                &uri.relative_ref(worksheet_name.base_uri()),
                false,
            )?;
        }
    }
    for id in relationships.keys() {
        if package.get_part(worksheet_name)?.rels().get(id).is_some() {
            return Err(invalid(format!(
                "worksheet relationship ID '{id}' already exists"
            )));
        }
    }
    package.get_part_mut(worksheet_name)?.set_blob(updated);
    for resource in parts.into_values() {
        let uri = PackURI::new(&resource.part_name).map_err(invalid)?;
        package.add_part(Box::new(BlobPart::new(
            uri,
            resource.content_type,
            resource.data,
        )));
    }
    for (id, (relationship_type, target, external)) in relationships {
        package
            .get_part_mut(worksheet_name)?
            .rels_mut()
            .add_relationship(relationship_type, target, id, external);
    }
    Ok(())
}

fn validate_payload(part: &dyn Part, kind: OleObjectRelationshipKind) -> Result<()> {
    if kind == OleObjectRelationshipKind::OleObject && part.content_type() != ct::OFC_OLE_OBJECT {
        return Err(invalid(format!(
            "OLE payload '{}' has invalid content type '{}'",
            part.partname(),
            part.content_type()
        )));
    }
    for relationship in part.rels().iter() {
        if !matches!(relationship.reltype(), rt::HYPERLINK | rt::STRICT_HYPERLINK) {
            return Err(invalid(format!(
                "embedded payload '{}' has forbidden outbound relationship",
                part.partname()
            )));
        }
    }
    Ok(())
}
fn embedded_kind(value: &str) -> Option<OleObjectRelationshipKind> {
    match value {
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT => Some(OleObjectRelationshipKind::OleObject),
        rt::PACKAGE | rt::STRICT_PACKAGE => Some(OleObjectRelationshipKind::Package),
        _ => None,
    }
}
fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a worksheet",
            part.partname()
        )))
    }
}
fn resource_uri(resource: &OleObjectResource, prefix: &str) -> Result<PackURI> {
    validate_resource(resource, prefix)?;
    PackURI::new(&resource.part_name).map_err(invalid)
}
fn add_part_plan(
    package: &OpcPackage,
    parts: &mut HashMap<String, OleObjectResource>,
    resource: &OleObjectResource,
) -> Result<()> {
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == resource.part_name)
    {
        return Err(invalid(format!(
            "resource part '{}' already exists",
            resource.part_name
        )));
    }
    if let Some(existing) = parts.get(&resource.part_name) {
        if existing != resource {
            return Err(invalid(format!(
                "conflicting resource part '{}'",
                resource.part_name
            )));
        }
    } else {
        parts.insert(resource.part_name.clone(), resource.clone());
    }
    Ok(())
}
fn add_relationship_plan(
    plans: &mut HashMap<String, (String, String, bool)>,
    id: &str,
    kind: &str,
    target: &str,
    external: bool,
) -> Result<()> {
    validate_id(id)?;
    let plan = (kind.to_owned(), target.to_owned(), external);
    if let Some(existing) = plans.get(id) {
        if existing != &plan {
            return Err(invalid(format!("conflicting relationship ID '{id}'")));
        }
    } else {
        plans.insert(id.to_owned(), plan);
    }
    Ok(())
}
