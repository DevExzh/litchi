//! Atomic slide-level package ownership for PresentationML model3d frames.

use litchi_drawingml::model3d as drawing;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};

use super::codec::{self, Inventory, Location};
use super::model::{
    Asset, Data, Link, Model, Origin, Preview, Relation, Scene, Shape, relationship_id,
};
use super::validation;
use super::{MAX_TOTAL_PAYLOAD_BYTES, MODEL_CONTENT_TYPE, MODEL_RELATIONSHIP};
use crate::shape::Key;
use crate::{Error, Result};

/// Load every model3d frame owned by one slide in source order.
pub(crate) fn load_all(package: &OpcPackage, slide_name: &PackURI) -> Result<Vec<Model>> {
    let source = slide_part(package, slide_name)?;
    let inventory = codec::locate(source.blob())?;
    load_inventory(package, slide_name, source, &inventory)
}

/// Load one model3d frame selected by its semantic shape anchor.
pub(crate) fn load(
    package: &OpcPackage,
    slide_name: &PackURI,
    key: Key<'_>,
) -> Result<Option<Model>> {
    let source = slide_part(package, slide_name)?;
    let inventory = codec::locate(source.blob())?;
    let Some(location) = select_location(&inventory, key)? else {
        return Ok(None);
    };
    let models = load_inventory(package, slide_name, source, &inventory)?;
    Ok(models.into_iter().find(|model| {
        model.shape.index() == location.shape_index
            && model
                .base_xml
                .as_slice()
                .get(location.range.clone())
                .is_some()
    }))
}

/// Replace one existing model3d frame transactionally.
pub(crate) fn put(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    key: Key<'_>,
    value: Model,
) -> Result<Option<Model>> {
    validation::validate(&value)?;
    let source = slide_part(package, slide_name)?;
    if value.base_xml.as_slice() != source.blob() {
        return Err(invalid("model3d edit is based on a stale slide snapshot"));
    }
    let inventory = codec::locate(source.blob())?;
    let Some(location) = select_location(&inventory, key)? else {
        return Err(invalid(
            "model3d put requires an existing graphic-frame owner",
        ));
    };
    let models = load_inventory(package, slide_name, source, &inventory)?;
    let previous = models
        .into_iter()
        .find(|model| model.shape.index() == location.shape_index)
        .ok_or_else(|| invalid("model3d shape anchor disappeared during preflight"))?;
    if previous.shape != value.shape {
        return Err(invalid(
            "model3d replacement targets a different shape anchor",
        ));
    }
    if value.semantic_eq(&previous) {
        return Ok(Some(previous));
    }

    let old_xml = source.blob_arc();
    let mut scene = value.scene.wire.clone();
    scene.reference.embedded = stage_embedded(
        package,
        slide_name,
        value.asset.embedded_data(),
        previous.asset.embedded_data(),
        previous.origin.model_embedded.as_ref(),
        MODEL_CONTENT_TYPE,
        "/ppt/media/model3d%d.glb",
        MODEL_RELATIONSHIP,
    )?;
    scene.reference.linked = stage_linked(
        package,
        slide_name,
        value.asset.linked_target(),
        previous.asset.linked_target(),
        previous.origin.model_linked.as_ref(),
        MODEL_RELATIONSHIP,
    )?;

    stage_preview(
        package,
        slide_name,
        &mut scene,
        value.preview.as_ref(),
        previous.preview.as_ref(),
        previous.origin.preview_embedded.as_ref(),
        previous.origin.preview_linked.as_ref(),
    )?;
    let fragment = codec::write(&Scene::from_wire(scene))?;
    let updated = codec::replace(old_xml.as_slice(), location.range.clone(), &fragment)?;
    package.get_part_mut(slide_name)?.set_blob(updated.clone());

    let old_relationships = old_model_relationships(&previous.origin);
    for relationship_id in old_relationships {
        cleanup_relationship(package, slide_name, &updated, &relationship_id)?;
    }
    package.unsign();
    Ok(Some(previous))
}

/// Remove one model3d frame and collect only resources no longer reachable.
pub(crate) fn remove(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    key: Key<'_>,
) -> Result<Option<Model>> {
    let source = slide_part(package, slide_name)?;
    let inventory = codec::locate(source.blob())?;
    let Some(location) = select_location(&inventory, key)? else {
        return Ok(None);
    };
    let models = load_inventory(package, slide_name, source, &inventory)?;
    let previous = models
        .into_iter()
        .find(|model| model.shape.index() == location.shape_index)
        .ok_or_else(|| invalid("model3d shape anchor disappeared during preflight"))?;
    let updated = codec::replace(source.blob(), location.range.clone(), &[])?;
    package.get_part_mut(slide_name)?.set_blob(updated.clone());
    for relationship_id in old_model_relationships(&previous.origin) {
        cleanup_relationship(package, slide_name, &updated, &relationship_id)?;
    }
    package.unsign();
    Ok(Some(previous))
}

fn load_inventory(
    package: &OpcPackage,
    slide_name: &PackURI,
    source: &dyn Part,
    inventory: &Inventory,
) -> Result<Vec<Model>> {
    let mut models = Vec::new();
    let mut total_payload = 0usize;
    for location in &inventory.locations {
        let model = load_location(package, slide_name, source, location)?;
        total_payload = total_payload
            .saturating_add(model.asset.embedded_data().map_or(0, Data::len))
            .saturating_add(
                model
                    .preview
                    .as_ref()
                    .and_then(Preview::data)
                    .map_or(0, Data::len),
            );
        if total_payload > MAX_TOTAL_PAYLOAD_BYTES {
            return Err(limit(
                "model3d slide payload bytes",
                MAX_TOTAL_PAYLOAD_BYTES,
            ));
        }
        models.push(model);
    }
    validation::validate_model_relationship_inventory(source, &models)?;
    Ok(models)
}

fn load_location(
    package: &OpcPackage,
    _slide_name: &PackURI,
    source: &dyn Part,
    location: &Location,
) -> Result<Model> {
    let xml = source
        .blob()
        .get(location.range.clone())
        .ok_or_else(|| invalid("model3d location is outside its slide snapshot"))?;
    let scene = codec::read(xml)?;
    let mut asset = Asset::default();
    let mut origin = Origin::default();

    if let Some(id) = scene.wire.reference.embedded.as_ref() {
        let relation = validation::relation_from_part(source, id)?;
        let data = load_internal_data(package, source, id)?;
        asset.set_embedded(Some(data));
        origin.model_embedded = Some(relation);
    }
    if let Some(id) = scene.wire.reference.linked.as_ref() {
        let relation = validation::relation_from_part(source, id)?;
        asset.set_linked(Some(Link::new(relation.target.as_str())?));
        origin.model_linked = Some(relation);
    }

    let preview = load_preview(package, source, &scene, &mut origin)?;
    origin.asset = asset.clone();
    origin.preview = preview.clone();
    let model = Model {
        scene,
        asset,
        preview,
        shape: Shape::from_location(location.shape_index, location.shape_name.clone()),
        base_xml: source.blob_arc(),
        origin,
    };
    validation::validate_loaded_graph(package, source, &model)?;
    Ok(model)
}

fn load_preview(
    package: &OpcPackage,
    source: &dyn Part,
    scene: &Scene,
    origin: &mut Origin,
) -> Result<Option<Preview>> {
    let Some(raster) = scene.wire.raster() else {
        return Ok(None);
    };
    let Some(blip) = raster.children.iter().find_map(|child| match child {
        drawing::RasterChild::Blip(blip) => Some(blip),
        drawing::RasterChild::Opaque(_) => None,
        _ => None,
    }) else {
        return Ok(None);
    };
    let mut preview = None;
    if let Some(id) = blip.reference.embedded.as_ref() {
        let relation = validation::relation_from_part(source, id)?;
        let data = load_internal_data(package, source, id)?;
        let target = source
            .rels()
            .get(id.as_str())
            .ok_or_else(|| invalid("model3d preview relationship disappeared"))?
            .target_partname()?;
        let part = package.get_part(&target)?;
        preview = Some(Preview::embedded(data, part.content_type())?);
        origin.preview_embedded = Some(relation);
    }
    if let Some(id) = blip.reference.linked.as_ref() {
        let relation = validation::relation_from_part(source, id)?;
        let link = Link::new(relation.target.as_str())?;
        if let Some(value) = preview.as_mut() {
            value.set_linked(Some(link));
        } else {
            preview = Some(Preview::linked(link.as_str())?);
        }
        origin.preview_linked = Some(relation);
    }
    Ok(preview)
}

fn load_internal_data(package: &OpcPackage, source: &dyn Part, id: &drawing::Id) -> Result<Data> {
    let relationship = source
        .rels()
        .get(id.as_str())
        .ok_or_else(|| invalid(format!("relationship '{}' is missing", id)))?;
    if relationship.is_external() {
        return Err(invalid(format!(
            "relationship '{}' unexpectedly targets an external resource",
            id
        )));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    Data::from_shared(part.blob_arc())
}

fn slide_part<'a>(package: &'a OpcPackage, slide_name: &PackURI) -> Result<&'a dyn Part> {
    let part = package.get_part(slide_name)?;
    crate::parts::SlidePart::from_part(part)?;
    Ok(part)
}

fn select_location<'a>(inventory: &'a Inventory, key: Key<'_>) -> Result<Option<&'a Location>> {
    let shape_index = match key {
        Key::Index(index) => {
            if index >= inventory.shape_names.len() {
                return Err(crate::shape::LookupError::IndexOutOfBounds {
                    index,
                    len: inventory.shape_names.len(),
                }
                .into());
            }
            index
        },
        Key::Name(name) => {
            let mut found = None;
            let mut matches = 0usize;
            for (index, value) in inventory.shape_names.iter().enumerate() {
                if value.as_deref() == Some(name) {
                    found = Some(index);
                    matches = matches.saturating_add(1);
                }
            }
            if matches > 1 {
                return Err(crate::shape::LookupError::AmbiguousName {
                    name: name.to_owned(),
                    matches,
                }
                .into());
            }
            let Some(index) = found else {
                return Err(crate::shape::LookupError::NameNotFound {
                    name: name.to_owned(),
                }
                .into());
            };
            index
        },
    };
    Ok(inventory
        .locations
        .iter()
        .find(|location| location.shape_index == shape_index))
}

fn stage_embedded(
    package: &mut OpcPackage,
    source_name: &PackURI,
    desired: Option<&Data>,
    previous: Option<&Data>,
    previous_relation: Option<&Relation>,
    content_type: &str,
    template: &str,
    relationship_type: &str,
) -> Result<Option<drawing::Id>> {
    let Some(data) = desired else {
        return Ok(None);
    };
    if previous == Some(data)
        && previous_relation.is_some_and(|relation| {
            package
                .get_part(source_name)
                .ok()
                .and_then(|part| part.rels().get(&relation.id))
                .is_some_and(|value| !value.is_external())
        })
    {
        return previous_relation
            .map(|relation| relationship_id(&relation.id))
            .transpose();
    }
    let target = package.next_partname(template)?;
    package.validate_new_part_name(&target)?;
    package.add_part(Box::new(BlobPart::new_shared(
        target.clone(),
        content_type.to_owned(),
        data.shared(),
    )));
    let target_ref = target.relative_ref(source_name.base_uri());
    let id = package
        .get_part_mut(source_name)?
        .relate_to(&target_ref, relationship_type);
    relationship_id(&id).map(Some)
}

fn stage_linked(
    package: &mut OpcPackage,
    source_name: &PackURI,
    desired: Option<&Link>,
    previous: Option<&Link>,
    previous_relation: Option<&Relation>,
    relationship_type: &str,
) -> Result<Option<drawing::Id>> {
    let Some(link) = desired else {
        return Ok(None);
    };
    if previous == Some(link)
        && previous_relation.is_some_and(|relation| {
            package
                .get_part(source_name)
                .ok()
                .and_then(|part| part.rels().get(&relation.id))
                .is_some_and(|value| value.is_external())
        })
    {
        return previous_relation
            .map(|relation| relationship_id(&relation.id))
            .transpose();
    }
    let id = package
        .get_part_mut(source_name)?
        .relate_to_ext(link.as_str(), relationship_type);
    relationship_id(&id).map(Some)
}

fn stage_preview(
    package: &mut OpcPackage,
    source_name: &PackURI,
    scene: &mut drawing::Metadata,
    desired: Option<&Preview>,
    previous: Option<&Preview>,
    previous_embedded: Option<&Relation>,
    previous_linked: Option<&Relation>,
) -> Result<()> {
    let embedded = stage_embedded(
        package,
        source_name,
        desired.and_then(Preview::data),
        previous.and_then(Preview::data),
        previous_embedded,
        desired
            .and_then(Preview::content_type)
            .unwrap_or("image/png"),
        "/ppt/media/model3d-preview%d.bin",
        rt::IMAGE,
    )?;
    let linked = stage_linked(
        package,
        source_name,
        desired.and_then(Preview::linked_target),
        previous.and_then(Preview::linked_target),
        previous_linked,
        rt::IMAGE,
    )?;
    if let Some(raster) = scene.children.iter_mut().find_map(|child| match child {
        drawing::Child::Raster(raster) => Some(raster),
        drawing::Child::Opaque(_) => None,
        _ => None,
    }) {
        if embedded.is_some() || linked.is_some() || desired.is_some() {
            let mut blip = raster.blip().cloned().unwrap_or_default();
            blip.reference.embedded = embedded;
            blip.reference.linked = linked;
            raster.set_blip(blip);
        } else if let Some(blip) = raster.children.iter_mut().find_map(|child| match child {
            drawing::RasterChild::Blip(blip) => Some(blip),
            drawing::RasterChild::Opaque(_) => None,
            _ => None,
        }) {
            blip.reference = drawing::Reference::none();
        }
    }
    Ok(())
}

fn old_model_relationships(origin: &Origin) -> Vec<String> {
    [
        origin.model_embedded.as_ref(),
        origin.model_linked.as_ref(),
        origin.preview_embedded.as_ref(),
        origin.preview_linked.as_ref(),
    ]
    .into_iter()
    .flatten()
    .map(|relation| relation.id.clone())
    .collect()
}

fn cleanup_relationship(
    package: &mut OpcPackage,
    source_name: &PackURI,
    updated_xml: &[u8],
    relationship_id: &str,
) -> Result<()> {
    if updated_xml
        .windows(relationship_id.len())
        .any(|window| window == relationship_id.as_bytes())
    {
        return Ok(());
    }
    let target = {
        let source = package.get_part(source_name)?;
        let Some(relationship) = source.rels().get(relationship_id) else {
            return Ok(());
        };
        (!relationship.is_external())
            .then(|| relationship.target_partname())
            .transpose()?
    };
    package
        .get_part_mut(source_name)?
        .rels_mut()
        .remove(relationship_id);
    if let Some(target) = target
        && !has_inbound(package, &target, source_name, relationship_id)?
    {
        package.remove_part(&target);
    }
    Ok(())
}

fn has_inbound(
    package: &OpcPackage,
    target: &PackURI,
    selected_source: &PackURI,
    selected_relationship: &str,
) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if source.partname().is_equivalent_to(selected_source)
                && relationship.r_id() == selected_relationship
            {
                continue;
            }
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(format!("PPTX model3d {}", message.into()))
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
