use super::{
    BTreeMap, BlobPart, Conformance, Data, Error, Event, Extension, HashSet, Kind, Limits, List,
    Namespace, NsReader, Offset, OpcPackage, PackURI, Part, ResolveResult, Resource, Result,
    STRICT_AUDIO_REL, STRICT_VIDEO_REL, TimeParseError, ct, document_conformance_with_limits,
    invalid, limit, parse, rt, validate_id, write_pictures, xml_error,
};
use crate::package::replay_index::ReplayIndex;
use litchi_core::xml::ReaderOrigin;

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, slide_name: &PackURI) -> Result<List> {
    load_with_limits(
        package,
        slide_name,
        Limits::from_source(package.read_limits()),
    )
}

/// Load one slide's media graph under a caller-selected semantic policy.
pub(crate) fn load_with_limits(
    package: &OpcPackage,
    slide_name: &PackURI,
    limits: Limits,
) -> Result<List> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_media_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source slide-media relationships",
        ));
    }
    let slide = package.get_part(slide_name)?;
    require_slide(slide)?;
    let mut value = parse(slide.blob(), limits)?;
    let conformance = document_conformance_with_limits(slide.blob(), limits)?;
    let mut total = 0usize;
    let mut loaded: BTreeMap<String, Resource> = BTreeMap::new();
    for picture in &mut value.pictures {
        let primary = checked_relationship(
            slide,
            &picture.relationship_id,
            conformance.media_rel(picture.kind),
        )?;
        // Office can retain an external legacy video placeholder (including
        // Target="NULL") alongside the actual embedded p14 media payload.
        // Read the embedded resource without following the legacy target.
        let target = if primary.is_external() {
            let id = picture
                .office_extension
                .as_ref()
                .and_then(|extension| extension.embed_relationship_id.as_deref())
                .ok_or_else(|| {
                    invalid(format!(
                        "external slide-media relationship '{}' is not fetched",
                        picture.relationship_id
                    ))
                })?;
            relationship_target(slide, id, rt::MEDIA)?
        } else {
            primary.target_partname()?
        };
        let resource = load_resource(
            package,
            &target,
            picture.kind,
            false,
            &mut total,
            &mut loaded,
            limits,
        )?;
        picture.resource = Some(resource.clone());
        if let Some(extension) = &picture.office_extension {
            for id in extension
                .embed_relationship_id
                .iter()
                .chain(extension.link_relationship_id.iter())
            {
                let extension_target = relationship_target(slide, id, rt::MEDIA)?;
                if extension_target != target {
                    return Err(invalid(format!(
                        "p14 media relationship '{id}' does not target the ISO media resource"
                    )));
                }
            }
        }
        if let Some(poster) = picture.poster.as_mut() {
            let target =
                relationship_target(slide, &poster.relationship_id, conformance.image_rel())?;
            poster.resource = Some(load_resource(
                package,
                &target,
                picture.kind,
                true,
                &mut total,
                &mut loaded,
                limits,
            )?);
        }
    }
    Ok(value)
}

/// Validate one payload changed by a typed patch against the recipient's
/// selected media policy. Replay already checks the exact before state; this
/// hook parses changed slide XML and bounds newly supplied media payloads
/// without scanning unrelated package parts.
pub(crate) fn validate_replay_part(
    package: &OpcPackage,
    part_name: &PackURI,
    part: &dyn Part,
    limits: Limits,
) -> Result<()> {
    if part.content_type() == ct::PML_SLIDE {
        load_with_limits(package, part_name, limits).map(drop)
    } else if part.blob().len() > limits.max_payload_bytes {
        Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            part.blob().len(),
            limits.max_payload_bytes,
        ))
    } else {
        Ok(())
    }
}

/// Revalidate unchanged slides that reference changed media resources.
///
/// A media-only patch can replace a resource while leaving its owning slide
/// XML byte-identical. In that case [`load_with_limits`] must still see the
/// recipient's complete referenced resource set so its aggregate payload
/// ceiling is enforced. The retained reverse relationship index avoids
/// scanning unrelated slides and skips slides already validated as changed
/// payloads.
pub(crate) fn validate_replay_references(
    package: &OpcPackage,
    index: &ReplayIndex,
    changed_resources: &[PackURI],
    changed_slides: &[PackURI],
    limits: Limits,
) -> Result<()> {
    if changed_resources.is_empty() {
        return Ok(());
    }
    let mut changed_slide_set = HashSet::new();
    changed_slide_set
        .try_reserve(changed_slides.len())
        .map_err(|source| Error::Allocation {
            resource: "media replay changed slide names",
            source,
        })?;
    changed_slide_set.extend(changed_slides.iter().cloned());
    let mut affected_slides = HashSet::new();
    affected_slides
        .try_reserve(changed_resources.len())
        .map_err(|source| Error::Allocation {
            resource: "media replay affected slide names",
            source,
        })?;
    for resource in changed_resources {
        let Some(slides) = index.media_slides(resource) else {
            continue;
        };
        for slide in slides.iter() {
            if changed_slide_set.contains(slide) {
                continue;
            }
            affected_slides
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "media replay affected slide names",
                    source,
                })?;
            affected_slides.insert(slide.clone());
        }
    }
    let mut affected_slides_ordered = Vec::new();
    affected_slides_ordered
        .try_reserve(affected_slides.len())
        .map_err(|source| Error::Allocation {
            resource: "media replay affected slide names",
            source,
        })?;
    affected_slides_ordered.extend(affected_slides);
    let mut affected_slides = affected_slides_ordered;
    affected_slides.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    for slide in affected_slides {
        load_with_limits(package, &slide, limits).map(drop)?;
    }
    Ok(())
}

fn checked_relationship<'a>(
    part: &'a dyn Part,
    id: &str,
    expected_type: &str,
) -> Result<&'a litchi_opc::Relationship> {
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid(format!("missing slide-media relationship '{id}'")))?;
    if relationship.reltype() != expected_type {
        return Err(invalid(format!(
            "relationship '{id}' has type '{}', expected '{expected_type}'",
            relationship.reltype()
        )));
    }
    Ok(relationship)
}

fn relationship_target(part: &dyn Part, id: &str, expected_type: &str) -> Result<PackURI> {
    let relationship = checked_relationship(part, id, expected_type)?;
    if relationship.is_external() {
        return Err(invalid(format!(
            "external slide-media relationship '{id}' is not fetched"
        )));
    }
    relationship.target_partname().map_err(Error::Opc)
}

fn load_resource(
    package: &OpcPackage,
    target: &PackURI,
    kind: Kind,
    image: bool,
    total: &mut usize,
    loaded: &mut BTreeMap<String, Resource>,
    limits: Limits,
) -> Result<Resource> {
    if let Some(value) = loaded.get(target.as_str()) {
        return Ok(value.clone());
    }
    if !target.as_str().starts_with("/ppt/media/") {
        return Err(invalid(format!(
            "slide media resource '{target}' is outside /ppt/media"
        )));
    }
    let part = package.get_part(target)?;
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "slide media resource '{target}' has forbidden outbound relationships"
        )));
    }
    if image {
        if !is_image_content_type(part.content_type()) {
            return Err(invalid(format!(
                "poster '{target}' has non-image content type '{}'",
                part.content_type()
            )));
        }
    } else if !is_media_content_type(part.content_type(), kind) {
        return Err(invalid(format!(
            "media '{target}' has content type '{}' inconsistent with its media kind",
            part.content_type()
        )));
    }
    add_payload(total, part.blob().len(), limits)?;
    let value = Resource {
        part_name: target.to_string(),
        content_type: part.content_type().to_owned(),
        data: Data::from_shared(part.blob_arc()),
    };
    loaded.insert(value.part_name.clone(), value.clone());
    Ok(value)
}

/// Adds a new set of media pictures and their inert internal resources to a Slide part.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
#[cfg(test)]
pub(crate) fn store(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: &List,
    conformance: Conformance,
) -> Result<()> {
    store_with(package, slide_name, value, conformance, false)
}

/// [`store`], or, with `append`, add the pictures after the slide's
/// existing media (record 0321). The caller chooses shape IDs, relationship
/// IDs and part names that are free on the slide; collisions are refused.
#[cfg(test)]
pub(crate) fn store_with(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: &List,
    conformance: Conformance,
    append: bool,
) -> Result<()> {
    store_with_limits(
        package,
        slide_name,
        value,
        conformance,
        append,
        Limits::from_source(package.read_limits()),
    )
}

/// [`store_with`] with an owner-retained semantic policy.
pub(crate) fn store_with_limits(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: &List,
    conformance: Conformance,
    append: bool,
    limits: Limits,
) -> Result<()> {
    validate_value_with_limits(value, true, limits)?;
    let slide = package.get_part(slide_name)?;
    require_slide(slide)?;
    if !append && !parse(slide.blob(), limits)?.pictures.is_empty() {
        return Err(invalid("slide already contains media pictures"));
    }
    if document_conformance_with_limits(slide.blob(), limits)? != conformance {
        return Err(invalid(
            "requested conformance does not match slide namespace",
        ));
    }
    let fragment = write_pictures(value, conformance, limits)?;
    let updated = insert_pictures(slide.blob(), &fragment, conformance, limits)?;
    let mut relationships: BTreeMap<String, (String, String)> = BTreeMap::new();
    let mut parts: BTreeMap<String, Resource> = BTreeMap::new();
    for picture in &value.pictures {
        let resource = picture
            .resource
            .as_ref()
            .ok_or_else(|| invalid("media resource is required for package storage"))?;
        let uri = resource_uri_with_limits(resource, false, Some(picture.kind), limits)?;
        add_part_plan(package, &mut parts, resource)?;
        let target = uri.relative_ref(slide_name.base_uri());
        add_relationship_plan(
            &mut relationships,
            &picture.relationship_id,
            conformance.media_rel(picture.kind),
            &target,
        )?;
        if let Some(extension) = &picture.office_extension {
            for id in extension
                .embed_relationship_id
                .iter()
                .chain(extension.link_relationship_id.iter())
            {
                add_relationship_plan(&mut relationships, id, rt::MEDIA, &target)?;
            }
        }
        if let Some(poster) = &picture.poster {
            let resource = poster
                .resource
                .as_ref()
                .ok_or_else(|| invalid("poster resource is required for package storage"))?;
            let uri = resource_uri_with_limits(resource, true, None, limits)?;
            add_part_plan(package, &mut parts, resource)?;
            add_relationship_plan(
                &mut relationships,
                &poster.relationship_id,
                conformance.image_rel(),
                &uri.relative_ref(slide_name.base_uri()),
            )?;
        }
    }
    for id in relationships.keys() {
        if slide.rels().get(id).is_some() {
            return Err(invalid(format!(
                "slide relationship ID '{id}' already exists"
            )));
        }
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
    for resource in parts.into_values() {
        let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
        package.add_part(Box::new(BlobPart::new_shared(
            uri,
            resource.content_type,
            resource.data.into_shared(),
        )));
    }
    for (id, (relationship_type, target)) in relationships {
        package
            .get_part_mut(slide_name)?
            .rels_mut()
            .add_relationship(relationship_type, target, id, false);
    }
    Ok(())
}

fn insert_pictures(
    xml: &[u8],
    fragment: &[u8],
    conformance: Conformance,
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut sp_tree_depth = None;
    let mut position = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.pml().as_bytes());
                if depth == 0 && (!core || element.local_name().as_ref() != b"sld") {
                    return Err(invalid("slide root does not match conformance"));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        limits.max_depth,
                    )
                })?;
                if depth > limits.max_depth {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        limits.max_depth,
                    ));
                }
                if core
                    && element.local_name().as_ref() == b"spTree"
                    && sp_tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            Event::Empty(element) if element.local_name().as_ref() == b"spTree" => {
                return Err(invalid("cannot insert into an empty shape tree"));
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected slide closing element"));
                }
                if sp_tree_depth == Some(depth) && element.local_name().as_ref() == b"spTree" {
                    position = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("invalid slide XML"));
    }
    let position = position.ok_or_else(|| invalid("slide is missing a shape tree"))?;
    let size = xml.len().checked_add(fragment.len()).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::XmlBytes,
            usize::MAX,
            limits.max_output_bytes,
        )
    })?;
    if size > limits.max_output_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            size,
            limits.max_output_bytes,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "media slide XML",
            source,
        })?;
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

#[cfg(test)]
pub(crate) fn validate_value(value: &List, require_resources: bool) -> Result<()> {
    validate_value_with_limits(value, require_resources, Limits::DEFAULT)
}

pub(crate) fn validate_value_with_limits(
    value: &List,
    require_resources: bool,
    limits: Limits,
) -> Result<()> {
    if value.pictures.len() > limits.max_media {
        return Err(limit(
            crate::error::CodecLimitKind::MediaPictures,
            value.pictures.len(),
            limits.max_media,
        ));
    }
    let mut shape_ids = HashSet::new();
    let mut resources: BTreeMap<String, Resource> = BTreeMap::new();
    let mut payload_allocations = HashSet::new();
    let mut resident_payload = 0usize;
    for picture in &value.pictures {
        if !(1..=67_098_623).contains(&picture.shape_id) {
            return Err(invalid(
                "media shape id is outside Office's supported range",
            ));
        }
        if !shape_ids.insert(picture.shape_id) {
            return Err(invalid(format!(
                "duplicate media shape id {}",
                picture.shape_id
            )));
        }
        bounded_with_limits(&picture.name, limits.max_string_bytes)?;
        validate_id(&picture.relationship_id)?;
        if require_resources && picture.resource.is_none() {
            return Err(invalid("media resource is required for package storage"));
        }
        if let Some(resource) = &picture.resource {
            resource_uri_with_limits(resource, false, Some(picture.kind), limits)?;
            count_payload_allocation(
                &mut payload_allocations,
                &mut resident_payload,
                &resource.data,
                limits,
            )?;
            merge_resource(&mut resources, resource)?;
        }
        if let Some(poster) = &picture.poster {
            validate_id(&poster.relationship_id)?;
            if require_resources && poster.resource.is_none() {
                return Err(invalid("poster resource is required for package storage"));
            }
            if let Some(resource) = &poster.resource {
                resource_uri_with_limits(resource, true, None, limits)?;
                count_payload_allocation(
                    &mut payload_allocations,
                    &mut resident_payload,
                    &resource.data,
                    limits,
                )?;
                merge_resource(&mut resources, resource)?;
            }
        }
        if let Some(extension) = &picture.office_extension {
            validate_extension_with_limits(extension, limits)?;
        }
    }
    let mut total = 0usize;
    for resource in resources.values() {
        add_payload(&mut total, resource.data.len(), limits)?;
    }
    Ok(())
}

fn count_payload_allocation(
    allocations: &mut HashSet<(usize, usize)>,
    total: &mut usize,
    data: &Data,
    limits: Limits,
) -> Result<()> {
    let allocation = (data.as_ptr() as usize, data.len());
    if allocations.insert(allocation) {
        add_payload(total, data.len(), limits)?;
    }
    Ok(())
}

fn validate_extension_with_limits(value: &Extension, limits: Limits) -> Result<()> {
    if value.embed_relationship_id.is_none() && value.link_relationship_id.is_none() {
        return Err(invalid(
            "p14 media extension requires embed or link relationship",
        ));
    }
    for id in value
        .embed_relationship_id
        .iter()
        .chain(value.link_relationship_id.iter())
    {
        validate_id(id)?;
    }
    if value.bookmarks.len() > limits.max_bookmarks {
        return Err(limit(
            crate::error::CodecLimitKind::Bookmarks,
            value.bookmarks.len(),
            limits.max_bookmarks,
        ));
    }
    let mut names = HashSet::new();
    let mut times = HashSet::new();
    for bookmark in &value.bookmarks {
        if let Some(name) = &bookmark.name {
            bounded_with_limits(name, limits.max_string_bytes)?;
            if !names.insert(name) {
                return Err(invalid(format!("duplicate media bookmark name '{name}'")));
            }
        }
        if let Some(time) = &bookmark.time
            && !times.insert(time)
        {
            return Err(invalid(format!("duplicate media bookmark time '{time}'")));
        }
    }
    Ok(())
}

pub(crate) fn parse_time(value: &str) -> Result<Offset> {
    Offset::parse(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> Error {
    invalid(format!("invalid media universal time offset: {error}"))
}

#[cfg(test)]
pub(crate) fn resource_uri(
    resource: &Resource,
    image: bool,
    kind: Option<Kind>,
) -> Result<PackURI> {
    resource_uri_with_limits(resource, image, kind, Limits::DEFAULT)
}

fn resource_uri_with_limits(
    resource: &Resource,
    image: bool,
    kind: Option<Kind>,
    limits: Limits,
) -> Result<PackURI> {
    let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
    if !uri.as_str().starts_with("/ppt/media/") {
        return Err(invalid(format!("resource '{uri}' is outside /ppt/media")));
    }
    if resource.content_type.is_empty()
        || resource
            .content_type
            .bytes()
            .any(|b| b.is_ascii_whitespace())
    {
        return Err(invalid("invalid media resource content type"));
    }
    if image {
        if !is_image_content_type(&resource.content_type) {
            return Err(invalid("poster resource has non-image content type"));
        }
    } else {
        let kind = kind.ok_or_else(|| invalid("non-image media resource requires a media kind"))?;
        if !is_media_content_type(&resource.content_type, kind) {
            return Err(invalid(
                "media resource content type is inconsistent with its kind",
            ));
        }
    }
    if resource.data.len() > limits.max_payload_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            resource.data.len(),
            limits.max_payload_bytes,
        ));
    }
    Ok(uri)
}

fn bounded_with_limits(value: &str, maximum: usize) -> Result<()> {
    if value.len() <= maximum {
        Ok(())
    } else {
        Err(limit(
            crate::error::CodecLimitKind::StringBytes,
            value.len(),
            maximum,
        ))
    }
}

fn merge_resource(resources: &mut BTreeMap<String, Resource>, resource: &Resource) -> Result<()> {
    if let Some(old) = resources.get(&resource.part_name) {
        if old != resource {
            return Err(invalid(format!(
                "conflicting resource part '{}'",
                resource.part_name
            )));
        }
    } else {
        resources.insert(resource.part_name.clone(), resource.clone());
    }
    Ok(())
}
fn add_part_plan(
    package: &OpcPackage,
    parts: &mut BTreeMap<String, Resource>,
    resource: &Resource,
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
    merge_resource(parts, resource)
}
fn add_relationship_plan(
    plans: &mut BTreeMap<String, (String, String)>,
    id: &str,
    kind: &str,
    target: &str,
) -> Result<()> {
    validate_id(id)?;
    let plan = (kind.to_owned(), target.to_owned());
    if let Some(old) = plans.get(id) {
        if old != &plan {
            return Err(invalid(format!("conflicting relationship ID '{id}'")));
        }
    } else {
        plans.insert(id.to_owned(), plan);
    }
    Ok(())
}
fn add_payload(total: &mut usize, size: usize, limits: Limits) -> Result<()> {
    if size > limits.max_payload_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            size,
            limits.max_payload_bytes,
        ));
    }
    *total = total.checked_add(size).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::PayloadBytes,
            usize::MAX,
            limits.max_total_payload_bytes,
        )
    })?;
    if *total > limits.max_total_payload_bytes {
        Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            *total,
            limits.max_total_payload_bytes,
        ))
    } else {
        Ok(())
    }
}
fn is_media_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::AUDIO | rt::VIDEO | rt::MEDIA | STRICT_AUDIO_REL | STRICT_VIDEO_REL
    )
}
fn is_media_content_type(value: &str, kind: Kind) -> bool {
    match kind {
        Kind::Audio => value.starts_with("audio/"),
        Kind::Video => value.starts_with("video/") || value == "application/vnd.ms-asf",
    }
}
fn is_image_content_type(value: &str) -> bool {
    value.starts_with("image/") || matches!(value, "application/x-emf" | "application/x-wmf")
}
fn require_slide(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::PML_SLIDE {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a slide",
            part.partname()
        )))
    }
}
