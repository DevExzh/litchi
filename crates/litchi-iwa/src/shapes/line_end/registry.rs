//! Shape-style variation allocation, replacement, and culling.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_external_reference, component_identifier_for_entry,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids,
};
use crate::protobuf::{tsd, tsp, tss, tswp};
use crate::wire::{
    append_repeated_length_delimited_field, parse_wire_fields, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::{
    LineEndpoint, LineEndpoints, SHAPE_INFO_MESSAGE_TYPE, SHAPE_STYLE_MESSAGE_TYPE, reference,
    shape_line_endpoints,
};

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_PROPERTIES_FIELD: u32 = 11;
const TSS_STYLE_PARENT_FIELD: u32 = 3;
const TSS_STYLE_VARIATION_FIELD: u32 = 4;
const TSS_STYLE_STYLESHEET_FIELD: u32 = 5;
const TSD_FILL_FIELD: u32 = 1;
const TSD_STROKE_FIELD: u32 = 2;
const TSD_OPACITY_FIELD: u32 = 3;
const TSD_REFLECTION_FIELD: u32 = 5;
const TSD_HEAD_LINE_END_FIELD: u32 = 6;
const TSD_TAIL_LINE_END_FIELD: u32 = 7;
const SHAPE_INFO_SUPER_FIELD: u32 = 1;
const DRAWABLE_STYLE_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_PARENT_CHILDREN_FIELD: u32 = 5;
const STYLE_CHILDREN_FIELD: u32 = 2;

#[derive(Debug, Clone, Default)]
pub(crate) struct ShapeStyleOverrides {
    pub(crate) fill: Option<tsd::FillArchive>,
    pub(crate) stroke: Option<tsd::StrokeArchive>,
    pub(crate) opacity: Option<f32>,
    pub(crate) reflection: Option<tsd::ReflectionArchive>,
    pub(crate) head_line_end: Option<tsd::LineEndArchive>,
    pub(crate) tail_line_end: Option<tsd::LineEndArchive>,
}

impl ShapeStyleOverrides {
    fn override_count(&self) -> u32 {
        u32::from(self.fill.is_some())
            + u32::from(self.stroke.is_some())
            + u32::from(self.opacity.is_some())
            + u32::from(self.reflection.is_some())
            + u32::from(self.head_line_end.is_some())
            + u32::from(self.tail_line_end.is_some())
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.override_count() == 0
    }
}

/// Return the direct overrides when this is a minimal variation owned by us.
///
/// Exact protobuf-field checks deliberately reject richer native styles so a
/// copy-on-write update can never discard properties that this crate does not
/// understand yet.
pub(crate) fn direct_shape_style_overrides(
    style: &tswp::ShapeStyleArchive,
    raw: &[u8],
) -> Result<Option<ShapeStyleOverrides>> {
    let Some(properties) = style.super_.shape_properties.as_ref() else {
        return Ok(None);
    };
    let overrides = ShapeStyleOverrides {
        fill: properties.fill.clone(),
        stroke: properties.stroke.clone(),
        opacity: properties.opacity,
        reflection: properties.reflection,
        head_line_end: properties.head_line_end.clone(),
        tail_line_end: properties.tail_line_end.clone(),
    };
    let override_count = overrides.override_count();
    let endpoints_are_paired =
        overrides.head_line_end.is_some() == overrides.tail_line_end.is_some();
    let semantic = override_count > 0
        && endpoints_are_paired
        && style.override_count == Some(override_count)
        && style
            .shape_properties
            .as_ref()
            .is_some_and(|properties| properties == &Default::default())
        && style.super_.override_count == Some(override_count)
        && properties.shadow.is_none()
        && style.super_.super_.name.is_none()
        && style.super_.super_.style_identifier.is_none()
        && style.super_.super_.parent.is_some()
        && style.super_.super_.is_variation == Some(true)
        && style.super_.super_.stylesheet.is_some();
    if !semantic {
        return Ok(None);
    }

    let tsd_raw = required_length_delimited_payload(raw, STYLE_SUPER_FIELD, "iWork shape style")?;
    let tss_raw =
        required_length_delimited_payload(tsd_raw, STYLE_SUPER_FIELD, "iWork shape style")?;
    let tsd_properties = required_length_delimited_payload(
        tsd_raw,
        STYLE_PROPERTIES_FIELD,
        "iWork shape style properties",
    )?;
    let tswp_properties = required_length_delimited_payload(
        raw,
        STYLE_PROPERTIES_FIELD,
        "iWork text shape properties",
    )?;
    let mut property_fields = Vec::with_capacity(6);
    if overrides.fill.is_some() {
        property_fields.push(TSD_FILL_FIELD);
    }
    if overrides.stroke.is_some() {
        property_fields.push(TSD_STROKE_FIELD);
    }
    if overrides.opacity.is_some() {
        property_fields.push(TSD_OPACITY_FIELD);
    }
    if overrides.reflection.is_some() {
        property_fields.push(TSD_REFLECTION_FIELD);
    }
    if overrides.head_line_end.is_some() {
        property_fields.push(TSD_HEAD_LINE_END_FIELD);
    }
    if overrides.tail_line_end.is_some() {
        property_fields.push(TSD_TAIL_LINE_END_FIELD);
    }
    let exact = has_exact_fields(
        raw,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_PROPERTIES_FIELD,
        ],
    )? && has_exact_fields(
        tsd_raw,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_PROPERTIES_FIELD,
        ],
    )? && has_exact_fields(
        tss_raw,
        &[
            TSS_STYLE_PARENT_FIELD,
            TSS_STYLE_VARIATION_FIELD,
            TSS_STYLE_STYLESHEET_FIELD,
        ],
    )? && has_exact_fields(tsd_properties, &property_fields)?
        && has_exact_fields(tswp_properties, &[])?;
    Ok(exact.then_some(overrides))
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

fn required_length_delimited_payload<'a>(
    data: &'a [u8],
    field_number: u32,
    context: &str,
) -> Result<&'a [u8]> {
    let payloads = repeated_length_delimited_payloads(data, field_number)?;
    let [payload] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain field {field_number} exactly once"
        )));
    };
    Ok(payload)
}

pub(crate) fn shape_style_is_exclusive(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    let mut shape_count = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in &object.messages {
                match message.type_ {
                    SHAPE_INFO_MESSAGE_TYPE => {
                        let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                        if shape
                            .super_
                            .style
                            .as_ref()
                            .is_some_and(|style| style.identifier == style_id)
                        {
                            shape_count += 1;
                        }
                    },
                    SHAPE_STYLE_MESSAGE_TYPE => {
                        let style = tswp::ShapeStyleArchive::decode(message.data.as_slice())?;
                        if style
                            .super_
                            .super_
                            .parent
                            .as_ref()
                            .is_some_and(|parent| parent.identifier == style_id)
                        {
                            return Ok(false);
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(shape_count == 1)
}

pub(crate) fn endpoint_from_archive(
    endpoint: Option<&tsd::LineEndArchive>,
) -> Result<LineEndpoint> {
    let Some(endpoint) = endpoint else {
        return Ok(LineEndpoint::None);
    };
    match endpoint.identifier.as_deref().unwrap_or("none") {
        "none" => Ok(LineEndpoint::None),
        "simple arrow" => Ok(LineEndpoint::SimpleArrow),
        "filled circle" => Ok(LineEndpoint::FilledCircle),
        "filled diamond" => Ok(LineEndpoint::FilledDiamond),
        "open arrow" => Ok(LineEndpoint::OpenArrow),
        "filled arrow" => Ok(LineEndpoint::FilledArrow),
        "filled square" => Ok(LineEndpoint::FilledSquare),
        "open square" => Ok(LineEndpoint::OpenSquare),
        "open circle" => Ok(LineEndpoint::OpenCircle),
        "inverted arrow" => Ok(LineEndpoint::InvertedArrow),
        "line" => Ok(LineEndpoint::Line),
        identifier => Err(Error::InvalidFormat(format!(
            "unsupported native iWork line endpoint {identifier:?}"
        ))),
    }
}

pub(crate) fn shape_style_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    overrides: ShapeStyleOverrides,
) -> Result<ArchiveObject> {
    let override_count = overrides.override_count();
    if override_count == 0 {
        return Err(Error::InvalidFormat(
            "an iWork shape-style variation must contain at least one override".to_owned(),
        ));
    }
    let data_identifier = overrides
        .fill
        .as_ref()
        .and_then(|fill| fill.image.as_ref())
        .and_then(|image| image.imagedata.as_ref())
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0);
    let data = tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(parent_style_id)),
                is_variation: Some(true),
                stylesheet: Some(reference(stylesheet_id)),
                ..Default::default()
            },
            override_count: Some(override_count),
            shape_properties: Some(tsd::ShapeStylePropertiesArchive {
                fill: overrides.fill,
                stroke: overrides.stroke,
                opacity: overrides.opacity,
                reflection: overrides.reflection,
                head_line_end: overrides.head_line_end,
                tail_line_end: overrides.tail_line_end,
                ..Default::default()
            }),
        },
        override_count: Some(override_count),
        shape_properties: Some(tswp::ShapeStylePropertiesArchive::default()),
    }
    .encode_to_vec();
    tswp::ShapeStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SHAPE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    if let Some(identifier) = data_identifier {
        object.archive_info.message_infos[0]
            .data_references
            .push(identifier);
    }
    Ok(object)
}

pub(crate) fn patch_shape_style_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork drawable {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
            )));
        };
        let original = &object.messages[*index];
        let data =
            transform_length_delimited_field(&original.data, SHAPE_INFO_SUPER_FIELD, |shape| {
                transform_length_delimited_field(shape, DRAWABLE_STYLE_FIELD, |style| {
                    let decoded = tsp::Reference::decode(style)?;
                    if decoded.identifier != old_style_id {
                        return Err(Error::InvalidFormat(format!(
                            "iWork drawable {drawable_id} style reference changed unexpectedly"
                        )));
                    }
                    patch_varint_field(style, REFERENCE_IDENTIFIER_FIELD, true, Some(new_style_id))
                })
            })?;
        object.replace_message(
            *index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        let mut replaced = 0usize;
        for reference in &mut info.object_references {
            if *reference == old_style_id {
                *reference = new_style_id;
                replaced += 1;
            }
        }
        for field in &mut info.field_infos {
            for reference in &mut field.object_references {
                if *reference == old_style_id {
                    *reference = new_style_id;
                }
            }
        }
        if replaced != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} metadata contains {replaced} style references"
            )));
        }
        Ok(())
    })
}

pub(crate) fn insert_style_variation(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    parent_style_id: u64,
    new_style_id: u64,
    new_style: ArchiveObject,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork stylesheet {stylesheet_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet {stylesheet_id} must have exactly one Stylesheet payload"
            )));
        };
        let original = &object.messages[*index];
        let data = append_style_registry_entry(&original.data, parent_style_id, new_style_id)?;
        object.replace_message(
            *index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        if info.object_references.contains(&new_style_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet already references style {new_style_id}"
            )));
        }
        info.object_references.push(new_style_id);
        archive.insert_object(new_style)?;
        Ok(())
    })
}

pub(crate) fn replace_style_variation(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    mut replacement: ArchiveObject,
) -> Result<()> {
    if replacement.messages.len() != 1 {
        return Err(Error::InvalidFormat(
            "replacement iWork line style must contain exactly one payload".to_owned(),
        ));
    }
    let Some(message) = replacement.messages.pop() else {
        return Err(Error::InvalidFormat(
            "replacement iWork line style payload disappeared".to_owned(),
        ));
    };
    let replacement_data_references = replacement
        .archive_info
        .message_infos
        .first()
        .map(|info| info.data_references.clone())
        .ok_or_else(|| {
            Error::InvalidFormat("replacement iWork line style metadata disappeared".to_owned())
        })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork shape style {style_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, candidate)| candidate.type_ == SHAPE_STYLE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork style {style_id} must have exactly one ShapeStyle payload"
            )));
        };
        object.replace_message(*index, message)?;
        object.archive_info.message_infos[*index].data_references = replacement_data_references;
        Ok(())
    })
}

pub(crate) struct ShapeStyleVariationLocation<'a> {
    pub(crate) drawable_archive_name: &'a str,
    pub(crate) style_archive_name: &'a str,
    pub(crate) drawable_id: u64,
    pub(crate) stylesheet_id: u64,
    pub(crate) style_id: u64,
    pub(crate) parent_style_id: u64,
}

pub(crate) fn collapse_line_end_variation(
    package: &mut IWorkPackage,
    location: ShapeStyleVariationLocation<'_>,
) -> Result<()> {
    let drawable_archive_name = location.drawable_archive_name.to_owned();
    let drawable_id = location.drawable_id;
    let mut staged = package.clone();
    collapse_style_variation(&mut staged, location)?;
    if shape_line_endpoints(&staged, &drawable_archive_name, drawable_id)?
        != LineEndpoints::default()
    {
        return Err(Error::InvalidFormat(
            "iWork line endpoint-style reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(crate) fn collapse_style_variation(
    package: &mut IWorkPackage,
    location: ShapeStyleVariationLocation<'_>,
) -> Result<()> {
    let ShapeStyleVariationLocation {
        drawable_archive_name,
        style_archive_name,
        drawable_id,
        stylesheet_id,
        style_id,
        parent_style_id,
    } = location;
    let mut staged = package.clone();
    patch_shape_style_reference(
        &mut staged,
        drawable_archive_name,
        drawable_id,
        style_id,
        parent_style_id,
    )?;
    remove_style_variation(
        &mut staged,
        style_archive_name,
        stylesheet_id,
        parent_style_id,
        style_id,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&staged, style_archive_name)? {
        remove_component_object_uuids(&mut staged, style_component, &[style_id])?;
        remove_component_external_references_to_object(&mut staged, style_component, style_id)?;
        if let Some(drawable_component) =
            component_identifier_for_entry(&staged, drawable_archive_name)?
            && drawable_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                drawable_component,
                style_component,
                parent_style_id,
            )?;
        }
    }
    release_package_identifier_suffix(&mut staged, &[style_id])?;
    *package = staged;
    Ok(())
}

fn remove_style_variation(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    parent_style_id: u64,
    style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let stylesheet = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork stylesheet {stylesheet_id} is missing"))
        })?;
        let indexes = stylesheet
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet {stylesheet_id} must have exactly one Stylesheet payload"
            )));
        };
        let data = remove_style_registry_entry(
            &stylesheet.messages[*index].data,
            parent_style_id,
            style_id,
        )?;
        stylesheet.replace_message(
            *index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut stylesheet.archive_info.message_infos[*index];
        let reference_count = info
            .object_references
            .iter()
            .filter(|&&reference| reference == style_id)
            .count();
        if reference_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet metadata references disposable style {style_id} {reference_count} times"
            )));
        }
        info.object_references
            .retain(|&reference| reference != style_id);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|&reference| reference != style_id);
        }
        archive.remove_object(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("disposable iWork shape style {style_id} is missing"))
        })?;
        Ok(())
    })
}

fn remove_style_registry_entry(
    data: &[u8],
    parent_style_id: u64,
    style_id: u64,
) -> Result<Vec<u8>> {
    let mut styles = repeated_length_delimited_payloads(data, STYLESHEET_STYLES_FIELD)?
        .into_iter()
        .map(|payload| {
            Ok((
                tsp::Reference::decode(payload)?.identifier,
                payload.to_vec(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    if styles.iter().filter(|(id, _)| *id == style_id).count() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet must contain disposable style {style_id} exactly once"
        )));
    }
    styles.retain(|(id, _)| *id != style_id);
    let styles = styles
        .into_iter()
        .map(|(_, payload)| payload)
        .collect::<Vec<_>>();
    let data = rewrite_repeated_length_delimited_fields(data, STYLESHEET_STYLES_FIELD, &styles)?;

    let mut removed_count = 0usize;
    let mut entries = Vec::new();
    for payload in repeated_length_delimited_payloads(&data, STYLESHEET_PARENT_CHILDREN_FIELD)? {
        let entry = tss::stylesheet_archive::StyleChildrenEntry::decode(payload)?;
        let mut children = repeated_length_delimited_payloads(payload, STYLE_CHILDREN_FIELD)?
            .into_iter()
            .map(|child| Ok((tsp::Reference::decode(child)?.identifier, child.to_vec())))
            .collect::<Result<Vec<_>>>()?;
        if children.iter().any(|(id, _)| *id == style_id)
            && entry.parent.identifier != parent_style_id
        {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet maps disposable style {style_id} under the wrong parent"
            )));
        }
        let before = children.len();
        children.retain(|(id, _)| *id != style_id);
        removed_count += before - children.len();
        if !children.is_empty() {
            let children = children
                .into_iter()
                .map(|(_, child)| child)
                .collect::<Vec<_>>();
            entries.push(rewrite_repeated_length_delimited_fields(
                payload,
                STYLE_CHILDREN_FIELD,
                &children,
            )?);
        }
    }
    if removed_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet child map references disposable style {style_id} {removed_count} times"
        )));
    }
    let data = rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_PARENT_CHILDREN_FIELD,
        &entries,
    )?;
    let verified = tss::StylesheetArchive::decode(data.as_slice())?;
    if verified
        .styles
        .iter()
        .any(|style| style.identifier == style_id)
        || verified
            .parent_to_children_style_map
            .iter()
            .flat_map(|entry| &entry.children)
            .any(|style| style.identifier == style_id)
    {
        return Err(Error::InvalidFormat(
            "iWork stylesheet retained a removed line style".to_owned(),
        ));
    }
    Ok(data)
}

fn append_style_registry_entry(
    data: &[u8],
    parent_style_id: u64,
    new_style_id: u64,
) -> Result<Vec<u8>> {
    let current = tss::StylesheetArchive::decode(data)?;
    if current
        .styles
        .iter()
        .any(|style| style.identifier == new_style_id)
        || current
            .parent_to_children_style_map
            .iter()
            .flat_map(|entry| &entry.children)
            .any(|style| style.identifier == new_style_id)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet already contains style {new_style_id}"
        )));
    }
    let mut data = append_repeated_length_delimited_field(
        data,
        STYLESHEET_STYLES_FIELD,
        &reference(new_style_id).encode_to_vec(),
    )?;
    let entries = repeated_length_delimited_payloads(&data, STYLESHEET_PARENT_CHILDREN_FIELD)?;
    let mut parent_matches = 0usize;
    let mut rewritten = Vec::with_capacity(entries.len() + 1);
    for entry in entries {
        let decoded = tss::stylesheet_archive::StyleChildrenEntry::decode(entry)?;
        if decoded.parent.identifier == parent_style_id {
            parent_matches += 1;
            rewritten.push(append_repeated_length_delimited_field(
                entry,
                STYLE_CHILDREN_FIELD,
                &reference(new_style_id).encode_to_vec(),
            )?);
        } else {
            rewritten.push(entry.to_vec());
        }
    }
    if parent_matches > 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet repeats parent style {parent_style_id}"
        )));
    }
    if parent_matches == 0 {
        rewritten.push(
            tss::stylesheet_archive::StyleChildrenEntry {
                parent: reference(parent_style_id),
                children: vec![reference(new_style_id)],
            }
            .encode_to_vec(),
        );
    }
    data = rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_PARENT_CHILDREN_FIELD,
        &rewritten,
    )?;
    let verified = tss::StylesheetArchive::decode(data.as_slice())?;
    if verified
        .styles
        .iter()
        .filter(|style| style.identifier == new_style_id)
        .count()
        != 1
        || verified
            .parent_to_children_style_map
            .iter()
            .filter(|entry| entry.parent.identifier == parent_style_id)
            .flat_map(|entry| &entry.children)
            .filter(|style| style.identifier == new_style_id)
            .count()
            != 1
    {
        return Err(Error::InvalidFormat(
            "iWork stylesheet insertion failed validation".to_owned(),
        ));
    }
    Ok(data)
}
