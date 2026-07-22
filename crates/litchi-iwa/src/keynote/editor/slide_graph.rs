//! Slide object-graph traversal, cloning, remapping, and deletion.

use super::*;

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MASK_MESSAGE_TYPE: u32 = 3_006;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STORAGELESS_PLACEHOLDER_STORAGE_ID: u64 = 0;

pub(super) struct ObjectGraph {
    pub(super) objects: HashMap<u64, Vec<RawMessage>>,
    pub(super) archives: HashMap<u64, String>,
}

impl ObjectGraph {
    pub(super) fn read(package: &IWorkPackage) -> Result<Self> {
        let mut objects = HashMap::new();
        let mut archives = HashMap::new();
        for name in package.iwa_entry_names() {
            for object in package.archive(name)?.objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                if objects.insert(identifier, object.messages).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Duplicate iWork object identifier {identifier}"
                    )));
                }
                archives.insert(identifier, name.to_owned());
            }
        }
        Ok(Self { objects, archives })
    }

    pub(super) fn decode<T: Message + Default>(
        &self,
        identifier: u64,
        type_name: &str,
    ) -> Result<T> {
        self.objects
            .get(&identifier)
            .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))?
            .iter()
            .find_map(|message| T::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {identifier} has no decodable {type_name} payload"
                ))
            })
    }

    pub(super) fn decode_type<T: Message + Default>(
        &self,
        identifier: u64,
        message_type: u32,
        type_name: &str,
    ) -> Result<T> {
        let messages = self
            .objects
            .get(&identifier)
            .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))?;
        let matches = messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Object {identifier} must contain exactly one {type_name} payload"
            )));
        }
        T::decode(matches[0].data.as_slice()).map_err(Into::into)
    }

    pub(super) fn message_data_type(
        &self,
        identifier: u64,
        message_type: u32,
        type_name: &str,
    ) -> Result<&[u8]> {
        let messages = self
            .objects
            .get(&identifier)
            .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))?;
        let mut matches = messages
            .iter()
            .filter(|message| message.type_ == message_type);
        let message = matches.next().ok_or_else(|| {
            Error::InvalidFormat(format!("Object {identifier} has no {type_name} payload"))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Object {identifier} repeats its {type_name} payload"
            )));
        }
        Ok(message.data.as_slice())
    }

    pub(super) fn drawable_storage(&self, identifier: u64) -> Result<Option<u64>> {
        let messages = self
            .objects
            .get(&identifier)
            .ok_or_else(|| Error::InvalidFormat(format!("Drawable {identifier} is missing")))?;
        Ok(messages.iter().find_map(|message| {
            kn::PlaceholderArchive::decode(message.data.as_slice())
                .ok()
                .and_then(|placeholder| placeholder.super_.owned_storage)
                .or_else(|| {
                    tswp::ShapeInfoArchive::decode(message.data.as_slice())
                        .ok()
                        .and_then(|shape| shape.owned_storage)
                })
                .map(|reference| reference.identifier)
                .filter(|identifier| *identifier != STORAGELESS_PLACEHOLDER_STORAGE_ID)
        }))
    }

    pub(super) fn storage_text(&self, identifier: u64) -> Result<String> {
        let storage: tswp::StorageArchive = self.decode(identifier, "TSWP.StorageArchive")?;
        Ok(storage.text.concat())
    }

    pub(super) fn archive_name(&self, identifier: u64) -> Result<&str> {
        self.archives
            .get(&identifier)
            .map(String::as_str)
            .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))
    }
}

pub(super) fn offset_keynote_drawable_clone(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    offset: f32,
) -> Result<()> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Keynote drawable duplicate offset must be finite".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote drawable {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote drawable {drawable_id} must have exactly one shape payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let shape = tswp::ShapeInfoArchive::decode(original)?;
        let position = shape
            .super_
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.position.as_ref())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote drawable {drawable_id} has no positioned geometry"
                ))
            })?;
        let x = position.x + offset;
        let y = position.y + offset;
        if !x.is_finite() || !y.is_finite() {
            return Err(Error::ParseError(
                "Keynote drawable duplicate position overflow".to_owned(),
            ));
        }
        let data = patch_nested_fixed32_field(original, &[1, 1, 1, 1, 1], true, Some(x.to_bits()))?;
        let data = patch_nested_fixed32_field(&data, &[1, 1, 1, 1, 2], true, Some(y.to_bits()))?;
        let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
        let verified_position = verified
            .super_
            .super_
            .geometry
            .and_then(|geometry| geometry.position)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote drawable offset removed its position".to_owned())
            })?;
        if verified_position.x != x || verified_position.y != y {
            return Err(Error::InvalidFormat(
                "Keynote drawable offset failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn patch_slide_drawable_references(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    remove: Option<u64>,
    add: Option<u64>,
) -> Result<()> {
    if remove.is_some() == add.is_some() {
        return Err(Error::InvalidFormat(
            "Keynote drawable ownership patch must add or remove exactly one object".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 5)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let slide = kn::SlideArchive::decode(original)?;
        let existing_drawables = slide
            .owned_drawables
            .iter()
            .chain(&slide.drawables_z_order)
            .map(|reference| reference.identifier)
            .collect::<HashSet<_>>();
        let mut data = original.to_vec();
        for field_number in [7, 42] {
            let raw = repeated_length_delimited_payloads(&data, field_number)?;
            let identifiers = raw
                .iter()
                .map(|reference| {
                    Ok(crate::protobuf::tsp::Reference::decode(*reference)?.identifier)
                })
                .collect::<Result<Vec<_>>>()?;
            if let Some(identifier) = remove {
                if identifiers
                    .iter()
                    .filter(|&&candidate| candidate == identifier)
                    .count()
                    != 1
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} field {field_number} must contain drawable {identifier} exactly once"
                    )));
                }
                data = remove_repeated_length_delimited_field_where(
                    &data,
                    field_number,
                    |reference| {
                        Ok(
                            crate::protobuf::tsp::Reference::decode(reference)?.identifier
                                == identifier,
                        )
                    },
                )?;
            } else if let Some(identifier) = add {
                if identifiers.contains(&identifier) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} already owns drawable {identifier} in field {field_number}"
                    )));
                }
                data = append_repeated_length_delimited_field(
                    &data,
                    field_number,
                    &crate::protobuf::tsp::Reference {
                        identifier,
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?;
            }
        }
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        for references in [&verified.owned_drawables, &verified.drawables_z_order] {
            if let Some(identifier) = remove
                && references
                    .iter()
                    .any(|reference| reference.identifier == identifier)
            {
                return Err(Error::InvalidFormat(
                    "Keynote drawable ownership removal failed validation".to_owned(),
                ));
            }
            if let Some(identifier) = add
                && references
                    .iter()
                    .filter(|reference| reference.identifier == identifier)
                    .count()
                    != 1
            {
                return Err(Error::InvalidFormat(
                    "Keynote drawable ownership insertion failed validation".to_owned(),
                ));
            }
        }
        object.replace_message(message_index, RawMessage { type_: 5, data })?;
        let info = &mut object.archive_info.message_infos[message_index];
        if let Some(identifier) = remove {
            info.object_references
                .retain(|candidate| *candidate != identifier);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|candidate| *candidate != identifier);
            }
        } else if let Some(identifier) = add {
            if !info.object_references.contains(&identifier) {
                info.object_references.push(identifier);
            }
            for field in &mut info.field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|candidate| existing_drawables.contains(candidate))
                    && !field.object_references.contains(&identifier)
                {
                    field.object_references.push(identifier);
                }
            }
        }
        Ok(())
    })
}

pub(super) fn package_references_object(package: &IWorkPackage, identifier: u64) -> Result<bool> {
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&identifier)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            }) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

pub(super) fn rewrite_show_slide_references(
    data: &[u8],
    previous: &[crate::protobuf::tsp::Reference],
    desired: &[u64],
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_field(data, 3, |slide_tree| {
        let raw_references = repeated_length_delimited_payloads(slide_tree, 2)?;
        if raw_references.len() != previous.len() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide tree has {} raw references but {} decoded references",
                raw_references.len(),
                previous.len()
            )));
        }

        let mut by_identifier = HashMap::with_capacity(previous.len());
        for (expected, raw) in previous.iter().zip(raw_references) {
            let decoded = crate::protobuf::tsp::Reference::decode(raw)?;
            if decoded.identifier != expected.identifier {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide-tree reference decoded as {} instead of {}",
                    decoded.identifier, expected.identifier
                )));
            }
            if by_identifier.insert(expected.identifier, raw).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide tree contains duplicate node {}",
                    expected.identifier
                )));
            }
        }

        let mut seen = std::collections::HashSet::with_capacity(desired.len());
        let replacements = desired
            .iter()
            .map(|identifier| {
                if !seen.insert(*identifier) {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide tree would contain duplicate node {identifier}"
                    )));
                }
                Ok(by_identifier.get(identifier).map_or_else(
                    || {
                        crate::protobuf::tsp::Reference {
                            identifier: *identifier,
                            ..Default::default()
                        }
                        .encode_to_vec()
                    },
                    |raw| raw.to_vec(),
                ))
            })
            .collect::<Result<Vec<_>>>()?;
        rewrite_repeated_length_delimited_fields(slide_tree, 2, &replacements)
    })?;

    let verified = kn::ShowArchive::decode(data.as_slice())?;
    if !verified
        .slide_tree
        .slides
        .iter()
        .map(|reference| reference.identifier)
        .eq(desired.iter().copied())
    {
        return Err(Error::InvalidFormat(
            "Keynote slide-tree wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_reference_paths(
    data: &[u8],
    paths: &[&[u32]],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    paths.iter().try_fold(data.to_vec(), |data, path| {
        transform_length_delimited_fields_at_path(&data, path, |reference| {
            let decoded = crate::protobuf::tsp::Reference::decode(reference)?;
            let Some(identifier) = remap.get(&decoded.identifier).copied() else {
                return Ok(reference.to_vec());
            };
            let data = patch_varint_field(reference, 1, true, Some(identifier))?;
            let verified = crate::protobuf::tsp::Reference::decode(data.as_slice())?;
            if verified.identifier != identifier {
                return Err(Error::InvalidFormat(
                    "Keynote reference wire remap failed validation".to_owned(),
                ));
            }
            Ok(data)
        })
    })
}

#[allow(deprecated)]
pub(super) fn remap_slide_archive_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1],
        &[2],
        &[43],
        &[5],
        &[6],
        &[30],
        &[20],
        &[7],
        &[42],
        &[28, 2],
        &[45, 1, 1],
        &[29],
        &[31],
        &[35],
        &[17],
        &[36],
        &[27],
        &[44],
        &[39],
        &[3, 1],
    ];
    let mut expected = kn::SlideArchive::decode(data)?;
    remap_slide_archive(&mut expected, remap);
    let data = remap_reference_paths(data, REFERENCE_PATHS, remap)?;
    if kn::SlideArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote SlideArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

#[allow(deprecated)]
pub(super) fn remap_shape_info_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 1, 2],
        &[1, 1, 6],
        &[1, 1, 9],
        &[1, 1, 10],
        &[1, 1, 11],
        &[1, 2],
        &[2],
        &[3],
        &[4],
    ];
    let mut expected = tswp::ShapeInfoArchive::decode(data)?;
    remap_shape_info(&mut expected, remap);
    let data = remap_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tswp::ShapeInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote ShapeInfoArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_placeholder_archive_wire(
    data: &[u8],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = kn::PlaceholderArchive::decode(data)?;
    remap_shape_info(&mut expected.super_, remap);
    let data = transform_length_delimited_fields_at_path(data, &[1], |shape| {
        remap_shape_info_wire(shape, remap)
    })?;
    if kn::PlaceholderArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote PlaceholderArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_drawable_archive(drawable: &mut tsd::DrawableArchive, remap: &HashMap<u64, u64>) {
    remap_optional_reference(&mut drawable.parent, remap);
    remap_optional_reference(&mut drawable.comment, remap);
    remap_references(&mut drawable.pencil_annotations, remap);
    remap_optional_reference(&mut drawable.title, remap);
    remap_optional_reference(&mut drawable.caption, remap);
}

pub(super) fn remap_image_archive_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    let mut expected = tsd::ImageArchive::decode(data)?;
    remap_drawable_archive(&mut expected.super_, remap);
    remap_optional_reference(&mut expected.database_data, remap);
    remap_optional_reference(&mut expected.style, remap);
    remap_optional_reference(&mut expected.mask, remap);
    remap_optional_reference(&mut expected.database_thumbnail_data, remap);
    remap_optional_reference(&mut expected.database_original_data, remap);
    let data = remap_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::ImageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote ImageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_mask_archive_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[&[1, 2], &[1, 6], &[1, 9], &[1, 10], &[1, 11]];
    let mut expected = tsd::MaskArchive::decode(data)?;
    remap_drawable_archive(&mut expected.super_, remap);
    let data = remap_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::MaskArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote MaskArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_movie_archive_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    let mut expected = tsd::MovieArchive::decode(data)?;
    remap_drawable_archive(&mut expected.super_, remap);
    remap_optional_reference(&mut expected.database_movie_data, remap);
    remap_optional_reference(&mut expected.database_poster_image_data, remap);
    remap_optional_reference(&mut expected.database_audio_only_image_data, remap);
    remap_optional_reference(&mut expected.style, remap);
    let data = remap_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::MovieArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote MovieArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_chart_archive_wire(
    data: &[u8],
    recorded_references: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = crate::charts::IWorkChartArchive::decode(data)?;
    expected.remap_references(remap, recorded_references)?;
    let data = expected.encode()?;
    if crate::charts::IWorkChartArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_note_archive_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    let mut expected = kn::NoteArchive::decode(data)?;
    remap_reference(&mut expected.contained_storage, remap);
    let data = remap_reference_paths(data, &[&[1]], remap)?;
    if kn::NoteArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote NoteArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_storage_archive_wire(
    data: &[u8],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    const OBJECT_TABLE_FIELDS: &[u32] = &[5, 7, 8, 9, 11, 12, 15, 16, 17, 18, 21, 22, 23, 27, 28];
    let mut expected = tswp::StorageArchive::decode(data)?;
    remap_storage_archive(&mut expected, remap);
    let mut data = remap_reference_paths(data, &[&[2]], remap)?;
    for field in OBJECT_TABLE_FIELDS {
        data = remap_reference_paths(&data, &[&[*field, 1, 2]], remap)?;
    }
    for field in [25, 26] {
        data = remap_reference_paths(&data, &[&[field, 1, 2]], remap)?;
    }
    if tswp::StorageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote StorageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn clone_slide_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let old_identifier = source
        .archive_info
        .identifier
        .ok_or_else(|| Error::InvalidFormat("Keynote slide object has no identifier".to_owned()))?;
    let new_identifier = *remap.get(&old_identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "No clone identifier allocated for object {old_identifier}"
        ))
    })?;
    let mut messages = Vec::with_capacity(source.messages.len());
    for (message, info) in source
        .messages
        .iter()
        .zip(&source.archive_info.message_infos)
    {
        let data = match message.type_ {
            crate::charts::source::CHART_MESSAGE_TYPE => {
                remap_chart_archive_wire(&message.data, &info.object_references, remap)?
            },
            crate::charts::source::CHART_PRESET_MESSAGE_TYPE => {
                crate::charts::source::remap_chart_preset_wire(
                    &message.data,
                    &info.object_references,
                    remap,
                )?
            },
            5 => remap_slide_archive_wire(&message.data, remap)?,
            7 => remap_placeholder_archive_wire(&message.data, remap)?,
            15 => remap_note_archive_wire(&message.data, remap)?,
            IMAGE_MESSAGE_TYPE => remap_image_archive_wire(&message.data, remap)?,
            MASK_MESSAGE_TYPE => remap_mask_archive_wire(&message.data, remap)?,
            MOVIE_MESSAGE_TYPE => remap_movie_archive_wire(&message.data, remap)?,
            2001 | 2022 => remap_storage_archive_wire(&message.data, remap)?,
            2011 => remap_shape_info_wire(&message.data, remap)?,
            _ => {
                if info
                    .object_references
                    .iter()
                    .any(|identifier| remap.contains_key(identifier))
                {
                    return Err(Error::InvalidFormat(format!(
                        "Cannot safely clone Keynote message type {} with internal references",
                        message.type_
                    )));
                }
                message.data.clone()
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    clone_object_metadata(source, new_identifier, messages, remap, false)
}

pub(super) fn clone_object_metadata(
    source: &ArchiveObject,
    new_identifier: u64,
    messages: Vec<RawMessage>,
    remap: &HashMap<u64, u64>,
    clear_data_references: bool,
) -> Result<ArchiveObject> {
    let mut cloned = ArchiveObject::new(new_identifier, messages)?;
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    for ((target, source), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&source.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_owned())
        })?;
        *target = source.clone();
        target.length = length;
        target.object_references = source
            .object_references
            .iter()
            .map(|identifier| remap.get(identifier).copied().unwrap_or(*identifier))
            .collect();
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
            if clear_data_references {
                field.data_references.clear();
            }
        }
        if clear_data_references {
            target.data_references.clear();
        }
    }
    Ok(cloned)
}

#[allow(deprecated)]
pub(super) fn clone_slide_node(
    source: &ArchiveObject,
    new_node_id: u64,
    old_slide_id: u64,
    new_slide_id: u64,
) -> Result<ArchiveObject> {
    let message_index = source
        .messages
        .iter()
        .position(|message| kn::SlideNodeArchive::decode(message.data.as_slice()).is_ok())
        .ok_or_else(|| Error::InvalidFormat("Keynote slide node payload is missing".to_owned()))?;
    let mut node = kn::SlideNodeArchive::decode(source.messages[message_index].data.as_slice())?;
    let mut removed_references = node
        .children
        .iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
    removed_references.push(old_slide_id);
    if let Some(reference) = &node.database_thumbnail {
        removed_references.push(reference.identifier);
    }
    removed_references.extend(
        node.database_thumbnails
            .iter()
            .map(|reference| reference.identifier),
    );
    node.children.clear();
    node.slide = Some(crate::protobuf::tsp::Reference {
        identifier: new_slide_id,
        ..Default::default()
    });
    node.thumbnails.clear();
    node.thumbnail_sizes.clear();
    node.thumbnails_are_dirty = Some(true);
    node.digests_for_datas_needing_download_for_thumbnail
        .clear();
    node.database_thumbnail = None;
    node.database_thumbnails.clear();
    node.unique_identifier = None;
    node.copy_from_slide_identifier = None;
    node.slide_specific_hyperlink_map.clear();
    node.slide_specific_hyperlink_count = Some(0);

    let original = source.messages[message_index].data.as_slice();
    let original_node = kn::SlideNodeArchive::decode(original)?;
    let remap = HashMap::from([(old_slide_id, new_slide_id)]);
    let mut data = original.to_vec();
    for field in [1, 9, 10, 16, 24, 25] {
        data = rewrite_repeated_length_delimited_fields(&data, field, &[])?;
    }
    data = remap_reference_paths(&data, &[&[2]], &remap)?;
    data = patch_varint_field(
        &data,
        14,
        original_node.thumbnails_are_dirty.is_some(),
        Some(1),
    )?;
    data =
        patch_length_delimited_field(&data, 3, original_node.database_thumbnail.is_some(), None)?;
    data =
        patch_length_delimited_field(&data, 11, original_node.unique_identifier.is_some(), None)?;
    data = patch_length_delimited_field(
        &data,
        12,
        original_node.copy_from_slide_identifier.is_some(),
        None,
    )?;
    data = patch_varint_field(
        &data,
        13,
        original_node.slide_specific_hyperlink_count.is_some(),
        Some(0),
    )?;
    if kn::SlideNodeArchive::decode(data.as_slice())? != node {
        return Err(Error::InvalidFormat(
            "Keynote SlideNodeArchive wire clone failed validation".to_owned(),
        ));
    }

    let mut messages = source.messages.clone();
    messages[message_index].data = data;
    let mut cloned = clone_object_metadata(source, new_node_id, messages, &remap, true)?;
    let references = &mut cloned.archive_info.message_infos[message_index].object_references;
    references.retain(|identifier| !removed_references.contains(identifier));
    if !references.contains(&new_slide_id) {
        references.push(new_slide_id);
    }
    let retained = references
        .iter()
        .copied()
        .collect::<std::collections::HashSet<_>>();
    for field in &mut cloned.archive_info.message_infos[message_index].field_infos {
        field
            .object_references
            .retain(|identifier| retained.contains(identifier));
    }
    Ok(cloned)
}

pub(super) fn remap_reference(
    reference: &mut crate::protobuf::tsp::Reference,
    remap: &HashMap<u64, u64>,
) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

pub(super) fn remap_optional_reference(
    reference: &mut Option<crate::protobuf::tsp::Reference>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(reference) = reference {
        remap_reference(reference, remap);
    }
}

pub(super) fn remap_references(
    references: &mut [crate::protobuf::tsp::Reference],
    remap: &HashMap<u64, u64>,
) {
    for reference in references {
        remap_reference(reference, remap);
    }
}

#[allow(deprecated)]
pub(super) fn remap_slide_archive(slide: &mut kn::SlideArchive, remap: &HashMap<u64, u64>) {
    remap_reference(&mut slide.style, remap);
    remap_references(&mut slide.builds, remap);
    for chunk in &mut slide.build_chunk_archives {
        remap_optional_reference(&mut chunk.build, remap);
    }
    remap_references(&mut slide.build_chunks, remap);
    remap_optional_reference(&mut slide.title_placeholder, remap);
    remap_optional_reference(&mut slide.body_placeholder, remap);
    remap_optional_reference(&mut slide.object_placeholder, remap);
    remap_optional_reference(&mut slide.slide_number_placeholder, remap);
    remap_references(&mut slide.owned_drawables, remap);
    remap_references(&mut slide.drawables_z_order, remap);
    for entry in &mut slide.sage_tag_to_info_map {
        remap_reference(&mut entry.info, remap);
    }
    if let Some(map) = &mut slide.instructional_text_map {
        for entry in &mut map.instructional_text_for_infos {
            remap_optional_reference(&mut entry.info, remap);
        }
    }
    remap_optional_reference(&mut slide.classic_stylesheet_record, remap);
    remap_references(&mut slide.body_paragraph_styles, remap);
    remap_references(&mut slide.body_list_styles, remap);
    remap_optional_reference(&mut slide.template_slide, remap);
    remap_optional_reference(&mut slide.user_defined_guide_storage, remap);
    remap_optional_reference(&mut slide.note, remap);
    remap_references(&mut slide.infos_using_object_placeholder_geometry, remap);
    remap_optional_reference(&mut slide.info_using_object_placeholder_geometry, remap);
}

#[allow(deprecated)]
pub(super) fn remap_shape_info(shape: &mut tswp::ShapeInfoArchive, remap: &HashMap<u64, u64>) {
    let drawable = &mut shape.super_.super_;
    remap_optional_reference(&mut drawable.parent, remap);
    remap_optional_reference(&mut drawable.comment, remap);
    remap_references(&mut drawable.pencil_annotations, remap);
    remap_optional_reference(&mut drawable.title, remap);
    remap_optional_reference(&mut drawable.caption, remap);
    remap_optional_reference(&mut shape.super_.style, remap);
    remap_optional_reference(&mut shape.deprecated_storage, remap);
    remap_optional_reference(&mut shape.text_flow, remap);
    remap_optional_reference(&mut shape.owned_storage, remap);
}

pub(super) fn remap_storage_archive(storage: &mut tswp::StorageArchive, remap: &HashMap<u64, u64>) {
    remap_optional_reference(&mut storage.style_sheet, remap);
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_attachment,
        &mut storage.table_smartfield,
        &mut storage.table_layout_style,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_section,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_highlight,
        &mut storage.table_tatechuyoko,
        &mut storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_optional_reference(&mut entry.object, remap);
        }
    }
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_reference(&mut entry.field, remap);
        }
    }
}

pub(super) fn remove_object(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        archive.remove_object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {identifier} is missing from {archive_name}"
            ))
        })?;
        Ok(())
    })
}
