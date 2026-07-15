//! Wire-preserving slide ordering for layout-owned media.

use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;

#[allow(clippy::too_many_arguments)]
pub(super) fn rewrite_slide_media_roots(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    remove: &[u64],
    target_roots: &[u64],
    add: &[u64],
    current: &kn::SlideArchive,
    target: &kn::SlideArchive,
) -> Result<()> {
    let replacements = target_roots
        .iter()
        .copied()
        .zip(add.iter().copied())
        .collect::<HashMap<_, _>>();
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(
                "Keynote slide must contain exactly one SlideArchive payload".to_owned(),
            ));
        };
        let index = *index;
        let original = object.messages[index].data.as_slice();
        let decoded = kn::SlideArchive::decode(original)?;
        let mut data = original.to_vec();
        let mut expected = decoded.clone();
        for (field, references) in [
            (SLIDE_OWNED_DRAWABLES_FIELD, &mut expected.owned_drawables),
            (
                SLIDE_DRAWABLES_Z_ORDER_FIELD,
                &mut expected.drawables_z_order,
            ),
        ] {
            let raw = repeated_length_delimited_payloads(&data, field)?;
            if raw.len() != references.len() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide field {field} reference count does not match its wire payload"
                )));
            }
            let mut preserved = HashMap::with_capacity(raw.len());
            for (reference, payload) in references.iter().zip(raw) {
                if preserved.insert(reference.identifier, payload).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide field {field} repeats drawable {}",
                        reference.identifier
                    )));
                }
            }
            let mut identifiers = references
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            for identifier in remove {
                if identifiers
                    .iter()
                    .filter(|candidate| *candidate == identifier)
                    .count()
                    != 1
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide field {field} must contain layout media object {identifier} exactly once"
                    )));
                }
                identifiers.retain(|candidate| candidate != identifier);
            }
            insert_target_media(&mut identifiers, target, current, &replacements)?;
            let payloads = identifiers
                .iter()
                .map(|identifier| {
                    preserved.get(identifier).map_or_else(
                        || {
                            tsp::Reference {
                                identifier: *identifier,
                                ..Default::default()
                            }
                            .encode_to_vec()
                        },
                        |payload| payload.to_vec(),
                    )
                })
                .collect::<Vec<_>>();
            data = rewrite_repeated_length_delimited_fields(&data, field, &payloads)?;
            *references = identifiers
                .into_iter()
                .map(|identifier| tsp::Reference {
                    identifier,
                    ..Default::default()
                })
                .collect();
        }
        if kn::SlideArchive::decode(data.as_slice())? != expected {
            return Err(Error::InvalidFormat(
                "Keynote slide layout-media ordering failed validation".to_owned(),
            ));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[index];
        info.object_references
            .retain(|identifier| !remove.contains(identifier));
        for identifier in add {
            if !info.object_references.contains(identifier) {
                info.object_references.push(*identifier);
            }
        }
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| !remove.contains(identifier));
        }
        Ok(())
    })
}

fn insert_target_media(
    identifiers: &mut Vec<u64>,
    target: &kn::SlideArchive,
    current: &kn::SlideArchive,
    replacements: &HashMap<u64, u64>,
) -> Result<()> {
    let anchor = |target_id: u64| {
        if target
            .title_placeholder
            .as_ref()
            .map(|item| item.identifier)
            == Some(target_id)
        {
            current
                .title_placeholder
                .as_ref()
                .map(|item| item.identifier)
        } else if target.body_placeholder.as_ref().map(|item| item.identifier) == Some(target_id) {
            current
                .body_placeholder
                .as_ref()
                .map(|item| item.identifier)
        } else {
            None
        }
    };
    let has_anchor = target
        .owned_drawables
        .iter()
        .any(|item| anchor(item.identifier).is_some());
    let mut leading = 0usize;
    for (index, item) in target.owned_drawables.iter().enumerate() {
        let Some(new_identifier) = replacements.get(&item.identifier).copied() else {
            continue;
        };
        if identifiers.contains(&new_identifier) {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide already owns materialized media object {new_identifier}"
            )));
        }
        let next_anchor = target.owned_drawables[index + 1..]
            .iter()
            .find_map(|candidate| anchor(candidate.identifier))
            .filter(|candidate| identifiers.contains(candidate));
        if let Some(next_anchor) = next_anchor {
            let position = identifiers
                .iter()
                .position(|identifier| *identifier == next_anchor)
                .ok_or_else(|| {
                    Error::InvalidFormat("Keynote layout media anchor disappeared".to_owned())
                })?;
            identifiers.insert(position, new_identifier);
        } else if has_anchor {
            identifiers.push(new_identifier);
        } else {
            identifiers.insert(leading, new_identifier);
            leading += 1;
        }
    }
    Ok(())
}
