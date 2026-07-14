//! Lossless mutation of Keynote placeholder ownership and z-order references.

use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;

pub(super) fn validate(
    slide_index: usize,
    slide: &kn::SlideArchive,
    placeholder_id: u64,
    label: &str,
) -> Result<bool> {
    let owned_count = slide
        .owned_drawables
        .iter()
        .filter(|reference| reference.identifier == placeholder_id)
        .count();
    let z_order_count = slide
        .drawables_z_order
        .iter()
        .filter(|reference| reference.identifier == placeholder_id)
        .count();
    if owned_count > 1 || z_order_count > 1 || owned_count != z_order_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide {slide_index} has inconsistent {label} placeholder ownership"
        )));
    }
    Ok(owned_count == 1)
}

pub(super) fn patch(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    reference_field: u32,
    placeholder_id: u64,
    visible: bool,
    label: &str,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive
            .object_mut(slide_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Keynote slide {slide_id} is missing")))?;
        let message_index = unique_message_index(object)?;
        let original = object.messages[message_index].data.as_slice();
        let raw_placeholder = repeated_length_delimited_payloads(original, reference_field)?;
        let [raw_placeholder] = raw_placeholder.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must contain exactly one raw {label} placeholder reference"
            )));
        };
        if tsp::Reference::decode(*raw_placeholder)?.identifier != placeholder_id {
            return Err(Error::InvalidFormat(format!(
                "Keynote {label} placeholder reference changed during update"
            )));
        }

        let mut data = original.to_vec();
        for field in [SLIDE_OWNED_DRAWABLES_FIELD, SLIDE_DRAWABLES_Z_ORDER_FIELD] {
            let matches = repeated_length_delimited_payloads(&data, field)?
                .into_iter()
                .try_fold(0usize, |count, payload| {
                    let identifier = tsp::Reference::decode(payload)?.identifier;
                    Ok::<_, Error>(count + usize::from(identifier == placeholder_id))
                })?;
            if visible {
                if matches != 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} already owns its hidden {label} placeholder"
                    )));
                }
                data = append_repeated_length_delimited_field(&data, field, raw_placeholder)?;
            } else {
                if matches != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} must own its visible {label} placeholder once"
                    )));
                }
                data = remove_repeated_length_delimited_field_where(&data, field, |payload| {
                    Ok(tsp::Reference::decode(payload)?.identifier == placeholder_id)
                })?;
            }
        }
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        let expected_count = usize::from(visible);
        for references in [&verified.owned_drawables, &verified.drawables_z_order] {
            if references
                .iter()
                .filter(|reference| reference.identifier == placeholder_id)
                .count()
                != expected_count
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote {label} placeholder ownership patch failed validation"
                )));
            }
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn unique_message_index(object: &ArchiveObject) -> Result<usize> {
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
        .map(|(index, _)| index);
    let Some(index) = indexes.next() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide has no message type {SLIDE_MESSAGE_TYPE} payload"
        )));
    };
    if indexes.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide repeats message type {SLIDE_MESSAGE_TYPE} payload"
        )));
    }
    Ok(index)
}
