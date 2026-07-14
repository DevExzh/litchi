//! Wire-preserving construction of empty Numbers sheet clones.

use super::*;

const SHEET_ARCHIVE_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_ARCHIVE_MESSAGE_TYPE: u32 = 3;
const SHEET_NAME_FIELD: u32 = 1;
const SHEET_DRAWABLES_FIELD: u32 = 2;
const FORM_SHEET_PAYLOAD_FIELD: u32 = 1;
const FORM_SHEET_UUID_FIELD: u32 = 2;
const CFUUID_BYTES_FIELD: u32 = 1;
const CFUUID_WORD0_FIELD: u32 = 2;
const CFUUID_WORD1_FIELD: u32 = 3;
const CFUUID_WORD2_FIELD: u32 = 4;
const CFUUID_WORD3_FIELD: u32 = 5;

pub(super) fn clone_empty_sheet_object(
    source: &ArchiveObject,
    message_index: usize,
    new_sheet_id: u64,
    name: &str,
    source_drawables: &[tsp::Reference],
) -> Result<ArchiveObject> {
    let source_sheet_id = source
        .archive_info
        .identifier
        .ok_or_else(|| Error::InvalidFormat("Numbers sheet object has no identifier".to_owned()))?;
    let source_message = source
        .messages
        .get(message_index)
        .ok_or_else(|| Error::InvalidFormat("Numbers sheet message index is invalid".to_owned()))?;
    let previous = source_drawables
        .iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
    let sheet_data =
        duplicate_sheet_wire(&source_message.data, source_message.type_, name, &previous)?;
    let mut sheet_data = Some(sheet_data);
    let messages = source
        .messages
        .iter()
        .enumerate()
        .map(|(index, message)| {
            Ok(RawMessage {
                type_: message.type_,
                data: if index == message_index {
                    sheet_data.take().ok_or_else(|| {
                        Error::InvalidFormat(
                            "Numbers sheet payload was selected more than once".to_owned(),
                        )
                    })?
                } else {
                    message.data.clone()
                },
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let remap = HashMap::from([(source_sheet_id, new_sheet_id)]);
    let mut cloned = clone_numbers_object_metadata(source, new_sheet_id, messages, &remap)?;
    let drawable_set = previous.into_iter().collect::<HashSet<_>>();
    let info = &mut cloned.archive_info.message_infos[message_index];
    info.object_references
        .retain(|identifier| !drawable_set.contains(identifier));
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|identifier| !drawable_set.contains(identifier));
    }
    Ok(cloned)
}

fn duplicate_sheet_wire(
    original: &[u8],
    message_type: u32,
    name: &str,
    previous_drawables: &[u64],
) -> Result<Vec<u8>> {
    let data = match message_type {
        SHEET_ARCHIVE_MESSAGE_TYPE => duplicate_sheet_payload(original, name, previous_drawables)?,
        FORM_BASED_SHEET_ARCHIVE_MESSAGE_TYPE => {
            let mut data =
                transform_length_delimited_field(original, FORM_SHEET_PAYLOAD_FIELD, |payload| {
                    duplicate_sheet_payload(payload, name, previous_drawables)
                })?;
            let form = tn::FormBasedSheetArchive::decode(data.as_slice())?;
            if form.table_id.is_some() {
                let replacement = fresh_uuid();
                data = transform_length_delimited_field(&data, FORM_SHEET_UUID_FIELD, |uuid| {
                    remap_cfuuid_wire(uuid, &replacement)
                })?;
            }
            data
        },
        other => {
            return Err(Error::InvalidFormat(format!(
                "Unsupported Numbers sheet message type {other}"
            )));
        },
    };
    let verified = if message_type == FORM_BASED_SHEET_ARCHIVE_MESSAGE_TYPE {
        tn::FormBasedSheetArchive::decode(data.as_slice())?.super_
    } else {
        tn::SheetArchive::decode(data.as_slice())?
    };
    if verified.name != name || !verified.drawable_infos.is_empty() {
        return Err(Error::InvalidFormat(
            "Numbers empty sheet clone failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn duplicate_sheet_payload(original: &[u8], name: &str, previous: &[u64]) -> Result<Vec<u8>> {
    let data =
        patch_length_delimited_field(original, SHEET_NAME_FIELD, true, Some(name.as_bytes()))?;
    rewrite_reference_list(&data, SHEET_DRAWABLES_FIELD, previous, &[])
}

fn fresh_uuid() -> tsp::Uuid {
    let bytes = litchi_core::id::generate_guid_bytes();
    let mut lower = [0u8; 8];
    lower.copy_from_slice(&bytes[..8]);
    let mut upper = [0u8; 8];
    upper.copy_from_slice(&bytes[8..]);
    tsp::Uuid {
        lower: u64::from_le_bytes(lower),
        upper: u64::from_le_bytes(upper),
    }
}

fn remap_cfuuid_wire(original: &[u8], replacement: &tsp::Uuid) -> Result<Vec<u8>> {
    let decoded = tsp::CfuuidArchive::decode(original)?;
    let mut data = original.to_vec();
    let replacement_bytes =
        ((u128::from(replacement.upper) << 64) | u128::from(replacement.lower)).to_be_bytes();
    if decoded.uuid_bytes.is_some() {
        data = patch_length_delimited_field(
            &data,
            CFUUID_BYTES_FIELD,
            true,
            Some(&replacement_bytes),
        )?;
    }
    for (field, present, value) in [
        (
            CFUUID_WORD0_FIELD,
            decoded.uuid_w0.is_some(),
            replacement.lower as u32,
        ),
        (
            CFUUID_WORD1_FIELD,
            decoded.uuid_w1.is_some(),
            (replacement.lower >> 32) as u32,
        ),
        (
            CFUUID_WORD2_FIELD,
            decoded.uuid_w2.is_some(),
            replacement.upper as u32,
        ),
        (
            CFUUID_WORD3_FIELD,
            decoded.uuid_w3.is_some(),
            (replacement.upper >> 32) as u32,
        ),
    ] {
        if present {
            data = patch_varint_field(&data, field, true, Some(u64::from(value)))?;
        }
    }
    if decoded.uuid_bytes.is_none()
        && decoded.uuid_w0.is_none()
        && decoded.uuid_w1.is_none()
        && decoded.uuid_w2.is_none()
        && decoded.uuid_w3.is_none()
    {
        return Err(Error::InvalidFormat(
            "Numbers form sheet UUID has no encoded representation".to_owned(),
        ));
    }
    let verified = tsp::CfuuidArchive::decode(data.as_slice())?;
    let valid = decoded
        .uuid_bytes
        .as_ref()
        .is_none_or(|_| verified.uuid_bytes.as_deref() == Some(replacement_bytes.as_slice()))
        && decoded
            .uuid_w0
            .is_none_or(|_| verified.uuid_w0 == Some(replacement.lower as u32))
        && decoded
            .uuid_w1
            .is_none_or(|_| verified.uuid_w1 == Some((replacement.lower >> 32) as u32))
        && decoded
            .uuid_w2
            .is_none_or(|_| verified.uuid_w2 == Some(replacement.upper as u32))
        && decoded
            .uuid_w3
            .is_none_or(|_| verified.uuid_w3 == Some((replacement.upper >> 32) as u32));
    if !valid {
        return Err(Error::InvalidFormat(
            "Numbers form sheet UUID remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}
