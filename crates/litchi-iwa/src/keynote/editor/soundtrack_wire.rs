//! Strict native object lookup and wire patching for Keynote soundtracks.

use super::*;
use crate::archive::FieldType;

const DOCUMENT_OBJECT_ID: u64 = 1;
const DOCUMENT_ARCHIVE_MESSAGE_TYPE: u32 = 1;
const SHOW_ARCHIVE_MESSAGE_TYPE: u32 = 2;
const SHOW_SOUNDTRACK_FIELD: u32 = 17;
const SOUNDTRACK_MESSAGE_TYPE: u32 = 21;
const VOLUME_FIELD: u32 = 1;
const MODE_FIELD: u32 = 2;
const MEDIA_FIELD: u32 = 3;
const FIXED64_WIRE_TYPE: u8 = 1;
const LENGTH_DELIMITED_WIRE_TYPE: u8 = 2;
const VARINT_WIRE_TYPE: u8 = 0;

pub(super) struct SoundtrackRecord<'a> {
    pub(super) id: u64,
    pub(super) data: &'a [u8],
}

pub(super) fn read_soundtrack<'a>(graph: &'a ObjectGraph) -> Result<Option<SoundtrackRecord<'a>>> {
    let document: kn::DocumentArchive = graph.decode_type(
        DOCUMENT_OBJECT_ID,
        DOCUMENT_ARCHIVE_MESSAGE_TYPE,
        "KN.DocumentArchive",
    )?;
    let show_data = graph.message_data_type(
        document.show.identifier,
        SHOW_ARCHIVE_MESSAGE_TYPE,
        "KN.ShowArchive",
    )?;
    let show = kn::ShowArchive::decode(show_data)?;
    validate_show_soundtrack_wire(show_data, show.soundtrack.is_some())?;
    let Some(reference) = show.soundtrack else {
        return Ok(None);
    };
    let data = graph.message_data_type(
        reference.identifier,
        SOUNDTRACK_MESSAGE_TYPE,
        "KN.Soundtrack",
    )?;
    decode_soundtrack(data)?;
    Ok(Some(SoundtrackRecord {
        id: reference.identifier,
        data,
    }))
}

fn decode_soundtrack(data: &[u8]) -> Result<kn::Soundtrack> {
    validate_soundtrack_wire(data)?;
    Ok(kn::Soundtrack::decode(data)?)
}

pub(super) fn soundtrack_media_payloads(data: &[u8]) -> Result<Vec<Vec<u8>>> {
    repeated_length_delimited_payloads(data, MEDIA_FIELD)?
        .into_iter()
        .map(|payload| {
            decode_media_reference_payload(payload)?;
            Ok(payload.to_vec())
        })
        .collect()
}

pub(super) fn soundtrack_media_identifiers(data: &[u8]) -> Result<Vec<u64>> {
    soundtrack_media_payloads(data)?
        .into_iter()
        .map(|payload| decode_media_reference_payload(&payload))
        .collect()
}

pub(super) fn encoded_media_reference(data_identifier: u64) -> Result<Vec<u8>> {
    if data_identifier == 0 {
        return Err(Error::InvalidFormat(
            "Keynote soundtrack data reference has identifier zero".to_owned(),
        ));
    }
    Ok(tsp::DataReference {
        identifier: data_identifier,
    }
    .encode_to_vec())
}

pub(super) fn rewrite_soundtrack_media(original: &[u8], payloads: &[Vec<u8>]) -> Result<Vec<u8>> {
    for payload in payloads {
        decode_media_reference_payload(payload)?;
    }
    let data = rewrite_repeated_length_delimited_fields(original, MEDIA_FIELD, payloads)?;
    if soundtrack_media_payloads(&data)? != payloads {
        return Err(Error::InvalidFormat(
            "Keynote soundtrack media rewrite failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn decode_media_reference_payload(payload: &[u8]) -> Result<u64> {
    let reference = tsp::DataReference::decode(payload)?;
    let _ = patch_varint_field(payload, 1, true, Some(reference.identifier))?;
    if reference.identifier == 0 {
        return Err(Error::InvalidFormat(
            "Keynote soundtrack data reference has identifier zero".to_owned(),
        ));
    }
    Ok(reference.identifier)
}

pub(super) fn replace_soundtrack_message(
    archive: &mut Archive,
    soundtrack_id: u64,
    data: Vec<u8>,
) -> Result<()> {
    let object = archive.object_mut(soundtrack_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote soundtrack object {soundtrack_id} is missing"
        ))
    })?;
    let mut message_indices = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == SOUNDTRACK_MESSAGE_TYPE).then_some(index));
    let Some(message_index) = message_indices.next() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote soundtrack object {soundtrack_id} has no Soundtrack payload"
        )));
    };
    if message_indices.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Keynote soundtrack object {soundtrack_id} repeats its Soundtrack payload"
        )));
    }
    let old_identifiers = soundtrack_media_identifiers(&object.messages[message_index].data)?;
    let new_identifiers = soundtrack_media_identifiers(&data)?;
    let info = &mut object.archive_info.message_infos[message_index];
    if info.data_references != old_identifiers {
        return Err(Error::InvalidFormat(format!(
            "Keynote soundtrack MessageInfo data references {:?} do not match payload references {old_identifiers:?}",
            info.data_references
        )));
    }
    for field_info in &mut info.field_infos {
        if field_info.r#type != Some(FieldType::DataReference) {
            continue;
        }
        if field_info.path.path != [MEDIA_FIELD] || field_info.data_references != old_identifiers {
            return Err(Error::InvalidFormat(
                "Keynote soundtrack has unsupported field-level data-reference metadata".to_owned(),
            ));
        }
        field_info.data_references.clone_from(&new_identifiers);
    }
    info.data_references.clone_from(&new_identifiers);
    object.replace_message(
        message_index,
        RawMessage {
            type_: SOUNDTRACK_MESSAGE_TYPE,
            data,
        },
    )?;
    Ok(())
}

fn validate_show_soundtrack_wire(data: &[u8], expected: bool) -> Result<()> {
    let mut count = 0usize;
    for field in parse_wire_fields(data)? {
        if field.number() != SHOW_SOUNDTRACK_FIELD {
            continue;
        }
        if field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "Keynote show soundtrack field uses wire type {}, not {LENGTH_DELIMITED_WIRE_TYPE}",
                field.wire_type()
            )));
        }
        count += 1;
    }
    if count != usize::from(expected) {
        return Err(Error::InvalidFormat(format!(
            "Keynote show has {count} soundtrack fields"
        )));
    }
    Ok(())
}

fn validate_soundtrack_wire(data: &[u8]) -> Result<()> {
    let mut volume_count = 0usize;
    let mut mode_count = 0usize;
    for field in parse_wire_fields(data)? {
        let (name, expected_wire, count) = match field.number() {
            VOLUME_FIELD => ("volume", FIXED64_WIRE_TYPE, Some(&mut volume_count)),
            MODE_FIELD => ("mode", VARINT_WIRE_TYPE, Some(&mut mode_count)),
            MEDIA_FIELD => ("media", LENGTH_DELIMITED_WIRE_TYPE, None),
            _ => continue,
        };
        if field.wire_type() != expected_wire {
            return Err(Error::InvalidFormat(format!(
                "Keynote soundtrack {name} field uses wire type {}, not {expected_wire}",
                field.wire_type()
            )));
        }
        if let Some(count) = count {
            *count += 1;
            if *count > 1 {
                return Err(Error::InvalidFormat(format!(
                    "Keynote soundtrack has duplicate {name} fields"
                )));
            }
        }
    }
    Ok(())
}
