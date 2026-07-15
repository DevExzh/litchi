//! Strict native object lookup and wire patching for Keynote soundtracks.

use super::*;

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
    pub(super) native: kn::Soundtrack,
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
    Ok(Some(SoundtrackRecord {
        id: reference.identifier,
        native: decode_soundtrack(data)?,
        data,
    }))
}

pub(super) fn decode_soundtrack(data: &[u8]) -> Result<kn::Soundtrack> {
    validate_soundtrack_wire(data)?;
    Ok(kn::Soundtrack::decode(data)?)
}

pub(super) fn patch_soundtrack_wire(
    original: &[u8],
    soundtrack: &kn::Soundtrack,
    settings: &KeynoteSoundtrackSettings,
) -> Result<Vec<u8>> {
    let data = patch_fixed64_field(
        original,
        VOLUME_FIELD,
        soundtrack.volume.is_some(),
        settings.volume.map(f64::to_bits),
    )?;
    patch_varint_field(
        &data,
        MODE_FIELD,
        soundtrack.mode.is_some(),
        settings.mode.map(|mode| i64::from(mode.as_raw()) as u64),
    )
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
        if field.number != SHOW_SOUNDTRACK_FIELD {
            continue;
        }
        if field.wire_type != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "Keynote show soundtrack field uses wire type {}, not {LENGTH_DELIMITED_WIRE_TYPE}",
                field.wire_type
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
        let (name, expected_wire, count) = match field.number {
            VOLUME_FIELD => ("volume", FIXED64_WIRE_TYPE, Some(&mut volume_count)),
            MODE_FIELD => ("mode", VARINT_WIRE_TYPE, Some(&mut mode_count)),
            MEDIA_FIELD => ("media", LENGTH_DELIMITED_WIRE_TYPE, None),
            _ => continue,
        };
        if field.wire_type != expected_wire {
            return Err(Error::InvalidFormat(format!(
                "Keynote soundtrack {name} field uses wire type {}, not {expected_wire}",
                field.wire_type
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
