//! `OfficeArt` `ClientData` child-record assembly.

use super::super::{Error, header_version, ppt_record_type};
use super::wire::EscherBuilder;
use litchi_odraw::write::record_type;

pub(super) fn legacy_hyperlink_interaction(
    hyperlink_id: u32,
    action: u8,
    jump: u8,
    hyperlink_type: u8,
) -> Result<crate::Interaction, Error> {
    let mut atom_data = [0u8; 16];
    atom_data[4..8].copy_from_slice(&hyperlink_id.to_le_bytes());
    atom_data[8] = action;
    atom_data[10] = jump;
    atom_data[12] = hyperlink_type;
    let atom = crate::InteractiveInfoAtom::parse_payload(&atom_data)
        .map_err(|error| std::io::Error::other(error.to_string()))?;
    Ok(crate::Interaction {
        trigger: crate::InteractionTrigger::Click,
        sound_id: atom.sound_id,
        hyperlink_id: atom.hyperlink_id,
        action: atom.action,
        ole_verb: atom.ole_verb,
        jump: atom.jump,
        animated: atom.animated,
        stop_sound: atom.stop_sound,
        custom_show_return: atom.custom_show_return,
        visited: atom.visited,
        link_target: atom.link_target,
        macro_name: None,
        unused: atom.unused,
        macro_name_data: None,
    })
}

pub(super) fn append_client_data_record_payload(
    client_data: &mut Option<Vec<u8>>,
    record: &[u8],
) -> Result<(), Error> {
    let payload_len = record
        .get(4..8)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .and_then(|length| usize::try_from(length).ok())
        .ok_or_else(|| std::io::Error::other("invalid ClientData record header"))?;
    let expected_options = 0x000fu16.to_le_bytes();
    let expected_type = record_type::CLIENT_DATA.to_le_bytes();
    if record.len() != payload_len.saturating_add(8)
        || record.get(0..2) != Some(expected_options.as_slice())
        || record.get(2..4) != Some(expected_type.as_slice())
    {
        return Err(std::io::Error::other("invalid ClientData record"));
    }
    append_client_data_payload(client_data, &record[8..])
}

pub(super) fn append_client_data_payload(
    client_data: &mut Option<Vec<u8>>,
    payload: &[u8],
) -> Result<(), Error> {
    let data = client_data.get_or_insert_with(|| {
        let mut bytes = Vec::with_capacity(8);
        bytes.extend_from_slice(&0x000fu16.to_le_bytes());
        bytes.extend_from_slice(&record_type::CLIENT_DATA.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes
    });
    let declared = data
        .get(4..8)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or_else(|| std::io::Error::other("invalid ClientData record header"))?;
    if data.len()
        != usize::try_from(declared)
            .unwrap_or(usize::MAX)
            .saturating_add(8)
    {
        return Err(std::io::Error::other(
            "ClientData record length does not match its payload",
        ));
    }
    let new_length = u32::try_from(data.len().saturating_sub(8))
        .ok()
        .and_then(|length| {
            u32::try_from(payload.len())
                .ok()
                .and_then(|addition| length.checked_add(addition))
        })
        .ok_or_else(|| std::io::Error::other("ClientData payload exceeds u32"))?;
    data.extend_from_slice(payload);
    data[4..8].copy_from_slice(&new_length.to_le_bytes());
    Ok(())
}

/// Builds `ClientData` with the legacy PPT hyperlink interaction atom.
#[cfg(test)]
pub(crate) fn build_client_data_with_hyperlink(
    hyperlink_id: u32,
    action: u8,
    jump: u8,
    hyperlink_type: u8,
) -> Result<Vec<u8>, Error> {
    let interaction = legacy_hyperlink_interaction(hyperlink_id, action, jump, hyperlink_type)?;
    let mut client_data = None;
    append_client_data_payload(
        &mut client_data,
        &interaction
            .to_bytes()
            .map_err(|error| std::io::Error::other(error.to_string()))?,
    )?;
    client_data.ok_or_else(|| std::io::Error::other("missing ClientData record"))
}

/// Builds `ClientData` containing one animation-information container.
pub(super) fn build_client_data_with_animation(
    animation_info: &crate::animation::AnimationInfo,
) -> Result<Vec<u8>, Error> {
    use crate::animation::writer::write_animation_info;

    let (animation_bytes, _sound_ref) = write_animation_info(animation_info).map_err(|error| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, error.to_string())
    })?;

    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&animation_bytes);
    client_data.build()
}

/// Builds `ClientData` with an `OEPlaceholderAtom` for a placeholder shape.
pub(crate) fn build_client_data_with_placeholder(placeholder_type: u8) -> Result<Vec<u8>, Error> {
    use crate::writer::records::RecordBuilder;

    let mut oe_atom = RecordBuilder::new(0x00, 0, ppt_record_type::OE_PLACEHOLDER_ATOM);
    oe_atom.write_data(&0u32.to_le_bytes());
    oe_atom.write_data(&[placeholder_type]);
    oe_atom.write_data(&[0x00]);
    oe_atom.write_data(&[0x00, 0x00]);
    let oe_bytes = oe_atom.build()?;

    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&oe_bytes);
    client_data.build()
}
