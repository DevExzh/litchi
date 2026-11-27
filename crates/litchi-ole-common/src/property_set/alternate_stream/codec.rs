//! Wire codec for the [MS-OLEPS] alternate-stream control packet.

use super::super::model::{Guid, invalid};
use super::AlternateStreamControl;
use litchi_cfb::OleError;

const CONTROL_PACKET_WITHOUT_CLSID_BYTES: usize = 8;
const CONTROL_PACKET_WITH_CLSID_BYTES: usize = 24;

pub(super) fn decode(bytes: &[u8]) -> Result<AlternateStreamControl, OleError> {
    if !matches!(
        bytes.len(),
        CONTROL_PACKET_WITHOUT_CLSID_BYTES | CONTROL_PACKET_WITH_CLSID_BYTES
    ) {
        return Err(invalid(
            "alternate-stream control packet must be exactly 8 or 24 bytes",
        ));
    }

    let reserved1 = u16::from_le_bytes([bytes[0], bytes[1]]);
    if reserved1 != 0 {
        return Err(invalid("alternate-stream control Reserved1 must be zero"));
    }
    let reserved2 = u16::from_le_bytes([bytes[2], bytes[3]]);
    let application_state = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    let class_identifier = if bytes.len() == CONTROL_PACKET_WITH_CLSID_BYTES {
        let mut value = [0u8; 16];
        value.copy_from_slice(&bytes[8..24]);
        Some(Guid::from_bytes(value))
    } else {
        None
    };
    Ok(AlternateStreamControl::from_wire(
        reserved2,
        application_state,
        class_identifier,
    ))
}

pub(super) fn encode(control: &AlternateStreamControl) -> Vec<u8> {
    let length = if control.class_identifier().is_some() {
        CONTROL_PACKET_WITH_CLSID_BYTES
    } else {
        CONTROL_PACKET_WITHOUT_CLSID_BYTES
    };
    let mut bytes = vec![0u8; length];
    bytes[2..4].copy_from_slice(&control.reserved2().to_le_bytes());
    bytes[4..8].copy_from_slice(&control.application_state().to_le_bytes());
    if let Some(class_identifier) = control.class_identifier() {
        bytes[8..24].copy_from_slice(class_identifier.as_bytes());
    }
    bytes
}
