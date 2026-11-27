//! OLE metadata stream codecs for the DOC writer.

use super::model::{CompObj, Ole};

/// Fixed length of the Word writer's `\x01CompObj` stream.
pub(super) const COMP_OBJ_LEN: usize = 106;
/// Fixed length of the Word writer's `\x01Ole` stream.
pub(super) const OLE_LEN: usize = 20;

pub(super) const COMPOBJ_VERSION: [u8; 4] = [0x01, 0x00, 0xFE, 0xFF];
pub(super) const COMPOBJ_RESERVED: [u8; 4] = [0x03, 0x0A, 0x00, 0x00];
pub(super) const COMPOBJ_RESERVED_MARKER: [u8; 4] = [0xFF, 0xFF, 0xFF, 0xFF];
pub(super) const UNICODE_MARKER: u32 = 0x71B2_39F4;

/// Encode the semantic `\x01CompObj` contents.
pub(super) fn write_comp_obj(value: CompObj) -> Vec<u8> {
    let mut data = Vec::with_capacity(COMP_OBJ_LEN);
    data.extend_from_slice(&COMPOBJ_VERSION);
    data.extend_from_slice(&COMPOBJ_RESERVED);
    data.extend_from_slice(&COMPOBJ_RESERVED_MARKER);
    data.extend_from_slice(value.class_id().as_bytes());
    write_ansi_string(&mut data, value.user_type());
    write_ansi_string(&mut data, value.clipboard_format());
    write_ansi_string(&mut data, value.prog_id());
    data.extend_from_slice(&UNICODE_MARKER.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data
}

/// Encode the semantic `\x01Ole` contents.
pub(super) fn write_ole(value: Ole) -> Vec<u8> {
    let mut data = Vec::with_capacity(OLE_LEN);
    data.extend_from_slice(&value.version().to_le_bytes());
    data.resize(OLE_LEN, 0);
    data
}

fn write_ansi_string(data: &mut Vec<u8>, value: &str) {
    let bytes = value.as_bytes();
    let length = bytes.len() + 1;
    data.extend_from_slice(&(length as u32).to_le_bytes());
    data.extend_from_slice(bytes);
    data.push(0);
}
