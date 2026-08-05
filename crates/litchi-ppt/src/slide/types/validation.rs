//! Narrow binary invariants shared by slide codecs.

use crate::records::Record;

pub(super) const MAX_SHAPE_DEPTH: usize = 256;

pub(super) fn is_ppt10_tag_name(record: &Record) -> bool {
    const PPT10: [u16; 8] = [0x5F, 0x5F, 0x5F, 0x50, 0x50, 0x54, 0x31, 0x30];
    record.version == 0
        && record.instance == 0
        && record.data.len() == 16
        && record
            .data
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .eq(PPT10)
}
