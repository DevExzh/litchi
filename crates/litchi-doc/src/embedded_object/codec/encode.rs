//! DOC and `ObjInfo` wire encoding.

use super::super::model::{FieldMarker, Info, RawPiece};
use super::{
    MAX_PIECES, ODTPERSIST1_RESERVED, ODTPERSIST2_RESERVED, SPRM_C_F_OBJ, SPRM_C_F_OLE2,
    SPRM_C_F_SPEC, SPRM_C_PIC_LOCATION, bit, corrupted, put_fib_pair,
};
use crate::package::Result;

impl Info {
    /// Serialize this ODT deterministically, retaining supported undefined
    /// bits and the presence of an all-zero optional `ODTPersist2`.
    ///
    /// This only writes passive `ObjInfo` metadata bytes. It never opens,
    /// instantiates, or otherwise activates the referenced OLE object.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.reserved_persist1 & !ODTPERSIST1_RESERVED != 0 {
            return Err(corrupted(
                "ObjInfo reserved_persist1 contains defined or MUST-be-zero bits",
            ));
        }
        if self.reserved_persist2 & !ODTPERSIST2_RESERVED != 0 {
            return Err(corrupted(
                "ObjInfo reserved_persist2 contains defined or MUST-be-zero bits",
            ));
        }
        if self.stream_control && !self.activex {
            return Err(corrupted("ObjInfo stream control requires ActiveX"));
        }

        let mut first = self.reserved_persist1;
        first |= bit(self.default_handler, 1);
        first |= bit(self.linked, 4);
        first |= bit(self.display_as_icon, 6);
        first |= bit(self.ole1, 7);
        first |= bit(self.manual_update, 8);
        first |= bit(self.recompose_on_resize, 9);
        first |= bit(self.activex, 12);
        first |= bit(self.stream_control, 13);
        first |= bit(self.view_object, 15);

        let mut second = self.reserved_persist2;
        second |= bit(self.enhanced_metafile, 0);
        second |= bit(self.queried_enhanced_metafile, 2);
        second |= bit(self.stored_as_enhanced_metafile, 3);

        // Inferring presence from non-zero fields keeps construction through
        // the original flattened API ergonomic while preserving an explicit
        // six-byte, all-zero optional member read from a stream.
        let persist2_present = self.persist2_present || second != 0;
        let mut output = Vec::with_capacity(if persist2_present { 6 } else { 4 });
        output.extend_from_slice(&first.to_le_bytes());
        output.extend_from_slice(&self.clipboard_format.to_le_bytes());
        if persist2_present {
            output.extend_from_slice(&second.to_le_bytes());
        }
        Ok(output)
    }
}

pub(in crate::embedded_object) fn serialize_clx(pieces: &[RawPiece]) -> Result<Vec<u8>> {
    if pieces.is_empty() || pieces.len() > MAX_PIECES {
        return Err(corrupted("piece table cardinality is invalid"));
    }
    let plc_size = pieces
        .len()
        .checked_mul(12)
        .and_then(|value| value.checked_add(4))
        .ok_or_else(|| corrupted("PlcPcd size overflow"))?;
    let mut output = vec![2];
    output.extend_from_slice(
        &u32::try_from(plc_size)
            .map_err(|_| corrupted("PlcPcd exceeds u32"))?
            .to_le_bytes(),
    );
    for piece in pieces {
        output.extend_from_slice(&piece.start.to_le_bytes());
    }
    let end = pieces
        .last()
        .ok_or_else(|| corrupted("piece table is empty"))?
        .end;
    output.extend_from_slice(&end.to_le_bytes());
    for piece in pieces {
        output.extend_from_slice(&piece.pcd_prefix);
        let raw_fc = if piece.unicode {
            piece.fc
        } else {
            piece
                .fc
                .checked_mul(2)
                .ok_or_else(|| corrupted("compressed FC overflow"))?
                | 0x4000_0000
        };
        if raw_fc & 0x8000_0000 != 0 {
            return Err(corrupted("FC uses reserved high bit"));
        }
        output.extend_from_slice(&raw_fc.to_le_bytes());
        output.extend_from_slice(&piece.prm);
    }
    Ok(output)
}

pub(in crate::embedded_object) fn serialize_fields(
    fields: &[FieldMarker],
    terminal: u32,
) -> Result<Vec<u8>> {
    super::validate_existing_fields(fields, terminal)?;
    let mut output = Vec::new();
    for marker in fields {
        output.extend_from_slice(&marker.cp.to_le_bytes());
    }
    output.extend_from_slice(&terminal.to_le_bytes());
    for marker in fields {
        output.extend_from_slice(&marker.descriptor);
    }
    Ok(output)
}

pub(in crate::embedded_object) fn object_separator_sprms(storage_id: u32) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
    output.extend_from_slice(&storage_id.to_le_bytes());
    for opcode in [SPRM_C_F_OLE2, SPRM_C_F_SPEC, SPRM_C_F_OBJ] {
        output.extend_from_slice(&opcode.to_le_bytes());
        output.push(1);
    }
    output
}

pub(in crate::embedded_object) fn object_preview_sprms(data_offset: u32) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
    output.extend_from_slice(&data_offset.to_le_bytes());
    output.extend_from_slice(&SPRM_C_F_SPEC.to_le_bytes());
    output.push(1);
    output
}

pub(in crate::embedded_object) fn append_table_block(
    word: &mut [u8],
    table: &mut Vec<u8>,
    index: usize,
    data: &[u8],
) -> Result<()> {
    let offset = u32::try_from(table.len()).map_err(|_| corrupted("Table stream exceeds u32"))?;
    table.extend_from_slice(data);
    put_fib_pair(
        word,
        index,
        offset,
        u32::try_from(data.len()).map_err(|_| corrupted("table block exceeds u32"))?,
    )
}
