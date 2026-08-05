//! Layered BIFF8 OLE-object owner for legacy XLS workbooks.
//!
//! Typed Obj/form-control models, BIFF codecs, and workbook/CFB transactions
//! stay in contextual layers while this module remains the ergonomic facade.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use litchi_ole_common::object::Limits;
pub use model::*;
pub use package::Editor;

const OBJ: u16 = 0x005D;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;
const BOUNDSHEET: u16 = 0x0085;
const EOF: u16 = 0x000A;
const FT_CMO: u16 = 0x0015;
const FT_CF: u16 = 0x0007;
const FT_PIO: u16 = 0x0008;
const FT_PICT_FMLA: u16 = 0x0009;
const FT_SBS: u16 = 0x000C;
const FT_GBO_DATA: u16 = 0x000F;
const FT_EDO_DATA: u16 = 0x0010;
const FT_RBO_DATA: u16 = 0x0011;
const FT_CBLS_DATA: u16 = 0x0012;
const FT_LBS_DATA: u16 = 0x0013;
const FT_END: u16 = 0;
/// MS-CFB directory-entry type for a user stream.
const CFB_STREAM: u8 = 0x02;

/// MS-XLS 2.5.141/2.5.145 `fNo3d`: the control is drawn without 3-D effects.
const NO_3D: u16 = 0x0001;
/// MS-XLS 2.5.154 FtSbs flag bits.
const SBS_DRAW: u16 = 0x0001;
const SBS_DRAW_SLIDER_ONLY: u16 = 0x0002;
const SBS_TRACK_ELEVATOR: u16 = 0x0004;
const SBS_NO_3D: u16 = 0x0008;
/// MS-XLS 2.5.147 FtLbsData flag bits.
const LBS_USE_CB: u16 = 0x0001;
const LBS_VALID_PLEX: u16 = 0x0002;
const LBS_VALID_IDS: u16 = 0x0004;
const LBS_NO_3D: u16 = 0x0008;
const LBS_SELECTION_TYPE_SHIFT: u16 = 4;
const LBS_SELECTION_TYPE_MASK: u16 = 0x0003;
const LBS_BEHAVIOR_CLASS_SHIFT: u16 = 8;
/// MS-XLS 2.5.171 LbsDropData flag bits.
const DROP_STYLE_MASK: u16 = 0x0003;
const DROP_FILTERED: u16 = 0x0004;
/// MS-XLS 2.5.294 XLUnicodeString option bits.
const XL_STRING_HIGH_BYTE: u8 = 0x01;
const XL_STRING_EXT: u8 = 0x04;
const XL_STRING_RICH: u8 = 0x08;
/// Size in bytes of one formatting run in a rich XLUnicodeString.
const FORMATTING_RUN_SIZE: usize = 4;

fn invalid(record_type: u16, message: impl Into<String>) -> crate::error::Error {
    crate::error::Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}
