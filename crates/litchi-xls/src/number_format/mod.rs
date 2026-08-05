//! Layered BIFF8 workbook-formatting owner.

mod codec;
mod model;

#[cfg(test)]
mod tests;

const DATE1904_RECORD: u16 = 0x0022;

/// MS-XLS 2.4.126 `Format` record type.
pub(crate) const FORMAT_RECORD: u16 = 0x041e;
/// MS-XLS 2.4.353 `XF` record type.
pub(crate) const XF_RECORD: u16 = 0x00e0;
/// MS-XLS 2.4.354 `XFCRC` record type.
pub(crate) const XFCRC_RECORD: u16 = 0x087c;

const MAX_DXF_RECORDS: usize = 65_536;
const MAX_FORMAT_RECORDS: usize = 218;
const MIN_XF_RECORDS: usize = 16;
const MAX_XF_RECORDS: usize = 65_536;

/// `fHighByte`: the characters in `rgb` are UTF-16 rather than compressed.
pub(crate) const XL_UNICODE_STRING_HIGH_BYTE: u8 = 0x01;
/// Bytes per character when `fHighByte` is set.
pub(crate) const UTF16_CHAR_BYTES: usize = 2;
/// Bytes per character in a compressed (single-byte) string.
pub(crate) const COMPRESSED_CHAR_BYTES: usize = 1;

pub use model::{
    DateSystem, EffectiveExtendedFormat, ExtendedFormat, ExtendedFormatApplications,
    ExtendedFormatKind, Formatting, NumberFormat,
};
