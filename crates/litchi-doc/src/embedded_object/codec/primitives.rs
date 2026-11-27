//! Bounds-checked DOC stream primitives shared by the wire layers.

use crate::package::{Error as PackageError, Result};

pub(in crate::embedded_object) const FIB_CCP_TEXT: usize = 76;
pub(in crate::embedded_object) const FIB_FC_LCB: usize = 154;
pub(in crate::embedded_object) const PLCFBTE_CHPX: usize = 12;
pub(in crate::embedded_object) const PLCFFLD_MOM: usize = 16;
pub(in crate::embedded_object) const CLX: usize = 33;
pub(in crate::embedded_object) const SPRM_C_PIC_LOCATION: u16 = 0x6A03;
pub(in crate::embedded_object) const SPRM_C_F_OLE2: u16 = 0x080A;
pub(in crate::embedded_object) const SPRM_C_F_SPEC: u16 = 0x0855;
pub(in crate::embedded_object) const SPRM_C_F_OBJ: u16 = 0x0856;
pub(in crate::embedded_object) const MAX_PIECES: usize = 65_536;
pub(in crate::embedded_object) const MAX_FIELDS: usize = 65_536;
pub(in crate::embedded_object) const MAX_PICF: usize = 128 * 1024 * 1024;
pub(in crate::embedded_object) const OBJ_INFO_STREAM: &str = "\u{3}ObjInfo";
pub(in crate::embedded_object) const ODTPERSIST1_MUST_BE_ZERO: u16 = (1 << 10) | (1 << 11);
pub(in crate::embedded_object) const ODTPERSIST1_RESERVED: u16 =
    (1 << 0) | (1 << 2) | (1 << 3) | (1 << 5) | (1 << 14);
pub(in crate::embedded_object) const ODTPERSIST2_MUST_BE_ZERO: u16 = 1 << 1;
pub(in crate::embedded_object) const ODTPERSIST2_RESERVED: u16 = 0xFFF0;

#[inline]
pub(in crate::embedded_object) const fn bit(value: bool, position: u16) -> u16 {
    if value { 1 << position } else { 0 }
}

pub(in crate::embedded_object) fn word(data: &[u8], offset: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(array_at(data, offset, "ObjInfo word")?))
}

pub(in crate::embedded_object) fn fib_pair(word: &[u8], index: usize) -> Result<(u32, u32)> {
    Ok((
        u32_at(word, FIB_FC_LCB + index * 8)?,
        u32_at(word, FIB_FC_LCB + index * 8 + 4)?,
    ))
}

pub(in crate::embedded_object) fn put_fib_pair(
    word: &mut [u8],
    index: usize,
    fc: u32,
    lcb: u32,
) -> Result<()> {
    put_u32(word, FIB_FC_LCB + index * 8, fc)?;
    put_u32(word, FIB_FC_LCB + index * 8 + 4, lcb)
}

pub(in crate::embedded_object) fn slice<'a>(
    data: &'a [u8],
    offset: u32,
    length: u32,
    name: &str,
) -> Result<&'a [u8]> {
    let start = offset as usize;
    let end = start
        .checked_add(length as usize)
        .ok_or_else(|| corrupted(format!("{name} range overflow")))?;
    data.get(start..end)
        .ok_or_else(|| corrupted(format!("{name} exceeds stream")))
}

pub(in crate::embedded_object) fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(array_at(data, offset, "u16")?))
}

pub(in crate::embedded_object) fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(array_at(data, offset, "u32")?))
}

pub(in crate::embedded_object) fn put_u32(
    data: &mut [u8],
    offset: usize,
    value: u32,
) -> Result<()> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| corrupted("FIB field offset overflow"))?;
    let slot = data
        .get_mut(offset..end)
        .ok_or_else(|| corrupted("truncated FIB field"))?;
    slot.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

pub(in crate::embedded_object) fn array_at<const N: usize>(
    data: &[u8],
    offset: usize,
    name: &str,
) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| corrupted(format!("{name} offset overflow")))?;
    data.get(offset..end)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| corrupted(format!("truncated {name}")))
}

pub(in crate::embedded_object) fn align2(value: usize) -> Result<usize> {
    value
        .checked_add(1)
        .map(|v| v & !1)
        .ok_or_else(|| corrupted("alignment overflow"))
}

pub(in crate::embedded_object) fn align4(value: usize) -> Result<usize> {
    value
        .checked_add(3)
        .map(|v| v & !3)
        .ok_or_else(|| corrupted("alignment overflow"))
}

pub(in crate::embedded_object) fn align512(value: usize) -> Result<usize> {
    value
        .checked_add(511)
        .map(|v| v & !511)
        .ok_or_else(|| corrupted("alignment overflow"))
}

pub(in crate::embedded_object) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
