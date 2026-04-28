#[inline(always)]
pub(crate) fn le_u64(d: &[u8]) -> u64 {
    u64::from_le_bytes([d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7]])
}

#[inline(always)]
pub(crate) fn le_u32(d: &[u8]) -> u32 {
    u32::from_le_bytes([d[0], d[1], d[2], d[3]])
}

#[inline(always)]
pub(crate) fn le_u16(d: &[u8]) -> u16 {
    u16::from_le_bytes([d[0], d[1]])
}
