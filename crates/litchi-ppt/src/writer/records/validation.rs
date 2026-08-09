use super::model::Error;

pub(super) fn validate_complete_record(
    data: &[u8],
    expected_version: u8,
    expected_type: u16,
) -> Result<(), Error> {
    let header = data
        .get(..8)
        .ok_or_else(|| std::io::Error::other("programmable tag record is truncated"))?;
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let record_type = u16::from_le_bytes([header[2], header[3]]);
    let length = usize::try_from(u32::from_le_bytes([
        header[4], header[5], header[6], header[7],
    ]))
    .map_err(|_err| std::io::Error::other("programmable tag length overflows usize"))?;
    if version_instance & 0x000f != u16::from(expected_version)
        || record_type != expected_type
        || length.checked_add(8) != Some(data.len())
    {
        return Err(std::io::Error::other(
            "invalid additional DocProgBinaryTag record",
        ));
    }
    Ok(())
}
