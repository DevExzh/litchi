use super::{CommentDateTime, WriteError};

pub(in crate::writer::core) fn utf16_code_unit_len(text: &str) -> Result<u32, WriteError> {
    let length = u32::try_from(text.encode_utf16().count())
        .map_err(|_| WriteError::InvalidData("DOC text exceeds the 32-bit CP range".to_string()))?;
    if length >= 0x7FFF_FFFF {
        return Err(WriteError::InvalidData(
            "DOC text exceeds the MS-DOC CP limit".to_string(),
        ));
    }
    Ok(length)
}

pub(crate) fn pack_dttm(value: Option<CommentDateTime>) -> Result<u32, WriteError> {
    let Some(value) = value else {
        return Ok(0);
    };
    if !(1900..=2411).contains(&value.year)
        || !(1..=12).contains(&value.month)
        || !(1..=31).contains(&value.day)
        || value.hour > 23
        || value.minute > 59
        || value.weekday > 6
    {
        return Err(WriteError::InvalidData(
            "DOC timestamp is outside the DTTM field ranges".to_string(),
        ));
    }
    Ok(u32::from(value.minute)
        | (u32::from(value.hour) << 6)
        | (u32::from(value.day) << 11)
        | (u32::from(value.month) << 16)
        | (u32::from(value.year - 1900) << 20)
        | (u32::from(value.weekday) << 29))
}
