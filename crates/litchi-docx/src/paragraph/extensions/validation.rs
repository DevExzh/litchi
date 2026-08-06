//! Checked lexical and structural rules for paragraph extensions.

use crate::error::{Error, Result};

use super::model::{Id, Ids};

pub(crate) fn validate_ids(value: &Ids) -> Result<()> {
    if value.text_id().is_some() && value.para_id().is_none() {
        return Err(Error::InvalidFormat(
            "Word textId requires a paragraph paraId".into(),
        ));
    }
    Ok(())
}

pub(crate) fn parse_id(value: &str, description: &str) -> Result<Id> {
    if value.len() != 8 || !value.bytes().all(is_ascii_hex_digit) {
        return Err(Error::InvalidFormat(format!(
            "invalid {description} '{value}'; expected eight hexadecimal digits"
        )));
    }
    let number = u32::from_str_radix(value, 16).map_err(|error| {
        Error::InvalidFormat(format!("invalid {description} '{value}': {error}"))
    })?;
    Id::new(number).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{description} '{value}' is outside the nonzero, below-0x80000000 range"
        ))
    })
}

pub(crate) fn parse_on_off(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid Word 2010 noSpellErr value '{value}'"
        ))),
    }
}

#[inline]
fn is_ascii_hex_digit(value: u8) -> bool {
    value.is_ascii_digit() || matches!(value, b'a'..=b'f' | b'A'..=b'F')
}
