//! Shared XML decoding helpers.

use crate::error::{OoxmlError, Result};
use quick_xml::events::BytesRef;

/// Decode a numeric or predefined XML entity reference.
pub(crate) fn decode_xml_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| OoxmlError::Xml(error.to_string()))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "unsupported XML entity reference '&{name};'"
        ))),
    }
}
