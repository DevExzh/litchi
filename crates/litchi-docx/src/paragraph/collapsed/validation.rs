//! Lexical and resource validation for the Word 2012 paragraph extension.

use crate::error::{Error, Result};

use super::model::Collapsed;

pub(crate) const WORD_2012_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2012/wordml";
pub(crate) const MAX_XML_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_XML_NODES: usize = 16 * 1024;
pub(crate) const MAX_XML_DEPTH: usize = 64;

pub(crate) fn validate(value: Option<Collapsed>) -> Result<()> {
    let _ = value;
    Ok(())
}

pub(crate) fn parse_on_off(value: Option<&str>) -> Result<Collapsed> {
    match value {
        None | Some("true") | Some("1") | Some("on") => Ok(Collapsed::Enabled),
        Some("false") | Some("0") | Some("off") => Ok(Collapsed::Disabled),
        Some(value) => Err(Error::InvalidFormat(format!(
            "invalid Word 2012 collapsed value '{value}'"
        ))),
    }
}
