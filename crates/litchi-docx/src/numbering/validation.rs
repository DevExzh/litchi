//! Bounded lexical and namespace validation for WordprocessingML numbering.

use crate::{Error, Result};

/// Word 2012 WordprocessingML namespace used by the numbering extension.
pub(crate) const WORD_2012_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2012/wordml";
/// Markup-compatibility namespace used by `mc:Ignorable`.
pub(crate) const MC_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
/// Maximum numbering part retained by a source-preserving snapshot.
pub(crate) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum element nesting accepted in one numbering part.
pub(crate) const MAX_XML_DEPTH: usize = 128;
/// Maximum element count accepted in one numbering part.
pub(crate) const MAX_XML_NODES: usize = 1_000_000;

pub(crate) fn validate_xml(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "Word numbering XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

pub(crate) fn parse_on_off(value: &str) -> Result<bool> {
    match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid Word numbering ST_OnOff value '{value}'"
        ))),
    }
}

pub(crate) fn validate_restart_numbering_after_break(value: Option<bool>) -> Result<()> {
    let _ = value;
    Ok(())
}

pub(crate) fn has_ignorable_prefix(value: &[Vec<u8>], prefix: &[u8]) -> bool {
    value.iter().any(|candidate| candidate.as_slice() == prefix)
}

pub(crate) fn parse_ignorable(value: &str) -> Result<Vec<Vec<u8>>> {
    let mut prefixes = Vec::new();
    for prefix in value.split_ascii_whitespace() {
        if prefix.is_empty()
            || !prefix
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
            || prefixes
                .iter()
                .any(|candidate: &Vec<u8>| candidate == prefix.as_bytes())
        {
            return Err(Error::InvalidFormat(
                "invalid or duplicate numbering mc:Ignorable prefix".into(),
            ));
        }
        prefixes.push(prefix.as_bytes().to_vec());
    }
    Ok(prefixes)
}
