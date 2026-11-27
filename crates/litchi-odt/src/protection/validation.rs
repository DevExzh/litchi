//! Bounds and lexical validation for the ODT protection owner.

use litchi_core::{Error, Result};

pub(super) fn validate_type(actual: &str, expected: &str, field: &str) -> Result<()> {
    if actual == expected {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{field} uses config:type '{actual}', expected '{expected}'"
        )))
    }
}

pub(super) fn validate_xml_size(xml: &[u8]) -> Result<()> {
    if xml.len() > super::MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODT protection settings exceed the {} byte limit",
            super::MAX_XML_BYTES
        )));
    }
    Ok(())
}
