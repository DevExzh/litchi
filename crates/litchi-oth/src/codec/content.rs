//! Web-template content validation.

use litchi_core::{Error, Result};

const BODY_MARKER: &str = "<office:text";
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "content.xml exceeds the family limit".to_string(),
        ));
    }
    if !xml.contains(BODY_MARKER) {
        return Err(Error::InvalidFormat(
            "content.xml has no text body".to_string(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate("<office:text/>").is_ok());
        assert!(validate("<office:chart/>").is_err());
    }
}
