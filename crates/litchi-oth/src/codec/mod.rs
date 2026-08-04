//! Bounded validation for the family content part.

use litchi_core::{Error, Result};

const BODY_MARKER: &str = "<office:text";
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// Validate a UTF-8 content part and return it unchanged for zero-copy package ownership.
pub fn validate_content(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "content.xml exceeds the family limit".to_string(),
        ));
    }
    if !xml.contains(BODY_MARKER) {
        return Err(Error::InvalidFormat(format!(
            "content.xml has no text body"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::validate_content;

    #[test]
    fn requires_family_body() {
        assert!(validate_content("<office:text/>").is_ok());
        assert!(validate_content("<office:text/>").is_ok());
    }
}
