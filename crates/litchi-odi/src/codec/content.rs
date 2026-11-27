//! Image content validation.

use litchi_core::{Error, Result};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "content.xml exceeds the family limit".to_string(),
        ));
    }
    crate::flat::validate_content_xml(xml)
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate(r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:image><draw:frame><draw:image xlink:href="x" xmlns:xlink="http://www.w3.org/1999/xlink"/></draw:frame></office:image></office:body></office:document-content>"#).is_ok());
        assert!(validate("<office:text/>").is_err());
    }
}
