//! Drawing content validation.

use litchi_core::Result;
use litchi_odf_common::core::validate_content_part;

const BODY_MARKER: &str = "<office:drawing";

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    validate_content_part(xml, BODY_MARKER, "ODG")
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate("<office:drawing/>").is_ok());
        assert!(validate("<office:text/>").is_err());
    }
}
