use crate::Result;

use super::support::invalid;

/// Maximum number of Unicode scalar values accepted in a smart-tag namespace URI.
pub const MAX_SMART_TAG_NAMESPACE_URI_CHARS: usize = 2083;
/// Maximum number of Unicode scalar values accepted in a smart-tag name.
pub const MAX_SMART_TAG_NAME_CHARS: usize = 255;
/// Maximum number of Unicode scalar values accepted in a smart-tag download URL.
pub const MAX_SMART_TAG_URL_CHARS: usize = 2083;

/// A smart-tag vocabulary declaration from a WordprocessingML settings part.
///
/// The value is deliberately package-neutral: relationship resolution,
/// document matching, and settings-part orchestration remain in the host
/// package facade.  The three attributes are required by the host settings
/// vocabulary; an empty-but-present attribute is retained for compatibility
/// with the historical parser.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagType {
    namespace_uri: String,
    name: String,
    url: String,
}

impl SmartTagType {
    /// Construct a validated smart-tag vocabulary declaration.
    pub fn new(
        namespace_uri: impl Into<String>,
        name: impl Into<String>,
        url: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            namespace_uri: namespace_uri.into(),
            name: name.into(),
            url: url.into(),
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate the client length limits for this smart-tag declaration.
    pub fn validate(&self) -> Result<()> {
        validate_smart_tag_type(&self.namespace_uri, &self.name, &self.url)
    }

    /// Return the smart-tag vocabulary namespace URI.
    #[inline]
    pub fn namespace_uri(&self) -> &str {
        &self.namespace_uri
    }

    /// Return the smart-tag type name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the vocabulary download URL.
    #[inline]
    pub fn url(&self) -> &str {
        &self.url
    }
}

/// Validate the bounded client representation of a smart-tag declaration.
pub fn validate_smart_tag_type(namespace_uri: &str, name: &str, url: &str) -> Result<()> {
    validate_smart_tag_value(
        namespace_uri,
        "namespace URI",
        MAX_SMART_TAG_NAMESPACE_URI_CHARS,
    )?;
    validate_smart_tag_value(name, "name", MAX_SMART_TAG_NAME_CHARS)?;
    validate_smart_tag_value(url, "URL", MAX_SMART_TAG_URL_CHARS)
}

fn validate_smart_tag_value(value: &str, description: &str, maximum: usize) -> Result<()> {
    if value.chars().count() > maximum {
        return Err(invalid(format!(
            "Word smart-tag {description} exceeds {maximum} characters"
        )));
    }
    Ok(())
}
