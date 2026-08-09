//! Web-template heading semantics.

use crate::link::Link;

/// A projected `text:h` block.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Heading {
    level: u8,
    links: Vec<Link>,
    style_name: Option<String>,
    text: String,
}

impl Heading {
    /// Creates a detached plain heading with a positive outline level.
    ///
    /// # Errors
    ///
    /// Returns an error when `level` is zero.
    pub fn new(level: u8, text: impl Into<String>) -> litchi_core::Result<Self> {
        validate_level(level)?;
        Ok(Self {
            level,
            links: Vec::new(),
            style_name: None,
            text: text.into(),
        })
    }

    /// Creates a detached heading with a style reference.
    ///
    /// # Errors
    ///
    /// Returns an error when `level` is zero.
    pub fn styled(
        level: u8,
        text: impl Into<String>,
        style_name: impl Into<String>,
    ) -> litchi_core::Result<Self> {
        validate_level(level)?;
        Ok(Self {
            level,
            links: Vec::new(),
            style_name: Some(style_name.into()),
            text: text.into(),
        })
    }

    pub(crate) const fn projected(
        level: u8,
        links: Vec<Link>,
        style_name: Option<String>,
        text: String,
    ) -> Self {
        Self {
            level,
            links,
            style_name,
            text,
        }
    }

    /// Returns the ODF outline level. Missing source levels project as level 1.
    #[must_use]
    pub const fn level(&self) -> u8 {
        self.level
    }

    /// Returns inert hyperlinks contained by the heading in document order.
    #[must_use]
    pub fn links(&self) -> &[Link] {
        &self.links
    }

    /// Returns the referenced paragraph style name, if present.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Returns projected character data, including ODF whitespace elements.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

fn validate_level(level: u8) -> litchi_core::Result<()> {
    if level == 0 {
        return Err(litchi_core::Error::InvalidFormat(
            "OTH heading outline level must be positive".to_string(),
        ));
    }
    Ok(())
}
