//! Read-only projection of ODF ruby annotations.

/// A `text:ruby` base/pronunciation pair.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Ruby {
    base: String,
    style_name: Option<String>,
    text: String,
    text_style_name: Option<String>,
}

impl Ruby {
    pub(crate) const fn projected(
        style_name: Option<String>,
        text_style_name: Option<String>,
        base: String,
        text: String,
    ) -> Self {
        Self {
            base,
            style_name,
            text,
            text_style_name,
        }
    }

    /// Ruby base text.
    #[must_use]
    pub fn base(&self) -> &str {
        &self.base
    }

    /// Ruby pronunciation text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Ruby style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Ruby text style reference.
    #[must_use]
    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }
}
