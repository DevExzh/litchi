//! Typed object model for an `OpenDocument` ruby annotation.

/// A ruby base/pronunciation annotation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Annotation {
    pub(super) style_name: Option<String>,
    pub(super) base: String,
    pub(super) text: String,
    pub(super) text_style_name: Option<String>,
}

impl Annotation {
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    pub fn base(&self) -> &str {
        &self.base
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }
}
