//! Read-only projection of ODF text sections.

use crate::paragraph::Paragraph;

/// A projected `text:section`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Section {
    condition: Option<String>,
    depth: usize,
    display: Option<bool>,
    name: Option<String>,
    paragraphs: Vec<Paragraph>,
    protected: bool,
    style_name: Option<String>,
    text: String,
}

impl Section {
    #[allow(
        clippy::too_many_arguments,
        reason = "projection fields mirror ODF section semantics"
    )]
    pub(crate) const fn projected(
        name: Option<String>,
        style_name: Option<String>,
        protected: bool,
        display: Option<bool>,
        condition: Option<String>,
        depth: usize,
        text: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            condition,
            depth,
            display,
            name,
            paragraphs,
            protected,
            style_name,
            text,
        }
    }

    /// Producer-visible section name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Section style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Whether the producer marked the section protected.
    #[must_use]
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Requested display state, when explicitly present.
    #[must_use]
    pub const fn display(&self) -> Option<bool> {
        self.display
    }

    /// Inert producer condition, when present.
    #[must_use]
    pub fn condition(&self) -> Option<&str> {
        self.condition.as_deref()
    }

    /// Nesting depth, starting at one.
    #[must_use]
    pub const fn depth(&self) -> usize {
        self.depth
    }

    /// Projected visible section text. Conditions are never evaluated.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Direct paragraphs in source order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}
