//! Detached text-body block values.

/// A detached block accepted by fresh and source-bound authoring.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Content {
    /// A heading.
    Heading(crate::heading::Heading),
    /// A flat list.
    List(crate::list::List),
    /// A paragraph.
    Paragraph(crate::paragraph::Paragraph),
}

impl From<crate::heading::Heading> for Content {
    fn from(value: crate::heading::Heading) -> Self {
        Self::Heading(value)
    }
}

impl From<crate::list::List> for Content {
    fn from(value: crate::list::List) -> Self {
        Self::List(value)
    }
}

impl From<crate::paragraph::Paragraph> for Content {
    fn from(value: crate::paragraph::Paragraph) -> Self {
        Self::Paragraph(value)
    }
}
