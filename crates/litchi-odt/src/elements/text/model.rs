//! Semantic values owned by the ODT text-element facade.
//!
//! The XML wrappers remain in the parent facade for a stable, discoverable
//! API.  Values that describe text content itself live here so the decoder
//! does not need to expose parser state or wire-level details.

use super::{Heading, Paragraph};
use litchi_core::Result;

/// A decoded block-level text element.
#[derive(Debug, Clone)]
pub enum Block {
    /// A `text:p` paragraph.
    Paragraph(Paragraph),
    /// A `text:h` heading.
    Heading(Heading),
}

impl Block {
    /// Return the semantic kind of this block.
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Paragraph(_) => Kind::Paragraph,
            Self::Heading(_) => Kind::Heading,
        }
    }

    /// Read the flattened visible text in this block.
    pub fn text(&self) -> Result<String> {
        match self {
            Self::Paragraph(paragraph) => paragraph.text(),
            Self::Heading(heading) => heading.text(),
        }
    }

    pub(crate) fn into_text(self) -> String {
        match self {
            Self::Paragraph(paragraph) => paragraph.into_text(),
            Self::Heading(heading) => heading.into_text(),
        }
    }

    /// Borrow the paragraph when this block is a paragraph.
    pub const fn as_paragraph(&self) -> Option<&Paragraph> {
        match self {
            Self::Paragraph(paragraph) => Some(paragraph),
            Self::Heading(_) => None,
        }
    }

    /// Borrow the heading when this block is a heading.
    pub const fn as_heading(&self) -> Option<&Heading> {
        match self {
            Self::Paragraph(_) => None,
            Self::Heading(heading) => Some(heading),
        }
    }

    /// Consume the block and return its paragraph, if it is one.
    pub fn into_paragraph(self) -> Option<Paragraph> {
        match self {
            Self::Paragraph(paragraph) => Some(paragraph),
            Self::Heading(_) => None,
        }
    }

    /// Consume the block and return its heading, if it is one.
    pub fn into_heading(self) -> Option<Heading> {
        match self {
            Self::Paragraph(_) => None,
            Self::Heading(heading) => Some(heading),
        }
    }
}

/// The block-level element kind represented by [`Block`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Paragraph,
    Heading,
}

/// Allowed `xlink:show` values for a simple ODF hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkShow {
    /// Open the target in a new frame or window.
    New,
    /// Replace the current frame or window.
    Replace,
}

impl LinkShow {
    /// Return the ODF/XLink lexical value.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::New => "new",
            Self::Replace => "replace",
        }
    }

    /// Parse an ODF/XLink lexical value.
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "new" => Some(Self::New),
            "replace" => Some(Self::Replace),
            _ => None,
        }
    }
}

/// Allowed explicit `xlink:actuate` values for a simple ODF hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkActuate {
    /// Activate the target only on an explicit user request.
    OnRequest,
}

impl LinkActuate {
    /// Return the ODF/XLink lexical value.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::OnRequest => "onRequest",
        }
    }

    /// Parse an ODF/XLink lexical value.
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "onRequest" => Some(Self::OnRequest),
            _ => None,
        }
    }
}
