//! Read-only projection of ODF footnotes and endnotes.

use crate::paragraph::Paragraph;

/// ODF note class.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum NoteClass {
    /// `text:note-class="footnote"`.
    Footnote,
    /// `text:note-class="endnote"`.
    Endnote,
    /// A valid producer value not known to this version.
    Other(String),
}

impl NoteClass {
    /// The lexical ODF value for the two standard note classes.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Other(value) => value,
        }
    }
}

/// One inert footnote or endnote.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Note {
    body: String,
    citation: String,
    class: NoteClass,
    id: Option<String>,
    label: Option<String>,
    paragraphs: Vec<Paragraph>,
}

impl Note {
    pub(crate) const fn projected(
        class: NoteClass,
        id: Option<String>,
        label: Option<String>,
        citation: String,
        body: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            body,
            citation,
            class,
            id,
            label,
            paragraphs,
        }
    }

    /// Note class.
    #[must_use]
    pub const fn class(&self) -> &NoteClass {
        &self.class
    }

    /// Optional producer note identity.
    #[must_use]
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// Citation text displayed at the note anchor.
    #[must_use]
    pub fn citation(&self) -> &str {
        &self.citation
    }

    /// Optional explicit citation label.
    #[must_use]
    pub fn label(&self) -> Option<&str> {
        self.label.as_deref()
    }

    /// Inert visible note body text.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }

    /// Direct body paragraphs.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}
