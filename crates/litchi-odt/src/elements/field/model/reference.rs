//! Reference, measure, and document-statistic semantics.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
/// Display mode permitted by `text:user-field-get`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UserFieldDisplay {
    Value,
    Formula,
    None,
}

/// Component displayed by a `text:measure` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MeasureKind {
    Value,
    Unit,
    Gap,
}

/// Display format shared by `text:reference-ref` and `text:bookmark-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CrossReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
    NumberNoSuperior,
    NumberAllSuperior,
    Number,
}

impl CrossReferenceFormat {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Chapter => "chapter",
            Self::Direction => "direction",
            Self::Text => "text",
            Self::NumberNoSuperior => "number-no-superior",
            Self::NumberAllSuperior => "number-all-superior",
            Self::Number => "number",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "chapter" => Ok(Self::Chapter),
            "direction" => Ok(Self::Direction),
            "text" => Ok(Self::Text),
            "number-no-superior" => Ok(Self::NumberNoSuperior),
            "number-all-superior" => Ok(Self::NumberAllSuperior),
            "number" => Ok(Self::Number),
            _ => Err(Error::InvalidFormat(format!(
                "invalid cross-reference format '{value}'"
            ))),
        }
    }
}

/// Display format permitted by `text:note-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NoteReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
}

impl NoteReferenceFormat {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Chapter => "chapter",
            Self::Direction => "direction",
            Self::Text => "text",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "chapter" => Ok(Self::Chapter),
            "direction" => Ok(Self::Direction),
            "text" => Ok(Self::Text),
            _ => Err(Error::InvalidFormat(format!(
                "invalid note reference format '{value}'"
            ))),
        }
    }
}

/// Note class selected by `text:note-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NoteReferenceClass {
    Footnote,
    Endnote,
}

/// Kind of cached ODF document statistic.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum StatisticKind {
    Page,
    Paragraph,
    Word,
    Character,
    Table,
    Image,
    Object,
}

impl StatisticKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::Page => "text:page-count",
            Self::Paragraph => "text:paragraph-count",
            Self::Word => "text:word-count",
            Self::Character => "text:character-count",
            Self::Table => "text:table-count",
            Self::Image => "text:image-count",
            Self::Object => "text:object-count",
        }
    }
}

impl NoteReferenceClass {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "footnote" => Ok(Self::Footnote),
            "endnote" => Ok(Self::Endnote),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:note-class '{value}'"
            ))),
        }
    }
}

impl MeasureKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Unit => "unit",
            Self::Gap => "gap",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "unit" => Ok(Self::Unit),
            "gap" => Ok(Self::Gap),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:measure kind '{value}'"
            ))),
        }
    }
}

impl UserFieldDisplay {
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Formula => "formula",
            Self::None => "none",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "formula" => Ok(Self::Formula),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid user-field-get text:display '{value}'"
            ))),
        }
    }
}
