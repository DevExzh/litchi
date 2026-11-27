//! Semantic model and lexical validation for ODF note configurations.

use litchi_core::Result;

use super::{MAX_VALUE_BYTES, invalid};
use crate::line_numbering::Format;

/// Note class selected by `text:note-class`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Class {
    Footnote,
    Endnote,
}

impl Class {
    pub const ALL: [Self; 2] = [Self::Footnote, Self::Endnote];

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "footnote" => Ok(Self::Footnote),
            "endnote" => Ok(Self::Endnote),
            _ => invalid(format!("unsupported text:note-class '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }
}

/// Scope at which note numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberingScope {
    Document,
    Chapter,
    Page,
}

impl NumberingScope {
    pub const ALL: [Self; 3] = [Self::Document, Self::Chapter, Self::Page];

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "document" => Ok(Self::Document),
            "chapter" => Ok(Self::Chapter),
            "page" => Ok(Self::Page),
            _ => invalid(format!("unsupported text:start-numbering-at '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Chapter => "chapter",
            Self::Page => "page",
        }
    }
}

/// Placement of footnotes in the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Position {
    Text,
    Page,
    Section,
    Document,
}

impl Position {
    pub const ALL: [Self; 4] = [Self::Text, Self::Page, Self::Section, Self::Document];

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "text" => Ok(Self::Text),
            "page" => Ok(Self::Page),
            "section" => Ok(Self::Section),
            "document" => Ok(Self::Document),
            _ => invalid(format!("unsupported text:footnotes-position '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Page => "page",
            Self::Section => "section",
            Self::Document => "document",
        }
    }
}

/// One `text:notes-configuration` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Configuration {
    pub note_class: Class,
    pub citation_style_name: Option<String>,
    pub citation_body_style_name: Option<String>,
    pub default_style_name: Option<String>,
    pub master_page_name: Option<String>,
    pub start_value: Option<u64>,
    pub number_prefix: Option<String>,
    pub number_suffix: Option<String>,
    pub number_format: Option<Format>,
    pub letter_sync: Option<bool>,
    pub start_numbering_at: Option<NumberingScope>,
    pub footnotes_position: Option<Position>,
    pub continuation_notice_forward: Option<String>,
    pub continuation_notice_backward: Option<String>,
}

impl Configuration {
    pub fn new(note_class: Class) -> Self {
        Self {
            note_class,
            citation_style_name: None,
            citation_body_style_name: None,
            default_style_name: None,
            master_page_name: None,
            start_value: None,
            number_prefix: None,
            number_suffix: None,
            number_format: None,
            letter_sync: None,
            start_numbering_at: None,
            footnotes_position: None,
            continuation_notice_forward: None,
            continuation_notice_backward: None,
        }
    }

    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !matches!(
                self.number_format,
                Some(Format::LowerAlpha | Format::UpperAlpha)
            )
        {
            return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
        }
        for (value, name) in [
            (
                self.citation_style_name.as_deref(),
                "text:citation-style-name",
            ),
            (
                self.citation_body_style_name.as_deref(),
                "text:citation-body-style-name",
            ),
            (
                self.default_style_name.as_deref(),
                "text:default-style-name",
            ),
            (self.master_page_name.as_deref(), "text:master-page-name"),
        ] {
            if let Some(value) = value {
                validate_style_name_ref(value, name)?;
            }
        }
        for (value, name, allow_empty) in [
            (self.number_prefix.as_deref(), "style:num-prefix", true),
            (self.number_suffix.as_deref(), "style:num-suffix", true),
            (
                self.continuation_notice_forward.as_deref(),
                "text:note-continuation-notice-forward",
                true,
            ),
            (
                self.continuation_notice_backward.as_deref(),
                "text:note-continuation-notice-backward",
                true,
            ),
        ] {
            if let Some(value) = value {
                validate_value(value, name, allow_empty)?;
            }
        }
        if let Some(format) = &self.number_format {
            validate_value(format.as_str(), "style:num-format", true)?;
        }
        Ok(())
    }
}

/// The at-most-one configuration for each standard note class.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Configurations {
    pub footnote: Option<Configuration>,
    pub endnote: Option<Configuration>,
}

impl Configurations {
    pub fn get(&self, note_class: Class) -> Option<&Configuration> {
        match note_class {
            Class::Footnote => self.footnote.as_ref(),
            Class::Endnote => self.endnote.as_ref(),
        }
    }

    pub fn validate(&self) -> Result<()> {
        if let Some(configuration) = &self.footnote {
            if configuration.note_class != Class::Footnote {
                return invalid("footnote slot contains an endnote configuration");
            }
            configuration.validate()?;
        }
        if let Some(configuration) = &self.endnote {
            if configuration.note_class != Class::Endnote {
                return invalid("endnote slot contains a footnote configuration");
            }
            configuration.validate()?;
        }
        Ok(())
    }
}

pub(super) fn validate_value(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

fn ncname_start(character: char) -> bool {
    matches!(character,
        'A'..='Z' | '_' | 'a'..='z'
        | '\u{c0}'..='\u{d6}' | '\u{d8}'..='\u{f6}' | '\u{f8}'..='\u{2ff}'
        | '\u{370}'..='\u{37d}' | '\u{37f}'..='\u{1fff}' | '\u{200c}'..='\u{200d}'
        | '\u{2070}'..='\u{218f}' | '\u{2c00}'..='\u{2fef}' | '\u{3001}'..='\u{d7ff}'
        | '\u{f900}'..='\u{fdcf}' | '\u{fdf0}'..='\u{fffd}' | '\u{10000}'..='\u{effff}'
    )
}

fn ncname_continue(character: char) -> bool {
    ncname_start(character)
        || matches!(character, '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}')
}

fn validate_style_name_ref(value: &str, name: &str) -> Result<()> {
    validate_value(value, name, true)?;
    if value.is_empty() {
        return Ok(());
    }
    let mut characters = value.chars();
    if !characters.next().is_some_and(ncname_start) || !characters.all(ncname_continue) {
        return invalid(format!("{name} must be an NCName or empty"));
    }
    Ok(())
}
