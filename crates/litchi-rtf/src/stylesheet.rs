//! RTF stylesheet support.
//!
//! This module provides support for RTF stylesheets and style definitions.

use super::types::{Formatting, Paragraph};
use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_STYLES: usize = 65_536;
pub(crate) const MAX_STYLE_NAME_BYTES: usize = 65_536;

/// Style type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum StyleType {
    /// Paragraph style
    #[default]
    Paragraph,
    /// Character style
    Character,
    /// Section style
    Section,
    /// Table style
    Table,
}

/// Inert conditional-formatting scope and banding metadata of a table style
/// definition (`\tsN`).
///
/// RTF 1.9.1 table-style definitions may declare which table regions the
/// conditional formatting targets (`\tscfirstrow`, `\tsclastrow`,
/// `\tscfirstcol`, `\tsclastcol`), odd/even banding scopes
/// (`\tscbandhorzodd`, `\tscbandhorzeven`, `\tscbandvertodd`,
/// `\tscbandverteven`), and band sizes (`\tscbandshN`, `\tscbandsvN`); the
/// `\tsrowd` marker closes the row-defaults portion of the definition.
///
/// The flags are passive metadata only: no conditional formatting is ever
/// evaluated or applied.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableStyleConditionalFormatting {
    /// Whether the definition carries the `\tsrowd` row-defaults marker.
    pub row_defaults_marker: bool,
    /// Conditional formatting targets the first row (`\tscfirstrow`).
    pub first_row: bool,
    /// Conditional formatting targets the last row (`\tsclastrow`).
    pub last_row: bool,
    /// Conditional formatting targets the first column (`\tscfirstcol`).
    pub first_column: bool,
    /// Conditional formatting targets the last column (`\tsclastcol`).
    pub last_column: bool,
    /// Odd horizontal banding scope (`\tscbandhorzodd`).
    pub band_horizontal_odd: bool,
    /// Even horizontal banding scope (`\tscbandhorzeven`).
    pub band_horizontal_even: bool,
    /// Odd vertical banding scope (`\tscbandvertodd`).
    pub band_vertical_odd: bool,
    /// Even vertical banding scope (`\tscbandverteven`).
    pub band_vertical_even: bool,
    /// Horizontal band size in rows (`\tscbandshN`).
    pub horizontal_band_size: Option<u16>,
    /// Vertical band size in columns (`\tscbandsvN`).
    pub vertical_band_size: Option<u16>,
}

impl TableStyleConditionalFormatting {
    /// Whether no conditional-formatting metadata is present.
    #[inline]
    pub const fn is_empty(&self) -> bool {
        !self.row_defaults_marker
            && !self.first_row
            && !self.last_row
            && !self.first_column
            && !self.last_column
            && !self.band_horizontal_odd
            && !self.band_horizontal_even
            && !self.band_vertical_odd
            && !self.band_vertical_even
            && self.horizontal_band_size.is_none()
            && self.vertical_band_size.is_none()
    }
}

/// RTF style definition
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style<'a> {
    /// Style index/ID
    pub id: u16,
    /// Style name
    pub name: Cow<'a, str>,
    /// Style type
    pub style_type: StyleType,
    /// Based-on style ID (parent style)
    pub based_on: Option<u16>,
    /// Next style ID (style for next paragraph)
    pub next_style: Option<u16>,
    /// Linked paragraph or character style ID
    pub linked_style: Option<u16>,
    /// Character formatting
    pub formatting: Formatting,
    /// Paragraph properties (for paragraph styles)
    pub paragraph: Option<Paragraph>,
    /// Table-style conditional formatting metadata (`\tsN` styles only).
    pub table_conditional: TableStyleConditionalFormatting,
    /// Whether this is a built-in style
    pub builtin: bool,
    /// Whether this style is hidden
    pub hidden: bool,
    /// Whether character properties are additive to the surrounding formatting
    pub additive: bool,
    /// Whether applications may automatically update the style definition
    pub auto_update: bool,
    /// Whether the style is locked against modification
    pub locked: bool,
    /// Whether the style is hidden until it is used
    pub semi_hidden: bool,
    /// Whether the style should become visible after it is used
    pub unhide_when_used: bool,
    /// Whether the style appears in the quick-style gallery
    pub quick_format: bool,
    /// Style UI sorting priority
    pub priority: Option<i32>,
    /// Style revision identifier
    pub revision_id: Option<i32>,
    /// Whether this is a personal e-mail style
    pub personal: bool,
    /// Whether this is an e-mail composition style
    pub compose: bool,
    /// Whether this is an e-mail reply style
    pub reply: bool,
}

impl<'a> Style<'a> {
    fn new(
        id: u16,
        name: Cow<'a, str>,
        style_type: StyleType,
        paragraph: Option<Paragraph>,
    ) -> Self {
        Self {
            id,
            name,
            style_type,
            based_on: None,
            next_style: None,
            linked_style: None,
            formatting: Formatting::default(),
            paragraph,
            table_conditional: TableStyleConditionalFormatting::default(),
            builtin: false,
            hidden: false,
            additive: false,
            auto_update: false,
            locked: false,
            semi_hidden: false,
            unhide_when_used: false,
            quick_format: false,
            priority: None,
            revision_id: None,
            personal: false,
            compose: false,
            reply: false,
        }
    }

    /// Create a new paragraph style
    #[inline]
    pub fn paragraph(id: u16, name: Cow<'a, str>) -> Self {
        Self::new(id, name, StyleType::Paragraph, Some(Paragraph::default()))
    }

    /// Create a new character style
    #[inline]
    pub fn character(id: u16, name: Cow<'a, str>) -> Self {
        Self::new(id, name, StyleType::Character, None)
    }

    /// Create a new section style.
    #[inline]
    pub fn section(id: u16, name: Cow<'a, str>) -> Self {
        Self::new(id, name, StyleType::Section, None)
    }

    /// Create a new table style.
    #[inline]
    pub fn table(id: u16, name: Cow<'a, str>) -> Self {
        Self::new(id, name, StyleType::Table, None)
    }

    /// Check if this is a paragraph style
    #[inline]
    pub fn is_paragraph_style(&self) -> bool {
        self.style_type == StyleType::Paragraph
    }

    /// Check if this is a character style
    #[inline]
    pub fn is_character_style(&self) -> bool {
        self.style_type == StyleType::Character
    }
}

/// Stylesheet containing all style definitions
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleSheet<'a> {
    /// Style definitions
    styles: Vec<Style<'a>>,
}

impl<'a> StyleSheet<'a> {
    /// Create a new stylesheet
    #[inline]
    pub fn new() -> Self {
        Self { styles: Vec::new() }
    }

    /// Add a style to the stylesheet
    #[inline]
    pub fn add(&mut self, style: Style<'a>) {
        self.styles.push(style);
    }

    /// Get a style by ID
    pub fn get(&self, id: u16) -> Option<&Style<'a>> {
        self.styles.iter().find(|s| s.id == id)
    }

    /// Get a style by type and ID.
    pub fn get_typed(&self, style_type: StyleType, id: u16) -> Option<&Style<'a>> {
        self.styles
            .iter()
            .find(|style| style.style_type == style_type && style.id == id)
    }

    /// Get a style by name
    pub fn get_by_name(&self, name: &str) -> Option<&Style<'a>> {
        self.styles.iter().find(|s| s.name.as_ref() == name)
    }

    /// Get all styles
    #[inline]
    pub fn styles(&self) -> &[Style<'a>] {
        &self.styles
    }

    /// Get all paragraph styles
    pub fn paragraph_styles(&self) -> Vec<&Style<'a>> {
        self.styles
            .iter()
            .filter(|s| s.is_paragraph_style())
            .collect()
    }

    /// Get all character styles
    pub fn character_styles(&self) -> Vec<&Style<'a>> {
        self.styles
            .iter()
            .filter(|s| s.is_character_style())
            .collect()
    }

    /// Return the based-on chain from the root ancestor to the selected style.
    ///
    /// Raw definitions remain unchanged so explicit resets survive writing.
    pub fn inheritance_chain(&self, style_type: StyleType, id: u16) -> RtfResult<Vec<&Style<'a>>> {
        let mut chain = Vec::new();
        let mut seen = HashSet::new();
        let mut current = self.get_typed(style_type, id);
        while let Some(style) = current {
            if !seen.insert((style.style_type, style.id)) {
                return Err(RtfError::MalformedDocument(
                    "RTF stylesheet contains a based-on cycle".to_string(),
                ));
            }
            chain.push(style);
            current = style
                .based_on
                .and_then(|parent| self.get_typed(style.style_type, parent));
        }
        chain.reverse();
        Ok(chain)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.styles.len() > MAX_STYLES {
            return Err(RtfError::MalformedDocument(
                "RTF stylesheet exceeds the supported style count".to_string(),
            ));
        }
        let mut ids = HashSet::with_capacity(self.styles.len());
        for style in &self.styles {
            if !ids.insert((style.style_type, style.id)) {
                return Err(RtfError::MalformedDocument(
                    "RTF stylesheet contains a duplicate typed style ID".to_string(),
                ));
            }
            if style.name.is_empty()
                || style.name.len() > MAX_STYLE_NAME_BYTES
                || style.name.contains(';')
            {
                return Err(RtfError::MalformedDocument(
                    "RTF style name is empty, too long, or contains a semicolon".to_string(),
                ));
            }
            if style
                .priority
                .is_some_and(|value| !(0..=99).contains(&value))
            {
                return Err(RtfError::MalformedDocument(
                    "RTF style priority must be in 0..=99".to_string(),
                ));
            }
            if style.revision_id.is_some_and(|value| value < 0) {
                return Err(RtfError::MalformedDocument(
                    "RTF style revision ID cannot be negative".to_string(),
                ));
            }
            self.inheritance_chain(style.style_type, style.id)?;
        }
        Ok(())
    }
}
