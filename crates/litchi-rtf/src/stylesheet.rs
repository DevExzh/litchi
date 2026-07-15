//! RTF stylesheet support.
//!
//! This module provides support for RTF stylesheets and style definitions.

use super::types::{Formatting, Paragraph};
use std::borrow::Cow;

/// Style type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
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

/// RTF style definition
#[derive(Debug, Clone)]
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
#[derive(Debug, Clone, Default)]
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
}
