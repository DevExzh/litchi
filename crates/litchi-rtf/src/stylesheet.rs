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
    /// Character formatting
    pub formatting: Formatting,
    /// Paragraph properties (for paragraph styles)
    pub paragraph: Option<Paragraph>,
    /// Whether this is a built-in style
    pub builtin: bool,
    /// Whether this style is hidden
    pub hidden: bool,
}

impl<'a> Style<'a> {
    /// Create a new paragraph style
    #[inline]
    pub fn paragraph(id: u16, name: Cow<'a, str>) -> Self {
        Self {
            id,
            name,
            style_type: StyleType::Paragraph,
            based_on: None,
            next_style: None,
            formatting: Formatting::default(),
            paragraph: Some(Paragraph::default()),
            builtin: false,
            hidden: false,
        }
    }

    /// Create a new character style
    #[inline]
    pub fn character(id: u16, name: Cow<'a, str>) -> Self {
        Self {
            id,
            name,
            style_type: StyleType::Character,
            based_on: None,
            next_style: None,
            formatting: Formatting::default(),
            paragraph: None,
            builtin: false,
            hidden: false,
        }
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
