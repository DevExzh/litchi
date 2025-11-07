//! RTF list and numbering support.
//!
//! This module provides support for bulleted and numbered lists in RTF documents.
//! RTF uses a complex two-table system: list table and list override table.

use std::borrow::Cow;

/// List level type (bullet or numbered)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ListLevelType {
    /// Arabic numerals (1, 2, 3...)
    Decimal,
    /// Uppercase Roman numerals (I, II, III...)
    UpperRoman,
    /// Lowercase Roman numerals (i, ii, iii...)
    LowerRoman,
    /// Uppercase letters (A, B, C...)
    UpperLetter,
    /// Lowercase letters (a, b, c...)
    LowerLetter,
    /// Ordinal numbers (1st, 2nd, 3rd...)
    Ordinal,
    /// Cardinal text (One, Two, Three...)
    CardinalText,
    /// Ordinal text (First, Second, Third...)
    OrdinalText,
    /// Bullet (•, ○, ■, etc.)
    #[default]
    Bullet,
    /// No numbering
    None,
}

/// List level justification
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ListJustification {
    /// Left-aligned
    #[default]
    Left,
    /// Right-aligned
    Right,
    /// Centered
    Center,
}

/// A single level in a list (for multi-level lists)
#[derive(Debug, Clone)]
pub struct ListLevel<'a> {
    /// Level number (0-8, where 0 is the top level)
    pub level: u8,
    /// Level type (bullet, decimal, etc.)
    pub level_type: ListLevelType,
    /// Number format text (template, e.g., "%1." for "1.", "%1.%2." for "1.1.")
    pub number_text: Cow<'a, str>,
    /// Start value for numbering
    pub start_at: i32,
    /// Justification
    pub justification: ListJustification,
    /// Whether to follow the previous level
    pub follow_previous: bool,
    /// Font for the number/bullet
    pub font_ref: super::types::FontRef,
    /// Indentation for this level (in twips)
    pub indent: i32,
    /// Space before the number/bullet (in twips)
    pub space: i32,
}

impl<'a> ListLevel<'a> {
    /// Create a new list level
    #[inline]
    pub fn new(level: u8) -> Self {
        Self {
            level,
            level_type: ListLevelType::default(),
            number_text: Cow::Borrowed(""),
            start_at: 1,
            justification: ListJustification::default(),
            follow_previous: false,
            font_ref: 0,
            indent: 0,
            space: 0,
        }
    }

    /// Check if this level is a bullet
    #[inline]
    pub fn is_bullet(&self) -> bool {
        matches!(self.level_type, ListLevelType::Bullet)
    }

    /// Check if this level is numbered
    #[inline]
    pub fn is_numbered(&self) -> bool {
        !self.is_bullet() && self.level_type != ListLevelType::None
    }
}

impl<'a> Default for ListLevel<'a> {
    fn default() -> Self {
        Self::new(0)
    }
}

/// RTF list definition
#[derive(Debug, Clone)]
pub struct List<'a> {
    /// Unique list identifier
    pub id: i32,
    /// List template ID
    pub template_id: i32,
    /// Whether this is a simple list (single level)
    pub simple: bool,
    /// List levels (up to 9 levels)
    pub levels: Vec<ListLevel<'a>>,
}

impl<'a> List<'a> {
    /// Create a new list
    #[inline]
    pub fn new(id: i32) -> Self {
        Self {
            id,
            template_id: id,
            simple: true,
            levels: Vec::new(),
        }
    }

    /// Add a level to the list
    #[inline]
    pub fn add_level(&mut self, level: ListLevel<'a>) {
        self.levels.push(level);
    }

    /// Get a level by index
    #[inline]
    pub fn get_level(&self, level: u8) -> Option<&ListLevel<'a>> {
        self.levels.iter().find(|l| l.level == level)
    }

    /// Get the number of levels
    #[inline]
    pub fn level_count(&self) -> usize {
        self.levels.len()
    }
}

/// List override entry (instance of a list)
#[derive(Debug, Clone)]
pub struct ListOverride {
    /// List override index
    pub index: i32,
    /// Original list ID this overrides
    pub list_id: i32,
    /// Override start value (if any)
    pub start_at_override: Option<i32>,
    /// Override level count (if any)
    pub level_count_override: Option<u8>,
}

impl ListOverride {
    /// Create a new list override
    #[inline]
    pub fn new(index: i32, list_id: i32) -> Self {
        Self {
            index,
            list_id,
            start_at_override: None,
            level_count_override: None,
        }
    }
}

/// List table containing all list definitions
#[derive(Debug, Clone, Default)]
pub struct ListTable<'a> {
    /// List definitions
    lists: Vec<List<'a>>,
}

impl<'a> ListTable<'a> {
    /// Create a new list table
    #[inline]
    pub fn new() -> Self {
        Self { lists: Vec::new() }
    }

    /// Add a list to the table
    #[inline]
    pub fn add(&mut self, list: List<'a>) {
        self.lists.push(list);
    }

    /// Get a list by ID
    #[inline]
    pub fn get(&self, id: i32) -> Option<&List<'a>> {
        self.lists.iter().find(|l| l.id == id)
    }

    /// Get all lists
    #[inline]
    pub fn lists(&self) -> &[List<'a>] {
        &self.lists
    }
}

/// List override table containing list instances
#[derive(Debug, Clone, Default)]
pub struct ListOverrideTable {
    /// List overrides
    overrides: Vec<ListOverride>,
}

impl ListOverrideTable {
    /// Create a new list override table
    #[inline]
    pub fn new() -> Self {
        Self {
            overrides: Vec::new(),
        }
    }

    /// Add a list override
    #[inline]
    pub fn add(&mut self, override_entry: ListOverride) {
        self.overrides.push(override_entry);
    }

    /// Get a list override by index
    #[inline]
    pub fn get(&self, index: i32) -> Option<&ListOverride> {
        self.overrides.iter().find(|o| o.index == index)
    }

    /// Get all overrides
    #[inline]
    pub fn overrides(&self) -> &[ListOverride] {
        &self.overrides
    }
}
