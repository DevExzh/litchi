//! RTF list and numbering support.
//!
//! This module provides support for bulleted and numbered lists in RTF documents.
//! RTF uses a complex two-table system: list table and list override table.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_LISTS: usize = 65_536;
pub(crate) const MAX_LIST_LEVELS: usize = 9;
pub(crate) const MAX_LIST_TEXT_BYTES: usize = 65_536;
pub(crate) const MAX_LIST_TABS: usize = 64;

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
    /// A numbering format not represented by a named variant
    Other(i32),
}

/// Text emitted after a list label.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ListFollow {
    /// Follow the label with a tab.
    #[default]
    Tab,
    /// Follow the label with a space.
    Space,
    /// Do not emit a following character.
    Nothing,
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
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListLevel<'a> {
    /// Level number (0-8, where 0 is the top level)
    pub level: u8,
    /// Level type (bullet, decimal, etc.)
    pub level_type: ListLevelType,
    /// Number format text (template, e.g., "%1." for "1.", "%1.%2." for "1.1.")
    pub number_text: Cow<'a, str>,
    /// Raw decoded position bytes from `\levelnumbers`.
    pub number_positions: Cow<'a, str>,
    /// Start value for numbering
    pub start_at: i32,
    /// Justification
    pub justification: ListJustification,
    /// Whether to follow the previous level
    pub follow_previous: bool,
    /// Character emitted after the generated label
    pub follow: ListFollow,
    /// Font for the number/bullet
    pub font_ref: super::types::FontRef,
    /// Indentation for this level (in twips)
    pub indent: i32,
    /// Space before the number/bullet (in twips)
    pub space: i32,
    pub left_indent: Option<i32>,
    pub first_line_indent: Option<i32>,
    pub tabs: Vec<i32>,
    /// Zero-based inert reference into the list-picture destination.
    pub picture_index: Option<u32>,
    /// Whether this level is tentative (`\lvltentative`).
    pub tentative: bool,
    /// Convert smaller levels' numbers to legal (Arabic) format (`\levellegal`).
    pub legal_format: bool,
    /// Do not restart this level's number after higher-level items
    /// (`\levelnorestart`).
    pub no_restart: bool,
    /// Level retained for backward compatibility (`\levelold`).
    pub legacy: bool,
    /// Include the previous level's number in the display (`\levelprev`).
    pub include_previous: bool,
    /// Include a space after the previous level's number (`\levelprevspace`).
    pub include_previous_space: bool,
    /// Level template identifier (`\leveltemplateid`).
    pub template_id: Option<i32>,
}

impl<'a> ListLevel<'a> {
    /// Create a new list level
    #[inline]
    pub fn new(level: u8) -> Self {
        Self {
            level,
            level_type: ListLevelType::default(),
            number_text: Cow::Borrowed(""),
            number_positions: Cow::Borrowed(""),
            start_at: 1,
            justification: ListJustification::default(),
            follow_previous: false,
            follow: ListFollow::default(),
            font_ref: 0,
            indent: 0,
            space: 0,
            left_indent: None,
            first_line_indent: None,
            tabs: Vec::new(),
            picture_index: None,
            tentative: false,
            legal_format: false,
            no_restart: false,
            legacy: false,
            include_previous: false,
            include_previous_space: false,
            template_id: None,
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
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct List<'a> {
    /// Unique list identifier
    pub id: i32,
    /// List template ID
    pub template_id: i32,
    /// Whether this is a simple list (single level)
    pub simple: bool,
    /// Whether this is a hybrid multilevel list
    pub hybrid: bool,
    /// Optional list name
    pub name: Cow<'a, str>,
    pub style_name: Cow<'a, str>,
    pub style_priority: Option<i32>,
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
            hybrid: false,
            name: Cow::Borrowed(""),
            style_name: Cow::Borrowed(""),
            style_priority: None,
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
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListOverrideLevel {
    pub level: u8,
    pub start_at: Option<i32>,
    pub format_override: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListOverride {
    /// List override index
    pub index: i32,
    /// Original list ID this overrides
    pub list_id: i32,
    /// Override start value (if any)
    pub start_at_override: Option<i32>,
    /// Override level count (if any)
    pub level_count_override: Option<u8>,
    pub levels: Vec<ListOverrideLevel>,
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
            levels: Vec::new(),
        }
    }
}

/// List table containing all list definitions
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ListTable<'a> {
    /// List definitions
    lists: Vec<List<'a>>,
    /// Number of inert picture-bullet records in `\listpicture`.
    pub picture_bullet_count: u32,
    /// Optional indices into the document picture store, in `levelpicture` order.
    picture_bullet_picture_indices: Vec<Option<usize>>,
}

impl<'a> ListTable<'a> {
    /// Create a new list table
    #[inline]
    pub fn new() -> Self {
        Self {
            lists: Vec::new(),
            picture_bullet_count: 0,
            picture_bullet_picture_indices: Vec::new(),
        }
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

    /// Return retained picture-store indices for nonempty picture-bullet records.
    pub fn picture_bullet_picture_indices(&self) -> &[Option<usize>] {
        &self.picture_bullet_picture_indices
    }

    pub(crate) fn set_picture_bullet_picture_indices(
        &mut self,
        indices: Vec<Option<usize>>,
    ) -> RtfResult<()> {
        if indices.len() > 65_536 {
            return Err(RtfError::MalformedDocument(
                "RTF list-picture record count exceeds the safety limit".to_string(),
            ));
        }
        self.picture_bullet_count = u32::try_from(indices.len()).map_err(|_| {
            RtfError::MalformedDocument("RTF list-picture count overflow".to_string())
        })?;
        self.picture_bullet_picture_indices = indices;
        Ok(())
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.lists.len() > MAX_LISTS {
            return Err(RtfError::MalformedDocument(
                "RTF list count exceeds the safety limit".to_string(),
            ));
        }
        if self.picture_bullet_count > 65_536
            || self.picture_bullet_picture_indices.len() > self.picture_bullet_count as usize
        {
            return Err(RtfError::MalformedDocument(
                "invalid or oversized RTF list-picture table".to_string(),
            ));
        }
        let mut ids = HashSet::with_capacity(self.lists.len());
        for list in &self.lists {
            if !ids.insert(list.id) {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF list ID".to_string(),
                ));
            }
            if list.levels.is_empty()
                || list.levels.len() > MAX_LIST_LEVELS
                || (list.simple && (list.hybrid || list.levels.len() > 1))
                || list.name.len() > MAX_LIST_TEXT_BYTES
                || list.style_name.len() > MAX_LIST_TEXT_BYTES
                || list
                    .style_priority
                    .is_some_and(|value| !(0..=99).contains(&value))
            {
                return Err(RtfError::MalformedDocument(
                    "invalid or oversized RTF list definition".to_string(),
                ));
            }
            let mut levels = HashSet::new();
            for level in &list.levels {
                if level.level > 8
                    || !levels.insert(level.level)
                    || level.number_text.len() > MAX_LIST_TEXT_BYTES
                    || level.number_positions.len() > MAX_LIST_TEXT_BYTES
                    || level.tabs.len() > MAX_LIST_TABS
                    || level
                        .picture_index
                        .is_some_and(|index| index >= self.picture_bullet_count)
                {
                    return Err(RtfError::MalformedDocument(
                        "invalid or oversized RTF list level".to_string(),
                    ));
                }
            }
        }
        Ok(())
    }
}

/// List override table containing list instances
#[derive(Debug, Clone, Default, PartialEq, Eq)]
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

    pub fn resolve<'a>(
        &'a self,
        index: i32,
        lists: &'a ListTable<'a>,
    ) -> Option<(&'a ListOverride, &'a List<'a>)> {
        let entry = self.get(index)?;
        Some((entry, lists.get(entry.list_id)?))
    }

    pub(crate) fn validate(&self, lists: &ListTable<'_>) -> RtfResult<()> {
        if self.overrides.len() > MAX_LISTS {
            return Err(RtfError::MalformedDocument(
                "RTF list override count exceeds the safety limit".to_string(),
            ));
        }
        // Partial documents can retain the complete override table while
        // retaining only a contiguous suffix of the corresponding definitions.
        // Accept precisely that shape; isolated and interleaved dangling IDs
        // remain malformed.
        let first_resolved = self
            .overrides
            .iter()
            .position(|entry| lists.get(entry.list_id).is_some());
        let partial_definition_suffix = first_resolved.is_some_and(|first| {
            first > 0
                && self.overrides.len() - first == lists.lists().len()
                && self.overrides[..first]
                    .iter()
                    .all(|entry| lists.get(entry.list_id).is_none())
                && self.overrides[first..]
                    .iter()
                    .all(|entry| lists.get(entry.list_id).is_some())
                && self
                    .overrides
                    .windows(2)
                    .all(|pair| pair[1].index == pair[0].index.saturating_add(1))
        });
        let mut indices = HashSet::with_capacity(self.overrides.len());
        for entry in &self.overrides {
            if !indices.insert(entry.index)
                || (lists.get(entry.list_id).is_none() && !partial_definition_suffix)
                || entry.levels.len() > MAX_LIST_LEVELS
                || entry
                    .level_count_override
                    .is_some_and(|count| usize::from(count) != entry.levels.len())
            {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF list override definition".to_string(),
                ));
            }
            if entry
                .levels
                .iter()
                .enumerate()
                .any(|(index, level)| level.level as usize != index)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF list override levels are out of order".to_string(),
                ));
            }
        }
        Ok(())
    }
}
