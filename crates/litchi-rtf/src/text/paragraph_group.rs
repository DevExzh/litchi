//! Paragraph-group property table (`pgptbl`) metadata.

use crate::{Borders, RtfError, RtfResult};
use std::collections::HashSet;

pub(crate) const MAX_PARAGRAPH_GROUP_PROPERTIES: usize = 4_096;
const MAX_LAYOUT_TWIPS: i32 = 10_000_000;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphGroupProperty {
    /// Implicit one-based ID assigned by table order.
    pub id: u32,
    /// Parent paragraph-group ID (`0` means no parent).
    pub parent_id: u32,
    /// Nested-table depth from `itap`.
    pub table_nesting_level: u8,
    pub left_indent: i32,
    pub right_indent: i32,
    pub space_before: i32,
    pub space_after: i32,
    pub borders: Borders,
}

impl ParagraphGroupProperty {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.id == 0 {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group IDs are one-based".to_string(),
            ));
        }
        if self.table_nesting_level > 63 {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group table nesting exceeds the safety limit".to_string(),
            ));
        }
        if [
            self.left_indent,
            self.right_indent,
            self.space_before,
            self.space_after,
        ]
        .into_iter()
        .any(|value| value.unsigned_abs() > MAX_LAYOUT_TWIPS as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group layout value exceeds the safety limit".to_string(),
            ));
        }
        for border in [
            self.borders.top,
            self.borders.bottom,
            self.borders.left,
            self.borders.right,
        ] {
            if border.width < 0
                || border.width > 10_000
                || border.space < 0
                || border.space > MAX_LAYOUT_TWIPS
            {
                return Err(RtfError::MalformedDocument(
                    "RTF paragraph-group border value exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphGroupPropertyTable {
    entries: Vec<ParagraphGroupProperty>,
}

impl ParagraphGroupPropertyTable {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn entries(&self) -> &[ParagraphGroupProperty] {
        &self.entries
    }

    #[must_use]
    pub fn get(&self, id: u32) -> Option<&ParagraphGroupProperty> {
        id.checked_sub(1)
            .and_then(|index| self.entries.get(index as usize))
    }

    #[must_use]
    pub fn parent_of(&self, id: u32) -> Option<&ParagraphGroupProperty> {
        let parent_id = self.get(id)?.parent_id;
        (parent_id != 0).then(|| self.get(parent_id)).flatten()
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push(&mut self, entry: ParagraphGroupProperty) -> RtfResult<()> {
        if self.entries.len() >= MAX_PARAGRAPH_GROUP_PROPERTIES {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group table exceeds the entry limit".to_string(),
            ));
        }
        if entry.id as usize != self.entries.len() + 1 {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group IDs are out of order".to_string(),
            ));
        }
        entry.validate()?;
        self.entries.push(entry);
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.entries.is_empty() || self.entries.len() > MAX_PARAGRAPH_GROUP_PROPERTIES {
            return Err(RtfError::MalformedDocument(
                "RTF paragraph-group table has an invalid entry count".to_string(),
            ));
        }
        for (index, entry) in self.entries.iter().enumerate() {
            entry.validate()?;
            if entry.id as usize != index + 1 || entry.parent_id as usize > self.entries.len() {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF paragraph-group ID or parent reference".to_string(),
                ));
            }
        }
        for entry in &self.entries {
            let mut visited = HashSet::new();
            let mut current = entry.id;
            while current != 0 {
                if !visited.insert(current) {
                    return Err(RtfError::MalformedDocument(
                        "RTF paragraph-group parent references contain a cycle".to_string(),
                    ));
                }
                current = self
                    .get(current)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "invalid RTF paragraph-group parent reference".to_string(),
                        )
                    })?
                    .parent_id;
            }
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> Self {
        self
    }
}
