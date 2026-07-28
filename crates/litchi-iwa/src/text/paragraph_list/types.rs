//! Strict list presets and nesting levels shared by Pages, Numbers, and Keynote.

use super::super::drop_cap::ParagraphStart;
use crate::{Error, Result};

const MAX_PARAGRAPH_LIST_LEVEL: u8 = 8;

/// A canonical paragraph-list presentation understood by all three iWork apps.
///
/// The presets describe the complete nine-level native list style rather than
/// exposing unvalidated protobuf integers or partial per-level state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphList {
    /// Ordinary paragraphs without labels.
    #[default]
    None,
    /// Apple’s standard bullet preset using the `•` marker.
    Bullet,
    /// Apple’s standard decimal-number preset.
    Numbered,
}

/// One list preset boundary at a validated UTF-16 paragraph start.
///
/// The preset remains effective until the next placement. A complete placement
/// list always begins at [`ParagraphStart::ZERO`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListPlacement {
    pub paragraph: ParagraphStart,
    pub list: ParagraphList,
}

impl ParagraphListPlacement {
    pub const fn new(paragraph: ParagraphStart, list: ParagraphList) -> Self {
        Self { paragraph, list }
    }
}

impl ParagraphList {
    pub(crate) const fn native_name(self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Bullet => "Bullet",
            Self::Numbered => "Numbered",
        }
    }

    pub(crate) const fn preset_index(self) -> usize {
        match self {
            Self::None => 0,
            Self::Bullet => 1,
            Self::Numbered => 2,
        }
    }
}

/// A zero-based nesting level in iWork's nine-level paragraph-list model.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct ParagraphListLevel(u8);

impl ParagraphListLevel {
    /// Top-level list item.
    pub const ZERO: Self = Self(0);
    /// First nested list level.
    pub const ONE: Self = Self(1);
    /// Deepest level supported by the native nine-level style model.
    pub const MAX: Self = Self(MAX_PARAGRAPH_LIST_LEVEL);

    /// Construct a validated zero-based nesting level.
    pub fn new(level: u8) -> Result<Self> {
        if level > MAX_PARAGRAPH_LIST_LEVEL {
            return Err(Error::InvalidFormat(format!(
                "paragraph list level must not exceed {MAX_PARAGRAPH_LIST_LEVEL}"
            )));
        }
        Ok(Self(level))
    }

    /// Return the zero-based native nesting level.
    pub const fn get(self) -> u8 {
        self.0
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        u8::try_from(value)
            .map_err(|_| {
                Error::InvalidFormat(format!("native paragraph list level {value} exceeds u8"))
            })
            .and_then(Self::new)
    }
}

/// One effective list-level boundary at a UTF-16 paragraph start.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListLevelPlacement {
    pub paragraph: ParagraphStart,
    pub level: ParagraphListLevel,
}

impl ParagraphListLevelPlacement {
    pub const fn new(paragraph: ParagraphStart, level: ParagraphListLevel) -> Self {
        Self { paragraph, level }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn list_levels_are_bounded_by_the_native_style_model() {
        assert_eq!(
            ParagraphListLevel::new(0).unwrap(),
            ParagraphListLevel::ZERO
        );
        assert_eq!(ParagraphListLevel::new(1).unwrap(), ParagraphListLevel::ONE);
        assert_eq!(ParagraphListLevel::new(8).unwrap(), ParagraphListLevel::MAX);
        assert!(ParagraphListLevel::new(9).is_err());
    }
}
