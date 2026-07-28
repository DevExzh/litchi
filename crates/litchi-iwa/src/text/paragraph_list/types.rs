//! Strict list presets and nesting levels shared by Pages, Numbers, and Keynote.

use super::super::drop_cap::ParagraphStart;
use crate::{Error, Result};

const MAX_PARAGRAPH_LIST_LEVEL: u8 = 8;
const MAX_PARAGRAPH_LIST_BULLET_CHARACTERS: usize = 32;

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

/// A positive starting number for a restarted numbered list.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ParagraphListStart(u32);

impl ParagraphListStart {
    /// The first positive list number.
    pub const ONE: Self = Self(1);

    /// Construct a validated positive starting number.
    pub fn new(number: u32) -> Result<Self> {
        if number == 0 {
            return Err(Error::InvalidFormat(
                "paragraph list starting number must be positive".to_owned(),
            ));
        }
        Ok(Self(number))
    }

    /// Return the native positive starting number.
    pub const fn get(self) -> u32 {
        self.0
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// How one numbered-list paragraph participates in the current sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphListNumbering {
    /// Continue the preceding numbered sequence.
    #[default]
    Continue,
    /// Restart numbering at the supplied positive number.
    StartAt(ParagraphListStart),
}

impl ParagraphListNumbering {
    pub(crate) const fn native_start(self) -> u32 {
        match self {
            Self::Continue => 0,
            Self::StartAt(start) => start.get(),
        }
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        if value == 0 {
            Ok(Self::Continue)
        } else {
            ParagraphListStart::from_native(value).map(Self::StartAt)
        }
    }
}

/// A validated text marker used by one bullet-list level.
///
/// iWork accepts a short sequence of printable characters, not just a single
/// Unicode scalar. Newlines and control characters are rejected because they
/// cannot be represented by the native bullet inspector.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ParagraphListBullet(Box<str>);

impl ParagraphListBullet {
    /// Apple's standard text bullet.
    pub const STANDARD: &'static str = "•";

    /// Construct a validated custom text bullet.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        let character_count = value.chars().count();
        if character_count == 0 {
            return Err(Error::InvalidFormat(
                "paragraph list bullet must not be empty".to_owned(),
            ));
        }
        if character_count > MAX_PARAGRAPH_LIST_BULLET_CHARACTERS {
            return Err(Error::InvalidFormat(format!(
                "paragraph list bullet must not exceed {MAX_PARAGRAPH_LIST_BULLET_CHARACTERS} characters"
            )));
        }
        if value.chars().any(char::is_control) {
            return Err(Error::InvalidFormat(
                "paragraph list bullet must not contain control characters".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Borrow the marker exactly as iWork displays it.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl Default for ParagraphListBullet {
    fn default() -> Self {
        Self(Self::STANDARD.into())
    }
}

impl AsRef<str> for ParagraphListBullet {
    fn as_ref(&self) -> &str {
        self.as_str()
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

    #[test]
    fn list_start_numbers_are_positive_and_numbering_is_typed() {
        assert!(ParagraphListStart::new(0).is_err());
        let seven = ParagraphListStart::new(7).unwrap();
        assert_eq!(seven.get(), 7);
        assert_eq!(
            ParagraphListNumbering::from_native(7).unwrap(),
            ParagraphListNumbering::StartAt(seven)
        );
        assert_eq!(
            ParagraphListNumbering::from_native(0).unwrap(),
            ParagraphListNumbering::Continue
        );
    }

    #[test]
    fn text_bullets_are_nonempty_printable_and_bounded() {
        assert_eq!(
            ParagraphListBullet::default().as_str(),
            ParagraphListBullet::STANDARD
        );
        assert_eq!(ParagraphListBullet::new("➡").unwrap().as_str(), "➡");
        assert!(ParagraphListBullet::new("").is_err());
        assert!(ParagraphListBullet::new("a\nb").is_err());
        assert!(ParagraphListBullet::new("x".repeat(33)).is_err());
    }
}
