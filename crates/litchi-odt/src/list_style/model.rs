//! Typed semantic values for ODF `text:list-style` declarations.

use super::{MAX_BINARY, MAX_LEVEL, bad, name_ok, percent};
use crate::outline_style::{NumberFormat, PositiveInteger};
use litchi_core::Result;
use std::collections::HashSet;

/// Valid ODF non-negative `percent` lexical value for
/// `text:bullet-relative-size`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BulletRelativeSize(String);

impl BulletRelativeSize {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if !percent(&value) {
            return Err(bad("text:bullet-relative-size must be an ODF percent"));
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Level numbering decoration of a `text:list-level-style-number`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NumberStyle {
    pub format: Option<NumberFormat>,
    pub prefix: Option<String>,
    pub suffix: Option<String>,
    pub letter_sync: Option<bool>,
    pub display_levels: Option<PositiveInteger>,
    pub start_value: Option<PositiveInteger>,
}

impl NumberStyle {
    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !self
                .format
                .as_ref()
                .is_some_and(|format| matches!(format.as_str(), "a" | "A"))
        {
            return Err(bad(
                "style:num-letter-sync requires style:num-format 'a' or 'A'",
            ));
        }
        if let Some(value) = &self.prefix {
            name_ok(value, "style:num-prefix")?;
        }
        if let Some(value) = &self.suffix {
            name_ok(value, "style:num-suffix")?;
        }
        Ok(())
    }
}

/// Bullet decoration of a `text:list-level-style-bullet`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BulletStyle {
    pub bullet_char: char,
    pub relative_size: Option<BulletRelativeSize>,
    pub prefix: Option<String>,
    pub suffix: Option<String>,
}

impl BulletStyle {
    pub fn new(bullet_char: char) -> Result<Self> {
        let result = Self {
            bullet_char,
            relative_size: None,
            prefix: None,
            suffix: None,
        };
        result.validate()?;
        Ok(result)
    }

    pub fn validate(&self) -> Result<()> {
        if self.bullet_char.is_control() {
            return Err(bad("text:bullet-char must not be a control character"));
        }
        if let Some(value) = &self.prefix {
            name_ok(value, "style:num-prefix")?;
        }
        if let Some(value) = &self.suffix {
            name_ok(value, "style:num-suffix")?;
        }
        Ok(())
    }
}

/// Image source of a `text:list-level-style-image`: linked or embedded binary
/// data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageSource {
    /// `xlink:href` reference to an image resource.
    Linked(String),
    /// Base64 content of an `office:binary-data` child.
    Embedded(String),
}

impl ImageSource {
    pub(super) fn validate(&self) -> Result<()> {
        match self {
            Self::Linked(href) => name_ok(href, "xlink:href"),
            Self::Embedded(data) => {
                if data.len() > MAX_BINARY
                    || !data.bytes().all(|byte| {
                        byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/' | b'=')
                    })
                {
                    return Err(bad("office:binary-data must be base64 text"));
                }
                Ok(())
            },
        }
    }
}

/// The level-specific decoration of one list level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Kind {
    Number(NumberStyle),
    Bullet(BulletStyle),
    Image(ImageSource),
}

/// One `text:list-level-style-*` declaration of a list style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LevelStyle {
    pub level: u16,
    pub style_name: Option<String>,
    pub kind: Kind,
}

impl LevelStyle {
    pub fn validate(&self) -> Result<()> {
        if !(1..=MAX_LEVEL).contains(&self.level) {
            return Err(bad("text:level is outside the supported range"));
        }
        if let Some(value) = &self.style_name {
            name_ok(value, "text:style-name")?;
        }
        match &self.kind {
            Kind::Number(number) => number.validate(),
            Kind::Bullet(bullet) => bullet.validate(),
            Kind::Image(source) => source.validate(),
        }
    }
}

/// One `text:list-style` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: String,
    pub display_name: Option<String>,
    pub consecutive_numbering: Option<bool>,
    pub levels: Vec<LevelStyle>,
}

impl Style {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let result = Self {
            name: name.into(),
            display_name: None,
            consecutive_numbering: None,
            levels: Vec::new(),
        };
        result.validate()?;
        Ok(result)
    }

    pub fn level(&self, level: u16) -> Option<&LevelStyle> {
        self.levels.iter().find(|entry| entry.level == level)
    }

    pub fn validate(&self) -> Result<()> {
        name_ok(&self.name, "list style name")?;
        if let Some(value) = &self.display_name {
            name_ok(value, "style:display-name")?;
        }
        let mut seen = HashSet::new();
        for level in &self.levels {
            level.validate()?;
            if !seen.insert(level.level) {
                return Err(bad("duplicate list level"));
            }
        }
        Ok(())
    }
}

/// All `text:list-style` declarations of one styles part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}

impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles.iter().find(|style| style.name == name)
    }
}
