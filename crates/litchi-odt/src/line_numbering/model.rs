//! Typed semantic model for ODF line-numbering metadata.

use litchi_core::Result;

use super::{MAX_VALUE_BYTES, invalid};

/// Numbering format for line numbers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Format {
    Empty,
    Arabic,
    LowerRoman,
    UpperRoman,
    LowerAlpha,
    UpperAlpha,
    Custom(String),
}

impl Format {
    pub fn parse(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE_BYTES {
            return invalid(format!(
                "style:num-format exceeds the {MAX_VALUE_BYTES} byte limit"
            ));
        }
        Ok(match value.as_str() {
            "" => Self::Empty,
            "1" => Self::Arabic,
            "i" => Self::LowerRoman,
            "I" => Self::UpperRoman,
            "a" => Self::LowerAlpha,
            "A" => Self::UpperAlpha,
            _ => Self::Custom(value),
        })
    }

    pub fn as_str(&self) -> &str {
        match self {
            Self::Empty => "",
            Self::Arabic => "1",
            Self::LowerRoman => "i",
            Self::UpperRoman => "I",
            Self::LowerAlpha => "a",
            Self::UpperAlpha => "A",
            Self::Custom(value) => value,
        }
    }

    fn permits_letter_sync(&self) -> bool {
        matches!(self, Self::LowerAlpha | Self::UpperAlpha)
    }
}

/// Placement of line numbers relative to the text area.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Position {
    Left,
    Right,
    Inner,
    Outer,
}

impl Position {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "inner" => Ok(Self::Inner),
            "outer" => Ok(Self::Outer),
            _ => invalid(format!("unsupported text:number-position '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Right => "right",
            Self::Inner => "inner",
            Self::Outer => "outer",
        }
    }
}

/// A validated ODF nonnegative length.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NonNegativeLength(String);

impl NonNegativeLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_nonnegative_length(&value)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Optional separator emitted after every configured number of lines.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Separator {
    pub increment: Option<u64>,
    pub text: String,
}

/// One standard `text:linenumbering-configuration` declaration.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Configuration {
    pub number_lines: Option<bool>,
    pub number_format: Option<Format>,
    pub letter_sync: Option<bool>,
    pub style_name: Option<String>,
    pub increment: Option<u64>,
    pub number_position: Option<Position>,
    pub offset: Option<NonNegativeLength>,
    pub count_empty_lines: Option<bool>,
    pub count_in_text_boxes: Option<bool>,
    pub restart_on_page: Option<bool>,
    pub separator: Option<Separator>,
}

impl Configuration {
    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !self
                .number_format
                .as_ref()
                .is_some_and(Format::permits_letter_sync)
        {
            return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
        }
        if let Some(format) = &self.number_format {
            validate_value(format.as_str(), "style:num-format", true)?;
        }
        if let Some(style_name) = &self.style_name {
            validate_value(style_name, "text:style-name", false)?;
        }
        if let Some(offset) = &self.offset {
            validate_nonnegative_length(offset.as_str())?;
        }
        if let Some(separator) = &self.separator {
            validate_value(&separator.text, "text:linenumbering-separator", true)?;
        }
        Ok(())
    }
}

fn validate_nonnegative_length(value: &str) -> Result<()> {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return invalid(format!("invalid nonnegative ODF length '{value}'"));
    };
    let mut dots = 0usize;
    let mut digits = 0usize;
    for byte in number.bytes() {
        match byte {
            b'.' => dots += 1,
            b'0'..=b'9' => digits += 1,
            _ => return invalid(format!("invalid nonnegative ODF length '{value}'")),
        }
    }
    if number.is_empty() || number == "." || dots > 1 || digits == 0 {
        return invalid(format!("invalid nonnegative ODF length '{value}'"));
    }
    validate_value(value, "text:offset", false)
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
