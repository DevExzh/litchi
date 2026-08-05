//! Semantic values and lexical validation for ODF footnote separators.

use litchi_core::{Error, Result};
use std::{fmt, str::FromStr};

use super::{MAX_VALUE_BYTES, invalid};

/// ODF `length` lexical value used by a footnote separator.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Length(String);

impl Length {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_measure(&value, false)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for Length {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl FromStr for Length {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// ODF `percent` lexical value used by `style:rel-width`.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Percent(String);

impl Percent {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_measure(&value, true)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for Percent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl FromStr for Percent {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum LineStyle {
    None,
    Solid,
    Dotted,
    Dash,
    LongDash,
    DotDash,
    DotDotDash,
    Wave,
}

impl LineStyle {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "solid" => Ok(Self::Solid),
            "dotted" => Ok(Self::Dotted),
            "dash" => Ok(Self::Dash),
            "long-dash" => Ok(Self::LongDash),
            "dot-dash" => Ok(Self::DotDash),
            "dot-dot-dash" => Ok(Self::DotDotDash),
            "wave" => Ok(Self::Wave),
            _ => invalid(format!("invalid style:line-style '{value}'")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::Dotted => "dotted",
            Self::Dash => "dash",
            Self::LongDash => "long-dash",
            Self::DotDash => "dot-dash",
            Self::DotDotDash => "dot-dot-dash",
            Self::Wave => "wave",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Adjustment {
    Left,
    Center,
    Right,
}

impl Adjustment {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            _ => invalid(format!("invalid style:adjustment '{value}'")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
        }
    }
}

/// One optional footnote-separator rule in `style:page-layout-properties`.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct Separator {
    pub width: Option<Length>,
    pub relative_width: Option<Percent>,
    pub color: Option<(u8, u8, u8)>,
    pub line_style: Option<LineStyle>,
    pub adjustment: Option<Adjustment>,
    pub distance_before: Option<Length>,
    pub distance_after: Option<Length>,
}

impl Separator {
    pub fn validate(&self) -> Result<()> {
        for value in [&self.width, &self.distance_before, &self.distance_after]
            .into_iter()
            .flatten()
        {
            validate_measure(value.as_str(), false)?;
        }
        if let Some(value) = &self.relative_width {
            validate_measure(value.as_str(), true)?;
        }
        Ok(())
    }
}

fn validate_measure(value: &str, percent: bool) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES {
        return invalid("invalid ODF measure");
    }
    let number = if percent {
        value
            .strip_suffix('%')
            .ok_or_else(|| Error::InvalidFormat(format!("invalid ODF percent '{value}'")))?
    } else {
        let unit = ["cm", "mm", "in", "pt", "pc", "px"]
            .into_iter()
            .find(|unit| value.ends_with(unit))
            .ok_or_else(|| Error::InvalidFormat(format!("invalid ODF length '{value}'")))?;
        &value[..value.len() - unit.len()]
    };
    let number = number.strip_prefix('-').unwrap_or(number);
    let mut parts = number.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    let valid = parts.next().is_none()
        && whole.bytes().all(|byte| byte.is_ascii_digit())
        && match fraction {
            Some(fraction) => !whole.is_empty() || !fraction.is_empty(),
            None => !whole.is_empty(),
        }
        && fraction.is_none_or(|fraction| fraction.bytes().all(|byte| byte.is_ascii_digit()));
    if !valid {
        return invalid(if percent {
            format!("invalid ODF percent '{value}'")
        } else {
            format!("invalid ODF length '{value}'")
        });
    }
    Ok(())
}
