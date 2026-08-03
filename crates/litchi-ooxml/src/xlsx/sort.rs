//! Worksheet sort state structures.
//!
//! Excel stores sort settings in the `<sortState>` element (typically within
//! `<autoFilter>` or table definitions). This module defines the data structures
//! used to parse and serialize that state.

use std::fmt;
use std::str::FromStr;

use super::conditional_formatting::{IconSet, TokenError};

/// Sort method for locale-sensitive ordering (`ST_SortMethod`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum SortMethod {
    None,
    Stroke,
    PinYin,
}

impl SortMethod {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Stroke => "stroke",
            Self::PinYin => "pinYin",
        }
    }
}

impl FromStr for SortMethod {
    type Err = TokenError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "none" => Ok(Self::None),
            "stroke" => Ok(Self::Stroke),
            "pinYin" => Ok(Self::PinYin),
            _ => Err(TokenError::new("sort method", value)),
        }
    }
}

impl fmt::Display for SortMethod {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Sort criterion (`ST_SortBy`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum SortBy {
    Value,
    CellColor,
    FontColor,
    Icon,
}

impl SortBy {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::CellColor => "cellColor",
            Self::FontColor => "fontColor",
            Self::Icon => "icon",
        }
    }
}

impl FromStr for SortBy {
    type Err = TokenError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "value" => Ok(Self::Value),
            "cellColor" => Ok(Self::CellColor),
            "fontColor" => Ok(Self::FontColor),
            "icon" => Ok(Self::Icon),
            _ => Err(TokenError::new("sort criterion", value)),
        }
    }
}

impl fmt::Display for SortBy {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// A single sort condition in a sort state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortCondition {
    /// Range that participates in this sort key (e.g. "A2:A20").
    pub ref_range: String,
    /// Whether the sort is descending.
    pub descending: Option<bool>,
    /// How the key should be sorted (value, color, icon).
    pub sort_by: Option<SortBy>,
    /// Custom list entries (comma separated) used for ordering.
    pub custom_list: Option<String>,
    /// Differential format ID used for color/icon sorting.
    pub dxf_id: Option<u32>,
    /// Core icon set for icon-based sorting.
    pub icon_set: Option<IconSet>,
    /// Icon index within the icon set.
    pub icon_id: Option<u32>,
}

impl SortCondition {
    /// Create a new sort condition for the given range.
    pub fn new(ref_range: impl Into<String>) -> Self {
        Self {
            ref_range: ref_range.into(),
            descending: None,
            sort_by: None,
            custom_list: None,
            dxf_id: None,
            icon_set: None,
            icon_id: None,
        }
    }

    /// Validate relationships between the selected criterion and its auxiliary fields.
    pub fn validate(&self) -> Result<(), Error> {
        let sort_by = self.sort_by.unwrap_or(SortBy::Value);
        match sort_by {
            SortBy::CellColor | SortBy::FontColor => {
                if self.dxf_id.is_none() {
                    return Err(Error::MissingDifferentialFormat(sort_by));
                }
                if self.icon_set.is_some() || self.icon_id.is_some() {
                    return Err(Error::UnexpectedIcon);
                }
            },
            SortBy::Icon => {
                if self.dxf_id.is_some() {
                    return Err(Error::UnexpectedDifferentialFormat);
                }
                match (self.icon_set, self.icon_id) {
                    (None, Some(_)) => return Err(Error::MissingIconSet),
                    (Some(set), Some(id)) if id >= u32::from(set.len()) => {
                        return Err(Error::IconOutOfRange { set, id });
                    },
                    _ => {},
                }
            },
            SortBy::Value => {
                if self.dxf_id.is_some() {
                    return Err(Error::UnexpectedDifferentialFormat);
                }
                if self.icon_set.is_some() || self.icon_id.is_some() {
                    return Err(Error::UnexpectedIcon);
                }
            },
        }
        Ok(())
    }
}

/// Invalid combination of fields in a sort condition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    MissingDifferentialFormat(SortBy),
    UnexpectedDifferentialFormat,
    MissingIconSet,
    UnexpectedIcon,
    IconOutOfRange { set: IconSet, id: u32 },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::MissingDifferentialFormat(sort_by) => {
                write!(formatter, "{} sorting requires dxfId", sort_by.as_str())
            },
            Self::UnexpectedDifferentialFormat => {
                formatter.write_str("dxfId is only valid for color sorting")
            },
            Self::MissingIconSet => formatter.write_str("iconId requires iconSet"),
            Self::UnexpectedIcon => {
                formatter.write_str("iconSet and iconId are only valid for icon sorting")
            },
            Self::IconOutOfRange { set, id } => write!(
                formatter,
                "iconId {id} exceeds {} cardinality {}",
                set.as_str(),
                set.len()
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Sort state associated with an auto-filter or table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortState {
    /// Range that covers the full sort operation.
    pub ref_range: String,
    /// Whether the sort is column-based rather than row-based.
    pub column_sort: Option<bool>,
    /// Whether the sort is case-sensitive.
    pub case_sensitive: Option<bool>,
    /// Locale sort method used for text.
    pub sort_method: Option<SortMethod>,
    /// Sort conditions (keys).
    pub conditions: Vec<SortCondition>,
}

impl SortState {
    /// Create a new sort state for the given range.
    pub fn new(ref_range: impl Into<String>) -> Self {
        Self {
            ref_range: ref_range.into(),
            column_sort: None,
            case_sensitive: None,
            sort_method: None,
            conditions: Vec::new(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sort_domains_are_compact_exact_and_strict() {
        for (value, token) in [
            (SortMethod::None, "none"),
            (SortMethod::Stroke, "stroke"),
            (SortMethod::PinYin, "pinYin"),
        ] {
            assert_eq!(value.as_str(), token);
            assert_eq!(value.to_string(), token);
            assert_eq!(token.parse::<SortMethod>(), Ok(value));
        }
        for (value, token) in [
            (SortBy::Value, "value"),
            (SortBy::CellColor, "cellColor"),
            (SortBy::FontColor, "fontColor"),
            (SortBy::Icon, "icon"),
        ] {
            assert_eq!(value.as_str(), token);
            assert_eq!(value.to_string(), token);
            assert_eq!(token.parse::<SortBy>(), Ok(value));
        }
        assert!("pinyin".parse::<SortMethod>().is_err());
        assert!("color".parse::<SortBy>().is_err());
        assert_eq!(std::mem::size_of::<SortMethod>(), 1);
        assert_eq!(std::mem::size_of::<SortBy>(), 1);
    }

    #[test]
    fn sort_condition_rejects_mismatched_auxiliary_fields() {
        let mut condition = SortCondition::new("A2:A10");
        condition.sort_by = Some(SortBy::Icon);
        condition.icon_set = Some(IconSet::ThreeArrows);
        condition.icon_id = Some(2);
        assert_eq!(condition.validate(), Ok(()));

        condition.icon_id = Some(3);
        assert_eq!(
            condition.validate(),
            Err(Error::IconOutOfRange {
                set: IconSet::ThreeArrows,
                id: 3,
            })
        );

        condition.sort_by = Some(SortBy::CellColor);
        condition.icon_set = None;
        condition.icon_id = None;
        assert_eq!(
            condition.validate(),
            Err(Error::MissingDifferentialFormat(SortBy::CellColor))
        );
    }
}
