//! Worksheet sort state structures.
//!
//! Excel stores sort settings in the `<sortState>` element (typically within
//! `<autoFilter>` or table definitions). This module defines the data structures
//! used to parse and serialize that state.

/// Sort method for locale-sensitive ordering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortMethod {
    /// Stroke order (used for some East Asian locales).
    Stroke,
    /// PinYin order.
    PinYin,
}

impl SortMethod {
    /// Parse a sort method from the OOXML attribute value.
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "stroke" => Some(Self::Stroke),
            "pinYin" => Some(Self::PinYin),
            _ => None,
        }
    }

    /// Return the OOXML attribute value for this sort method.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Stroke => "stroke",
            Self::PinYin => "pinYin",
        }
    }
}

/// Sort criterion type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortBy {
    /// Sort by the cell value (default).
    Value,
    /// Sort by cell fill color.
    CellColor,
    /// Sort by font color.
    FontColor,
    /// Sort by conditional formatting icon.
    Icon,
}

impl SortBy {
    /// Parse a sort criterion from the OOXML attribute value.
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "value" => Some(Self::Value),
            "cellColor" => Some(Self::CellColor),
            "fontColor" => Some(Self::FontColor),
            "icon" => Some(Self::Icon),
            _ => None,
        }
    }

    /// Return the OOXML attribute value for this criterion.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::CellColor => "cellColor",
            Self::FontColor => "fontColor",
            Self::Icon => "icon",
        }
    }
}

/// A single sort condition in a sort state.
#[derive(Debug, Clone)]
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
    /// Icon set name for icon-based sorting.
    pub icon_set: Option<String>,
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
}

/// Sort state associated with an auto-filter or table.
#[derive(Debug, Clone)]
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
