//! Typed SpreadsheetML timeline cache and worksheet timeline values.

use crate::package::error::{Error, Result};
use chrono::{DateTime, NaiveDateTime, Utc};
use std::collections::HashSet;

pub(crate) const MAX_CACHES: usize = 4096;
pub(crate) const MAX_VIEWS: usize = 16_384;
pub(crate) const MAX_PIVOT_TABLES: usize = 65_536;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;

/// Timeline grouping level from `[MS-XLSX]`/`[MS-XLSB]` timeline XML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Level {
    /// Calendar year.
    Year,
    /// Calendar quarter.
    Quarter,
    /// Calendar month.
    Month,
    /// Calendar day.
    Day,
}

impl Level {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "0" => Ok(Self::Year),
            "1" => Ok(Self::Quarter),
            "2" => Ok(Self::Month),
            "3" => Ok(Self::Day),
            _ => Err(Error::Unrecognized {
                typ: "timeline level".to_string(),
                val: value.to_string(),
            }),
        }
    }

    pub(crate) const fn wire(self) -> u8 {
        match self {
            Self::Year => 0,
            Self::Quarter => 1,
            Self::Month => 2,
            Self::Day => 3,
        }
    }
}

/// Date-range selection in a timeline cache state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
    /// Inclusive start dateTime lexical value.
    pub start_date: String,
    /// Inclusive end dateTime lexical value.
    pub end_date: String,
}

impl Range {
    /// Construct and validate a date range.
    pub fn new(start_date: impl Into<String>, end_date: impl Into<String>) -> Result<Self> {
        let value = Self {
            start_date: start_date.into(),
            end_date: end_date.into(),
        };
        validate_range(&value)?;
        Ok(value)
    }
}

/// Supported timeline filter vocabulary. Filter execution is intentionally
/// inert; this enum only validates and preserves the XML spelling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterType {
    /// No recognized filter.
    Unknown,
    /// Exact date.
    DateEqual,
    /// Date interval.
    DateBetween,
    /// Inverse date interval.
    DateNotBetween,
    /// Relative day/month/quarter/year filters.
    Today,
    Yesterday,
    Tomorrow,
    ThisWeek,
    LastWeek,
    NextWeek,
    ThisMonth,
    LastMonth,
    NextMonth,
    ThisQuarter,
    LastQuarter,
    NextQuarter,
    ThisYear,
    LastYear,
    NextYear,
    YearToDate,
    Q1,
    Q2,
    Q3,
    Q4,
    M1,
    M2,
    M3,
    M4,
    M5,
    M6,
    M7,
    M8,
    M9,
    M10,
    M11,
    M12,
}

impl FilterType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        let parsed = match value {
            "unknown" => Self::Unknown,
            "dateEqual" => Self::DateEqual,
            "dateBetween" => Self::DateBetween,
            "dateNotBetween" => Self::DateNotBetween,
            "today" => Self::Today,
            "yesterday" => Self::Yesterday,
            "tomorrow" => Self::Tomorrow,
            "thisWeek" => Self::ThisWeek,
            "lastWeek" => Self::LastWeek,
            "nextWeek" => Self::NextWeek,
            "thisMonth" => Self::ThisMonth,
            "lastMonth" => Self::LastMonth,
            "nextMonth" => Self::NextMonth,
            "thisQuarter" => Self::ThisQuarter,
            "lastQuarter" => Self::LastQuarter,
            "nextQuarter" => Self::NextQuarter,
            "thisYear" => Self::ThisYear,
            "lastYear" => Self::LastYear,
            "nextYear" => Self::NextYear,
            "yearToDate" => Self::YearToDate,
            "Q1" => Self::Q1,
            "Q2" => Self::Q2,
            "Q3" => Self::Q3,
            "Q4" => Self::Q4,
            "M1" => Self::M1,
            "M2" => Self::M2,
            "M3" => Self::M3,
            "M4" => Self::M4,
            "M5" => Self::M5,
            "M6" => Self::M6,
            "M7" => Self::M7,
            "M8" => Self::M8,
            "M9" => Self::M9,
            "M10" => Self::M10,
            "M11" => Self::M11,
            "M12" => Self::M12,
            _ => {
                return Err(Error::Unrecognized {
                    typ: "timeline filter type".to_string(),
                    val: value.to_string(),
                });
            },
        };
        Ok(parsed)
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Unknown => "unknown",
            Self::DateEqual => "dateEqual",
            Self::DateBetween => "dateBetween",
            Self::DateNotBetween => "dateNotBetween",
            Self::Today => "today",
            Self::Yesterday => "yesterday",
            Self::Tomorrow => "tomorrow",
            Self::ThisWeek => "thisWeek",
            Self::LastWeek => "lastWeek",
            Self::NextWeek => "nextWeek",
            Self::ThisMonth => "thisMonth",
            Self::LastMonth => "lastMonth",
            Self::NextMonth => "nextMonth",
            Self::ThisQuarter => "thisQuarter",
            Self::LastQuarter => "lastQuarter",
            Self::NextQuarter => "nextQuarter",
            Self::ThisYear => "thisYear",
            Self::LastYear => "lastYear",
            Self::NextYear => "nextYear",
            Self::YearToDate => "yearToDate",
            Self::Q1 => "Q1",
            Self::Q2 => "Q2",
            Self::Q3 => "Q3",
            Self::Q4 => "Q4",
            Self::M1 => "M1",
            Self::M2 => "M2",
            Self::M3 => "M3",
            Self::M4 => "M4",
            Self::M5 => "M5",
            Self::M6 => "M6",
            Self::M7 => "M7",
            Self::M8 => "M8",
            Self::M9 => "M9",
            Self::M10 => "M10",
            Self::M11 => "M11",
            Self::M12 => "M12",
        }
    }
}

/// Timeline cache state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct State {
    /// Optional selected date range.
    pub selection: Option<Range>,
    /// Complete available date range.
    pub bounds: Range,
    /// Whether one contiguous selection is required.
    pub single_range_filter_state: bool,
    /// Minimum refresh version advertised by the producer.
    pub minimal_refresh_version: u32,
    /// Last refresh version advertised by the producer.
    pub last_refresh_version: u32,
    /// Associated PivotCache identifier.
    pub pivot_cache_id: u32,
    /// Inert filter vocabulary value.
    pub filter_type: FilterType,
}

impl State {
    /// Construct a state with default refresh metadata.
    #[must_use]
    pub fn new(bounds: Range, pivot_cache_id: u32, filter_type: FilterType) -> Self {
        Self {
            selection: None,
            bounds,
            single_range_filter_state: true,
            minimal_refresh_version: 0,
            last_refresh_version: 0,
            pivot_cache_id,
            filter_type,
        }
    }
}

/// PivotTable reference in a timeline cache definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTable {
    /// Worksheet tab identifier.
    pub tab_id: u32,
    /// PivotTable name.
    pub name: String,
}

/// Optional inert timeline PivotFilter metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Filter {
    /// Whether date matching uses the whole day.
    pub use_whole_day: bool,
    /// Pivot field identifier.
    pub field: u32,
    /// Filter identifier.
    pub id: u32,
    /// Optional display name.
    pub name: Option<String>,
    /// Optional description.
    pub description: Option<String>,
}

impl Filter {
    /// Construct a minimal inert PivotFilter.
    #[must_use]
    pub const fn new(field: u32, id: u32) -> Self {
        Self {
            use_whole_day: false,
            field,
            id,
            name: None,
            description: None,
        }
    }
}

/// Timeline cache definition stored as XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    /// Unique cache name.
    pub name: String,
    /// Source column name.
    pub source_name: String,
    /// Optional revision GUID preserved as text.
    pub uid: Option<String>,
    /// PivotTables using this cache.
    pub pivot_tables: Vec<PivotTable>,
    /// Cache state.
    pub state: State,
    /// Optional inert PivotFilter.
    pub filter: Option<Filter>,
}

impl Cache {
    /// Construct a timeline cache definition.
    #[must_use]
    pub fn new(name: impl Into<String>, source_name: impl Into<String>, state: State) -> Self {
        Self {
            name: name.into(),
            source_name: source_name.into(),
            uid: None,
            pivot_tables: Vec::new(),
            state,
            filter: None,
        }
    }
}

/// One worksheet timeline view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    /// Unique worksheet timeline name.
    pub name: String,
    /// Referenced cache name.
    pub cache: String,
    /// Display grouping level.
    pub level: Level,
    /// Selection grouping level.
    pub selection_level: Level,
    /// Optional display caption.
    pub caption: Option<String>,
    /// Whether the header is shown.
    pub show_header: bool,
    /// Whether the selection label is shown.
    pub show_selection_label: bool,
    /// Whether the time-level label is shown.
    pub show_time_level: bool,
    /// Whether the horizontal scrollbar is shown.
    pub show_horizontal_scrollbar: bool,
    /// Optional lexical dateTime scroll position.
    pub scroll_position: Option<String>,
    /// Optional timeline style name.
    pub style: Option<String>,
    /// Optional revision GUID preserved as text.
    pub uid: Option<String>,
}

impl View {
    /// Construct a timeline view with SpreadsheetML defaults.
    #[must_use]
    pub fn new(name: impl Into<String>, cache: impl Into<String>, level: Level) -> Self {
        Self {
            name: name.into(),
            cache: cache.into(),
            level,
            selection_level: level,
            caption: None,
            show_header: true,
            show_selection_label: true,
            show_time_level: true,
            show_horizontal_scrollbar: true,
            scroll_position: None,
            style: None,
            uid: None,
        }
    }
}

/// Timelines stored in one worksheet timeline part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Views {
    /// Views in XML order.
    pub items: Vec<View>,
}

impl Views {
    /// Construct an empty timeline collection.
    #[must_use]
    pub const fn new() -> Self {
        Self { items: Vec::new() }
    }
}

fn bounded(value: &str, field: &'static str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > MAX_STRING_BYTES || value.contains('\0')
    {
        return Err(Error::Unrecognized {
            typ: field.to_string(),
            val: "empty, oversized, or NUL-containing value".to_string(),
        });
    }
    Ok(())
}

fn parse_datetime(value: &str) -> Option<i64> {
    if let Ok(parsed) = DateTime::parse_from_rfc3339(value) {
        return Some(parsed.timestamp());
    }
    NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
        .ok()
        .map(|parsed| DateTime::<Utc>::from_naive_utc_and_offset(parsed, Utc).timestamp())
}

pub(crate) fn validate_range(value: &Range) -> Result<()> {
    bounded(&value.start_date, "timeline startDate", false)?;
    bounded(&value.end_date, "timeline endDate", false)?;
    let start = parse_datetime(&value.start_date).ok_or_else(|| Error::Unrecognized {
        typ: "timeline startDate".to_string(),
        val: value.start_date.clone(),
    })?;
    let end = parse_datetime(&value.end_date).ok_or_else(|| Error::Unrecognized {
        typ: "timeline endDate".to_string(),
        val: value.end_date.clone(),
    })?;
    if start > end {
        return Err(Error::Unrecognized {
            typ: "timeline range".to_string(),
            val: "startDate is after endDate".to_string(),
        });
    }
    Ok(())
}

pub(crate) fn validate_cache(value: &Cache) -> Result<()> {
    bounded(&value.name, "timeline cache name", false)?;
    bounded(&value.source_name, "timeline sourceName", false)?;
    if let Some(uid) = &value.uid {
        bounded(uid, "timeline cache uid", false)?;
    }
    if value.pivot_tables.len() > MAX_PIVOT_TABLES {
        return Err(Error::InvalidLength {
            expected: MAX_PIVOT_TABLES,
            found: value.pivot_tables.len(),
        });
    }
    let mut pivots = HashSet::with_capacity(value.pivot_tables.len());
    for pivot in &value.pivot_tables {
        bounded(&pivot.name, "timeline PivotTable name", false)?;
        if !pivots.insert((pivot.tab_id, pivot.name.to_ascii_lowercase())) {
            return Err(Error::Unrecognized {
                typ: "timeline PivotTable collection".to_string(),
                val: "duplicate PivotTable reference".to_string(),
            });
        }
    }
    validate_state(&value.state)?;
    if let Some(filter) = &value.filter {
        bounded(
            filter.name.as_deref().unwrap_or(""),
            "timeline filter name",
            true,
        )?;
        bounded(
            filter.description.as_deref().unwrap_or(""),
            "timeline filter description",
            true,
        )?;
    }
    Ok(())
}

pub(crate) fn validate_state(value: &State) -> Result<()> {
    validate_range(&value.bounds)?;
    if let Some(selection) = &value.selection {
        validate_range(selection)?;
        let bounds_start =
            parse_datetime(&value.bounds.start_date).ok_or_else(|| Error::Unrecognized {
                typ: "timeline bounds".to_string(),
                val: "invalid start dateTime".to_string(),
            })?;
        let bounds_end =
            parse_datetime(&value.bounds.end_date).ok_or_else(|| Error::Unrecognized {
                typ: "timeline bounds".to_string(),
                val: "invalid end dateTime".to_string(),
            })?;
        let selection_start =
            parse_datetime(&selection.start_date).ok_or_else(|| Error::Unrecognized {
                typ: "timeline selection".to_string(),
                val: "invalid start dateTime".to_string(),
            })?;
        let selection_end =
            parse_datetime(&selection.end_date).ok_or_else(|| Error::Unrecognized {
                typ: "timeline selection".to_string(),
                val: "invalid end dateTime".to_string(),
            })?;
        if selection_start < bounds_start || selection_end > bounds_end {
            return Err(Error::Unrecognized {
                typ: "timeline selection".to_string(),
                val: "selection lies outside bounds".to_string(),
            });
        }
    }
    Ok(())
}

pub(crate) fn validate_views(value: &Views) -> Result<()> {
    if value.items.len() > MAX_VIEWS {
        return Err(Error::InvalidLength {
            expected: MAX_VIEWS,
            found: value.items.len(),
        });
    }
    let mut names = HashSet::with_capacity(value.items.len());
    for view in &value.items {
        bounded(&view.name, "timeline name", false)?;
        bounded(&view.cache, "timeline cache", false)?;
        if !names.insert(view.name.to_ascii_lowercase()) {
            return Err(Error::Unrecognized {
                typ: "timeline collection".to_string(),
                val: "duplicate timeline name".to_string(),
            });
        }
        if let Some(value) = &view.uid {
            bounded(value, "timeline uid", false)?;
        }
        if let Some(value) = &view.caption {
            bounded(value, "timeline caption", true)?;
        }
        if let Some(value) = &view.scroll_position {
            let range = Range::new(value, value)?;
            let _ = range;
        }
        if let Some(value) = &view.style {
            bounded(value, "timeline style", false)?;
        }
    }
    Ok(())
}
