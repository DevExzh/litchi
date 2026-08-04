//! Semantic XLSX timeline cache, state, filter, and worksheet-view values.

use super::codec::validate_opaque_kind;
use super::{
    MAX_CACHES, MAX_PIVOT_TABLES, MAX_TIMELINES, MAX_TOTAL_OPAQUE_BYTES, SML, STRICT_SML, X15,
    bounded, bounded_nonempty, invalid, limit,
};
use crate::auto_filter::{AutoFilterDefinition, write_auto_filter_fragment};
use crate::error::Result;
use litchi_ooxml_common::custom_xml::valid_guid;
use std::collections::HashSet;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueXml {
    /// A single, self-contained XML element. It is parsed and bounded but not interpreted.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachePivotTable {
    pub tab_id: u32,
    pub name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CacheDefinition {
    pub name: String,
    pub uid: Option<String>,
    pub source_name: String,
    pub pivot_tables: Vec<CachePivotTable>,
    pub state: State,
    pub timeline_pivot_filter: Option<PivotFilter>,
    pub extension_list: Option<OpaqueXml>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
    pub start_date: String,
    pub end_date: String,
}
impl Range {
    pub fn new(start: impl Into<String>, end: impl Into<String>) -> Result<Self> {
        let v = Self {
            start_date: start.into(),
            end_date: end.into(),
        };
        validate_range(&v)?;
        Ok(v)
    }
    pub fn start_date(&self) -> &str {
        &self.start_date
    }
    pub fn end_date(&self) -> &str {
        &self.end_date
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FilterType {
    Unknown,
    Count,
    Percent,
    Sum,
    CaptionEqual,
    CaptionNotEqual,
    CaptionBeginsWith,
    CaptionNotBeginsWith,
    CaptionEndsWith,
    CaptionNotEndsWith,
    CaptionContains,
    CaptionNotContains,
    CaptionGreaterThan,
    CaptionGreaterThanOrEqual,
    CaptionLessThan,
    CaptionLessThanOrEqual,
    CaptionBetween,
    CaptionNotBetween,
    ValueEqual,
    ValueNotEqual,
    ValueGreaterThan,
    ValueGreaterThanOrEqual,
    ValueLessThan,
    ValueLessThanOrEqual,
    ValueBetween,
    ValueNotBetween,
    DateEqual,
    DateNotEqual,
    DateOlderThan,
    DateOlderThanOrEqual,
    DateNewerThan,
    DateNewerThanOrEqual,
    DateBetween,
    DateNotBetween,
    Tomorrow,
    Today,
    Yesterday,
    NextWeek,
    ThisWeek,
    LastWeek,
    NextMonth,
    ThisMonth,
    LastMonth,
    NextQuarter,
    ThisQuarter,
    LastQuarter,
    NextYear,
    ThisYear,
    LastYear,
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
    pub(super) fn parse(v: &str) -> Result<Self> {
        use FilterType::*;
        Ok(match v {
            "unknown" => Unknown,
            "count" => Count,
            "percent" => Percent,
            "sum" => Sum,
            "captionEqual" => CaptionEqual,
            "captionNotEqual" => CaptionNotEqual,
            "captionBeginsWith" => CaptionBeginsWith,
            "captionNotBeginsWith" => CaptionNotBeginsWith,
            "captionEndsWith" => CaptionEndsWith,
            "captionNotEndsWith" => CaptionNotEndsWith,
            "captionContains" => CaptionContains,
            "captionNotContains" => CaptionNotContains,
            "captionGreaterThan" => CaptionGreaterThan,
            "captionGreaterThanOrEqual" => CaptionGreaterThanOrEqual,
            "captionLessThan" => CaptionLessThan,
            "captionLessThanOrEqual" => CaptionLessThanOrEqual,
            "captionBetween" => CaptionBetween,
            "captionNotBetween" => CaptionNotBetween,
            "valueEqual" => ValueEqual,
            "valueNotEqual" => ValueNotEqual,
            "valueGreaterThan" => ValueGreaterThan,
            "valueGreaterThanOrEqual" => ValueGreaterThanOrEqual,
            "valueLessThan" => ValueLessThan,
            "valueLessThanOrEqual" => ValueLessThanOrEqual,
            "valueBetween" => ValueBetween,
            "valueNotBetween" => ValueNotBetween,
            "dateEqual" => DateEqual,
            "dateNotEqual" => DateNotEqual,
            "dateOlderThan" => DateOlderThan,
            "dateOlderThanOrEqual" => DateOlderThanOrEqual,
            "dateNewerThan" => DateNewerThan,
            "dateNewerThanOrEqual" => DateNewerThanOrEqual,
            "dateBetween" => DateBetween,
            "dateNotBetween" => DateNotBetween,
            "tomorrow" => Tomorrow,
            "today" => Today,
            "yesterday" => Yesterday,
            "nextWeek" => NextWeek,
            "thisWeek" => ThisWeek,
            "lastWeek" => LastWeek,
            "nextMonth" => NextMonth,
            "thisMonth" => ThisMonth,
            "lastMonth" => LastMonth,
            "nextQuarter" => NextQuarter,
            "thisQuarter" => ThisQuarter,
            "lastQuarter" => LastQuarter,
            "nextYear" => NextYear,
            "thisYear" => ThisYear,
            "lastYear" => LastYear,
            "yearToDate" => YearToDate,
            "Q1" => Q1,
            "Q2" => Q2,
            "Q3" => Q3,
            "Q4" => Q4,
            "M1" => M1,
            "M2" => M2,
            "M3" => M3,
            "M4" => M4,
            "M5" => M5,
            "M6" => M6,
            "M7" => M7,
            "M8" => M8,
            "M9" => M9,
            "M10" => M10,
            "M11" => M11,
            "M12" => M12,
            _ => return Err(invalid(format!("invalid pivot filter type '{v}'"))),
        })
    }
    pub(super) fn as_str(self) -> &'static str {
        use FilterType::*;
        match self {
            Unknown => "unknown",
            Count => "count",
            Percent => "percent",
            Sum => "sum",
            CaptionEqual => "captionEqual",
            CaptionNotEqual => "captionNotEqual",
            CaptionBeginsWith => "captionBeginsWith",
            CaptionNotBeginsWith => "captionNotBeginsWith",
            CaptionEndsWith => "captionEndsWith",
            CaptionNotEndsWith => "captionNotEndsWith",
            CaptionContains => "captionContains",
            CaptionNotContains => "captionNotContains",
            CaptionGreaterThan => "captionGreaterThan",
            CaptionGreaterThanOrEqual => "captionGreaterThanOrEqual",
            CaptionLessThan => "captionLessThan",
            CaptionLessThanOrEqual => "captionLessThanOrEqual",
            CaptionBetween => "captionBetween",
            CaptionNotBetween => "captionNotBetween",
            ValueEqual => "valueEqual",
            ValueNotEqual => "valueNotEqual",
            ValueGreaterThan => "valueGreaterThan",
            ValueGreaterThanOrEqual => "valueGreaterThanOrEqual",
            ValueLessThan => "valueLessThan",
            ValueLessThanOrEqual => "valueLessThanOrEqual",
            ValueBetween => "valueBetween",
            ValueNotBetween => "valueNotBetween",
            DateEqual => "dateEqual",
            DateNotEqual => "dateNotEqual",
            DateOlderThan => "dateOlderThan",
            DateOlderThanOrEqual => "dateOlderThanOrEqual",
            DateNewerThan => "dateNewerThan",
            DateNewerThanOrEqual => "dateNewerThanOrEqual",
            DateBetween => "dateBetween",
            DateNotBetween => "dateNotBetween",
            Tomorrow => "tomorrow",
            Today => "today",
            Yesterday => "yesterday",
            NextWeek => "nextWeek",
            ThisWeek => "thisWeek",
            LastWeek => "lastWeek",
            NextMonth => "nextMonth",
            ThisMonth => "thisMonth",
            LastMonth => "lastMonth",
            NextQuarter => "nextQuarter",
            ThisQuarter => "thisQuarter",
            LastQuarter => "lastQuarter",
            NextYear => "nextYear",
            ThisYear => "thisYear",
            LastYear => "lastYear",
            YearToDate => "yearToDate",
            Q1 => "Q1",
            Q2 => "Q2",
            Q3 => "Q3",
            Q4 => "Q4",
            M1 => "M1",
            M2 => "M2",
            M3 => "M3",
            M4 => "M4",
            M5 => "M5",
            M6 => "M6",
            M7 => "M7",
            M8 => "M8",
            M9 => "M9",
            M10 => "M10",
            M11 => "M11",
            M12 => "M12",
        }
    }
    fn permits_timeline_filter(self) -> bool {
        !matches!(self, Self::Unknown | Self::DateEqual | Self::DateBetween)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct State {
    pub selection: Option<Range>,
    pub bounds: Range,
    pub extension_list: Option<OpaqueXml>,
    pub single_range_filter_state: Option<bool>,
    pub minimal_refresh_version: u32,
    pub last_refresh_version: u32,
    pub pivot_cache_id: u32,
    pub filter_type: FilterType,
}
impl State {
    pub fn new(bounds: Range, pivot_cache_id: u32, filter_type: FilterType) -> Self {
        Self {
            selection: None,
            bounds,
            extension_list: None,
            single_range_filter_state: None,
            minimal_refresh_version: 0,
            last_refresh_version: 0,
            pivot_cache_id,
            filter_type,
        }
    }
    pub fn effective_single_range_filter_state(&self) -> bool {
        self.single_range_filter_state.unwrap_or(true)
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct PivotFilter {
    pub use_whole_day: Option<bool>,
    pub field: u32,
    pub id: u32,
    pub name: Option<String>,
    pub description: Option<String>,
    pub auto_filter: Option<AutoFilterDefinition>,
}
impl PivotFilter {
    pub fn new(field: u32, id: u32) -> Self {
        Self {
            use_whole_day: None,
            field,
            id,
            name: None,
            description: None,
            auto_filter: None,
        }
    }
    pub fn effective_use_whole_day(&self) -> bool {
        self.use_whole_day.unwrap_or(false)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cache {
    pub relationship_id: String,
    pub part_name: String,
    pub definition: CacheDefinition,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Level {
    Year,
    Quarter,
    Month,
    Day,
}

impl Level {
    pub(super) fn parse(value: &str, name: &str) -> Result<Self> {
        match value {
            "0" => Ok(Self::Year),
            "1" => Ok(Self::Quarter),
            "2" => Ok(Self::Month),
            "3" => Ok(Self::Day),
            _ => Err(invalid(format!("{name} must be between 0 and 3"))),
        }
    }
    pub(super) fn number(self) -> &'static str {
        match self {
            Self::Year => "0",
            Self::Quarter => "1",
            Self::Month => "2",
            Self::Day => "3",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    pub name: String,
    pub uid: Option<String>,
    pub cache: String,
    pub caption: Option<String>,
    pub show_header: Option<bool>,
    pub show_selection_label: Option<bool>,
    pub show_time_level: Option<bool>,
    pub show_horizontal_scrollbar: Option<bool>,
    pub level: Level,
    pub selection_level: Level,
    pub scroll_position: Option<String>,
    pub style: Option<String>,
    pub extension_list: Option<OpaqueXml>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Views {
    pub timelines: Vec<View>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetView {
    pub worksheet_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub timelines: Views,
}

pub(super) fn validate_cache_definition(value: &CacheDefinition) -> Result<()> {
    validate_defined_name(&value.name)?;
    bounded_nonempty(&value.source_name, "sourceName")?;
    if let Some(uid) = &value.uid {
        validate_guid(uid)?;
    }
    if value.pivot_tables.len() > MAX_PIVOT_TABLES {
        return Err(limit("pivot table count"));
    }
    let mut bindings = HashSet::new();
    for pivot in &value.pivot_tables {
        bounded_nonempty(&pivot.name, "pivot table name")?;
        if !bindings.insert((pivot.tab_id, pivot.name.to_lowercase())) {
            return Err(invalid("duplicate View Cache pivotTable binding"));
        }
    }
    validate_state(&value.state)?;
    if let Some(payload) = &value.timeline_pivot_filter {
        validate_timeline_filter(payload)?;
        if !value.state.filter_type.permits_timeline_filter() {
            return Err(invalid(
                "timelinePivotFilter is forbidden for unknown, dateEqual, and dateBetween",
            ));
        }
    }
    if let Some(payload) = &value.extension_list {
        validate_opaque_kind(payload, &[X15, SML, STRICT_SML], "extLst")?;
    }
    let total = value
        .state
        .extension_list
        .as_ref()
        .map_or(0, |x| x.xml.len())
        .checked_add(value.extension_list.as_ref().map_or(0, |x| x.xml.len()))
        .ok_or_else(|| limit("opaque XML bytes"))?;
    if total > MAX_TOTAL_OPAQUE_BYTES {
        return Err(limit("total opaque XML bytes"));
    }
    Ok(())
}

pub(super) fn validate_state(v: &State) -> Result<()> {
    validate_range(&v.bounds)?;
    if let Some(s) = &v.selection {
        validate_range(s)?;
        if compare_datetime(&s.start_date, &v.bounds.start_date)? == std::cmp::Ordering::Less
            || compare_datetime(&s.end_date, &v.bounds.end_date)? == std::cmp::Ordering::Greater
        {
            return Err(invalid("View selection must be within bounds"));
        }
    }
    if let Some(e) = &v.extension_list {
        validate_opaque_kind(e, &[X15, SML, STRICT_SML], "extLst")?
    }
    Ok(())
}
pub(super) fn validate_timeline_filter(v: &PivotFilter) -> Result<()> {
    if let Some(n) = &v.name {
        bounded(n, "timeline pivot filter name")?;
        if n.chars().count() > 65_535 {
            return Err(limit("timeline pivot filter name characters"));
        }
    }
    if let Some(n) = &v.description {
        bounded(n, "timeline pivot filter description")?;
        if n.chars().count() > 65_535 {
            return Err(limit("timeline pivot filter description characters"));
        }
    }
    if let Some(a) = &v.auto_filter {
        write_auto_filter_fragment(a)?;
    }
    Ok(())
}
pub(super) fn validate_range(v: &Range) -> Result<()> {
    validate_datetime(&v.start_date)?;
    validate_datetime(&v.end_date)?;
    if compare_datetime(&v.start_date, &v.end_date)? == std::cmp::Ordering::Greater {
        return Err(invalid("View range startDate exceeds endDate"));
    }
    Ok(())
}
pub(super) fn compare_datetime(a: &str, b: &str) -> Result<std::cmp::Ordering> {
    let a = parse_xsd_datetime(a)?;
    let b = parse_xsd_datetime(b)?;
    match (a.timezone_minutes, b.timezone_minutes) {
        (Some(_), Some(_)) => Ok(a.utc_normalized().cmp_value(&b.utc_normalized())),
        (None, None) => Ok(a.cmp_value(&b)),
        _ => Err(invalid(
            "cannot compare timezone-aware and local View dates",
        )),
    }
}

pub(super) fn validate_cache_set(caches: &[Cache]) -> Result<()> {
    if caches.len() > MAX_CACHES {
        return Err(limit("cache count"));
    }
    let mut names = HashSet::new();
    let mut ids = HashSet::new();
    let mut targets = HashSet::new();
    let mut uids = HashSet::new();
    let any_uid = caches.iter().any(|cache| cache.definition.uid.is_some());
    for cache in caches {
        validate_cache_definition(&cache.definition)?;
        validate_relationship_id(&cache.relationship_id)?;
        if !names.insert(cache.definition.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive View Cache name '{}'",
                cache.definition.name
            )));
        }
        if !ids.insert(cache.relationship_id.clone()) {
            return Err(invalid("duplicate View Cache relationship ID"));
        }
        if !targets.insert(cache.part_name.clone()) {
            return Err(invalid("duplicate View Cache part name"));
        }
        if any_uid && cache.definition.uid.is_none() {
            return Err(invalid(
                "View Cache UIDs must be specified on all caches or none",
            ));
        }
        if let Some(uid) = &cache.definition.uid
            && !uids.insert(uid.to_ascii_lowercase())
        {
            return Err(invalid("duplicate View Cache UID"));
        }
    }
    Ok(())
}

pub(super) fn validate_views_local(value: &Views) -> Result<()> {
    if value.timelines.is_empty() {
        return Err(invalid("Views part must contain at least one timeline"));
    }
    if value.timelines.len() > MAX_TIMELINES {
        return Err(limit("timeline count"));
    }
    let mut names = HashSet::new();
    for view in &value.timelines {
        bounded_nonempty(&view.name, "timeline name")?;
        bounded_nonempty(&view.cache, "timeline cache name")?;
        if view.name.chars().count() > 32767 {
            return Err(invalid("timeline name exceeds 32767 characters"));
        }
        if !names.insert(view.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive timeline name '{}'",
                view.name
            )));
        }
        if let Some(uid) = &view.uid {
            validate_guid(uid)?;
        }
        if let Some(caption) = &view.caption {
            bounded_nonempty(caption, "timeline caption")?;
        }
        if let Some(position) = &view.scroll_position {
            validate_datetime(position)?;
        }
        if let Some(style) = &view.style {
            bounded_nonempty(style, "timeline style")?;
        }
        if let Some(payload) = &view.extension_list {
            validate_opaque_kind(payload, &[X15, SML, STRICT_SML], "extLst")?;
        }
    }
    Ok(())
}

pub(super) fn validate_global_views(values: &[WorksheetView]) -> Result<()> {
    let count: usize = values
        .iter()
        .map(|value| value.timelines.timelines.len())
        .sum();
    if count > MAX_TIMELINES {
        return Err(limit("workbook timeline count"));
    }
    let any_uid = values
        .iter()
        .flat_map(|value| &value.timelines.timelines)
        .any(|view| view.uid.is_some());
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    for view in values.iter().flat_map(|value| &value.timelines.timelines) {
        if !names.insert(view.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate workbook-wide timeline name '{}'",
                view.name
            )));
        }
        if any_uid && view.uid.is_none() {
            return Err(invalid(
                "timeline UIDs must be specified on all views or none",
            ));
        }
        if let Some(uid) = &view.uid
            && !uids.insert(uid.to_ascii_lowercase())
        {
            return Err(invalid("duplicate timeline UID"));
        }
    }
    Ok(())
}

fn validate_defined_name(value: &str) -> Result<()> {
    bounded_nonempty(value, "View Cache name")?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first.is_alphabetic() || matches!(first, '_' | '\\'))
        || !chars
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.' | '\\'))
    {
        return Err(invalid(format!(
            "invalid defined-name-shaped View Cache name '{value}'"
        )));
    }
    let upper = value.to_ascii_uppercase();
    if looks_like_a1(&upper) || looks_like_r1c1(&upper) {
        return Err(invalid(format!(
            "View Cache name '{value}' conflicts with a cell reference"
        )));
    }
    Ok(())
}
fn looks_like_a1(value: &str) -> bool {
    let split = value.bytes().position(|byte| byte.is_ascii_digit());
    split.is_some_and(|at| {
        at > 0
            && at <= 3
            && value[..at].bytes().all(|b| b.is_ascii_uppercase())
            && value[at..]
                .parse::<u32>()
                .is_ok_and(|row| (1..=1_048_576).contains(&row))
    })
}
fn looks_like_r1c1(value: &str) -> bool {
    let Some(rest) = value.strip_prefix('R') else {
        return false;
    };
    let Some((row, col)) = rest.split_once('C') else {
        return false;
    };
    !row.is_empty()
        && !col.is_empty()
        && row.bytes().all(|b| b.is_ascii_digit())
        && col.bytes().all(|b| b.is_ascii_digit())
}
fn validate_guid(value: &str) -> Result<()> {
    if !valid_guid(value) {
        Err(invalid(format!("invalid GUID '{value}'")))
    } else {
        Ok(())
    }
}
#[derive(Clone, Debug, Eq, PartialEq)]
struct XsdYear {
    negative: bool,
    digits: String,
}

impl XsdYear {
    fn magnitude(&self) -> &str {
        self.digits.trim_start_matches('0')
    }
    fn cmp_value(&self, other: &Self) -> std::cmp::Ordering {
        use std::cmp::Ordering;
        match (self.negative, other.negative) {
            (true, false) => Ordering::Less,
            (false, true) => Ordering::Greater,
            (false, false) => cmp_decimal(self.magnitude(), other.magnitude()),
            (true, true) => cmp_decimal(other.magnitude(), self.magnitude()),
        }
    }
    fn next(&mut self) {
        if self.negative {
            if decimal_is_one(&self.digits) {
                self.negative = false;
                self.digits = "0001".into();
            } else {
                decimal_sub_one(&mut self.digits);
            }
        } else {
            decimal_add_one(&mut self.digits);
        }
    }
    fn previous(&mut self) {
        if self.negative {
            decimal_add_one(&mut self.digits);
        } else if decimal_is_one(&self.digits) {
            self.negative = true;
            self.digits = "0001".into();
        } else {
            decimal_sub_one(&mut self.digits);
        }
    }
    fn astronomical_mod_400(&self) -> u16 {
        let magnitude = decimal_mod(&self.digits, 400);
        if self.negative {
            (401 - magnitude) % 400
        } else {
            magnitude
        }
    }
}

#[derive(Clone, Debug)]
struct XsdDateTime {
    year: XsdYear,
    month: u8,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
    fraction: String,
    timezone_minutes: Option<i16>,
}

impl XsdDateTime {
    fn shift_day(&mut self, direction: i8) {
        if direction > 0 {
            let last = days_in_month(&self.year, self.month);
            if self.day < last {
                self.day += 1;
            } else {
                self.day = 1;
                if self.month < 12 {
                    self.month += 1;
                } else {
                    self.month = 1;
                    self.year.next();
                }
            }
        } else if self.day > 1 {
            self.day -= 1;
        } else {
            if self.month > 1 {
                self.month -= 1;
            } else {
                self.month = 12;
                self.year.previous();
            }
            self.day = days_in_month(&self.year, self.month);
        }
    }
    fn utc_normalized(&self) -> Self {
        let mut result = self.clone();
        let mut minutes = i16::from(result.hour) * 60 + i16::from(result.minute)
            - result.timezone_minutes.unwrap_or(0);
        if minutes < 0 {
            minutes += 1440;
            result.shift_day(-1);
        } else if minutes >= 1440 {
            minutes -= 1440;
            result.shift_day(1);
        }
        result.hour = (minutes / 60) as u8;
        result.minute = (minutes % 60) as u8;
        result.timezone_minutes = Some(0);
        result
    }
    fn cmp_value(&self, other: &Self) -> std::cmp::Ordering {
        self.year
            .cmp_value(&other.year)
            .then_with(|| self.month.cmp(&other.month))
            .then_with(|| self.day.cmp(&other.day))
            .then_with(|| self.hour.cmp(&other.hour))
            .then_with(|| self.minute.cmp(&other.minute))
            .then_with(|| self.second.cmp(&other.second))
            .then_with(|| cmp_fraction(&self.fraction, &other.fraction))
    }
}

fn cmp_decimal(a: &str, b: &str) -> std::cmp::Ordering {
    a.len()
        .cmp(&b.len())
        .then_with(|| a.as_bytes().cmp(b.as_bytes()))
}
fn decimal_is_one(value: &str) -> bool {
    value.trim_start_matches('0') == "1"
}
fn decimal_mod(value: &str, modulus: u16) -> u16 {
    value
        .bytes()
        .fold(0, |acc, b| (acc * 10 + u16::from(b - b'0')) % modulus)
}
fn decimal_add_one(value: &mut String) {
    let mut bytes = Vec::with_capacity(value.len().saturating_add(1));
    bytes.extend_from_slice(value.as_bytes());
    let mut carry = true;
    for byte in bytes.iter_mut().rev() {
        if !carry {
            break;
        }
        if *byte == b'9' {
            *byte = b'0';
        } else {
            *byte += 1;
            carry = false;
        }
    }
    if carry {
        bytes.insert(0, b'1');
    }
    *value = String::from_utf8(bytes).expect("decimal digits remain valid UTF-8");
}
fn decimal_sub_one(value: &mut String) {
    let mut bytes = value.as_bytes().to_vec();
    for byte in bytes.iter_mut().rev() {
        if *byte == b'0' {
            *byte = b'9';
        } else {
            *byte -= 1;
            break;
        }
    }
    if bytes.len() > 4 && bytes.first() == Some(&b'0') {
        bytes.remove(0);
    }
    *value = String::from_utf8(bytes).expect("decimal digits remain valid UTF-8");
}
fn is_leap_year(year: &XsdYear) -> bool {
    let value = year.astronomical_mod_400();
    value.is_multiple_of(4) && (!value.is_multiple_of(100) || value == 0)
}
fn days_in_month(year: &XsdYear, month: u8) -> u8 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    }
}
fn cmp_fraction(a: &str, b: &str) -> std::cmp::Ordering {
    let length = a.len().max(b.len());
    (0..length)
        .map(|index| {
            a.as_bytes()
                .get(index)
                .copied()
                .unwrap_or(b'0')
                .cmp(&b.as_bytes().get(index).copied().unwrap_or(b'0'))
        })
        .find(|order| !order.is_eq())
        .unwrap_or(std::cmp::Ordering::Equal)
}

fn parse_xsd_datetime(value: &str) -> Result<XsdDateTime> {
    bounded_nonempty(value, "dateTime")?;
    let t = value
        .find('T')
        .ok_or_else(|| invalid(format!("invalid xsd:dateTime '{value}'")))?;
    if value[t + 1..].contains('T') {
        return Err(invalid(format!("invalid xsd:dateTime '{value}'")));
    }
    let date = &value[..t];
    let mut time = &value[t + 1..];
    let timezone_minutes = if let Some(stripped) = time.strip_suffix('Z') {
        time = stripped;
        Some(0)
    } else {
        let sign = time.rfind(['+', '-']);
        if let Some(index) = sign {
            let zone = &time[index..];
            time = &time[..index];
            if zone.len() != 6
                || zone.as_bytes()[3] != b':'
                || !zone[1..3].bytes().all(|b| b.is_ascii_digit())
                || !zone[4..].bytes().all(|b| b.is_ascii_digit())
            {
                return Err(invalid("invalid xsd:dateTime timezone"));
            }
            let hours = zone[1..3].parse::<i16>().unwrap();
            let minutes = zone[4..].parse::<i16>().unwrap();
            if hours > 14 || minutes > 59 || (hours == 14 && minutes != 0) {
                return Err(invalid("xsd:dateTime timezone exceeds 14:00"));
            }
            let offset = hours * 60 + minutes;
            Some(if zone.starts_with('-') {
                -offset
            } else {
                offset
            })
        } else {
            None
        }
    };
    let negative = date.starts_with('-');
    let year_start = usize::from(negative);
    let year_end = date[year_start..]
        .find('-')
        .map(|index| year_start + index)
        .ok_or_else(|| invalid(format!("invalid xsd:dateTime '{value}'")))?;
    let year_digits = &date[year_start..year_end];
    if year_digits.len() < 4
        || !year_digits.bytes().all(|b| b.is_ascii_digit())
        || year_digits.bytes().all(|b| b == b'0')
        || (year_digits.len() > 4 && year_digits.starts_with('0'))
    {
        return Err(invalid("invalid xsd:dateTime year"));
    }
    let date_tail = &date[year_end + 1..];
    if date_tail.len() != 5
        || date_tail.as_bytes()[2] != b'-'
        || !date_tail[..2].bytes().all(|b| b.is_ascii_digit())
        || !date_tail[3..].bytes().all(|b| b.is_ascii_digit())
    {
        return Err(invalid(format!("invalid xsd:dateTime '{value}'")));
    }
    if time.len() < 8
        || time.as_bytes().get(2) != Some(&b':')
        || time.as_bytes().get(5) != Some(&b':')
        || !time[..2].bytes().all(|b| b.is_ascii_digit())
        || !time[3..5].bytes().all(|b| b.is_ascii_digit())
        || !time[6..8].bytes().all(|b| b.is_ascii_digit())
    {
        return Err(invalid(format!("invalid xsd:dateTime '{value}'")));
    }
    let fraction = if time.len() == 8 {
        ""
    } else if time.as_bytes().get(8) == Some(&b'.')
        && time.len() > 9
        && time[9..].bytes().all(|b| b.is_ascii_digit())
    {
        &time[9..]
    } else {
        return Err(invalid(format!("invalid xsd:dateTime '{value}'")));
    };
    let month = date_tail[..2].parse::<u8>().unwrap();
    let day = date_tail[3..].parse::<u8>().unwrap();
    let hour = time[..2].parse::<u8>().unwrap();
    let minute = time[3..5].parse::<u8>().unwrap();
    let second = time[6..8].parse::<u8>().unwrap();
    let year = XsdYear {
        negative,
        digits: year_digits.into(),
    };
    if month == 0
        || month > 12
        || day == 0
        || day > days_in_month(&year, month)
        || hour > 24
        || minute > 59
        || second > 59
        || (hour == 24 && (minute != 0 || second != 0 || fraction.bytes().any(|b| b != b'0')))
    {
        return Err(invalid(format!("invalid xsd:dateTime '{value}'")));
    }
    let mut result = XsdDateTime {
        year,
        month,
        day,
        hour,
        minute,
        second,
        fraction: fraction.into(),
        timezone_minutes,
    };
    if result.hour == 24 {
        result.hour = 0;
        result.shift_day(1);
    }
    Ok(result)
}

pub(super) fn validate_datetime(value: &str) -> Result<()> {
    parse_xsd_datetime(value).map(|_| ())
}
pub(super) fn validate_relationship_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
