//! Immutable static worksheet auto-filter and sort-state read model.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use crate::xlsx::sort::{SortBy, SortMethod};
use quick_xml::Writer;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_COLUMNS: usize = 16_384;
const MAX_ITEMS: usize = 10_000;
const MAX_SORT_CONDITIONS: usize = 64;
const MAX_FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FilterRange(String);
impl FilterRange {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        parse_range(&value)?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalendarType {
    None,
    Gregorian,
    GregorianUs,
    GregorianMeFrench,
    GregorianArabic,
    Hijri,
    Hebrew,
    Taiwan,
    Japan,
    Thai,
    Korea,
    Saka,
}
impl CalendarType {
    fn parse(v: &str) -> Result<Self> {
        match v {
            "none" => Ok(Self::None),
            "gregorian" => Ok(Self::Gregorian),
            "gregorianUs" => Ok(Self::GregorianUs),
            "gregorianMeFrench" => Ok(Self::GregorianMeFrench),
            "gregorianArabic" => Ok(Self::GregorianArabic),
            "hijri" => Ok(Self::Hijri),
            "hebrew" => Ok(Self::Hebrew),
            "taiwan" => Ok(Self::Taiwan),
            "japan" => Ok(Self::Japan),
            "thai" => Ok(Self::Thai),
            "korea" => Ok(Self::Korea),
            "saka" => Ok(Self::Saka),
            _ => Err(invalid(format!("invalid calendarType '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateTimeGrouping {
    Year,
    Month,
    Day,
    Hour,
    Minute,
    Second,
}
impl DateTimeGrouping {
    fn parse(v: &str) -> Result<Self> {
        match v {
            "year" => Ok(Self::Year),
            "month" => Ok(Self::Month),
            "day" => Ok(Self::Day),
            "hour" => Ok(Self::Hour),
            "minute" => Ok(Self::Minute),
            "second" => Ok(Self::Second),
            _ => Err(invalid(format!("invalid dateTimeGrouping '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateGroupItem {
    year: u16,
    month: Option<u8>,
    day: Option<u8>,
    hour: Option<u8>,
    minute: Option<u8>,
    second: Option<u8>,
    grouping: DateTimeGrouping,
}
impl DateGroupItem {
    pub fn year(&self) -> u16 {
        self.year
    }
    pub fn month(&self) -> Option<u8> {
        self.month
    }
    pub fn day(&self) -> Option<u8> {
        self.day
    }
    pub fn hour(&self) -> Option<u8> {
        self.hour
    }
    pub fn minute(&self) -> Option<u8> {
        self.minute
    }
    pub fn second(&self) -> Option<u8> {
        self.second
    }
    pub fn grouping(&self) -> DateTimeGrouping {
        self.grouping
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FilterItem {
    Value(String),
    DateGroup(DateGroupItem),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FilterValues {
    blank: bool,
    calendar_type: CalendarType,
    items: Vec<FilterItem>,
}
impl FilterValues {
    pub fn blank(&self) -> bool {
        self.blank
    }
    pub fn calendar_type(&self) -> CalendarType {
        self.calendar_type
    }
    pub fn items(&self) -> &[FilterItem] {
        &self.items
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomFilterOperator {
    LessThan,
    LessThanOrEqual,
    NotEqual,
    Equal,
    GreaterThanOrEqual,
    GreaterThan,
}
impl CustomFilterOperator {
    fn parse(v: &str) -> Result<Self> {
        match v {
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "notEqual" => Ok(Self::NotEqual),
            "equal" => Ok(Self::Equal),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            _ => Err(invalid(format!("invalid custom-filter operator '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomFilter {
    operator: CustomFilterOperator,
    value: String,
}
impl CustomFilter {
    pub fn operator(&self) -> CustomFilterOperator {
        self.operator
    }
    pub fn value(&self) -> &str {
        &self.value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomFilters {
    and: bool,
    filters: Vec<CustomFilter>,
}
impl CustomFilters {
    pub fn and(&self) -> bool {
        self.and
    }
    pub fn filters(&self) -> &[CustomFilter] {
        &self.filters
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DynamicFilterType {
    AboveAverage,
    BelowAverage,
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
    Null,
}
impl DynamicFilterType {
    fn parse(v: &str) -> Result<Self> {
        use DynamicFilterType::*;
        match v {
            "aboveAverage" => Ok(AboveAverage),
            "belowAverage" => Ok(BelowAverage),
            "tomorrow" => Ok(Tomorrow),
            "today" => Ok(Today),
            "yesterday" => Ok(Yesterday),
            "nextWeek" => Ok(NextWeek),
            "thisWeek" => Ok(ThisWeek),
            "lastWeek" => Ok(LastWeek),
            "nextMonth" => Ok(NextMonth),
            "thisMonth" => Ok(ThisMonth),
            "lastMonth" => Ok(LastMonth),
            "nextQuarter" => Ok(NextQuarter),
            "thisQuarter" => Ok(ThisQuarter),
            "lastQuarter" => Ok(LastQuarter),
            "nextYear" => Ok(NextYear),
            "thisYear" => Ok(ThisYear),
            "lastYear" => Ok(LastYear),
            "yearToDate" => Ok(YearToDate),
            "Q1" => Ok(Q1),
            "Q2" => Ok(Q2),
            "Q3" => Ok(Q3),
            "Q4" => Ok(Q4),
            "M1" => Ok(M1),
            "M2" => Ok(M2),
            "M3" => Ok(M3),
            "M4" => Ok(M4),
            "M5" => Ok(M5),
            "M6" => Ok(M6),
            "M7" => Ok(M7),
            "M8" => Ok(M8),
            "M9" => Ok(M9),
            "M10" => Ok(M10),
            "M11" => Ok(M11),
            "M12" => Ok(M12),
            "null" => Ok(Null),
            _ => Err(invalid(format!("invalid dynamic-filter type '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DynamicFilter {
    filter_type: DynamicFilterType,
    value: Option<f64>,
    max_value: Option<f64>,
}
impl DynamicFilter {
    pub fn filter_type(&self) -> DynamicFilterType {
        self.filter_type
    }
    pub fn value(&self) -> Option<f64> {
        self.value
    }
    pub fn max_value(&self) -> Option<f64> {
        self.max_value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorFilter {
    differential_format_id: u32,
    cell_color: bool,
}
impl ColorFilter {
    pub fn differential_format_id(&self) -> u32 {
        self.differential_format_id
    }
    pub fn cell_color(&self) -> bool {
        self.cell_color
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterIconSet {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}
impl FilterIconSet {
    fn parse(v: &str) -> Result<Self> {
        use FilterIconSet::*;
        match v {
            "3Arrows" => Ok(ThreeArrows),
            "3ArrowsGray" => Ok(ThreeArrowsGray),
            "3Flags" => Ok(ThreeFlags),
            "3TrafficLights1" => Ok(ThreeTrafficLights1),
            "3TrafficLights2" => Ok(ThreeTrafficLights2),
            "3Signs" => Ok(ThreeSigns),
            "3Symbols" => Ok(ThreeSymbols),
            "3Symbols2" => Ok(ThreeSymbols2),
            "4Arrows" => Ok(FourArrows),
            "4ArrowsGray" => Ok(FourArrowsGray),
            "4RedToBlack" => Ok(FourRedToBlack),
            "4Rating" => Ok(FourRating),
            "4TrafficLights" => Ok(FourTrafficLights),
            "5Arrows" => Ok(FiveArrows),
            "5ArrowsGray" => Ok(FiveArrowsGray),
            "5Rating" => Ok(FiveRating),
            "5Quarters" => Ok(FiveQuarters),
            _ => Err(invalid(format!("invalid icon set '{v}'"))),
        }
    }
    fn cardinality(self) -> u32 {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => 3,
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => 4,
            _ => 5,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IconFilter {
    icon_set: FilterIconSet,
    icon_id: Option<u32>,
}
impl IconFilter {
    pub fn icon_set(&self) -> FilterIconSet {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Top10Filter {
    top: bool,
    percent: bool,
    value: f64,
    filter_value: Option<f64>,
}
impl Top10Filter {
    pub fn top(&self) -> bool {
        self.top
    }
    pub fn percent(&self) -> bool {
        self.percent
    }
    pub fn value(&self) -> f64 {
        self.value
    }
    pub fn filter_value(&self) -> Option<f64> {
        self.filter_value
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FilterColumnPayload {
    Values(FilterValues),
    Custom(CustomFilters),
    Dynamic(DynamicFilter),
    Color(ColorFilter),
    Icon(IconFilter),
    Top10(Top10Filter),
}

#[derive(Debug, Clone, PartialEq)]
pub struct FilterColumnDefinition {
    pub column_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    pub payload: Option<FilterColumnPayload>,
}
impl FilterColumnDefinition {
    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    pub fn hidden_button(&self) -> bool {
        self.hidden_button
    }
    pub fn show_button(&self) -> bool {
        self.show_button
    }
    pub fn payload(&self) -> Option<&FilterColumnPayload> {
        self.payload.as_ref()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortConditionDefinition {
    reference: FilterRange,
    descending: bool,
    sort_by: SortBy,
    custom_list: Option<String>,
    differential_format_id: Option<u32>,
    icon_set: Option<FilterIconSet>,
    icon_id: Option<u32>,
}
impl SortConditionDefinition {
    pub fn reference(&self) -> &FilterRange {
        &self.reference
    }
    pub fn descending(&self) -> bool {
        self.descending
    }
    pub fn sort_by(&self) -> SortBy {
        self.sort_by
    }
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    pub fn icon_set(&self) -> Option<FilterIconSet> {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortStateDefinition {
    reference: FilterRange,
    column_sort: bool,
    case_sensitive: bool,
    sort_method: Option<SortMethod>,
    conditions: Vec<SortConditionDefinition>,
}
impl SortStateDefinition {
    pub fn reference(&self) -> &FilterRange {
        &self.reference
    }
    pub fn column_sort(&self) -> bool {
        self.column_sort
    }
    pub fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }
    pub fn sort_method(&self) -> Option<SortMethod> {
        self.sort_method
    }
    pub fn conditions(&self) -> &[SortConditionDefinition] {
        &self.conditions
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilterDefinition {
    pub reference: Option<FilterRange>,
    pub columns: Vec<FilterColumnDefinition>,
    pub sort_state: Option<SortStateDefinition>,
}
impl AutoFilterDefinition {
    pub fn new(reference: Option<FilterRange>) -> Self {
        Self {
            reference,
            columns: Vec::new(),
            sort_state: None,
        }
    }
    pub fn reference(&self) -> Option<&FilterRange> {
        self.reference.as_ref()
    }
    pub fn columns(&self) -> &[FilterColumnDefinition] {
        &self.columns
    }
    pub fn sort_state(&self) -> Option<&SortStateDefinition> {
        self.sort_state.as_ref()
    }
}

struct ColumnBuilder {
    column_id: u32,
    hidden_button: bool,
    show_button: bool,
    payload: Option<FilterColumnPayload>,
}
struct ValuesBuilder {
    blank: bool,
    calendar_type: CalendarType,
    items: Vec<FilterItem>,
}
struct CustomBuilder {
    and: bool,
    filters: Vec<CustomFilter>,
}
struct SortBuilder {
    reference: FilterRange,
    column_sort: bool,
    case_sensitive: bool,
    sort_method: Option<SortMethod>,
    conditions: Vec<SortConditionDefinition>,
}

pub(crate) fn parse_auto_filter(xml: &[u8]) -> Result<Option<AutoFilterDefinition>> {
    let processed = process_ooxml(xml)?;
    let Some(fragment) = capture(processed.as_ref())? else {
        return Ok(None);
    };
    parse_fragment(&fragment).map(Some)
}
pub(crate) fn parse_auto_filter_fragment(xml: &[u8]) -> Result<AutoFilterDefinition> {
    if xml.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("autoFilter is too large"));
    }
    parse_fragment(xml)
}

pub(crate) fn write_auto_filter_fragment(value: &AutoFilterDefinition) -> Result<Vec<u8>> {
    let mut x = Vec::new();
    x.extend_from_slice(
        b"<x:autoFilter xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    );
    if let Some(r) = &value.reference {
        parse_range(r.as_str())?;
        a(&mut x, "ref", r.as_str());
    }
    if value.columns.is_empty() && value.sort_state.is_none() {
        x.extend_from_slice(b"/>");
        return Ok(x);
    }
    x.push(b'>');
    let mut ids = HashSet::new();
    for c in &value.columns {
        if c.column_id >= MAX_COLUMNS as u32 || !ids.insert(c.column_id) {
            return Err(invalid("invalid or duplicate filterColumn colId"));
        }
        x.extend_from_slice(b"<x:filterColumn");
        a(&mut x, "colId", &c.column_id.to_string());
        if c.hidden_button {
            a(&mut x, "hiddenButton", "1")
        }
        if !c.show_button {
            a(&mut x, "showButton", "0")
        }
        if let Some(p) = &c.payload {
            x.push(b'>');
            write_payload(&mut x, p)?;
            x.extend_from_slice(b"</x:filterColumn>")
        } else {
            x.extend_from_slice(b"/>")
        }
    }
    if let Some(s) = &value.sort_state {
        x.extend_from_slice(b"<x:sortState");
        a(&mut x, "ref", s.reference.as_str());
        if s.column_sort {
            a(&mut x, "columnSort", "1")
        }
        if s.case_sensitive {
            a(&mut x, "caseSensitive", "1")
        }
        if let Some(v) = s.sort_method {
            a(&mut x, "sortMethod", v.as_str())
        }
        if s.conditions.is_empty() {
            x.extend_from_slice(b"/>")
        } else {
            x.push(b'>');
            for c in &s.conditions {
                x.extend_from_slice(b"<x:sortCondition");
                a(&mut x, "ref", c.reference.as_str());
                if c.descending {
                    a(&mut x, "descending", "1")
                }
                if c.sort_by != SortBy::Value {
                    a(&mut x, "sortBy", c.sort_by.as_str())
                }
                if let Some(v) = &c.custom_list {
                    a(&mut x, "customList", v)
                }
                if let Some(v) = c.differential_format_id {
                    a(&mut x, "dxfId", &v.to_string())
                }
                if let Some(v) = c.icon_set {
                    a(&mut x, "iconSet", icon(v))
                }
                if let Some(v) = c.icon_id {
                    a(&mut x, "iconId", &v.to_string())
                }
                x.extend_from_slice(b"/>");
            }
            x.extend_from_slice(b"</x:sortState>")
        }
    }
    x.extend_from_slice(b"</x:autoFilter>");
    if x.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("autoFilter is too large"));
    }
    Ok(x)
}
fn write_payload(x: &mut Vec<u8>, p: &FilterColumnPayload) -> Result<()> {
    match p {
        FilterColumnPayload::Values(v) => {
            x.extend_from_slice(b"<x:filters");
            if v.blank {
                a(x, "blank", "1")
            }
            if v.calendar_type != CalendarType::None {
                a(x, "calendarType", calendar(v.calendar_type))
            }
            if v.items.is_empty() {
                x.extend_from_slice(b"/>")
            } else {
                x.push(b'>');
                for i in &v.items {
                    match i {
                        FilterItem::Value(v) => {
                            x.extend_from_slice(b"<x:filter");
                            a(x, "val", v);
                            x.extend_from_slice(b"/>")
                        },
                        FilterItem::DateGroup(v) => {
                            x.extend_from_slice(b"<x:dateGroupItem");
                            a(x, "year", &v.year.to_string());
                            for (n, o) in [
                                ("month", v.month),
                                ("day", v.day),
                                ("hour", v.hour),
                                ("minute", v.minute),
                                ("second", v.second),
                            ] {
                                if let Some(q) = o {
                                    a(x, n, &q.to_string())
                                }
                            }
                            a(x, "dateTimeGrouping", group(v.grouping));
                            x.extend_from_slice(b"/>")
                        },
                    }
                }
                x.extend_from_slice(b"</x:filters>")
            }
        },
        FilterColumnPayload::Custom(v) => {
            if !(1..=2).contains(&v.filters.len()) {
                return Err(invalid("customFilters requires one or two filters"));
            }
            x.extend_from_slice(b"<x:customFilters");
            if v.and {
                a(x, "and", "1")
            }
            x.push(b'>');
            for f in &v.filters {
                x.extend_from_slice(b"<x:customFilter");
                if f.operator != CustomFilterOperator::Equal {
                    a(x, "operator", custom(f.operator))
                }
                a(x, "val", &f.value);
                x.extend_from_slice(b"/>")
            }
            x.extend_from_slice(b"</x:customFilters>")
        },
        FilterColumnPayload::Dynamic(v) => {
            x.extend_from_slice(b"<x:dynamicFilter");
            a(x, "type", dynamic(v.filter_type));
            if let Some(q) = v.value {
                a(x, "val", &q.to_string())
            }
            if let Some(q) = v.max_value {
                a(x, "maxVal", &q.to_string())
            }
            x.extend_from_slice(b"/>")
        },
        FilterColumnPayload::Color(v) => {
            x.extend_from_slice(b"<x:colorFilter");
            a(x, "dxfId", &v.differential_format_id.to_string());
            if !v.cell_color {
                a(x, "cellColor", "0")
            }
            x.extend_from_slice(b"/>")
        },
        FilterColumnPayload::Icon(v) => {
            x.extend_from_slice(b"<x:iconFilter");
            a(x, "iconSet", icon(v.icon_set));
            if let Some(q) = v.icon_id {
                a(x, "iconId", &q.to_string())
            }
            x.extend_from_slice(b"/>")
        },
        FilterColumnPayload::Top10(v) => {
            x.extend_from_slice(b"<x:top10");
            if !v.top {
                a(x, "top", "0")
            }
            if v.percent {
                a(x, "percent", "1")
            }
            a(x, "val", &v.value.to_string());
            if let Some(q) = v.filter_value {
                a(x, "filterVal", &q.to_string())
            }
            x.extend_from_slice(b"/>")
        },
    }
    Ok(())
}
fn a(x: &mut Vec<u8>, n: &str, v: &str) {
    x.push(b' ');
    x.extend_from_slice(n.as_bytes());
    x.extend_from_slice(b"=\"");
    for c in v.bytes() {
        match c {
            b'&' => x.extend_from_slice(b"&amp;"),
            b'<' => x.extend_from_slice(b"&lt;"),
            b'"' => x.extend_from_slice(b"&quot;"),
            _ => x.push(c),
        }
    }
    x.push(b'"')
}
fn calendar(v: CalendarType) -> &'static str {
    match v {
        CalendarType::None => "none",
        CalendarType::Gregorian => "gregorian",
        CalendarType::GregorianUs => "gregorianUs",
        CalendarType::GregorianMeFrench => "gregorianMeFrench",
        CalendarType::GregorianArabic => "gregorianArabic",
        CalendarType::Hijri => "hijri",
        CalendarType::Hebrew => "hebrew",
        CalendarType::Taiwan => "taiwan",
        CalendarType::Japan => "japan",
        CalendarType::Thai => "thai",
        CalendarType::Korea => "korea",
        CalendarType::Saka => "saka",
    }
}
fn group(v: DateTimeGrouping) -> &'static str {
    match v {
        DateTimeGrouping::Year => "year",
        DateTimeGrouping::Month => "month",
        DateTimeGrouping::Day => "day",
        DateTimeGrouping::Hour => "hour",
        DateTimeGrouping::Minute => "minute",
        DateTimeGrouping::Second => "second",
    }
}
fn custom(v: CustomFilterOperator) -> &'static str {
    match v {
        CustomFilterOperator::LessThan => "lessThan",
        CustomFilterOperator::LessThanOrEqual => "lessThanOrEqual",
        CustomFilterOperator::NotEqual => "notEqual",
        CustomFilterOperator::Equal => "equal",
        CustomFilterOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        CustomFilterOperator::GreaterThan => "greaterThan",
    }
}
fn dynamic(v: DynamicFilterType) -> &'static str {
    use DynamicFilterType::*;
    match v {
        AboveAverage => "aboveAverage",
        BelowAverage => "belowAverage",
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
        Null => "null",
    }
}
fn icon(v: FilterIconSet) -> &'static str {
    use FilterIconSet::*;
    match v {
        ThreeArrows => "3Arrows",
        ThreeArrowsGray => "3ArrowsGray",
        ThreeFlags => "3Flags",
        ThreeTrafficLights1 => "3TrafficLights1",
        ThreeTrafficLights2 => "3TrafficLights2",
        ThreeSigns => "3Signs",
        ThreeSymbols => "3Symbols",
        ThreeSymbols2 => "3Symbols2",
        FourArrows => "4Arrows",
        FourArrowsGray => "4ArrowsGray",
        FourRedToBlack => "4RedToBlack",
        FourRating => "4Rating",
        FourTrafficLights => "4TrafficLights",
        FiveArrows => "5Arrows",
        FiveArrowsGray => "5ArrowsGray",
        FiveRating => "5Rating",
        FiveQuarters => "5Quarters",
    }
}

fn capture(xml: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut reader = NsReader::from_reader(xml);
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut result = None;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((depth, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *depth += 1,
                Event::End(_) => *depth -= 1,
                _ => {},
            }
            if *depth == 0 {
                let (_, writer) = capture.take().unwrap();
                let value = writer.into_inner();
                if value.len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("autoFilter is too large"));
                }
                if result.replace(value).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            }
            continue;
        }
        match event {
            Event::Start(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Start(e)).map_err(xml_error)?;
                capture = Some((1, writer));
            },
            Event::Empty(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Empty(e)).map_err(xml_error)?;
                if result.replace(writer.into_inner()).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if capture.is_some() {
        return Err(invalid("unterminated autoFilter"));
    }
    Ok(result)
}

fn parse_fragment(fragment: &[u8]) -> Result<AutoFilterDefinition> {
    let wrapped = wrap(fragment);
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut depth = 0usize;
    let mut root = None;
    let mut closed = false;
    let mut reference = None;
    let mut width = None;
    let mut columns = Vec::new();
    let mut column: Option<(usize, ColumnBuilder)> = None;
    let mut values: Option<(usize, ValuesBuilder)> = None;
    let mut custom: Option<(usize, CustomBuilder)> = None;
    let mut sort: Option<(usize, SortBuilder)> = None;
    let mut sort_state = None;
    let mut phase = 0u8;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) => {
                let name = e.local_name();
                if spreadsheet(&namespace) && name.as_ref() == b"autoFilter" && root.is_none() {
                    depth += 1;
                    root = Some(depth);
                    if let Some(v) = optional_attr(&e, b"ref", decoder)? {
                        let parsed = parse_range(&v)?;
                        width = Some(parsed.2 - parsed.0 + 1);
                        reference = Some(FilterRange(v));
                    }
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"filterColumn"
                    && root == Some(depth)
                {
                    if phase > 0 {
                        return Err(invalid("filterColumn must precede sortState"));
                    }
                    let builder = parse_column(&e, decoder, width)?;
                    depth += 1;
                    column = Some((depth, builder));
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"filters"
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                {
                    ensure_payload_empty(&column)?;
                    depth += 1;
                    values = Some((
                        depth,
                        ValuesBuilder {
                            blank: optional_bool(&e, b"blank", decoder)?.unwrap_or(false),
                            calendar_type: CalendarType::parse(
                                optional_attr(&e, b"calendarType", decoder)?
                                    .as_deref()
                                    .unwrap_or("none"),
                            )?,
                            items: Vec::new(),
                        },
                    ));
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"customFilters"
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                {
                    ensure_payload_empty(&column)?;
                    depth += 1;
                    custom = Some((
                        depth,
                        CustomBuilder {
                            and: optional_bool(&e, b"and", decoder)?.unwrap_or(false),
                            filters: Vec::new(),
                        },
                    ));
                } else if spreadsheet(&namespace)
                    && values.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(name.as_ref(), b"filter" | b"dateGroupItem")
                {
                    push_value(values.as_mut().unwrap(), name.as_ref(), &e, decoder)?;
                    depth += 1;
                } else if spreadsheet(&namespace)
                    && custom.as_ref().is_some_and(|v| v.0 == depth)
                    && name.as_ref() == b"customFilter"
                {
                    push_custom(custom.as_mut().unwrap(), &e, decoder)?;
                    depth += 1;
                } else if spreadsheet(&namespace)
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(
                        name.as_ref(),
                        b"dynamicFilter" | b"colorFilter" | b"iconFilter" | b"top10"
                    )
                {
                    set_simple_payload(column.as_mut().unwrap(), name.as_ref(), &e, decoder)?;
                    depth += 1;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortState"
                    && root == Some(depth)
                {
                    if phase > 0 || sort_state.is_some() {
                        return Err(invalid("duplicate sortState"));
                    }
                    phase = 1;
                    depth += 1;
                    sort = Some((depth, parse_sort_state(&e, decoder)?));
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortCondition"
                    && sort.as_ref().is_some_and(|v| v.0 == depth)
                {
                    push_sort(sort.as_mut().unwrap(), &e, decoder)?;
                    depth += 1;
                } else {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("autoFilter nesting is too deep"))?;
                }
            },
            Event::Empty(e) => {
                let name = e.local_name();
                if spreadsheet(&namespace) && name.as_ref() == b"autoFilter" && root.is_none() {
                    if let Some(v) = optional_attr(&e, b"ref", decoder)? {
                        parse_range(&v)?;
                        reference = Some(FilterRange(v));
                    }
                    closed = true;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"filterColumn"
                    && root == Some(depth)
                {
                    if phase > 0 {
                        return Err(invalid("filterColumn must precede sortState"));
                    }
                    let b = parse_column(&e, decoder, width)?;
                    columns.push(finish_column(b)?);
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"filters"
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                {
                    ensure_payload_empty(&column)?;
                    column.as_mut().unwrap().1.payload =
                        Some(FilterColumnPayload::Values(FilterValues {
                            blank: optional_bool(&e, b"blank", decoder)?.unwrap_or(false),
                            calendar_type: CalendarType::parse(
                                optional_attr(&e, b"calendarType", decoder)?
                                    .as_deref()
                                    .unwrap_or("none"),
                            )?,
                            items: Vec::new(),
                        }));
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"customFilters"
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                {
                    return Err(invalid(
                        "customFilters requires one or two customFilter children",
                    ));
                } else if spreadsheet(&namespace)
                    && values.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(name.as_ref(), b"filter" | b"dateGroupItem")
                {
                    push_value(values.as_mut().unwrap(), name.as_ref(), &e, decoder)?;
                } else if spreadsheet(&namespace)
                    && custom.as_ref().is_some_and(|v| v.0 == depth)
                    && name.as_ref() == b"customFilter"
                {
                    push_custom(custom.as_mut().unwrap(), &e, decoder)?;
                } else if spreadsheet(&namespace)
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(
                        name.as_ref(),
                        b"dynamicFilter" | b"colorFilter" | b"iconFilter" | b"top10"
                    )
                {
                    set_simple_payload(column.as_mut().unwrap(), name.as_ref(), &e, decoder)?;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortState"
                    && root == Some(depth)
                {
                    if phase > 0 || sort_state.is_some() {
                        return Err(invalid("duplicate sortState"));
                    }
                    phase = 1;
                    sort_state = Some(finish_sort(parse_sort_state(&e, decoder)?));
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortCondition"
                    && sort.as_ref().is_some_and(|v| v.0 == depth)
                {
                    push_sort(sort.as_mut().unwrap(), &e, decoder)?;
                }
            },
            Event::End(e) => {
                if values.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"filters"
                {
                    let (_, b) = values.take().unwrap();
                    column.as_mut().unwrap().1.payload =
                        Some(FilterColumnPayload::Values(FilterValues {
                            blank: b.blank,
                            calendar_type: b.calendar_type,
                            items: b.items,
                        }));
                }
                if custom.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"customFilters"
                {
                    let (_, b) = custom.take().unwrap();
                    if !(1..=2).contains(&b.filters.len()) {
                        return Err(invalid(
                            "customFilters requires one or two customFilter children",
                        ));
                    }
                    column.as_mut().unwrap().1.payload =
                        Some(FilterColumnPayload::Custom(CustomFilters {
                            and: b.and,
                            filters: b.filters,
                        }));
                }
                if column.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"filterColumn"
                {
                    let (_, b) = column.take().unwrap();
                    columns.push(finish_column(b)?);
                    if columns.len() > MAX_COLUMNS {
                        return Err(invalid("too many filter columns"));
                    }
                }
                if sort.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"sortState"
                {
                    let (_, b) = sort.take().unwrap();
                    sort_state = Some(finish_sort(b));
                }
                if root == Some(depth) && e.local_name().as_ref() == b"autoFilter" {
                    closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid autoFilter nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !closed || column.is_some() || values.is_some() || custom.is_some() || sort.is_some() {
        return Err(invalid("unterminated autoFilter"));
    }
    let mut ids = HashSet::with_capacity(columns.len());
    if columns.iter().any(|v| !ids.insert(v.column_id)) {
        return Err(invalid("duplicate filterColumn colId"));
    }
    Ok(AutoFilterDefinition {
        reference,
        columns,
        sort_state,
    })
}

fn parse_column(e: &BytesStart<'_>, d: Decoder, width: Option<u32>) -> Result<ColumnBuilder> {
    let id = required_u32(e, b"colId", d)?;
    if id >= MAX_COLUMNS as u32 || width.is_some_and(|w| id >= w) {
        return Err(invalid("filterColumn colId is outside autoFilter range"));
    }
    Ok(ColumnBuilder {
        column_id: id,
        hidden_button: optional_bool(e, b"hiddenButton", d)?.unwrap_or(false),
        show_button: optional_bool(e, b"showButton", d)?.unwrap_or(true),
        payload: None,
    })
}
fn finish_column(b: ColumnBuilder) -> Result<FilterColumnDefinition> {
    Ok(FilterColumnDefinition {
        column_id: b.column_id,
        hidden_button: b.hidden_button,
        show_button: b.show_button,
        payload: b.payload,
    })
}
fn ensure_payload_empty(c: &Option<(usize, ColumnBuilder)>) -> Result<()> {
    if c.as_ref().is_some_and(|v| v.1.payload.is_some()) {
        Err(invalid("filterColumn has multiple filter payloads"))
    } else {
        Ok(())
    }
}

fn push_value(
    v: &mut (usize, ValuesBuilder),
    name: &[u8],
    e: &BytesStart<'_>,
    d: Decoder,
) -> Result<()> {
    if v.1.items.len() == MAX_ITEMS {
        return Err(invalid("too many filter values"));
    }
    if name == b"filter" {
        let value = required_attr(e, b"val", d)?;
        bounded(&value)?;
        v.1.items.push(FilterItem::Value(value));
    } else {
        v.1.items
            .push(FilterItem::DateGroup(parse_date_group(e, d)?));
    }
    Ok(())
}
fn parse_date_group(e: &BytesStart<'_>, d: Decoder) -> Result<DateGroupItem> {
    let grouping = DateTimeGrouping::parse(&required_attr(e, b"dateTimeGrouping", d)?)?;
    let year = required_u32(e, b"year", d)?;
    if year > 9999 {
        return Err(invalid("date-group year is out of range"));
    }
    let month = small(e, b"month", d, 1, 12)?;
    let day = small(e, b"day", d, 1, 31)?;
    let hour = small(e, b"hour", d, 0, 23)?;
    let minute = small(e, b"minute", d, 0, 59)?;
    let second = small(e, b"second", d, 0, 59)?;
    let required = match grouping {
        DateTimeGrouping::Year => 0,
        DateTimeGrouping::Month => 1,
        DateTimeGrouping::Day => 2,
        DateTimeGrouping::Hour => 3,
        DateTimeGrouping::Minute => 4,
        DateTimeGrouping::Second => 5,
    };
    let present = [month, day, hour, minute, second]
        .iter()
        .take(required)
        .all(Option::is_some);
    if !present {
        return Err(invalid("date-group components do not match grouping"));
    }
    Ok(DateGroupItem {
        year: year as u16,
        month,
        day,
        hour,
        minute,
        second,
        grouping,
    })
}
fn small(e: &BytesStart<'_>, n: &[u8], d: Decoder, min: u8, max: u8) -> Result<Option<u8>> {
    optional_u32(e, n, d)?
        .map(|v| {
            u8::try_from(v)
                .ok()
                .filter(|v| (*v >= min) && (*v <= max))
                .ok_or_else(|| invalid(format!("{} is out of range", String::from_utf8_lossy(n))))
        })
        .transpose()
}

fn push_custom(v: &mut (usize, CustomBuilder), e: &BytesStart<'_>, d: Decoder) -> Result<()> {
    if v.1.filters.len() == 2 {
        return Err(invalid("customFilters has more than two conditions"));
    }
    let value = required_attr(e, b"val", d)?;
    bounded(&value)?;
    v.1.filters.push(CustomFilter {
        operator: CustomFilterOperator::parse(
            optional_attr(e, b"operator", d)?
                .as_deref()
                .unwrap_or("equal"),
        )?,
        value,
    });
    Ok(())
}
fn set_simple_payload(
    c: &mut (usize, ColumnBuilder),
    name: &[u8],
    e: &BytesStart<'_>,
    d: Decoder,
) -> Result<()> {
    if c.1.payload.is_some() {
        return Err(invalid("filterColumn has multiple filter payloads"));
    }
    c.1.payload = Some(match name {
        b"dynamicFilter" => {
            let value = optional_f64(e, b"val", d)?;
            let max_value = optional_f64(e, b"maxVal", d)?;
            FilterColumnPayload::Dynamic(DynamicFilter {
                filter_type: DynamicFilterType::parse(&required_attr(e, b"type", d)?)?,
                value,
                max_value,
            })
        },
        b"colorFilter" => FilterColumnPayload::Color(ColorFilter {
            differential_format_id: required_u32(e, b"dxfId", d)?,
            cell_color: optional_bool(e, b"cellColor", d)?.unwrap_or(true),
        }),
        b"iconFilter" => {
            let set = FilterIconSet::parse(&required_attr(e, b"iconSet", d)?)?;
            let id = optional_u32(e, b"iconId", d)?;
            if id.is_some_and(|v| v >= set.cardinality()) {
                return Err(invalid("iconFilter iconId exceeds icon-set cardinality"));
            }
            FilterColumnPayload::Icon(IconFilter {
                icon_set: set,
                icon_id: id,
            })
        },
        b"top10" => {
            let value = required_f64(e, b"val", d)?;
            let percent = optional_bool(e, b"percent", d)?.unwrap_or(false);
            if value < 0.0 || (percent && value > 100.0) {
                return Err(invalid("top10 val is out of range"));
            }
            FilterColumnPayload::Top10(Top10Filter {
                top: optional_bool(e, b"top", d)?.unwrap_or(true),
                percent,
                value,
                filter_value: optional_f64(e, b"filterVal", d)?,
            })
        },
        _ => unreachable!(),
    });
    Ok(())
}

fn parse_sort_state(e: &BytesStart<'_>, d: Decoder) -> Result<SortBuilder> {
    let reference = FilterRange(required_attr(e, b"ref", d)?);
    parse_range(reference.as_str())?;
    let method = optional_attr(e, b"sortMethod", d)?
        .map(|v| SortMethod::parse(&v).ok_or_else(|| invalid(format!("invalid sortMethod '{v}'"))))
        .transpose()?;
    Ok(SortBuilder {
        reference,
        column_sort: optional_bool(e, b"columnSort", d)?.unwrap_or(false),
        case_sensitive: optional_bool(e, b"caseSensitive", d)?.unwrap_or(false),
        sort_method: method,
        conditions: Vec::new(),
    })
}
fn push_sort(s: &mut (usize, SortBuilder), e: &BytesStart<'_>, d: Decoder) -> Result<()> {
    if s.1.conditions.len() == MAX_SORT_CONDITIONS {
        return Err(invalid("too many sort conditions"));
    }
    let reference = FilterRange(required_attr(e, b"ref", d)?);
    parse_range(reference.as_str())?;
    let sort_by = optional_attr(e, b"sortBy", d)?
        .map(|v| SortBy::parse(&v).ok_or_else(|| invalid(format!("invalid sortBy '{v}'"))))
        .transpose()?
        .unwrap_or(SortBy::Value);
    let dxf = optional_u32(e, b"dxfId", d)?;
    let icon = optional_attr(e, b"iconSet", d)?
        .map(|v| FilterIconSet::parse(&v))
        .transpose()?;
    let icon_id = optional_u32(e, b"iconId", d)?;
    match sort_by {
        SortBy::CellColor | SortBy::FontColor if dxf.is_none() => {
            return Err(invalid("color sort requires dxfId"));
        },
        SortBy::Icon if icon.is_none() => return Err(invalid("icon sort requires iconSet")),
        SortBy::Icon => {
            if icon_id.is_some_and(|v| v >= icon.unwrap().cardinality()) {
                return Err(invalid("sort iconId exceeds icon-set cardinality"));
            }
        },
        _ => {},
    }
    let custom = optional_attr(e, b"customList", d)?;
    if custom.as_ref().is_some_and(|v| v.len() > MAX_TEXT_BYTES) {
        return Err(invalid("custom sort list is too large"));
    }
    s.1.conditions.push(SortConditionDefinition {
        reference,
        descending: optional_bool(e, b"descending", d)?.unwrap_or(false),
        sort_by,
        custom_list: custom,
        differential_format_id: dxf,
        icon_set: icon,
        icon_id,
    });
    Ok(())
}
fn finish_sort(s: SortBuilder) -> SortStateDefinition {
    SortStateDefinition {
        reference: s.reference,
        column_sort: s.column_sort,
        case_sensitive: s.case_sensitive,
        sort_method: s.sort_method,
        conditions: s.conditions,
    }
}

fn parse_range(v: &str) -> Result<(u32, u32, u32, u32)> {
    let mut p = v.split(':');
    let a = parse_cell(p.next().unwrap_or(""))?;
    let b = p.next().map(parse_cell).transpose()?.unwrap_or(a);
    if p.next().is_some() || a.0 > b.0 || a.1 > b.1 {
        return Err(invalid(format!("invalid filter range '{v}'")));
    }
    Ok((a.0, a.1, b.0, b.1))
}
fn parse_cell(v: &str) -> Result<(u32, u32)> {
    let b = v.as_bytes();
    let mut i = 0;
    if i < b.len() && b[i] == b'$' {
        i += 1;
    }
    let start = i;
    while i < b.len() && b[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return Err(invalid("invalid cell reference"));
    }
    let mut col = 0u32;
    for x in &b[start..i] {
        col = col
            .saturating_mul(26)
            .saturating_add(u32::from(x.to_ascii_uppercase() - b'A' + 1));
    }
    if i < b.len() && b[i] == b'$' {
        i += 1;
    }
    let row = std::str::from_utf8(&b[i..])
        .ok()
        .and_then(|v| v.parse::<u32>().ok())
        .ok_or_else(|| invalid("invalid cell row"))?;
    if !(1..=16384).contains(&col) || !(1..=1_048_576).contains(&row) {
        return Err(invalid("cell reference is out of range"));
    }
    Ok((col, row))
}

fn optional_attr(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<Option<String>> {
    let mut r = None;
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() == n {
            if r.is_some() {
                return Err(invalid("duplicate XML attribute"));
            }
            r = Some(
                a.decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, d)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(r)
}
fn required_attr(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<String> {
    optional_attr(e, n, d)?.ok_or_else(|| {
        invalid(format!(
            "missing '{}' attribute",
            String::from_utf8_lossy(n)
        ))
    })
}
fn optional_u32(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<Option<u32>> {
    optional_attr(e, n, d)?
        .map(|v| {
            v.parse()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
fn required_u32(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<u32> {
    optional_u32(e, n, d)?.ok_or_else(|| {
        invalid(format!(
            "missing '{}' attribute",
            String::from_utf8_lossy(n)
        ))
    })
}
fn optional_bool(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<Option<bool>> {
    optional_attr(e, n, d)?
        .map(|v| match v.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
fn optional_f64(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<Option<f64>> {
    optional_attr(e, n, d)?
        .map(|v| {
            let x = v
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid number '{v}'")))?;
            if x.is_finite() {
                Ok(x)
            } else {
                Err(invalid("non-finite filter number"))
            }
        })
        .transpose()
}
fn required_f64(e: &BytesStart<'_>, n: &[u8], d: Decoder) -> Result<f64> {
    optional_f64(e, n, d)?.ok_or_else(|| {
        invalid(format!(
            "missing '{}' attribute",
            String::from_utf8_lossy(n)
        ))
    })
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_TEXT_BYTES {
        Err(invalid("filter value is too large"))
    } else {
        Ok(())
    }
}
fn wrap(f: &[u8]) -> Vec<u8> {
    let mut v=br#"<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main">"#.to_vec();
    v.extend_from_slice(f);
    v.extend_from_slice(b"</root>");
    v
}
fn spreadsheet(ns: &ResolveResult<'_>) -> bool {
    matches!(ns,ResolveResult::Bound(v)if v.as_ref()==CORE||v.as_ref()==STRICT)
}
fn xml_error(e: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(e.to_string())
}
fn invalid(e: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(e.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    fn fixture(path: &str) -> AutoFilterDefinition {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(path)).unwrap();
        let u = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_auto_filter(&p.blob_for(&u).unwrap())
            .unwrap()
            .unwrap()
    }
    #[test]
    fn parses_bundled_fixtures() {
        let custom = fixture("test-data/ooxml/xlsx/autofilter.xlsx");
        assert_eq!(custom.reference.unwrap().as_str(), "A1:C5");
        assert!(matches!(
            custom.columns[0].payload,
            Some(FilterColumnPayload::Custom(_))
        ));
        let values = fixture("test-data/ooxml/xlsx/autofilternamedrange.xlsx");
        assert!(
            matches!(&values.columns[0].payload,Some(FilterColumnPayload::Values(v))if v.items.len()==2)
        );
        let date = fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx");
        assert!(
            matches!(&date.columns[0].payload,Some(FilterColumnPayload::Values(v))if v.items.len()==2)
        );
        let top =
            fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf143068_top10filter.xlsx");
        assert!(
            matches!(&top.columns[0].payload,Some(FilterColumnPayload::Top10(v))if v.value==4.0&&v.filter_value==Some(7.0))
        );
        let buttons =
            fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/autofilterShowButton.xlsx");
        assert_eq!(buttons.columns.len(), 4);
        assert!(buttons.columns.iter().all(|v| !v.show_button));
        for f in [
            "test-data/ooxml/xlsx/sortconditionref.xlsx",
            "test-data/ooxml/xlsx/sortconditionref2.xlsx",
        ] {
            let v = fixture(f);
            assert_eq!(v.sort_state.unwrap().conditions.len(), 1);
        }
    }
    #[test]
    fn parses_all_variants_strict_and_mce() {
        let xml=br#"<s:worksheet xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><s:autoFilter ref="A1:F9"><s:filterColumn colId="0"><s:filters blank="1"><s:filter val="x"/><s:dateGroupItem year="2024" month="2" dateTimeGrouping="month"/></s:filters></s:filterColumn><s:filterColumn colId="1"><s:customFilters and="1"><s:customFilter operator="greaterThan" val="2"/></s:customFilters></s:filterColumn><s:filterColumn colId="2"><s:dynamicFilter type="today" val="4"/></s:filterColumn><s:filterColumn colId="3"><s:colorFilter dxfId="2" cellColor="0"/></s:filterColumn><s:filterColumn colId="4"><mc:AlternateContent><mc:Choice Requires="x14"><x14:iconFilter iconSet="3Arrows" iconId="2"/></mc:Choice><mc:Fallback><s:customFilters><s:customFilter val="fallback"/></s:customFilters></mc:Fallback></mc:AlternateContent></s:filterColumn><s:filterColumn colId="5"><s:top10 percent="1" val="10"/></s:filterColumn><s:sortState ref="A2:F9" caseSensitive="1" sortMethod="none"><s:sortCondition ref="D2:D9" sortBy="cellColor" dxfId="2"/><s:sortCondition ref="E2:E9" sortBy="icon" iconSet="3Arrows" iconId="1"/></s:sortState></s:autoFilter></s:worksheet>"#;
        let v = parse_auto_filter(xml).unwrap().unwrap();
        assert_eq!(v.columns.len(), 6);
        assert!(matches!(
            v.columns[4].payload,
            Some(FilterColumnPayload::Custom(_))
        ));
        let sort = v.sort_state.unwrap();
        assert_eq!(sort.sort_method, Some(SortMethod::None));
        assert_eq!(sort.conditions.len(), 2);
    }
    #[test]
    fn rejects_malformed_and_security_cases() {
        for xml in [
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="B2:A1"/></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:A2"><filterColumn colId="1"/></autoFilter></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"/><filterColumn colId="0"/></autoFilter></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"><customFilters/></filterColumn></autoFilter></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"><top10 percent="1" val="101"/></filterColumn></autoFilter></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter><sortState ref="A1"><sortCondition ref="A1" sortBy="icon"/></sortState></autoFilter></worksheet>"#,
            r#"<!DOCTYPE x><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            r#"<?bad x?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1"><filterColumn colId="0"><filters><filter val="&bogus;"/></filters></filterColumn></autoFilter></worksheet>"#,
        ] {
            assert!(parse_auto_filter(xml.as_bytes()).is_err(), "{xml}");
        }
        let conditions = "<sortCondition ref=\"A1\"/>".repeat(MAX_SORT_CONDITIONS + 1);
        let xml = format!(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><autoFilter><sortState ref=\"A1\">{conditions}</sortState></autoFilter></worksheet>"
        );
        assert!(parse_auto_filter(xml.as_bytes()).is_err());
    }
}
