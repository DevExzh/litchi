//! Typed MS-XLSX Timeline Cache and Timelines parts with exact OPC graph validation.

use super::auto_filter::{
    AutoFilterDefinition, parse_auto_filter_fragment, write_auto_filter_fragment,
};
use crate::error::{OoxmlError, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const XR10: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";

pub const TIMELINE_CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.TimelineCache+xml";
pub const TIMELINE_CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/TimelineCache";
pub const TIMELINES_CONTENT_TYPE: &str = "application/vnd.ms-excel.Timeline+xml";
pub const TIMELINES_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/Timeline";
pub const TIMELINE_CACHE_EXTENSION_URI: &str = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}";
pub const TIMELINES_EXTENSION_URI: &str = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}";

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
const MAX_OPAQUE_BYTES: usize = 4 * 1024 * 1024;
const MAX_TOTAL_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;
const MAX_CACHES: usize = 4096;
const MAX_TIMELINES: usize = 16_384;
const MAX_PIVOT_TABLES: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimelineOpaqueXml {
    /// A single, self-contained XML element. It is parsed and bounded but not interpreted.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimelineCachePivotTable {
    pub tab_id: u32,
    pub name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TimelineCacheDefinition {
    pub name: String,
    pub uid: Option<String>,
    pub source_name: String,
    pub pivot_tables: Vec<TimelineCachePivotTable>,
    pub state: TimelineState,
    pub timeline_pivot_filter: Option<TimelinePivotFilter>,
    pub extension_list: Option<TimelineOpaqueXml>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimelineRange {
    pub start_date: String,
    pub end_date: String,
}
impl TimelineRange {
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
pub enum PivotFilterType {
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
impl PivotFilterType {
    fn parse(v: &str) -> Result<Self> {
        use PivotFilterType::*;
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
    fn as_str(self) -> &'static str {
        use PivotFilterType::*;
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
pub struct TimelineState {
    pub selection: Option<TimelineRange>,
    pub bounds: TimelineRange,
    pub extension_list: Option<TimelineOpaqueXml>,
    pub single_range_filter_state: Option<bool>,
    pub minimal_refresh_version: u32,
    pub last_refresh_version: u32,
    pub pivot_cache_id: u32,
    pub filter_type: PivotFilterType,
}
impl TimelineState {
    pub fn new(bounds: TimelineRange, pivot_cache_id: u32, filter_type: PivotFilterType) -> Self {
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
pub struct TimelinePivotFilter {
    pub use_whole_day: Option<bool>,
    pub field: u32,
    pub id: u32,
    pub name: Option<String>,
    pub description: Option<String>,
    pub auto_filter: Option<AutoFilterDefinition>,
}
impl TimelinePivotFilter {
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
pub struct WorkbookTimelineCache {
    pub relationship_id: String,
    pub part_name: String,
    pub definition: TimelineCacheDefinition,
}

fn parse_range(n: &Node) -> Result<TimelineRange> {
    no_attributes(n, &[("", "startDate"), ("", "endDate")])?;
    empty(n)?;
    TimelineRange::new(required(n, "", "startDate")?, required(n, "", "endDate")?)
}
fn parse_state(n: &Node) -> Result<TimelineState> {
    require(n, X15, "state")?;
    whitespace(n)?;
    no_attributes(
        n,
        &[
            ("", "singleRangeFilterState"),
            ("", "minimalRefreshVersion"),
            ("", "lastRefreshVersion"),
            ("", "pivotCacheId"),
            ("", "filterType"),
        ],
    )?;
    let mut selection = None;
    let mut bounds = None;
    let mut extension_list = None;
    let mut stage = 0;
    for c in &n.children {
        match (c.namespace.as_str(), c.name.as_str()) {
            (X15, "selection") if stage == 0 => {
                selection = Some(parse_range(c)?);
                stage = 1
            },
            (X15, "bounds") if stage <= 1 && bounds.is_none() => {
                bounds = Some(parse_range(c)?);
                stage = 2
            },
            (ns, "extLst") if stage == 2 && (ns == X15 || ns == SML || ns == STRICT_SML) => {
                extension_list = Some(opaque_from_node(c)?);
                stage = 3
            },
            _ => return Err(invalid("invalid or out-of-order Timeline state child")),
        }
    }
    let state = TimelineState {
        selection,
        bounds: bounds.ok_or_else(|| invalid("Timeline state requires bounds"))?,
        extension_list,
        single_range_filter_state: optional_bool(n, "singleRangeFilterState")?,
        minimal_refresh_version: required(n, "", "minimalRefreshVersion")?
            .parse()
            .map_err(|_| invalid("invalid minimalRefreshVersion"))?,
        last_refresh_version: required(n, "", "lastRefreshVersion")?
            .parse()
            .map_err(|_| invalid("invalid lastRefreshVersion"))?,
        pivot_cache_id: required(n, "", "pivotCacheId")?
            .parse()
            .map_err(|_| invalid("invalid pivotCacheId"))?,
        filter_type: PivotFilterType::parse(required(n, "", "filterType")?)?,
    };
    validate_state(&state)?;
    Ok(state)
}
fn parse_timeline_pivot_filter(n: &Node) -> Result<TimelinePivotFilter> {
    require(n, X15, "timelinePivotFilter")?;
    whitespace(n)?;
    no_attributes(
        n,
        &[
            ("", "useWholeDay"),
            ("", "fld"),
            ("", "id"),
            ("", "name"),
            ("", "description"),
        ],
    )?;
    if n.children.len() > 1 {
        return Err(invalid(
            "timelinePivotFilter permits at most one autoFilter",
        ));
    }
    let auto_filter = n
        .children
        .first()
        .map(|c| {
            if !matches!(c.namespace.as_str(), SML | STRICT_SML) || c.name != "autoFilter" {
                return Err(invalid("timelinePivotFilter child must be autoFilter"));
            }
            parse_auto_filter_fragment(&serialize_node(c)?)
        })
        .transpose()?;
    let v = TimelinePivotFilter {
        use_whole_day: optional_bool(n, "useWholeDay")?,
        field: required(n, "", "fld")?
            .parse()
            .map_err(|_| invalid("invalid timeline filter fld"))?,
        id: required(n, "", "id")?
            .parse()
            .map_err(|_| invalid("invalid timeline filter id"))?,
        name: optional(n, "", "name").map(str::to_owned),
        description: optional(n, "", "description").map(str::to_owned),
        auto_filter,
    };
    validate_timeline_filter(&v)?;
    Ok(v)
}
fn write_range(x: &mut Vec<u8>, name: &str, v: &TimelineRange) -> Result<()> {
    validate_range(v)?;
    x.extend_from_slice(b"<x15:");
    x.extend_from_slice(name.as_bytes());
    attr(x, "startDate", &v.start_date);
    attr(x, "endDate", &v.end_date);
    x.extend_from_slice(b"/>");
    Ok(())
}
fn write_state(x: &mut Vec<u8>, v: &TimelineState) -> Result<()> {
    validate_state(v)?;
    x.extend_from_slice(b"<x15:state");
    if let Some(q) = v.single_range_filter_state {
        bool_attr(x, "singleRangeFilterState", q)
    }
    attr(
        x,
        "minimalRefreshVersion",
        &v.minimal_refresh_version.to_string(),
    );
    attr(x, "lastRefreshVersion", &v.last_refresh_version.to_string());
    attr(x, "pivotCacheId", &v.pivot_cache_id.to_string());
    attr(x, "filterType", v.filter_type.as_str());
    x.push(b'>');
    if let Some(q) = &v.selection {
        write_range(x, "selection", q)?
    }
    write_range(x, "bounds", &v.bounds)?;
    if let Some(q) = &v.extension_list {
        append_opaque_any_namespace(x, q, "extLst")?
    }
    x.extend_from_slice(b"</x15:state>");
    Ok(())
}
fn write_timeline_pivot_filter(x: &mut Vec<u8>, v: &TimelinePivotFilter) -> Result<()> {
    validate_timeline_filter(v)?;
    x.extend_from_slice(b"<x15:timelinePivotFilter");
    if let Some(q) = v.use_whole_day {
        bool_attr(x, "useWholeDay", q)
    }
    attr(x, "fld", &v.field.to_string());
    attr(x, "id", &v.id.to_string());
    if let Some(q) = &v.name {
        attr(x, "name", q)
    }
    if let Some(q) = &v.description {
        attr(x, "description", q)
    }
    if let Some(q) = &v.auto_filter {
        x.push(b'>');
        x.extend_from_slice(&write_auto_filter_fragment(q)?);
        x.extend_from_slice(b"</x15:timelinePivotFilter>")
    } else {
        x.extend_from_slice(b"/>")
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimelineLevel {
    Year,
    Quarter,
    Month,
    Day,
}

impl TimelineLevel {
    fn parse(value: &str, name: &str) -> Result<Self> {
        match value {
            "0" => Ok(Self::Year),
            "1" => Ok(Self::Quarter),
            "2" => Ok(Self::Month),
            "3" => Ok(Self::Day),
            _ => Err(invalid(format!("{name} must be between 0 and 3"))),
        }
    }
    fn number(self) -> &'static str {
        match self {
            Self::Year => "0",
            Self::Quarter => "1",
            Self::Month => "2",
            Self::Day => "3",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Timeline {
    pub name: String,
    pub uid: Option<String>,
    pub cache: String,
    pub caption: Option<String>,
    pub show_header: Option<bool>,
    pub show_selection_label: Option<bool>,
    pub show_time_level: Option<bool>,
    pub show_horizontal_scrollbar: Option<bool>,
    pub level: TimelineLevel,
    pub selection_level: TimelineLevel,
    pub scroll_position: Option<String>,
    pub style: Option<String>,
    pub extension_list: Option<TimelineOpaqueXml>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Timelines {
    pub timelines: Vec<Timeline>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetTimelines {
    pub worksheet_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub timelines: Timelines,
}

#[derive(Clone, Debug)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone, Debug)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

/// Parse one MS-XLSX Timeline Cache part.
pub fn parse_timeline_cache_definition(xml: &[u8]) -> Result<TimelineCacheDefinition> {
    let root = parse_document(xml)?;
    require(&root, X15, "timelineCacheDefinition")?;
    whitespace(&root)?;
    no_attributes(&root, &[("", "name"), ("", "sourceName"), (XR10, "uid")])?;
    let name = required(&root, "", "name")?.to_owned();
    let source_name = required(&root, "", "sourceName")?.to_owned();
    let uid = optional(&root, XR10, "uid").map(str::to_owned);
    let mut pivot_tables = Vec::new();
    let mut state = None;
    let mut timeline_pivot_filter = None;
    let mut extension_list = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "pivotTables" if child.namespace == X15 && stage == 0 => {
                stage = 1;
                parse_pivot_tables(child, &mut pivot_tables)?;
            },
            "state" if child.namespace == X15 && stage <= 1 && state.is_none() => {
                stage = 2;
                state = Some(parse_state(child)?);
            },
            "timelinePivotFilter"
                if child.namespace == X15 && stage == 2 && timeline_pivot_filter.is_none() =>
            {
                stage = 3;
                timeline_pivot_filter = Some(parse_timeline_pivot_filter(child)?);
            },
            "extLst"
                if (child.namespace == X15
                    || child.namespace == SML
                    || child.namespace == STRICT_SML)
                    && stage >= 2
                    && extension_list.is_none() =>
            {
                stage = 4;
                extension_list = Some(opaque_from_node(child)?);
            },
            _ => {
                return Err(invalid(format!(
                    "unexpected or out-of-order Timeline Cache child '{}'",
                    child.name
                )));
            },
        }
    }
    let value = TimelineCacheDefinition {
        name,
        uid,
        source_name,
        pivot_tables,
        state: state.ok_or_else(|| invalid("Timeline Cache requires exactly one state element"))?,
        timeline_pivot_filter,
        extension_list,
    };
    validate_cache_definition(&value)?;
    Ok(value)
}

fn parse_pivot_tables(node: &Node, output: &mut Vec<TimelineCachePivotTable>) -> Result<()> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid("pivotTables must contain at least one pivotTable"));
    }
    if node.children.len() > MAX_PIVOT_TABLES {
        return Err(limit("pivot table count"));
    }
    let mut bindings = HashSet::new();
    for child in &node.children {
        require(child, X15, "pivotTable")?;
        no_attributes(child, &[("", "tabId"), ("", "name")])?;
        empty(child)?;
        let tab_id = required(child, "", "tabId")?
            .parse::<u32>()
            .map_err(|_| invalid("invalid Timeline Cache pivotTable tabId"))?;
        let name = required(child, "", "name")?.to_owned();
        bounded(&name, "pivot table name")?;
        if !bindings.insert((tab_id, name.to_lowercase())) {
            return Err(invalid("duplicate Timeline Cache pivotTable binding"));
        }
        output.push(TimelineCachePivotTable { tab_id, name });
    }
    Ok(())
}

/// Deterministically serialize one Timeline Cache part.
pub fn write_timeline_cache_definition(value: &TimelineCacheDefinition) -> Result<Vec<u8>> {
    validate_cache_definition(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:timelineCacheDefinition xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:xr10=\"");
    escape(&mut output, XR10);
    output.push(b'\"');
    attr(&mut output, "name", &value.name);
    attr(&mut output, "sourceName", &value.source_name);
    if let Some(uid) = &value.uid {
        attr(&mut output, "xr10:uid", uid);
    }
    output.push(b'>');
    if !value.pivot_tables.is_empty() {
        output.extend_from_slice(b"<x15:pivotTables>");
        for pivot in &value.pivot_tables {
            output.extend_from_slice(b"<x15:pivotTable");
            attr(&mut output, "tabId", &pivot.tab_id.to_string());
            attr(&mut output, "name", &pivot.name);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:pivotTables>");
    }
    write_state(&mut output, &value.state)?;
    if let Some(payload) = &value.timeline_pivot_filter {
        write_timeline_pivot_filter(&mut output, payload)?;
    }
    if let Some(payload) = &value.extension_list {
        append_opaque_any_namespace(&mut output, payload, "extLst")?;
    }
    output.extend_from_slice(b"</x15:timelineCacheDefinition>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized cache XML bytes"));
    }
    Ok(output)
}

/// Parse one worksheet-scoped Timelines part.
pub fn parse_timelines(xml: &[u8]) -> Result<Timelines> {
    let root = parse_document(xml)?;
    require(&root, X15, "timelines")?;
    no_attributes(&root, &[])?;
    whitespace(&root)?;
    if root.children.is_empty() {
        return Err(invalid("Timelines part must contain at least one timeline"));
    }
    if root.children.len() > MAX_TIMELINES {
        return Err(limit("timeline count"));
    }
    let mut timelines = Vec::with_capacity(root.children.len());
    for child in &root.children {
        timelines.push(parse_timeline(child)?);
    }
    let value = Timelines { timelines };
    validate_timelines_local(&value)?;
    Ok(value)
}

fn parse_timeline(node: &Node) -> Result<Timeline> {
    require(node, X15, "timeline")?;
    whitespace(node)?;
    no_attributes(
        node,
        &[
            ("", "name"),
            (XR10, "uid"),
            ("", "cache"),
            ("", "caption"),
            ("", "showHeader"),
            ("", "showSelectionLabel"),
            ("", "showTimeLevel"),
            ("", "showHorizontalScrollbar"),
            ("", "level"),
            ("", "selectionLevel"),
            ("", "scrollPosition"),
            ("", "style"),
        ],
    )?;
    if node.children.len() > 1 {
        return Err(invalid("timeline permits at most one extLst"));
    }
    let extension_list = node
        .children
        .first()
        .map(|child| {
            if child.name != "extLst"
                || !(child.namespace == X15
                    || child.namespace == SML
                    || child.namespace == STRICT_SML)
            {
                return Err(invalid("timeline child must be extLst"));
            }
            opaque_from_node(child)
        })
        .transpose()?;
    Ok(Timeline {
        name: required(node, "", "name")?.to_owned(),
        uid: optional(node, XR10, "uid").map(str::to_owned),
        cache: required(node, "", "cache")?.to_owned(),
        caption: optional(node, "", "caption").map(str::to_owned),
        show_header: optional_bool(node, "showHeader")?,
        show_selection_label: optional_bool(node, "showSelectionLabel")?,
        show_time_level: optional_bool(node, "showTimeLevel")?,
        show_horizontal_scrollbar: optional_bool(node, "showHorizontalScrollbar")?,
        level: TimelineLevel::parse(required(node, "", "level")?, "level")?,
        selection_level: TimelineLevel::parse(
            required(node, "", "selectionLevel")?,
            "selectionLevel",
        )?,
        scroll_position: optional(node, "", "scrollPosition").map(str::to_owned),
        style: optional(node, "", "style").map(str::to_owned),
        extension_list,
    })
}

/// Deterministically serialize one worksheet-scoped Timelines part.
pub fn write_timelines(value: &Timelines) -> Result<Vec<u8>> {
    validate_timelines_local(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:timelines xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:xr10=\"");
    escape(&mut output, XR10);
    output.extend_from_slice(b"\">");
    for timeline in &value.timelines {
        output.extend_from_slice(b"<x15:timeline");
        attr(&mut output, "name", &timeline.name);
        if let Some(uid) = &timeline.uid {
            attr(&mut output, "xr10:uid", uid);
        }
        attr(&mut output, "cache", &timeline.cache);
        if let Some(value) = &timeline.caption {
            attr(&mut output, "caption", value);
        }
        for (name, value) in [
            ("showHeader", timeline.show_header),
            ("showSelectionLabel", timeline.show_selection_label),
            ("showTimeLevel", timeline.show_time_level),
            (
                "showHorizontalScrollbar",
                timeline.show_horizontal_scrollbar,
            ),
        ] {
            if let Some(value) = value {
                bool_attr(&mut output, name, value);
            }
        }
        attr(&mut output, "level", timeline.level.number());
        attr(
            &mut output,
            "selectionLevel",
            timeline.selection_level.number(),
        );
        if let Some(value) = &timeline.scroll_position {
            attr(&mut output, "scrollPosition", value);
        }
        if let Some(value) = &timeline.style {
            attr(&mut output, "style", value);
        }
        if let Some(payload) = &timeline.extension_list {
            output.push(b'>');
            append_opaque_any_namespace(&mut output, payload, "extLst")?;
            output.extend_from_slice(b"</x15:timeline>");
        } else {
            output.extend_from_slice(b"/>");
        }
    }
    output.extend_from_slice(b"</x15:timelines>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized timelines XML bytes"));
    }
    Ok(output)
}

/// Load and validate all workbook Timeline Cache parts.
pub fn load_timeline_caches(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Vec<WorkbookTimelineCache>> {
    reject_root_relationships(package, TIMELINE_CACHE_RELATIONSHIP_TYPE, "Timeline Cache")?;
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = parse_document(workbook.blob())?;
    let (core, rel) = source_namespaces(&workbook_root, "workbook")?;
    let refs = integration_refs(
        &workbook_root,
        core,
        rel,
        TIMELINE_CACHE_EXTENSION_URI,
        "timelineCacheRefs",
        "timelineCacheRef",
    )?
    .unwrap_or_default();
    if refs.len() > MAX_CACHES {
        return Err(limit("cache reference count"));
    }
    for part in package.iter_parts() {
        if part.partname().as_str() != workbook_name.as_str()
            && part
                .rels()
                .iter()
                .any(|relationship| relationship.reltype() == TIMELINE_CACHE_RELATIONSHIP_TYPE)
        {
            return Err(invalid(format!(
                "non-workbook part '{}' sources a Timeline Cache relationship",
                part.partname()
            )));
        }
    }
    let mut ids = HashSet::new();
    let mut targets = HashSet::new();
    let mut output = Vec::with_capacity(refs.len());
    for id in refs {
        validate_relationship_id(&id)?;
        if !ids.insert(id.clone()) {
            return Err(invalid(format!(
                "duplicate Timeline Cache reference '{id}'"
            )));
        }
        let relationship = workbook
            .rels()
            .get(&id)
            .ok_or_else(|| invalid(format!("missing Timeline Cache relationship '{id}'")))?;
        if relationship.reltype() != TIMELINE_CACHE_RELATIONSHIP_TYPE || relationship.is_external()
        {
            return Err(invalid(format!(
                "Timeline Cache reference '{id}' must target an internal Timeline Cache relationship"
            )));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple Timeline Cache references target '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != TIMELINE_CACHE_CONTENT_TYPE {
            return Err(invalid(format!(
                "Timeline Cache part '{target}' has content type '{}'",
                part.content_type()
            )));
        }
        if !part.rels().is_empty() {
            return Err(invalid(format!(
                "Timeline Cache part '{target}' has forbidden outbound relationships"
            )));
        }
        output.push(WorkbookTimelineCache {
            relationship_id: id,
            part_name: target.to_string(),
            definition: parse_timeline_cache_definition(part.blob())?,
        });
    }
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TIMELINE_CACHE_RELATIONSHIP_TYPE)
    {
        if !ids.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced Timeline Cache relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == TIMELINE_CACHE_CONTENT_TYPE)
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan Timeline Cache part '{}'",
                part.partname()
            )));
        }
    }
    validate_cache_set(&output)?;
    Ok(output)
}

/// Store a complete workbook Timeline Cache set and its single integration extension.
pub fn store_timeline_caches(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    caches: &[WorkbookTimelineCache],
) -> Result<()> {
    if caches.is_empty() {
        return Err(invalid(
            "at least one Timeline Cache is required for storage",
        ));
    }
    validate_cache_set(caches)?;
    if !load_timeline_caches(package, workbook_name)?.is_empty() {
        return Err(invalid("workbook already contains Timeline Caches"));
    }
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    let (core, rel) = source_namespaces(&root, "workbook")?;
    let mut targets = HashSet::new();
    let mut ids = HashSet::new();
    let mut plans = Vec::with_capacity(caches.len());
    for cache in caches {
        validate_relationship_id(&cache.relationship_id)?;
        if !ids.insert(cache.relationship_id.clone()) {
            return Err(invalid("duplicate Timeline Cache relationship ID"));
        }
        if workbook.rels().get(&cache.relationship_id).is_some() {
            return Err(invalid(format!(
                "workbook relationship ID '{}' already exists",
                cache.relationship_id
            )));
        }
        let uri = PackURI::new(&cache.part_name).map_err(OoxmlError::InvalidUri)?;
        if !uri.as_str().starts_with("/xl/timelineCaches/") || !uri.as_str().ends_with(".xml") {
            return Err(invalid(format!(
                "Timeline Cache part '{uri}' must be under /xl/timelineCaches and end in .xml"
            )));
        }
        if !targets.insert(uri.to_string())
            || package
                .iter_parts()
                .any(|part| part.partname().as_str() == uri.as_str())
        {
            return Err(invalid(format!(
                "Timeline Cache target '{uri}' already exists"
            )));
        }
        plans.push((
            cache.relationship_id.clone(),
            uri,
            write_timeline_cache_definition(&cache.definition)?,
        ));
    }
    let refs: Vec<String> = plans.iter().map(|plan| plan.0.clone()).collect();
    let fragment = integration_extension(
        core,
        rel,
        TIMELINE_CACHE_EXTENSION_URI,
        "timelineCacheRefs",
        "timelineCacheRef",
        &refs,
    );
    let updated = insert_extension(
        workbook.blob(),
        &root,
        core,
        TIMELINE_CACHE_EXTENSION_URI,
        &fragment,
    )?;
    for (_, uri, xml) in &plans {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            TIMELINE_CACHE_CONTENT_TYPE.into(),
            xml.clone(),
        )));
    }
    for (id, uri, _) in &plans {
        package
            .get_part_mut(workbook_name)?
            .rels_mut()
            .add_relationship(
                TIMELINE_CACHE_RELATIONSHIP_TYPE.into(),
                uri.relative_ref(workbook_name.base_uri()),
                id.clone(),
                false,
            );
    }
    package.get_part_mut(workbook_name)?.set_blob(updated);
    Ok(())
}

/// Load every worksheet Timelines part and cross-validate views against Timeline Caches.
pub fn load_timelines(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Vec<WorksheetTimelines>> {
    reject_root_relationships(package, TIMELINES_RELATIONSHIP_TYPE, "Timelines")?;
    let caches = load_timeline_caches(package, workbook_name)?;
    let cache_names: HashSet<String> = caches
        .iter()
        .map(|cache| cache.definition.name.to_lowercase())
        .collect();
    let mut targets = HashSet::new();
    let mut output = Vec::new();
    for part in package.iter_parts() {
        let timeline_relationships: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == TIMELINES_RELATIONSHIP_TYPE)
            .collect();
        if part.content_type() != ct::SML_WORKSHEET {
            if !timeline_relationships.is_empty() {
                return Err(invalid(format!(
                    "non-worksheet part '{}' sources a Timelines relationship",
                    part.partname()
                )));
            }
            continue;
        }
        let root = parse_document(part.blob())?;
        let (core, rel) = source_namespaces(&root, "worksheet")?;
        let refs = integration_refs(
            &root,
            core,
            rel,
            TIMELINES_EXTENSION_URI,
            "timelineRefs",
            "timelineRef",
        )?
        .unwrap_or_default();
        if refs.is_empty() {
            if !timeline_relationships.is_empty() {
                return Err(invalid(format!(
                    "worksheet '{}' has a Timelines relationship without timelineRefs",
                    part.partname()
                )));
            }
            continue;
        }
        if refs.len() != 1 {
            return Err(invalid(
                "worksheet timelineRefs must contain exactly one timelineRef",
            ));
        }
        let id = refs[0].clone();
        validate_relationship_id(&id)?;
        let relationship = part
            .rels()
            .get(&id)
            .ok_or_else(|| invalid(format!("missing Timelines relationship '{id}'")))?;
        if relationship.reltype() != TIMELINES_RELATIONSHIP_TYPE || relationship.is_external() {
            return Err(invalid(format!(
                "Timelines reference '{id}' must target an internal Timelines relationship"
            )));
        }
        if timeline_relationships.len() != 1 {
            return Err(invalid(format!(
                "worksheet '{}' has unreferenced or duplicate Timelines relationships",
                part.partname()
            )));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple worksheets target Timelines part '{target}'"
            )));
        }
        let timeline_part = package.get_part(&target)?;
        if timeline_part.content_type() != TIMELINES_CONTENT_TYPE {
            return Err(invalid(format!(
                "Timelines part '{target}' has content type '{}'",
                timeline_part.content_type()
            )));
        }
        if !timeline_part.rels().is_empty() {
            return Err(invalid(format!(
                "Timelines part '{target}' has forbidden outbound relationships"
            )));
        }
        let timelines = parse_timelines(timeline_part.blob())?;
        for view in &timelines.timelines {
            if !cache_names.contains(&view.cache.to_lowercase()) {
                return Err(invalid(format!(
                    "timeline '{}' references unknown cache '{}'",
                    view.name, view.cache
                )));
            }
        }
        output.push(WorksheetTimelines {
            worksheet_part_name: part.partname().to_string(),
            relationship_id: id,
            part_name: target.to_string(),
            timelines,
        });
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == TIMELINES_CONTENT_TYPE)
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan Timelines part '{}'",
                part.partname()
            )));
        }
    }
    validate_global_timelines(&output)?;
    Ok(output)
}

/// Store one worksheet's Timelines part, then validate it against all workbook caches and views.
pub fn store_worksheet_timelines(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &WorksheetTimelines,
) -> Result<()> {
    validate_timelines_local(&value.timelines)?;
    let caches = load_timeline_caches(package, workbook_name)?;
    let cache_names: HashSet<String> = caches
        .iter()
        .map(|cache| cache.definition.name.to_lowercase())
        .collect();
    for view in &value.timelines.timelines {
        if !cache_names.contains(&view.cache.to_lowercase()) {
            return Err(invalid(format!(
                "timeline '{}' references unknown cache '{}'",
                view.name, view.cache
            )));
        }
    }
    let existing = load_timelines(package, workbook_name)?;
    if existing
        .iter()
        .any(|sheet| sheet.worksheet_part_name == value.worksheet_part_name)
    {
        return Err(invalid("worksheet already contains a Timelines part"));
    }
    let mut combined = existing;
    combined.push(value.clone());
    validate_global_timelines(&combined)?;
    validate_relationship_id(&value.relationship_id)?;
    let worksheet_uri = PackURI::new(&value.worksheet_part_name).map_err(OoxmlError::InvalidUri)?;
    let target_uri = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    if !target_uri.as_str().starts_with("/xl/timelines/") || !target_uri.as_str().ends_with(".xml")
    {
        return Err(invalid(
            "Timelines part must be under /xl/timelines and end in .xml",
        ));
    }
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == target_uri.as_str())
    {
        return Err(invalid(format!(
            "Timelines part '{target_uri}' already exists"
        )));
    }
    let worksheet = package.get_part(&worksheet_uri)?;
    if worksheet.content_type() != ct::SML_WORKSHEET {
        return Err(invalid(format!(
            "part '{worksheet_uri}' is not a worksheet"
        )));
    }
    if worksheet.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(format!(
            "worksheet relationship ID '{}' already exists",
            value.relationship_id
        )));
    }
    let root = parse_document(worksheet.blob())?;
    let (core, rel) = source_namespaces(&root, "worksheet")?;
    let fragment = integration_extension(
        core,
        rel,
        TIMELINES_EXTENSION_URI,
        "timelineRefs",
        "timelineRef",
        std::slice::from_ref(&value.relationship_id),
    );
    let updated = insert_extension(
        worksheet.blob(),
        &root,
        core,
        TIMELINES_EXTENSION_URI,
        &fragment,
    )?;
    let xml = write_timelines(&value.timelines)?;
    package.add_part(Box::new(BlobPart::new(
        target_uri.clone(),
        TIMELINES_CONTENT_TYPE.into(),
        xml,
    )));
    package
        .get_part_mut(&worksheet_uri)?
        .rels_mut()
        .add_relationship(
            TIMELINES_RELATIONSHIP_TYPE.into(),
            target_uri.relative_ref(worksheet_uri.base_uri()),
            value.relationship_id.clone(),
            false,
        );
    package.get_part_mut(&worksheet_uri)?.set_blob(updated);
    Ok(())
}

fn validate_cache_definition(value: &TimelineCacheDefinition) -> Result<()> {
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
            return Err(invalid("duplicate Timeline Cache pivotTable binding"));
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

fn validate_state(v: &TimelineState) -> Result<()> {
    validate_range(&v.bounds)?;
    if let Some(s) = &v.selection {
        validate_range(s)?;
        if compare_datetime(&s.start_date, &v.bounds.start_date)? == std::cmp::Ordering::Less
            || compare_datetime(&s.end_date, &v.bounds.end_date)? == std::cmp::Ordering::Greater
        {
            return Err(invalid("Timeline selection must be within bounds"));
        }
    }
    if let Some(e) = &v.extension_list {
        validate_opaque_kind(e, &[X15, SML, STRICT_SML], "extLst")?
    }
    Ok(())
}
fn validate_timeline_filter(v: &TimelinePivotFilter) -> Result<()> {
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
fn validate_range(v: &TimelineRange) -> Result<()> {
    validate_datetime(&v.start_date)?;
    validate_datetime(&v.end_date)?;
    if compare_datetime(&v.start_date, &v.end_date)? == std::cmp::Ordering::Greater {
        return Err(invalid("Timeline range startDate exceeds endDate"));
    }
    Ok(())
}
fn compare_datetime(a: &str, b: &str) -> Result<std::cmp::Ordering> {
    let a = parse_xsd_datetime(a)?;
    let b = parse_xsd_datetime(b)?;
    match (a.timezone_minutes, b.timezone_minutes) {
        (Some(_), Some(_)) => Ok(a.utc_normalized().cmp_value(&b.utc_normalized())),
        (None, None) => Ok(a.cmp_value(&b)),
        _ => Err(invalid(
            "cannot compare timezone-aware and local Timeline dates",
        )),
    }
}

fn validate_cache_set(caches: &[WorkbookTimelineCache]) -> Result<()> {
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
                "duplicate case-insensitive Timeline Cache name '{}'",
                cache.definition.name
            )));
        }
        if !ids.insert(cache.relationship_id.clone()) {
            return Err(invalid("duplicate Timeline Cache relationship ID"));
        }
        if !targets.insert(cache.part_name.clone()) {
            return Err(invalid("duplicate Timeline Cache part name"));
        }
        if any_uid && cache.definition.uid.is_none() {
            return Err(invalid(
                "Timeline Cache UIDs must be specified on all caches or none",
            ));
        }
        if let Some(uid) = &cache.definition.uid {
            if !uids.insert(uid.to_ascii_lowercase()) {
                return Err(invalid("duplicate Timeline Cache UID"));
            }
        }
    }
    Ok(())
}

fn validate_timelines_local(value: &Timelines) -> Result<()> {
    if value.timelines.is_empty() {
        return Err(invalid("Timelines part must contain at least one timeline"));
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

fn validate_global_timelines(values: &[WorksheetTimelines]) -> Result<()> {
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
        if let Some(uid) = &view.uid {
            if !uids.insert(uid.to_ascii_lowercase()) {
                return Err(invalid("duplicate timeline UID"));
            }
        }
    }
    Ok(())
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn opaque_from_node(node: &Node) -> Result<TimelineOpaqueXml> {
    let xml = serialize_node(node)?;
    if xml.len() > MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML bytes"));
    }
    Ok(TimelineOpaqueXml { xml })
}

fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::<String, String>::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            X15 => "x15".into(),
            SML | STRICT_SML => "x".into(),
            REL | STRICT_REL => "r".into(),
            XR10 => "xr10".into(),
            _ => {
                let value = format!("n{next}");
                next += 1;
                value
            },
        };
        prefixes.insert(namespace.clone(), prefix);
    }
    let mut output = Vec::new();
    write_node(node, &prefixes, true, &mut output);
    Ok(output)
}

fn collect_namespaces(node: &Node, output: &mut BTreeMap<String, String>) {
    if !node.namespace.is_empty() {
        output.insert(node.namespace.clone(), String::new());
    }
    for attr in &node.attributes {
        if !attr.namespace.is_empty() {
            output.insert(attr.namespace.clone(), String::new());
        }
    }
    for child in &node.children {
        collect_namespaces(child, output);
    }
}
fn write_node(node: &Node, prefixes: &HashMap<String, String>, root: bool, output: &mut Vec<u8>) {
    output.push(b'<');
    qname(output, &node.namespace, &node.name, prefixes);
    if root {
        let mut entries: Vec<_> = prefixes.iter().collect();
        entries.sort_by(|a, b| a.1.cmp(b.1));
        for (namespace, prefix) in entries {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
            escape(output, namespace);
            output.push(b'\"');
        }
    }
    for attr_value in &node.attributes {
        output.push(b' ');
        qname(output, &attr_value.namespace, &attr_value.name, prefixes);
        output.extend_from_slice(b"=\"");
        escape(output, &attr_value.value);
        output.push(b'\"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_node(child, prefixes, false, output);
    }
    output.extend_from_slice(b"</");
    qname(output, &node.namespace, &node.name, prefixes);
    output.push(b'>');
}
fn qname(output: &mut Vec<u8>, namespace: &str, name: &str, prefixes: &HashMap<String, String>) {
    if !namespace.is_empty() {
        output.extend_from_slice(prefixes[namespace].as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}

fn append_opaque_any_namespace(
    output: &mut Vec<u8>,
    payload: &TimelineOpaqueXml,
    name: &str,
) -> Result<()> {
    validate_opaque_kind(payload, &[X15, SML, STRICT_SML], name)?;
    output.extend_from_slice(&payload.xml);
    Ok(())
}
fn validate_opaque_kind(
    payload: &TimelineOpaqueXml,
    namespaces: &[&str],
    name: &str,
) -> Result<()> {
    if payload.xml.len() > MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML bytes"));
    }
    let root = parse_document(&payload.xml)?;
    if root.name != name || !namespaces.contains(&root.namespace.as_str()) {
        return Err(invalid(format!(
            "opaque XML must be a {name} element in its normative namespace"
        )));
    }
    Ok(())
}

fn integration_refs(
    root: &Node,
    core: &str,
    rel: &str,
    uri: &str,
    list_name: &str,
    ref_name: &str,
) -> Result<Option<Vec<String>>> {
    let mut found = None;
    for ext_lst in root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
    {
        whitespace(ext_lst)?;
        for ext in ext_lst.children.iter().filter(|child| {
            child.namespace == core
                && child.name == "ext"
                && optional(child, "", "uri") == Some(uri)
        }) {
            if found.is_some() {
                return Err(invalid(format!(
                    "duplicate integration extension URI '{uri}'"
                )));
            }
            no_attributes(ext, &[("", "uri")])?;
            whitespace(ext)?;
            if ext.children.len() != 1 {
                return Err(invalid(format!(
                    "integration extension '{uri}' must contain exactly one child"
                )));
            }
            let list = &ext.children[0];
            require(list, X15, list_name)?;
            no_attributes(list, &[])?;
            whitespace(list)?;
            if list.children.is_empty() {
                return Err(invalid(format!("{list_name} must not be empty")));
            }
            let mut ids = Vec::with_capacity(list.children.len());
            for item in &list.children {
                require(item, X15, ref_name)?;
                no_attributes(item, &[(rel, "id")])?;
                empty(item)?;
                ids.push(required(item, rel, "id")?.to_owned());
            }
            found = Some(ids);
        }
    }
    Ok(found)
}

fn integration_extension(
    core: &str,
    rel: &str,
    uri: &str,
    list_name: &str,
    ref_name: &str,
    ids: &[String],
) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<ext xmlns=\"");
    escape(&mut output, core);
    output.extend_from_slice(b"\" uri=\"");
    escape(&mut output, uri);
    output.extend_from_slice(b"\"><x15:");
    output.extend_from_slice(list_name.as_bytes());
    output.extend_from_slice(b" xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, rel);
    output.extend_from_slice(b"\">");
    for id in ids {
        output.extend_from_slice(b"<x15:");
        output.extend_from_slice(ref_name.as_bytes());
        attr(&mut output, "r:id", id);
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</x15:");
    output.extend_from_slice(list_name.as_bytes());
    output.extend_from_slice(b"></ext>");
    output
}

fn insert_extension(
    xml: &[u8],
    root: &Node,
    core: &str,
    uri: &str,
    fragment: &[u8],
) -> Result<Vec<u8>> {
    if root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
        .flat_map(|list| &list.children)
        .any(|child| {
            child.namespace == core
                && child.name == "ext"
                && optional(child, "", "uri") == Some(uri)
        })
    {
        return Err(invalid(format!(
            "integration extension '{uri}' already exists"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut open_ext = None;
    let mut empty_ext = None;
    let mut root_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let mut empty_candidate = None;
        match event {
            Event::Start(element) => {
                let is_core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == core.as_bytes());
                if depth == 1 && is_core && element.local_name().as_ref() == b"extLst" {
                    open_ext = Some(2usize);
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
            },
            Event::Empty(element) => {
                let is_core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == core.as_bytes());
                if depth == 1 && is_core && element.local_name().as_ref() == b"extLst" {
                    empty_candidate = Some((start, element.name().as_ref().to_vec()));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected closing element"));
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                if open_ext == Some(depth) && element.local_name().as_ref() == b"extLst" {
                    let size = xml
                        .len()
                        .checked_add(fragment.len())
                        .ok_or_else(|| limit("rewrite bytes"))?;
                    if size > MAX_REWRITE_BYTES {
                        return Err(limit("rewrite bytes"));
                    }
                    let mut output = Vec::with_capacity(size);
                    output.extend_from_slice(&xml[..start]);
                    output.extend_from_slice(fragment);
                    output.extend_from_slice(&xml[start..]);
                    return Ok(output);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        if let Some((start, qname)) = empty_candidate {
            let end = usize::try_from(reader.buffer_position())
                .map_err(|_| invalid("XML offset overflow"))?;
            empty_ext = Some((start, end, qname));
        }
    }
    if let Some((start, end, qname)) = empty_ext {
        let size = xml
            .len()
            .checked_add(fragment.len() + qname.len() + 2)
            .ok_or_else(|| limit("rewrite bytes"))?;
        if size > MAX_REWRITE_BYTES {
            return Err(limit("rewrite bytes"));
        }
        let raw = &xml[start..end];
        let close =
            memchr::memmem::rfind(raw, b"/>").ok_or_else(|| invalid("invalid empty extLst"))?;
        let mut output = Vec::with_capacity(size);
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(&raw[..close]);
        output.push(b'>');
        output.extend_from_slice(fragment);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&qname);
        output.push(b'>');
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    let position = root_close.ok_or_else(|| invalid("missing source root closing element"))?;
    let mut wrapper = Vec::new();
    wrapper.extend_from_slice(b"<extLst xmlns=\"");
    escape(&mut wrapper, core);
    wrapper.extend_from_slice(b"\">");
    wrapper.extend_from_slice(fragment);
    wrapper.extend_from_slice(b"</extLst>");
    let size = xml
        .len()
        .checked_add(wrapper.len())
        .ok_or_else(|| limit("rewrite bytes"))?;
    if size > MAX_REWRITE_BYTES {
        return Err(limit("rewrite bytes"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(&wrapper);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn source_namespaces<'a>(root: &'a Node, name: &str) -> Result<(&'a str, &'static str)> {
    if root.name != name {
        return Err(invalid(format!("expected {name} source root")));
    }
    match root.namespace.as_str() {
        SML => Ok((SML, REL)),
        STRICT_SML => Ok((STRICT_SML, STRICT_REL)),
        _ => Err(invalid(format!("unsupported {name} namespace"))),
    }
}
fn reject_root_relationships(package: &OpcPackage, kind: &str, name: &str) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == kind)
    {
        Err(invalid(format!(
            "package root cannot source {name} relationships"
        )))
    } else {
        Ok(())
    }
}

fn validate_defined_name(value: &str) -> Result<()> {
    bounded_nonempty(value, "Timeline Cache name")?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first.is_alphabetic() || matches!(first, '_' | '\\'))
        || !chars
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.' | '\\'))
    {
        return Err(invalid(format!(
            "invalid defined-name-shaped Timeline Cache name '{value}'"
        )));
    }
    let upper = value.to_ascii_uppercase();
    if looks_like_a1(&upper) || looks_like_r1c1(&upper) {
        return Err(invalid(format!(
            "Timeline Cache name '{value}' conflicts with a cell reference"
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
    let inner = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
        .unwrap_or(value);
    let groups: Vec<_> = inner.split('-').collect();
    if groups.len() != 5
        || groups
            .iter()
            .zip([8, 4, 4, 4, 12])
            .any(|(group, len)| group.len() != len || !group.bytes().all(|b| b.is_ascii_hexdigit()))
    {
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
    value % 4 == 0 && (value % 100 != 0 || value == 0)
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

fn validate_datetime(value: &str) -> Result<()> {
    parse_xsd_datetime(value).map(|_| ())
}
fn validate_relationship_id(value: &str) -> Result<()> {
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

fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected {{{namespace}}}{name}, got {{{}}}{}",
            node.namespace, node.name
        )))
    }
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn optional_bool(node: &Node, name: &str) -> Result<Option<bool>> {
    optional(node, "", name)
        .map(|value| parse_bool(value, name))
        .transpose()
}
fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{value}' for {name}"))),
    }
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn empty(node: &Node) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} must be empty", node.name)))
    }
}
fn bounded(value: &str, name: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit(name))
    }
}
fn bounded_nonempty(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{name} cannot be empty")));
    }
    bounded(value, name)
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value)
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn bool_attr(output: &mut Vec<u8>, name: &str, value: bool) {
    attr(output, name, if value { "1" } else { "0" });
}
fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
}
fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("Timeline {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn state() -> TimelineState {
        TimelineState {
            selection: Some(
                TimelineRange::new("2024-01-02T00:00:00Z", "2024-01-31T23:59:59Z").unwrap(),
            ),
            bounds: TimelineRange::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z").unwrap(),
            extension_list: None,
            single_range_filter_state: None,
            minimal_refresh_version: 0,
            last_refresh_version: 1,
            pivot_cache_id: 1,
            filter_type: PivotFilterType::DateBetween,
        }
    }
    fn cache() -> WorkbookTimelineCache {
        WorkbookTimelineCache {
            relationship_id: "rIdTimelineCache1".into(),
            part_name: "/xl/timelineCaches/timelineCache1.xml".into(),
            definition: TimelineCacheDefinition {
                name: "Timeline_Date".into(),
                uid: Some("{11111111-1111-1111-1111-111111111111}".into()),
                source_name: "Date".into(),
                pivot_tables: vec![TimelineCachePivotTable {
                    tab_id: 1,
                    name: "PivotTable1".into(),
                }],
                state: state(),
                timeline_pivot_filter: None,
                extension_list: None,
            },
        }
    }
    fn views() -> Timelines {
        Timelines {
            timelines: vec![Timeline {
                name: "Timeline_Date_View".into(),
                uid: Some("{22222222-2222-2222-2222-222222222222}".into()),
                cache: "Timeline_Date".into(),
                caption: Some("Date".into()),
                show_header: Some(true),
                show_selection_label: Some(false),
                show_time_level: None,
                show_horizontal_scrollbar: Some(true),
                level: TimelineLevel::Month,
                selection_level: TimelineLevel::Day,
                scroll_position: Some("2024-01-02T03:04:05Z".into()),
                style: Some("TimelineStyleLight1".into()),
                extension_list: None,
            }],
        }
    }
    fn package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let sheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            sheet,
            ct::SML_WORKSHEET.into(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        )));
        (package, workbook)
    }

    #[test]
    fn typed_protocol_shaped_parts_round_trip() {
        let cache = cache().definition;
        let cache_xml = write_timeline_cache_definition(&cache).unwrap();
        assert_eq!(parse_timeline_cache_definition(&cache_xml).unwrap(), cache);
        let timelines = views();
        let xml = write_timelines(&timelines).unwrap();
        assert_eq!(parse_timelines(&xml).unwrap(), timelines);
    }

    #[test]
    fn typed_state_pivot_filter_and_all_filter_types_round_trip() {
        let tokens = [
            "unknown",
            "count",
            "percent",
            "sum",
            "captionEqual",
            "captionNotEqual",
            "captionBeginsWith",
            "captionNotBeginsWith",
            "captionEndsWith",
            "captionNotEndsWith",
            "captionContains",
            "captionNotContains",
            "captionGreaterThan",
            "captionGreaterThanOrEqual",
            "captionLessThan",
            "captionLessThanOrEqual",
            "captionBetween",
            "captionNotBetween",
            "valueEqual",
            "valueNotEqual",
            "valueGreaterThan",
            "valueGreaterThanOrEqual",
            "valueLessThan",
            "valueLessThanOrEqual",
            "valueBetween",
            "valueNotBetween",
            "dateEqual",
            "dateNotEqual",
            "dateOlderThan",
            "dateOlderThanOrEqual",
            "dateNewerThan",
            "dateNewerThanOrEqual",
            "dateBetween",
            "dateNotBetween",
            "tomorrow",
            "today",
            "yesterday",
            "nextWeek",
            "thisWeek",
            "lastWeek",
            "nextMonth",
            "thisMonth",
            "lastMonth",
            "nextQuarter",
            "thisQuarter",
            "lastQuarter",
            "nextYear",
            "thisYear",
            "lastYear",
            "yearToDate",
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "M1",
            "M2",
            "M3",
            "M4",
            "M5",
            "M6",
            "M7",
            "M8",
            "M9",
            "M10",
            "M11",
            "M12",
        ];
        for token in tokens {
            assert_eq!(PivotFilterType::parse(token).unwrap().as_str(), token)
        }
        let xml = format!(
            r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" xmlns:x="{SML}" name="Timeline_Date" sourceName="Date"><x15:state minimalRefreshVersion="1" lastRefreshVersion="2" pivotCacheId="3" filterType="count"><x15:selection startDate="2024-02-01T00:00:00+02:00" endDate="2024-02-29T23:59:59+02:00"/><x15:bounds startDate="2024-01-01T00:00:00+02:00" endDate="2024-12-31T23:59:59+02:00"/></x15:state><x15:timelinePivotFilter useWholeDay="1" fld="2" id="7" name="Recent"><x:autoFilter ref="A1:B9"><x:filterColumn colId="0"><x:customFilters><x:customFilter operator="greaterThan" val="2"/></x:customFilters></x:filterColumn></x:autoFilter></x15:timelinePivotFilter></x15:timelineCacheDefinition>"#
        );
        let parsed = parse_timeline_cache_definition(xml.as_bytes()).unwrap();
        assert!(
            parsed
                .timeline_pivot_filter
                .as_ref()
                .unwrap()
                .auto_filter
                .is_some()
        );
        let written = write_timeline_cache_definition(&parsed).unwrap();
        assert_eq!(parse_timeline_cache_definition(&written).unwrap(), parsed);
    }

    #[test]
    fn rejects_state_calendar_timezone_range_order_and_filter_presence() {
        let state = |filter: &str, selection: &str, bounds: &str, pivot: &str| {
            format!(
                r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" name="Timeline_Date" sourceName="Date"><x15:state minimalRefreshVersion="0" lastRefreshVersion="0" pivotCacheId="1" filterType="{filter}">{selection}{bounds}</x15:state>{pivot}</x15:timelineCacheDefinition>"#
            )
        };
        let good_bounds =
            r#"<x15:bounds startDate="2024-01-01T00:00:00Z" endDate="2024-12-31T23:59:59Z"/>"#;
        for xml in [
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2024-02-30T00:00:00Z" endDate="2024-12-31T00:00:00Z"/>"#,
                "",
            ),
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2024-01-01T00:00:00+15:00" endDate="2024-12-31T00:00:00+15:00"/>"#,
                "",
            ),
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2025-01-01T00:00:00Z" endDate="2024-01-01T00:00:00Z"/>"#,
                "",
            ),
            state(
                "dateBetween",
                r#"<x15:selection startDate="2023-01-01T00:00:00Z" endDate="2024-02-01T00:00:00Z"/>"#,
                good_bounds,
                "",
            ),
            state(
                "dateBetween",
                "",
                good_bounds,
                r#"<x15:timelinePivotFilter fld="0" id="1"/>"#,
            ),
            state("bogus", "", good_bounds, ""),
        ] {
            assert!(
                parse_timeline_cache_definition(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn package_store_load_and_cross_validation_round_trip() {
        let (mut package, workbook) = package();
        let expected_cache = cache();
        store_timeline_caches(
            &mut package,
            &workbook,
            std::slice::from_ref(&expected_cache),
        )
        .unwrap();
        assert_eq!(
            load_timeline_caches(&package, &workbook).unwrap(),
            vec![expected_cache]
        );
        let expected = WorksheetTimelines {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdTimeline1".into(),
            part_name: "/xl/timelines/timeline1.xml".into(),
            timelines: views(),
        };
        store_worksheet_timelines(&mut package, &workbook, &expected).unwrap();
        assert_eq!(load_timelines(&package, &workbook).unwrap(), vec![expected]);
    }

    #[test]
    fn rejects_hostile_grammar_identity_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x15:timelines xmlns:x15="{X15}"/>"#),
            format!(
                r#"<x15:timelines xmlns:x15="{X15}"><x15:timeline name="x" cache="c" level="4" selectionLevel="0"/></x15:timelines>"#
            ),
            format!(
                r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" name="A1" sourceName="Date"><x15:state/></x15:timelineCacheDefinition>"#
            ),
        ] {
            assert!(
                parse_timelines(xml.as_bytes()).is_err()
                    || parse_timeline_cache_definition(xml.as_bytes()).is_err()
            );
        }
        assert!(parse_timelines(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let mut duplicate = views();
        duplicate.timelines.push(duplicate.timelines[0].clone());
        assert!(write_timelines(&duplicate).is_err());
    }

    #[test]
    fn rejects_package_graph_and_unknown_cache_errors() {
        let (mut package, workbook) = package();
        store_timeline_caches(&mut package, &workbook, &[cache()]).unwrap();
        let mut bad = WorksheetTimelines {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdTimeline1".into(),
            part_name: "/xl/timelines/timeline1.xml".into(),
            timelines: views(),
        };
        bad.timelines.timelines[0].cache = "Missing".into();
        assert!(store_worksheet_timelines(&mut package, &workbook, &bad).is_err());
        let target = PackURI::new("/xl/timelineCaches/timelineCache1.xml").unwrap();
        package
            .get_part_mut(&target)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "x.xml".into(),
                "rIdBad".into(),
                false,
            );
        assert!(load_timeline_caches(&package, &workbook).is_err());
    }
}

#[cfg(test)]
mod xsd_datetime_full_lexical_tests {
    use super::*;

    #[test]
    fn accepts_and_compares_full_applicable_xsd_datetime_space() {
        for value in [
            "2000-02-29T24:00:00Z",
            "-0001-02-29T24:00:00-14:00",
            "12345-12-31T23:59:59.12345678901234567890+14:00",
            "-12345-01-01T00:00:00.00000000000000000001",
        ] {
            validate_datetime(value).unwrap();
        }
        assert_eq!(
            compare_datetime("2000-02-29T24:00:00Z", "2000-03-01T00:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("0001-01-01T00:00:00+14:00", "-0001-12-31T10:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("9999-12-31T24:00:00+14:00", "9999-12-31T10:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("-10000-01-01T00:00:00Z", "-9999-01-01T00:00:00Z").unwrap(),
            std::cmp::Ordering::Less
        );
    }

    #[test]
    fn rejects_non_xsd_datetime_lexemes_and_indeterminate_mixed_zones() {
        for value in [
            "0000-01-01T00:00:00Z",
            "-0000-01-01T00:00:00Z",
            "999-01-01T00:00:00Z",
            "01234-01-01T00:00:00Z",
            "0001-02-29T00:00:00Z",
            "-0002-02-29T00:00:00Z",
            "2000-01-01T24:00:00.1Z",
            "2000-01-01T24:00:01Z",
            "2000-01-01t00:00:00Z",
            "2000-01-01T00:00:00z",
            "2000-01-01T00:00:60Z",
            "2000-01-01T00:00:00+14:01",
            "2000-01-01T00:00:00-15:00",
        ] {
            assert!(validate_datetime(value).is_err(), "accepted {value}");
        }
        assert!(compare_datetime("2000-01-01T00:00:00", "2000-01-01T00:00:00Z").is_err());
    }

    #[test]
    fn timeline_range_preserves_extended_lexemes_round_trip() {
        let range = TimelineRange::new("-12345-01-01T24:00:00", "12345-12-31T24:00:00").unwrap();
        assert_eq!(range.start_date(), "-12345-01-01T24:00:00");
        assert_eq!(range.end_date(), "12345-12-31T24:00:00");
    }
}
