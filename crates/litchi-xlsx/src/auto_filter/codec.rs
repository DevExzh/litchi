//! Bounded `SpreadsheetML` auto-filter fragment codec.

use crate::error::{Error, Result};
use crate::sort::{SortBy, SortMethod};
use quick_xml::Writer;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::model::{
    CORE, Calendar, ChildOrder, Color, Column, Condition, Custom, Customs, DateGroup, Definition,
    Dynamic, DynamicType, Grouping, Icon, IconSet, Item, MAX_COLUMNS, MAX_FRAGMENT_BYTES,
    MAX_ITEMS, MAX_SORT_CONDITIONS, MAX_TEXT_CHARS, MAX_UNKNOWN_BYTES, OpaqueFields, Operator,
    Payload, Range, STRICT, State, Top10, UnknownAttribute, UnknownElement, Values, opaque_mut,
};

struct ColumnBuilder {
    column_id: u32,
    hidden_button: bool,
    show_button: bool,
    payload: Option<Payload>,
    opaque: Option<Box<OpaqueFields>>,
}
struct ValuesBuilder {
    blank: bool,
    calendar_type: Calendar,
    items: Vec<Item>,
    opaque: Option<Box<OpaqueFields>>,
}
struct CustomBuilder {
    and: bool,
    filters: Vec<Custom>,
    opaque: Option<Box<OpaqueFields>>,
}
struct SortBuilder {
    reference: Range,
    column_sort: bool,
    case_sensitive: bool,
    sort_method: Option<SortMethod>,
    conditions: Vec<Condition>,
    opaque: Option<Box<OpaqueFields>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum UnknownOwner {
    Root,
    Column,
    Values,
    Customs,
    Sort,
    Payload,
}

pub fn parse_auto_filter_fragment(xml: &[u8]) -> Result<Definition> {
    if xml.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("autoFilter is too large"));
    }
    parse_fragment(xml)
}

pub fn write_auto_filter_fragment(value: &Definition) -> Result<Vec<u8>> {
    let mut x = Vec::new();
    x.extend_from_slice(
        b"<x:autoFilter xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    );
    if let Some(r) = &value.reference {
        parse_range(r.as_str())?;
        a(&mut x, "ref", r.as_str());
    }
    write_unknown_attributes(&mut x, value.opaque.as_deref())?;
    if value.columns.is_empty()
        && value.sort_state.is_none()
        && !has_unknown_children(value.opaque.as_deref())
    {
        x.extend_from_slice(b"/>");
        return Ok(x);
    }
    x.push(b'>');
    let mut ids = HashSet::with_capacity(value.columns.len());
    let max_columns = u32::try_from(MAX_COLUMNS)
        .map_err(|_source| invalid("worksheet column limit exceeds the filterColumn wire type"))?;
    if value
        .columns
        .iter()
        .any(|column| column.column_id >= max_columns || !ids.insert(column.column_id))
    {
        return Err(invalid("invalid or duplicate filterColumn colId"));
    }
    let mut columns = vec![false; value.columns.len()];
    let mut sort_written = false;
    if let Some(opaque) = value.opaque.as_deref() {
        for order in &opaque.order {
            match *order {
                ChildOrder::Column(index) if index < value.columns.len() && !columns[index] => {
                    write_column(&mut x, &value.columns[index])?;
                    columns[index] = true;
                },
                ChildOrder::SortState if !sort_written => {
                    if let Some(state) = value.sort_state.as_ref() {
                        write_state(&mut x, state)?;
                        sort_written = true;
                    }
                },
                ChildOrder::Unknown(index) => write_unknown_index(&mut x, opaque, index)?,
                ChildOrder::Column(_)
                | ChildOrder::SortState
                | ChildOrder::Payload
                | ChildOrder::Item(_)
                | ChildOrder::Custom(_)
                | ChildOrder::Condition(_) => {},
            }
        }
        if opaque.order.is_empty() {
            write_unknown_elements(&mut x, opaque)?;
        }
    }
    for (index, column) in value.columns.iter().enumerate() {
        if !columns[index] {
            write_column(&mut x, column)?;
        }
    }
    if !sort_written && let Some(state) = value.sort_state.as_ref() {
        write_state(&mut x, state)?;
    }
    x.extend_from_slice(b"</x:autoFilter>");
    if x.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("autoFilter is too large"));
    }
    Ok(x)
}
fn write_payload(x: &mut Vec<u8>, p: &Payload) -> Result<()> {
    match p {
        Payload::Values(v) => {
            x.extend_from_slice(b"<x:filters");
            if v.blank {
                a(x, "blank", "1");
            }
            if v.calendar_type != Calendar::None {
                a(x, "calendarType", calendar(v.calendar_type));
            }
            write_unknown_attributes(x, v.opaque.as_deref())?;
            if v.items.is_empty() && !has_unknown_children(v.opaque.as_deref()) {
                x.extend_from_slice(b"/>");
            } else {
                x.push(b'>');
                let mut written = vec![false; v.items.len()];
                if let Some(opaque) = v.opaque.as_deref() {
                    for order in &opaque.order {
                        match *order {
                            ChildOrder::Item(index) if index < v.items.len() && !written[index] => {
                                write_item(x, &v.items[index])?;
                                written[index] = true;
                            },
                            ChildOrder::Unknown(index) => write_unknown_index(x, opaque, index)?,
                            ChildOrder::Column(_)
                            | ChildOrder::SortState
                            | ChildOrder::Payload
                            | ChildOrder::Item(_)
                            | ChildOrder::Custom(_)
                            | ChildOrder::Condition(_) => {},
                        }
                    }
                }
                for (index, item) in v.items.iter().enumerate() {
                    if !written[index] {
                        write_item(x, item)?;
                    }
                }
                x.extend_from_slice(b"</x:filters>");
            }
        },
        Payload::Custom(v) => {
            if !(1..=2).contains(&v.filters.len()) {
                return Err(invalid("customFilters requires one or two filters"));
            }
            x.extend_from_slice(b"<x:customFilters");
            if v.and {
                a(x, "and", "1");
            }
            write_unknown_attributes(x, v.opaque.as_deref())?;
            x.push(b'>');
            let mut written = vec![false; v.filters.len()];
            if let Some(opaque) = v.opaque.as_deref() {
                for order in &opaque.order {
                    match *order {
                        ChildOrder::Custom(index) if index < v.filters.len() && !written[index] => {
                            write_custom(x, &v.filters[index])?;
                            written[index] = true;
                        },
                        ChildOrder::Unknown(index) => write_unknown_index(x, opaque, index)?,
                        ChildOrder::Column(_)
                        | ChildOrder::SortState
                        | ChildOrder::Payload
                        | ChildOrder::Item(_)
                        | ChildOrder::Custom(_)
                        | ChildOrder::Condition(_) => {},
                    }
                }
            }
            for (index, filter) in v.filters.iter().enumerate() {
                if !written[index] {
                    write_custom(x, filter)?;
                }
            }
            x.extend_from_slice(b"</x:customFilters>");
        },
        Payload::Dynamic(v) => {
            x.extend_from_slice(b"<x:dynamicFilter");
            a(x, "type", dynamic(v.filter_type));
            if let Some(q) = v.value {
                a(x, "val", &q.to_string());
            }
            if let Some(q) = v.max_value {
                a(x, "maxVal", &q.to_string());
            }
            write_unknown_attributes(x, v.opaque.as_deref())?;
            x.extend_from_slice(b"/>");
        },
        Payload::Color(v) => {
            x.extend_from_slice(b"<x:colorFilter");
            a(x, "dxfId", &v.differential_format_id.to_string());
            if !v.cell_color {
                a(x, "cellColor", "0");
            }
            write_unknown_attributes(x, v.opaque.as_deref())?;
            x.extend_from_slice(b"/>");
        },
        Payload::Icon(v) => {
            x.extend_from_slice(b"<x:iconFilter");
            a(x, "iconSet", icon(v.icon_set));
            a(x, "iconId", &v.icon_id.to_string());
            write_unknown_attributes(x, v.opaque.as_deref())?;
            x.extend_from_slice(b"/>");
        },
        Payload::Top10(v) => {
            x.extend_from_slice(b"<x:top10");
            if !v.top {
                a(x, "top", "0");
            }
            if v.percent {
                a(x, "percent", "1");
            }
            a(x, "val", &v.value.to_string());
            if let Some(q) = v.filter_value {
                a(x, "filterVal", &q.to_string());
            }
            write_unknown_attributes(x, v.opaque.as_deref())?;
            x.extend_from_slice(b"/>");
        },
    }
    Ok(())
}

fn write_column(x: &mut Vec<u8>, value: &Column) -> Result<()> {
    x.extend_from_slice(b"<x:filterColumn");
    a(x, "colId", &value.column_id.to_string());
    if value.hidden_button {
        a(x, "hiddenButton", "1");
    }
    if !value.show_button {
        a(x, "showButton", "0");
    }
    write_unknown_attributes(x, value.opaque.as_deref())?;
    if value.payload.is_none() && !has_unknown_children(value.opaque.as_deref()) {
        x.extend_from_slice(b"/>");
        return Ok(());
    }
    x.push(b'>');
    let mut payload_written = false;
    if let Some(opaque) = value.opaque.as_deref() {
        for order in &opaque.order {
            match *order {
                ChildOrder::Payload if !payload_written => {
                    if let Some(payload) = value.payload.as_ref() {
                        write_payload(x, payload)?;
                        payload_written = true;
                    }
                },
                ChildOrder::Unknown(index) => write_unknown_index(x, opaque, index)?,
                ChildOrder::Column(_)
                | ChildOrder::SortState
                | ChildOrder::Payload
                | ChildOrder::Item(_)
                | ChildOrder::Custom(_)
                | ChildOrder::Condition(_) => {},
            }
        }
    }
    if !payload_written && let Some(payload) = value.payload.as_ref() {
        write_payload(x, payload)?;
    }
    x.extend_from_slice(b"</x:filterColumn>");
    Ok(())
}

fn write_state(x: &mut Vec<u8>, value: &State) -> Result<()> {
    x.extend_from_slice(b"<x:sortState");
    a(x, "ref", value.reference.as_str());
    if value.column_sort {
        a(x, "columnSort", "1");
    }
    if value.case_sensitive {
        a(x, "caseSensitive", "1");
    }
    if let Some(method) = value.sort_method {
        a(x, "sortMethod", method.as_str());
    }
    write_unknown_attributes(x, value.opaque.as_deref())?;
    if value.conditions.is_empty() && !has_unknown_children(value.opaque.as_deref()) {
        x.extend_from_slice(b"/>");
        return Ok(());
    }
    x.push(b'>');
    let mut written = vec![false; value.conditions.len()];
    if let Some(opaque) = value.opaque.as_deref() {
        for order in &opaque.order {
            match *order {
                ChildOrder::Condition(index)
                    if index < value.conditions.len() && !written[index] =>
                {
                    write_condition(x, value, &value.conditions[index])?;
                    written[index] = true;
                },
                ChildOrder::Unknown(index) => write_unknown_index(x, opaque, index)?,
                ChildOrder::Column(_)
                | ChildOrder::SortState
                | ChildOrder::Payload
                | ChildOrder::Item(_)
                | ChildOrder::Custom(_)
                | ChildOrder::Condition(_) => {},
            }
        }
    }
    for (index, condition) in value.conditions.iter().enumerate() {
        if !written[index] {
            write_condition(x, value, condition)?;
        }
    }
    x.extend_from_slice(b"</x:sortState>");
    Ok(())
}

fn write_condition(x: &mut Vec<u8>, state: &State, value: &Condition) -> Result<()> {
    validate_sort_condition(
        state.reference.as_str(),
        state.column_sort,
        value.reference.as_str(),
        value.sort_by,
        value.custom_list.is_some(),
        value.differential_format_id.is_some(),
        value.icon_set,
        value.icon_id,
    )?;
    x.extend_from_slice(b"<x:sortCondition");
    a(x, "ref", value.reference.as_str());
    if value.descending {
        a(x, "descending", "1");
    }
    if value.sort_by != SortBy::Value {
        a(x, "sortBy", value.sort_by.as_str());
    }
    if let Some(custom_list) = value.custom_list.as_deref() {
        a(x, "customList", custom_list);
    }
    if let Some(dxf) = value.differential_format_id {
        a(x, "dxfId", &dxf.to_string());
    }
    if let Some(icon_set) = value.icon_set {
        a(x, "iconSet", icon(icon_set));
    }
    if let Some(icon_id) = value.icon_id {
        a(x, "iconId", &icon_id.to_string());
    }
    write_unknown_attributes(x, value.opaque.as_deref())?;
    if has_unknown_children(value.opaque.as_deref()) {
        x.push(b'>');
        write_unknown_elements(
            x,
            value.opaque.as_deref().unwrap_or_else(|| {
                crate::error::panic_missing_invariant(
                    "required value was checked before extraction",
                )
            }),
        )?;
        x.extend_from_slice(b"</x:sortCondition>");
    } else {
        x.extend_from_slice(b"/>");
    }
    Ok(())
}

fn write_item(x: &mut Vec<u8>, value: &Item) -> Result<()> {
    match value {
        Item::Value(value) => {
            x.extend_from_slice(b"<x:filter");
            a(x, "val", value);
            x.extend_from_slice(b"/>");
        },
        Item::DateGroup(value) => write_date_group(x, value)?,
    }
    Ok(())
}

fn write_date_group(x: &mut Vec<u8>, value: &DateGroup) -> Result<()> {
    x.extend_from_slice(b"<x:dateGroupItem");
    a(x, "year", &value.year.to_string());
    for (name, field) in [
        ("month", value.month),
        ("day", value.day),
        ("hour", value.hour),
        ("minute", value.minute),
        ("second", value.second),
    ] {
        if let Some(value) = field {
            a(x, name, &value.to_string());
        }
    }
    a(x, "dateTimeGrouping", group(value.grouping));
    write_unknown_attributes(x, value.opaque.as_deref())?;
    if has_unknown_children(value.opaque.as_deref()) {
        x.push(b'>');
        write_unknown_elements(
            x,
            value.opaque.as_deref().unwrap_or_else(|| {
                crate::error::panic_missing_invariant(
                    "required value was checked before extraction",
                )
            }),
        )?;
        x.extend_from_slice(b"</x:dateGroupItem>");
    } else {
        x.extend_from_slice(b"/>");
    }
    Ok(())
}

fn write_custom(x: &mut Vec<u8>, value: &Custom) -> Result<()> {
    x.extend_from_slice(b"<x:customFilter");
    if value.operator != Operator::Equal {
        a(x, "operator", custom(value.operator));
    }
    a(x, "val", &value.value);
    write_unknown_attributes(x, value.opaque.as_deref())?;
    if has_unknown_children(value.opaque.as_deref()) {
        x.push(b'>');
        write_unknown_elements(
            x,
            value.opaque.as_deref().unwrap_or_else(|| {
                crate::error::panic_missing_invariant(
                    "required value was checked before extraction",
                )
            }),
        )?;
        x.extend_from_slice(b"</x:customFilter>");
    } else {
        x.extend_from_slice(b"/>");
    }
    Ok(())
}

fn has_unknown_children(value: Option<&OpaqueFields>) -> bool {
    value.is_some_and(|value| !value.elements.is_empty())
}

fn write_unknown_attributes(x: &mut Vec<u8>, value: Option<&OpaqueFields>) -> Result<()> {
    if let Some(value) = value {
        for attribute in &value.attributes {
            a(x, attribute.name(), attribute.value());
        }
    }
    Ok(())
}

fn write_unknown_elements(x: &mut Vec<u8>, value: &OpaqueFields) -> Result<()> {
    for element in &value.elements {
        x.extend_from_slice(element.as_xml());
        if x.len() > MAX_FRAGMENT_BYTES {
            return Err(invalid("autoFilter is too large"));
        }
    }
    Ok(())
}

fn write_unknown_index(x: &mut Vec<u8>, value: &OpaqueFields, index: usize) -> Result<()> {
    let element = value
        .elements
        .get(index)
        .ok_or_else(|| invalid("invalid autoFilter unknown-child index"))?;
    x.extend_from_slice(element.as_xml());
    if x.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("autoFilter is too large"));
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
    x.push(b'"');
}
fn calendar(v: Calendar) -> &'static str {
    match v {
        Calendar::None => "none",
        Calendar::Gregorian => "gregorian",
        Calendar::GregorianUs => "gregorianUs",
        Calendar::GregorianMeFrench => "gregorianMeFrench",
        Calendar::GregorianArabic => "gregorianArabic",
        Calendar::Hijri => "hijri",
        Calendar::Hebrew => "hebrew",
        Calendar::Taiwan => "taiwan",
        Calendar::Japan => "japan",
        Calendar::Thai => "thai",
        Calendar::Korea => "korea",
        Calendar::Saka => "saka",
    }
}
fn group(v: Grouping) -> &'static str {
    match v {
        Grouping::Year => "year",
        Grouping::Month => "month",
        Grouping::Day => "day",
        Grouping::Hour => "hour",
        Grouping::Minute => "minute",
        Grouping::Second => "second",
    }
}
fn custom(v: Operator) -> &'static str {
    match v {
        Operator::LessThan => "lessThan",
        Operator::LessThanOrEqual => "lessThanOrEqual",
        Operator::NotEqual => "notEqual",
        Operator::Equal => "equal",
        Operator::GreaterThanOrEqual => "greaterThanOrEqual",
        Operator::GreaterThan => "greaterThan",
    }
}
fn dynamic(v: DynamicType) -> &'static str {
    use DynamicType::{
        AboveAverage, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, M1, M2, M3, M4, M5,
        M6, M7, M8, M9, M10, M11, M12, NextMonth, NextQuarter, NextWeek, NextYear, Null, Q1, Q2,
        Q3, Q4, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday,
    };
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
fn icon(v: IconSet) -> &'static str {
    use IconSet::{
        FiveArrows, FiveArrowsGray, FiveBoxes, FiveQuarters, FiveRating, FourArrows,
        FourArrowsGray, FourRating, FourRedToBlack, FourTrafficLights, NoIcons, ThreeArrows,
        ThreeArrowsGray, ThreeFlags, ThreeSigns, ThreeStars, ThreeSymbols, ThreeSymbols2,
        ThreeTrafficLights1, ThreeTrafficLights2, ThreeTriangles,
    };
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
        ThreeStars => "3Stars",
        ThreeTriangles => "3Triangles",
        FiveBoxes => "5Boxes",
        NoIcons => "NoIcons",
    }
}
fn parse_fragment(fragment: &[u8]) -> Result<Definition> {
    let wrapped = wrap(fragment);
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut depth = 0usize;
    let mut root = None;
    let mut closed = false;
    let mut reference = None;
    let mut width = None;
    let mut root_opaque = None;
    let mut columns = Vec::new();
    let mut column: Option<(usize, ColumnBuilder)> = None;
    let mut values: Option<(usize, ValuesBuilder)> = None;
    let mut custom: Option<(usize, CustomBuilder)> = None;
    let mut sort: Option<(usize, SortBuilder)> = None;
    let mut sort_state = None;
    let mut payload_depth = None;
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
                    root_opaque = unknown_attributes(&e, decoder, &[b"ref"])?;
                    if let Some(v) = optional_attr(&e, b"ref", decoder)? {
                        let parsed = parse_range(&v)?;
                        width = Some(parsed.2 - parsed.0 + 1);
                        reference = Some(Range::from_parsed(v));
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
                            calendar_type: Calendar::parse(
                                optional_attr(&e, b"calendarType", decoder)?
                                    .as_deref()
                                    .unwrap_or("none"),
                            )?,
                            items: Vec::new(),
                            opaque: unknown_attributes(&e, decoder, &[b"blank", b"calendarType"])?
                                .or_else(|| Some(Box::new(OpaqueFields::default()))),
                        },
                    ));
                    opaque_mut(
                        &mut column
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Payload)?;
                    payload_depth = Some(depth);
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
                            opaque: unknown_attributes(&e, decoder, &[b"and"])?
                                .or_else(|| Some(Box::new(OpaqueFields::default()))),
                        },
                    ));
                    opaque_mut(
                        &mut column
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Payload)?;
                    payload_depth = Some(depth);
                } else if spreadsheet(&namespace)
                    && values.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(name.as_ref(), b"filter" | b"dateGroupItem")
                {
                    let index = push_value(
                        values.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        name.as_ref(),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut values
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Item(index))?;
                    depth += 1;
                } else if spreadsheet(&namespace)
                    && custom.as_ref().is_some_and(|v| v.0 == depth)
                    && name.as_ref() == b"customFilter"
                {
                    let index = push_custom(
                        custom.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut custom
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Custom(index))?;
                    depth += 1;
                } else if spreadsheet(&namespace)
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(
                        name.as_ref(),
                        b"dynamicFilter" | b"colorFilter" | b"iconFilter" | b"top10"
                    )
                {
                    set_simple_payload(
                        column.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        name.as_ref(),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut column
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Payload)?;
                    depth += 1;
                    payload_depth = Some(depth);
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
                    opaque_mut(&mut root_opaque).push_order(ChildOrder::SortState)?;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortCondition"
                    && sort.as_ref().is_some_and(|v| v.0 == depth)
                {
                    let index = push_sort(
                        sort.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut sort
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Condition(index))?;
                    depth += 1;
                } else if let Some(owner) = unknown_owner(
                    depth,
                    root,
                    column.as_ref(),
                    values.as_ref(),
                    custom.as_ref(),
                    sort.as_ref(),
                    payload_depth,
                ) {
                    let unknown = capture_unknown(&mut reader, Event::Start(e))?;
                    attach_unknown(
                        owner,
                        unknown,
                        &mut root_opaque,
                        &mut column,
                        &mut values,
                        &mut custom,
                        &mut sort,
                    )?;
                } else {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("autoFilter nesting is too deep"))?;
                }
            },
            Event::Empty(e) => {
                let name = e.local_name();
                if spreadsheet(&namespace) && name.as_ref() == b"autoFilter" && root.is_none() {
                    root = Some(0);
                    root_opaque = unknown_attributes(&e, decoder, &[b"ref"])?;
                    if let Some(v) = optional_attr(&e, b"ref", decoder)? {
                        parse_range(&v)?;
                        reference = Some(Range::from_parsed(v));
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
                    opaque_mut(&mut root_opaque)
                        .push_order(ChildOrder::Column(columns.len() - 1))?;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"filters"
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                {
                    ensure_payload_empty(&column)?;
                    column
                        .as_mut()
                        .unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        })
                        .1
                        .payload = Some(Payload::Values(Values {
                        blank: optional_bool(&e, b"blank", decoder)?.unwrap_or(false),
                        calendar_type: Calendar::parse(
                            optional_attr(&e, b"calendarType", decoder)?
                                .as_deref()
                                .unwrap_or("none"),
                        )?,
                        items: Vec::new(),
                        opaque: None,
                    }));
                    opaque_mut(
                        &mut column
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Payload)?;
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
                    let index = push_value(
                        values.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        name.as_ref(),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut values
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Item(index))?;
                } else if spreadsheet(&namespace)
                    && custom.as_ref().is_some_and(|v| v.0 == depth)
                    && name.as_ref() == b"customFilter"
                {
                    let index = push_custom(
                        custom.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut custom
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Custom(index))?;
                } else if spreadsheet(&namespace)
                    && column.as_ref().is_some_and(|v| v.0 == depth)
                    && matches!(
                        name.as_ref(),
                        b"dynamicFilter" | b"colorFilter" | b"iconFilter" | b"top10"
                    )
                {
                    set_simple_payload(
                        column.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        name.as_ref(),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut column
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Payload)?;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortState"
                    && root == Some(depth)
                {
                    if phase > 0 || sort_state.is_some() {
                        return Err(invalid("duplicate sortState"));
                    }
                    phase = 1;
                    sort_state = Some(finish_sort(parse_sort_state(&e, decoder)?));
                    opaque_mut(&mut root_opaque).push_order(ChildOrder::SortState)?;
                } else if spreadsheet(&namespace)
                    && name.as_ref() == b"sortCondition"
                    && sort.as_ref().is_some_and(|v| v.0 == depth)
                {
                    let index = push_sort(
                        sort.as_mut().unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        }),
                        &e,
                        decoder,
                    )?;
                    opaque_mut(
                        &mut sort
                            .as_mut()
                            .unwrap_or_else(|| {
                                crate::error::panic_missing_invariant(
                                    "required value was checked before extraction",
                                )
                            })
                            .1
                            .opaque,
                    )
                    .push_order(ChildOrder::Condition(index))?;
                } else if let Some(owner) = unknown_owner(
                    depth,
                    root,
                    column.as_ref(),
                    values.as_ref(),
                    custom.as_ref(),
                    sort.as_ref(),
                    payload_depth,
                ) {
                    let unknown = UnknownElement::new(e.to_vec())?;
                    attach_unknown(
                        owner,
                        unknown,
                        &mut root_opaque,
                        &mut column,
                        &mut values,
                        &mut custom,
                        &mut sort,
                    )?;
                }
            },
            Event::End(e) => {
                if values.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"filters"
                {
                    let (_, b) = values.take().unwrap_or_else(|| {
                        crate::error::panic_missing_invariant(
                            "required value was checked before extraction",
                        )
                    });
                    let order = (0..b.items.len()).map(ChildOrder::Item).collect::<Vec<_>>();
                    column
                        .as_mut()
                        .unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        })
                        .1
                        .payload = Some(Payload::Values(Values {
                        blank: b.blank,
                        calendar_type: b.calendar_type,
                        items: b.items,
                        opaque: normalize_opaque(b.opaque, &order),
                    }));
                    payload_depth = None;
                }
                if custom.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"customFilters"
                {
                    let (_, b) = custom.take().unwrap_or_else(|| {
                        crate::error::panic_missing_invariant(
                            "required value was checked before extraction",
                        )
                    });
                    if !(1..=2).contains(&b.filters.len()) {
                        return Err(invalid(
                            "customFilters requires one or two customFilter children",
                        ));
                    }
                    let order = (0..b.filters.len())
                        .map(ChildOrder::Custom)
                        .collect::<Vec<_>>();
                    column
                        .as_mut()
                        .unwrap_or_else(|| {
                            crate::error::panic_missing_invariant(
                                "required value was checked before extraction",
                            )
                        })
                        .1
                        .payload = Some(Payload::Custom(Customs {
                        and: b.and,
                        filters: b.filters,
                        opaque: normalize_opaque(b.opaque, &order),
                    }));
                    payload_depth = None;
                }
                if column.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"filterColumn"
                {
                    let (_, b) = column.take().unwrap_or_else(|| {
                        crate::error::panic_missing_invariant(
                            "required value was checked before extraction",
                        )
                    });
                    columns.push(finish_column(b)?);
                    opaque_mut(&mut root_opaque)
                        .push_order(ChildOrder::Column(columns.len() - 1))?;
                    if columns.len() > MAX_COLUMNS {
                        return Err(invalid("too many filter columns"));
                    }
                }
                if sort.as_ref().is_some_and(|v| v.0 == depth)
                    && e.local_name().as_ref() == b"sortState"
                {
                    let (_, b) = sort.take().unwrap_or_else(|| {
                        crate::error::panic_missing_invariant(
                            "required value was checked before extraction",
                        )
                    });
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !closed || column.is_some() || values.is_some() || custom.is_some() || sort.is_some() {
        return Err(invalid("unterminated autoFilter"));
    }
    let mut ids = HashSet::with_capacity(columns.len());
    if columns.iter().any(|v| !ids.insert(v.column_id)) {
        return Err(invalid("duplicate filterColumn colId"));
    }
    let mut root_order = (0..columns.len())
        .map(ChildOrder::Column)
        .collect::<Vec<_>>();
    if sort_state.is_some() {
        root_order.push(ChildOrder::SortState);
    }
    Ok(Definition {
        reference,
        columns,
        sort_state,
        opaque: normalize_opaque(root_opaque, &root_order),
    })
}

fn parse_column(e: &BytesStart<'_>, d: Decoder, width: Option<u32>) -> Result<ColumnBuilder> {
    let id = required_u32(e, b"colId", d)?;
    let max_columns = u32::try_from(MAX_COLUMNS)
        .map_err(|_source| invalid("worksheet column limit exceeds the filterColumn wire type"))?;
    if id >= max_columns || width.is_some_and(|w| id >= w) {
        return Err(invalid("filterColumn colId is outside autoFilter range"));
    }
    Ok(ColumnBuilder {
        column_id: id,
        hidden_button: optional_bool(e, b"hiddenButton", d)?.unwrap_or(false),
        show_button: optional_bool(e, b"showButton", d)?.unwrap_or(true),
        payload: None,
        opaque: unknown_attributes(e, d, &[b"colId", b"hiddenButton", b"showButton"])?,
    })
}

fn unknown_attributes(
    e: &BytesStart<'_>,
    d: Decoder,
    known: &[&[u8]],
) -> Result<Option<Box<OpaqueFields>>> {
    let mut opaque = OpaqueFields::default();
    for attribute in e.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") || known.contains(&name) {
            continue;
        }
        let name = std::str::from_utf8(name)
            .map_err(|_source| invalid("autoFilter attribute name is not UTF-8"))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        opaque.push_attribute(UnknownAttribute::new(name, value)?)?;
    }
    Ok((!opaque.attributes.is_empty()).then(|| Box::new(opaque)))
}

fn unknown_owner(
    depth: usize,
    root: Option<usize>,
    column: Option<&(usize, ColumnBuilder)>,
    values: Option<&(usize, ValuesBuilder)>,
    custom: Option<&(usize, CustomBuilder)>,
    sort: Option<&(usize, SortBuilder)>,
    payload_depth: Option<usize>,
) -> Option<UnknownOwner> {
    if values.is_some_and(|value| value.0 == depth) {
        Some(UnknownOwner::Values)
    } else if custom.is_some_and(|value| value.0 == depth) {
        Some(UnknownOwner::Customs)
    } else if sort.is_some_and(|value| value.0 == depth) {
        Some(UnknownOwner::Sort)
    } else if payload_depth == Some(depth) {
        Some(UnknownOwner::Payload)
    } else if column.is_some_and(|value| value.0 == depth) {
        Some(UnknownOwner::Column)
    } else if root == Some(depth) {
        Some(UnknownOwner::Root)
    } else {
        None
    }
}

fn capture_unknown(reader: &mut NsReader<&[u8]>, first: Event<'static>) -> Result<UnknownElement> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(first.clone()).map_err(xml_error)?;
    let mut depth = match first {
        Event::Start(_) => 1usize,
        Event::Empty(_) => 0,
        Event::End(_)
        | Event::Text(_)
        | Event::CData(_)
        | Event::Comment(_)
        | Event::Decl(_)
        | Event::PI(_)
        | Event::DocType(_)
        | Event::GeneralRef(_)
        | Event::Eof => {
            return Err(invalid(
                "unknown autoFilter capture did not start at an element",
            ));
        },
    };
    while depth != 0 {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        writer.write_event(event.clone()).map_err(xml_error)?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("unknown autoFilter nesting is too deep"))?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unknown autoFilter nesting underflow"))?;
            },
            Event::Eof => return Err(invalid("unterminated unknown autoFilter element")),
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        if writer.get_ref().len() > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown autoFilter element is too large"));
        }
    }
    UnknownElement::new(writer.into_inner())
}

fn attach_unknown(
    owner: UnknownOwner,
    element: UnknownElement,
    root: &mut Option<Box<OpaqueFields>>,
    column: &mut Option<(usize, ColumnBuilder)>,
    values: &mut Option<(usize, ValuesBuilder)>,
    custom: &mut Option<(usize, CustomBuilder)>,
    sort: &mut Option<(usize, SortBuilder)>,
) -> Result<()> {
    let target = match owner {
        UnknownOwner::Root => root,
        UnknownOwner::Column => {
            &mut column
                .as_mut()
                .unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                })
                .1
                .opaque
        },
        UnknownOwner::Values => {
            &mut values
                .as_mut()
                .unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                })
                .1
                .opaque
        },
        UnknownOwner::Customs => {
            &mut custom
                .as_mut()
                .unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                })
                .1
                .opaque
        },
        UnknownOwner::Sort => {
            &mut sort
                .as_mut()
                .unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                })
                .1
                .opaque
        },
        UnknownOwner::Payload => {
            let payload = column
                .as_mut()
                .and_then(|value| value.1.payload.as_mut())
                .ok_or_else(|| invalid("unknown autoFilter payload has no owner"))?;
            let opaque = match payload {
                Payload::Values(value) => &mut value.opaque,
                Payload::Custom(value) => &mut value.opaque,
                Payload::Dynamic(value) => &mut value.opaque,
                Payload::Color(value) => &mut value.opaque,
                Payload::Icon(value) => &mut value.opaque,
                Payload::Top10(value) => &mut value.opaque,
            };
            let index = opaque_mut(opaque).push_element(element)?;
            opaque_mut(opaque).push_order(ChildOrder::Unknown(index))?;
            return Ok(());
        },
    };
    let index = opaque_mut(target).push_element(element)?;
    opaque_mut(target).push_order(ChildOrder::Unknown(index))
}
fn finish_column(b: ColumnBuilder) -> Result<Column> {
    let order = if b.payload.is_some() {
        vec![ChildOrder::Payload]
    } else {
        Vec::new()
    };
    Ok(Column {
        column_id: b.column_id,
        hidden_button: b.hidden_button,
        show_button: b.show_button,
        payload: b.payload,
        opaque: normalize_opaque(b.opaque, &order),
    })
}

fn normalize_opaque(
    opaque: Option<Box<OpaqueFields>>,
    canonical_order: &[ChildOrder],
) -> Option<Box<OpaqueFields>> {
    match opaque {
        Some(value)
            if value.attributes.is_empty()
                && value.elements.is_empty()
                && value.order == canonical_order =>
        {
            None
        },
        other => other,
    }
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
) -> Result<usize> {
    if v.1.items.len() == MAX_ITEMS {
        return Err(invalid("too many filter values"));
    }
    if name == b"filter" {
        let value = required_attr(e, b"val", d)?;
        bounded(&value)?;
        v.1.items.push(Item::Value(value));
    } else {
        v.1.items.push(Item::DateGroup(parse_date_group(e, d)?));
    }
    Ok(v.1.items.len() - 1)
}
fn parse_date_group(e: &BytesStart<'_>, d: Decoder) -> Result<DateGroup> {
    let grouping = Grouping::parse(&required_attr(e, b"dateTimeGrouping", d)?)?;
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
        Grouping::Year => 0,
        Grouping::Month => 1,
        Grouping::Day => 2,
        Grouping::Hour => 3,
        Grouping::Minute => 4,
        Grouping::Second => 5,
    };
    let present = [month, day, hour, minute, second]
        .iter()
        .take(required)
        .all(Option::is_some);
    if !present {
        return Err(invalid("date-group components do not match grouping"));
    }
    Ok(DateGroup {
        year: u16::try_from(year)
            .map_err(|_source| invalid("date-group year exceeds the unsigned 16-bit wire type"))?,
        month,
        day,
        hour,
        minute,
        second,
        grouping,
        opaque: unknown_attributes(
            e,
            d,
            &[
                b"year",
                b"month",
                b"day",
                b"hour",
                b"minute",
                b"second",
                b"dateTimeGrouping",
            ],
        )?,
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

fn push_custom(v: &mut (usize, CustomBuilder), e: &BytesStart<'_>, d: Decoder) -> Result<usize> {
    if v.1.filters.len() == 2 {
        return Err(invalid("customFilters has more than two conditions"));
    }
    let value = required_attr(e, b"val", d)?;
    bounded(&value)?;
    v.1.filters.push(Custom {
        operator: Operator::parse(
            optional_attr(e, b"operator", d)?
                .as_deref()
                .unwrap_or("equal"),
        )?,
        value,
        opaque: unknown_attributes(e, d, &[b"operator", b"val"])?,
    });
    Ok(v.1.filters.len() - 1)
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
            if max_value.is_some() && value.is_none() {
                return Err(invalid("dynamicFilter maxVal requires val"));
            }
            if let (Some(value), Some(max_value)) = (value, max_value)
                && value >= max_value
            {
                return Err(invalid("dynamicFilter val must be less than maxVal"));
            }
            Payload::Dynamic(Dynamic {
                filter_type: DynamicType::parse(&required_attr(e, b"type", d)?)?,
                value,
                max_value,
                opaque: unknown_attributes(e, d, &[b"type", b"val", b"maxVal"])?,
            })
        },
        b"colorFilter" => Payload::Color(Color {
            differential_format_id: required_u32(e, b"dxfId", d)?,
            cell_color: optional_bool(e, b"cellColor", d)?.unwrap_or(true),
            opaque: unknown_attributes(e, d, &[b"dxfId", b"cellColor"])?,
        }),
        b"iconFilter" => {
            let set = IconSet::parse(&required_attr(e, b"iconSet", d)?)?;
            let id = required_u32(e, b"iconId", d)?;
            if (set == IconSet::NoIcons && id != 0)
                || (set != IconSet::NoIcons && id >= set.cardinality())
            {
                return Err(invalid("iconFilter iconId exceeds icon-set cardinality"));
            }
            Payload::Icon(Icon {
                icon_set: set,
                icon_id: id,
                opaque: unknown_attributes(e, d, &[b"iconSet", b"iconId"])?,
            })
        },
        b"top10" => {
            let value = required_f64(e, b"val", d)?;
            let percent = optional_bool(e, b"percent", d)?.unwrap_or(false);
            if (!percent && !(1.0..=500.0).contains(&value))
                || (percent && !(0.0..=100.0).contains(&value))
            {
                return Err(invalid("top10 val is out of range"));
            }
            Payload::Top10(Top10 {
                top: optional_bool(e, b"top", d)?.unwrap_or(true),
                percent,
                value,
                filter_value: optional_f64(e, b"filterVal", d)?,
                opaque: unknown_attributes(e, d, &[b"top", b"percent", b"val", b"filterVal"])?,
            })
        },
        _ => return Err(invalid("unsupported filterColumn payload")),
    });
    Ok(())
}

fn parse_sort_state(e: &BytesStart<'_>, d: Decoder) -> Result<SortBuilder> {
    let reference = Range::from_parsed(required_attr(e, b"ref", d)?);
    parse_range(reference.as_str())?;
    let method = optional_attr(e, b"sortMethod", d)?
        .map(|value| {
            value
                .parse::<SortMethod>()
                .map_err(|error| invalid(error.to_string()))
        })
        .transpose()?;
    Ok(SortBuilder {
        reference,
        column_sort: optional_bool(e, b"columnSort", d)?.unwrap_or(false),
        case_sensitive: optional_bool(e, b"caseSensitive", d)?.unwrap_or(false),
        sort_method: method,
        conditions: Vec::new(),
        opaque: unknown_attributes(
            e,
            d,
            &[b"ref", b"columnSort", b"caseSensitive", b"sortMethod"],
        )?,
    })
}
fn push_sort(s: &mut (usize, SortBuilder), e: &BytesStart<'_>, d: Decoder) -> Result<usize> {
    if s.1.conditions.len() == MAX_SORT_CONDITIONS {
        return Err(invalid("too many sort conditions"));
    }
    let reference = Range::from_parsed(required_attr(e, b"ref", d)?);
    parse_range(reference.as_str())?;
    let sort_by = optional_attr(e, b"sortBy", d)?
        .map(|value| {
            value
                .parse::<SortBy>()
                .map_err(|error| invalid(error.to_string()))
        })
        .transpose()?
        .unwrap_or(SortBy::Value);
    let dxf = optional_u32(e, b"dxfId", d)?;
    let icon = optional_attr(e, b"iconSet", d)?
        .map(|v| IconSet::parse(&v))
        .transpose()?;
    let icon_id = optional_u32(e, b"iconId", d)?;
    match sort_by {
        SortBy::CellColor | SortBy::FontColor if dxf.is_none() => {
            return Err(invalid("color sort requires dxfId"));
        },
        SortBy::Icon if icon.is_none() => return Err(invalid("icon sort requires iconSet")),
        SortBy::Icon
            if icon_id.is_some_and(|v| {
                v >= icon
                    .unwrap_or_else(|| {
                        crate::error::panic_missing_invariant(
                            "required value was checked before extraction",
                        )
                    })
                    .cardinality()
            }) =>
        {
            return Err(invalid("sort iconId exceeds icon-set cardinality"));
        },
        SortBy::Value | SortBy::CellColor | SortBy::FontColor | SortBy::Icon => {},
    }
    let custom = optional_attr(e, b"customList", d)?;
    if custom
        .as_ref()
        .is_some_and(|v| v.chars().count() > MAX_TEXT_CHARS)
    {
        return Err(invalid("custom sort list is too large"));
    }
    if let Some(custom) = custom.as_deref() {
        bounded(custom)?;
    }
    validate_sort_condition(
        s.1.reference.as_str(),
        s.1.column_sort,
        reference.as_str(),
        sort_by,
        custom.is_some(),
        dxf.is_some(),
        icon,
        icon_id,
    )?;
    s.1.conditions.push(Condition {
        reference,
        descending: optional_bool(e, b"descending", d)?.unwrap_or(false),
        sort_by,
        custom_list: custom,
        differential_format_id: dxf,
        icon_set: icon,
        icon_id,
        opaque: unknown_attributes(
            e,
            d,
            &[
                b"ref",
                b"descending",
                b"sortBy",
                b"customList",
                b"dxfId",
                b"iconSet",
                b"iconId",
            ],
        )?,
    });
    Ok(s.1.conditions.len() - 1)
}

fn validate_sort_condition(
    state_reference: &str,
    column_sort: bool,
    condition_reference: &str,
    sort_by: SortBy,
    has_custom_list: bool,
    has_dxf: bool,
    icon_set: Option<IconSet>,
    icon_id: Option<u32>,
) -> Result<()> {
    parse_range(state_reference)?;
    let condition = parse_range(condition_reference)?;
    if column_sort {
        if condition.1 != condition.3 {
            return Err(invalid("row sortCondition must select one row"));
        }
    } else if condition.0 != condition.2 {
        return Err(invalid("column sortCondition must select one column"));
    }
    match sort_by {
        SortBy::Value => {
            if has_dxf || icon_set.is_some() || icon_id.is_some() {
                return Err(invalid(
                    "value sortCondition cannot specify color or icon metadata",
                ));
            }
        },
        SortBy::CellColor | SortBy::FontColor => {
            if !has_dxf || has_custom_list || icon_set.is_some() || icon_id.is_some() {
                return Err(invalid("color sortCondition metadata is invalid"));
            }
        },
        SortBy::Icon => {
            if has_custom_list || has_dxf || icon_id.is_some_and(|_| icon_set.is_none()) {
                return Err(invalid("icon sortCondition metadata is invalid"));
            }
            if let (Some(set), Some(id)) = (icon_set, icon_id) {
                if set == IconSet::NoIcons {
                    if id != 0 {
                        return Err(invalid("NoIcons sortCondition requires iconId 0"));
                    }
                } else if id >= set.cardinality() {
                    return Err(invalid("sort iconId exceeds icon-set cardinality"));
                }
            }
        },
    }
    Ok(())
}
fn finish_sort(s: SortBuilder) -> State {
    let order = (0..s.conditions.len())
        .map(ChildOrder::Condition)
        .collect::<Vec<_>>();
    State {
        reference: s.reference,
        column_sort: s.column_sort,
        case_sensitive: s.case_sensitive,
        sort_method: s.sort_method,
        conditions: s.conditions,
        opaque: normalize_opaque(s.opaque, &order),
    }
}

pub(super) fn parse_range(v: &str) -> Result<(u32, u32, u32, u32)> {
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
                .map_err(|_source| invalid(format!("invalid unsigned integer '{v}'")))
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
                .map_err(|_source| invalid(format!("invalid number '{v}'")))?;
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
    if v.chars().count() > MAX_TEXT_CHARS {
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
fn xml_error(e: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(e.to_string()))
}
fn invalid(e: impl Into<String>) -> Error {
    Error::Invalid(e.into())
}
