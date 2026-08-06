//! Parser for data-pilot levels and their display, sort, and layout metadata.

use litchi_core::Result;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

use crate::model::data_pilot::model::{
    DisplayInfo, DisplayMemberMode, LayoutInfo, LayoutMode, Level, Member, SortInfo, SortMode,
    SortOrder,
};

use super::super::xml::{
    CALC_EXT_NAMESPACE, is_table, is_table_namespace, optional_attr, optional_bool,
    optional_ns_bool, required_attr, required_bool, required_u64, set_once, skip_foreign_element,
    text_is_whitespace,
};
use super::support::{parse_empty_children, xml_error};
use crate::model::data_pilot::invalid_message;

pub(super) fn level_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Level> {
    Ok(Level {
        show_empty: optional_bool(reader, start, b"show-empty")?,
        repeat_item_labels: optional_ns_bool(
            reader,
            start,
            CALC_EXT_NAMESPACE,
            b"repeat-item-labels",
        )?,
        ..Default::default()
    })
}

pub(super) fn parse_level(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Level> {
    let mut level = level_from_start(reader, start)?;
    let mut subtotals_seen = false;
    let mut members_seen = false;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
                level.subtotals = parse_subtotals(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
                level.members = parse_members(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info") =>
            {
                set_once(
                    &mut level.display,
                    parse_display(reader, element)?,
                    "display info",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-sort-info") =>
            {
                set_once(&mut level.sort, parse_sort(reader, element)?, "sort info")?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-layout-info") =>
            {
                set_once(
                    &mut level.layout,
                    parse_layout(reader, element)?,
                    "layout info",
                )?;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-level") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info")
                    || is_table(&namespace, element, b"data-pilot-sort-info")
                    || is_table(&namespace, element, b"data-pilot-layout-info") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Start(ref element) if !is_table_namespace(&namespace) => {
                skip_foreign_element(reader, element)?
            },
            Event::Empty(_) if !is_table_namespace(&namespace) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-level")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-level")),
        }
        buf.clear();
    }
    Ok(level)
}

fn parse_subtotals(reader: &mut NsReader<&[u8]>) -> Result<Vec<String>> {
    parse_empty_children(
        reader,
        b"data-pilot-subtotals",
        b"data-pilot-subtotal",
        b"function",
    )
}

fn parse_members(reader: &mut NsReader<&[u8]>) -> Result<Vec<Member>> {
    let mut members = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-member") =>
            {
                members.push(Member {
                    name: required_attr(reader, element, b"name")?,
                    display: optional_bool(reader, element, b"display")?,
                    show_details: optional_bool(reader, element, b"show-details")?,
                })
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                break;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-member") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-members")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-members")),
        }
        buf.clear();
    }
    Ok(members)
}

fn parse_display(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<DisplayInfo> {
    Ok(DisplayInfo {
        enabled: required_bool(reader, element, b"enabled")?,
        data_field: required_attr(reader, element, b"data-field")?,
        member_count: required_u64(reader, element, b"member-count")?,
        mode: DisplayMemberMode::parse(&required_attr(reader, element, b"display-member-mode")?)?,
    })
}

fn parse_sort(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SortInfo> {
    Ok(SortInfo {
        mode: SortMode::parse(&required_attr(reader, element, b"sort-mode")?)?,
        data_field: optional_attr(reader, element, b"data-field")?,
        order: SortOrder::parse(&required_attr(reader, element, b"order")?)?,
    })
}

fn parse_layout(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<LayoutInfo> {
    Ok(LayoutInfo {
        mode: LayoutMode::parse(&required_attr(reader, element, b"layout-mode")?)?,
        add_empty_lines: required_bool(reader, element, b"add-empty-lines")?,
    })
}
