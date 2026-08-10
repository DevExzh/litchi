//! Bounded, namespace-aware annotation scanner.

use super::model::{
    ActiveBuilder, Annotation, AnnotationHost, AnnotationPosition, EndMarker, Frame, FrameKind,
    Record, Scan, Site, Span,
};
use super::{invalid, invalid_error};
use litchi_core::{Error, Result};
use litchi_odf_common::annotation::Builder;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_ANNOTATIONS: usize = 65_536;
const MAX_EVENTS: usize = 1_000_000;

pub(crate) fn scan(xml: &str, host: AnnotationHost) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("annotation host XML exceeds size limit");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut frames: Vec<Frame> = Vec::new();
    let mut namespaces = BTreeMap::new();
    let mut builders: Vec<ActiveBuilder> = Vec::new();
    let mut records: Vec<Record> = Vec::new();
    let mut ends = Vec::new();
    let mut sites = Vec::new();
    let mut next_sheet = 0usize;
    let mut next_page = 0usize;
    let mut next_paragraph = 0usize;
    let mut events = 0usize;
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid annotation host XML: {error}"))
            })?;
        let ns = namespace(&resolved);
        let end = position(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("annotation event overflow"))?;
        if events > MAX_EVENTS {
            return invalid("annotation host exceeds event limit");
        }
        match event {
            Event::Start(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let local = local.as_ref();
                if ns == Ns::Office && local == b"annotation-end" {
                    return invalid("office:annotation-end must be an empty element");
                }
                if ns == Ns::Office && local == b"annotation" {
                    for active in &mut builders {
                        active.builder.start(&element, reader.decoder())?;
                    }
                    if records.len() >= MAX_ANNOTATIONS {
                        return invalid("document exceeds annotation limit");
                    }
                    let anchor = current_position(&frames, &sites, host)?;
                    let record = records.len();
                    records.push(Record {
                        span: Span {
                            start,
                            end: 0,
                            close_start: None,
                            qname: qname(element.name().as_ref())?,
                        },
                        parent_start: frames.last().map_or(0, |frame| frame.start),
                        annotation: None,
                        start_position: anchor,
                        end: None,
                    });
                    builders.push(ActiveBuilder {
                        record,
                        builder: Builder::new(&element, reader.decoder(), namespaces.clone())?,
                    });
                    frames.push(Frame {
                        start,
                        kind: FrameKind::Annotation { record },
                        namespace_changes: changes,
                    });
                } else {
                    for active in &mut builders {
                        active.builder.start(&element, reader.decoder())?;
                    }
                    let kind = structural_kind(
                        &reader,
                        &element,
                        ns,
                        start,
                        end,
                        false,
                        host,
                        &mut frames,
                        &mut sites,
                        &mut next_sheet,
                        &mut next_page,
                        &mut next_paragraph,
                    )?;
                    frames.push(Frame {
                        start,
                        kind,
                        namespace_changes: changes,
                    });
                }
                if frames.len() > MAX_DEPTH {
                    return invalid("annotation XML nesting exceeds limit");
                }
            },
            Event::Empty(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let local = local.as_ref();
                if ns == Ns::Office && local == b"annotation" {
                    for active in &mut builders {
                        active.builder.empty(&element, reader.decoder())?;
                    }
                    if records.len() >= MAX_ANNOTATIONS {
                        return invalid("document exceeds annotation limit");
                    }
                    let anchor = current_position(&frames, &sites, host)?;
                    let builder = Builder::new(&element, reader.decoder(), namespaces.clone())?;
                    records.push(Record {
                        span: Span {
                            start,
                            end,
                            close_start: None,
                            qname: qname(element.name().as_ref())?,
                        },
                        parent_start: frames.last().map_or(0, |frame| frame.start),
                        annotation: Some(builder.finish()?),
                        start_position: anchor,
                        end: None,
                    });
                } else if ns == Ns::Office && local == b"annotation-end" {
                    for active in &mut builders {
                        active.builder.empty(&element, reader.decoder())?;
                    }
                    let name = required_attribute(
                        &reader,
                        &element,
                        OFFICE_NS,
                        b"name",
                        "annotation end",
                    )?;
                    ends.push(EndMarker {
                        span: Span {
                            start,
                            end,
                            close_start: None,
                            qname: qname(element.name().as_ref())?,
                        },
                        name,
                        position: current_position(&frames, &sites, host)?,
                    });
                } else {
                    for active in &mut builders {
                        active.builder.empty(&element, reader.decoder())?;
                    }
                    let _kind = structural_kind(
                        &reader,
                        &element,
                        ns,
                        start,
                        end,
                        true,
                        host,
                        &mut frames,
                        &mut sites,
                        &mut next_sheet,
                        &mut next_page,
                        &mut next_paragraph,
                    )?;
                }
                restore_namespaces(changes, &mut namespaces);
            },
            Event::End(element) => {
                let local = element.local_name();
                let local = local.as_ref();
                if ns == Ns::Office && local == b"annotation" {
                    let frame = frames
                        .pop()
                        .ok_or_else(|| invalid_error("annotation XML depth underflow"))?;
                    let FrameKind::Annotation { record } = frame.kind else {
                        return invalid("mismatched office:annotation end");
                    };
                    let finished = builders
                        .pop()
                        .ok_or_else(|| invalid_error("missing annotation builder"))?;
                    if finished.record != record {
                        return invalid("mismatched nested annotation");
                    }
                    for active in &mut builders {
                        active.builder.end_element()?;
                    }
                    records[record].span.end = end;
                    records[record].span.close_start = Some(start);
                    records[record].annotation = Some(finished.builder.finish()?);
                    restore_namespaces(frame.namespace_changes, &mut namespaces);
                } else {
                    for active in &mut builders {
                        active.builder.end_element()?;
                    }
                    let frame = frames
                        .pop()
                        .ok_or_else(|| invalid_error("annotation XML depth underflow"))?;
                    finish_site(&frame.kind, start, end, &mut sites);
                    restore_namespaces(frame.namespace_changes, &mut namespaces);
                }
            },
            Event::Text(value) => {
                for active in &mut builders {
                    active.builder.text(&value)?;
                }
            },
            Event::CData(value) => {
                for active in &mut builders {
                    active.builder.cdata(&value)?;
                }
            },
            Event::GeneralRef(value) => {
                for active in &mut builders {
                    active.builder.reference(&value)?;
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in annotation host XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !frames.is_empty() || !builders.is_empty() {
        return invalid("unterminated annotation host XML");
    }
    pair_ranges(&mut records, ends)?;
    sites.sort_by_key(|site| site.span.start);
    Ok(Scan { records, sites })
}

#[allow(clippy::too_many_arguments)]
fn structural_kind(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    ns: Ns,
    start: usize,
    end: usize,
    empty: bool,
    host: AnnotationHost,
    frames: &mut [Frame],
    sites: &mut Vec<Site>,
    next_sheet: &mut usize,
    next_page: &mut usize,
    next_paragraph: &mut usize,
) -> Result<FrameKind> {
    let local = element.local_name();
    let local = local.as_ref();
    if host == AnnotationHost::Spreadsheet && ns == Ns::Table && local == b"table" {
        let sheet = *next_sheet;
        *next_sheet = next_sheet
            .checked_add(1)
            .ok_or_else(|| invalid_error("sheet index overflow"))?;
        return Ok(FrameKind::Table { sheet, next_row: 0 });
    }
    if host == AnnotationHost::Spreadsheet && ns == Ns::Table && local == b"table-row" {
        let repeat =
            optional_usize(reader, element, TABLE_NS, b"number-rows-repeated")?.unwrap_or(1);
        let (sheet, row) = take_table_row(frames, repeat)?;
        return Ok(FrameKind::Row {
            sheet,
            row,
            next_column: 0,
        });
    }
    if host == AnnotationHost::Spreadsheet
        && ns == Ns::Table
        && matches!(local, b"table-cell" | b"covered-table-cell")
    {
        let repeat =
            optional_usize(reader, element, TABLE_NS, b"number-columns-repeated")?.unwrap_or(1);
        let (sheet, row, column) = take_row_cell(frames, repeat)?;
        let position = AnnotationPosition::SpreadsheetCell {
            sheet_index: sheet,
            row,
            column,
        };
        let site = push_site(sites, position, start, end, empty, element)?;
        return Ok(FrameKind::Cell { site });
    }
    if host == AnnotationHost::Presentation && ns == Ns::Draw && local == b"page" {
        let page = *next_page;
        *next_page = next_page
            .checked_add(1)
            .ok_or_else(|| invalid_error("page index overflow"))?;
        let site = push_site(
            sites,
            AnnotationPosition::PresentationPage { page_index: page },
            start,
            end,
            empty,
            element,
        )?;
        return Ok(FrameKind::Page { site, page });
    }
    if host == AnnotationHost::Presentation
        && ns == Ns::Draw
        && local != b"page"
        && let Some(name) = optional_attribute(reader, element, DRAW_NS, b"name")?
        && let Some(page) = current_page(frames)
    {
        let site = push_site(
            sites,
            AnnotationPosition::PresentationShape {
                page_index: page,
                shape_name: name,
            },
            start,
            end,
            empty,
            element,
        )?;
        return Ok(FrameKind::Shape { site });
    }
    if ns == Ns::Text && matches!(local, b"p" | b"h") {
        let position = if let Some(annotation) = current_annotation(frames) {
            Some(AnnotationPosition::AnnotationBody {
                annotation_index: annotation,
            })
        } else if host == AnnotationHost::Text {
            let index = *next_paragraph;
            *next_paragraph = next_paragraph
                .checked_add(1)
                .ok_or_else(|| invalid_error("paragraph index overflow"))?;
            Some(AnnotationPosition::TextParagraph {
                paragraph_index: index,
            })
        } else {
            None
        };
        let site = position
            .map(|position| push_site(sites, position, start, end, empty, element))
            .transpose()?;
        return Ok(FrameKind::Paragraph { site });
    }
    Ok(FrameKind::Other)
}

fn take_table_row(frames: &mut [Frame], repeat: usize) -> Result<(usize, usize)> {
    if repeat == 0 {
        return invalid("table row repeat count cannot be zero");
    }
    for frame in frames.iter_mut().rev() {
        if let FrameKind::Table { sheet, next_row } = &mut frame.kind {
            let row = *next_row;
            *next_row = next_row
                .checked_add(repeat)
                .ok_or_else(|| invalid_error("table row index overflow"))?;
            return Ok((*sheet, row));
        }
    }
    invalid("table row is outside a table")
}

fn take_row_cell(frames: &mut [Frame], repeat: usize) -> Result<(usize, usize, usize)> {
    if repeat == 0 {
        return invalid("table column repeat count cannot be zero");
    }
    for frame in frames.iter_mut().rev() {
        if let FrameKind::Row {
            sheet,
            row,
            next_column,
        } = &mut frame.kind
        {
            let column = *next_column;
            *next_column = next_column
                .checked_add(repeat)
                .ok_or_else(|| invalid_error("table column index overflow"))?;
            return Ok((*sheet, *row, column));
        }
    }
    invalid("table cell is outside a row")
}

fn push_site(
    sites: &mut Vec<Site>,
    position: AnnotationPosition,
    start: usize,
    end: usize,
    empty: bool,
    element: &BytesStart<'_>,
) -> Result<usize> {
    let index = sites.len();
    sites.push(Site {
        position,
        span: Span {
            start,
            end: if empty { end } else { 0 },
            close_start: None,
            qname: qname(element.name().as_ref())?,
        },
    });
    Ok(index)
}

fn finish_site(kind: &FrameKind, close_start: usize, end: usize, sites: &mut [Site]) {
    let site = match kind {
        FrameKind::Cell { site } | FrameKind::Page { site, .. } | FrameKind::Shape { site } => {
            Some(*site)
        },
        FrameKind::Paragraph { site } => *site,
        _ => None,
    };
    if let Some(site) = site {
        sites[site].span.end = end;
        sites[site].span.close_start = Some(close_start);
    }
}

fn current_position(
    frames: &[Frame],
    sites: &[Site],
    host: AnnotationHost,
) -> Result<AnnotationPosition> {
    if let Some(annotation) = current_annotation(frames) {
        return Ok(AnnotationPosition::AnnotationBody {
            annotation_index: annotation,
        });
    }
    match host {
        AnnotationHost::Text => frames
            .iter()
            .rev()
            .find_map(|frame| match &frame.kind {
                FrameKind::Paragraph { site: Some(site) } => Some(*site),
                _ => None,
            })
            .and_then(|site| sites.get(site).map(|site| site.position.clone()))
            .ok_or_else(|| invalid_error("text annotation is outside a paragraph")),
        AnnotationHost::Spreadsheet => frames
            .iter()
            .rev()
            .find_map(|frame| match &frame.kind {
                FrameKind::Cell { site } => Some(*site),
                _ => None,
            })
            .and_then(|site| sites.get(site).map(|site| site.position.clone()))
            .ok_or_else(|| invalid_error("spreadsheet annotation is outside a cell")),
        AnnotationHost::Presentation => frames
            .iter()
            .rev()
            .find_map(|frame| match &frame.kind {
                FrameKind::Shape { site } => Some(*site),
                FrameKind::Page { site, .. } => Some(*site),
                _ => None,
            })
            .and_then(|site| sites.get(site).map(|site| site.position.clone()))
            .ok_or_else(|| invalid_error("presentation annotation is outside a page")),
    }
}

fn current_annotation(frames: &[Frame]) -> Option<usize> {
    frames.iter().rev().find_map(|frame| match &frame.kind {
        FrameKind::Annotation { record } => Some(*record),
        _ => None,
    })
}

fn current_page(frames: &[Frame]) -> Option<usize> {
    frames.iter().rev().find_map(|frame| match &frame.kind {
        FrameKind::Page { page, .. } => Some(*page),
        _ => None,
    })
}

fn pair_ranges(records: &mut [Record], ends: Vec<EndMarker>) -> Result<()> {
    let mut starts = HashMap::new();
    for (index, record) in records.iter().enumerate() {
        if let Some(name) = record.annotation.as_ref().and_then(Annotation::name) {
            if name.is_empty() {
                return invalid("annotation office:name cannot be empty");
            }
            if starts.insert(name.to_string(), index).is_some() {
                return invalid(format!("duplicate annotation name '{name}'"));
            }
        }
    }
    for marker in ends {
        let index = *starts.get(&marker.name).ok_or_else(|| {
            invalid_error(format!(
                "annotation end '{}' has no matching start",
                marker.name
            ))
        })?;
        if records[index].end.is_some() {
            return invalid(format!(
                "annotation '{}' has multiple end markers",
                marker.name
            ));
        }
        if marker.span.start <= records[index].span.start {
            return invalid(format!(
                "annotation end '{}' precedes its start",
                marker.name
            ));
        }
        if marker.span.start < records[index].span.end {
            return invalid(format!(
                "annotation end '{}' occurs inside its annotation body",
                marker.name
            ));
        }
        records[index].end = Some((marker.span, marker.position));
    }
    let mut ranges: Vec<(usize, usize)> = records
        .iter()
        .filter_map(|record| {
            record
                .end
                .as_ref()
                .map(|(end, _)| (record.span.start, end.start))
        })
        .collect();
    ranges.sort_unstable();
    let mut stack: Vec<usize> = Vec::new();
    for (start, end) in ranges {
        while stack.last().is_some_and(|outer_end| start > *outer_end) {
            stack.pop();
        }
        if stack.last().is_some_and(|outer_end| end > *outer_end) {
            return invalid("annotation ranges cross; only properly nested ranges are allowed");
        }
        stack.push(end);
    }
    Ok(())
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    decoder: quick_xml::Decoder,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid annotation namespace: {error}")))?;
        let raw = qname(attribute.key.as_ref())?;
        let prefix = if raw == "xmlns" {
            Some(String::new())
        } else {
            raw.strip_prefix("xmlns:").map(str::to_string)
        };
        let Some(prefix) = prefix else { continue };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid_error(format!("invalid annotation namespace value: {error}")))?
            .into_owned();
        let previous = namespaces.insert(prefix.clone(), value);
        changes.push((prefix, previous));
    }
    Ok(changes)
}

fn restore_namespaces(
    changes: Vec<(String, Option<String>)>,
    namespaces: &mut BTreeMap<String, String>,
) {
    for (prefix, previous) in changes.into_iter().rev() {
        if let Some(previous) = previous {
            namespaces.insert(prefix, previous);
        } else {
            namespaces.remove(&prefix);
        }
    }
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    crate::elements::xml::namespaced_attribute(reader, element, namespace, local, "annotation")
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
    label: &str,
) -> Result<String> {
    optional_attribute(reader, element, namespace, local)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid_error(format!("{label} requires a non-empty office:name")))
}

fn optional_usize(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<usize>> {
    optional_attribute(reader, element, namespace, local)?
        .map(|value| {
            value
                .parse::<usize>()
                .map_err(|_error| invalid_error("invalid repeated table count"))
        })
        .transpose()
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Ns {
    Office,
    Text,
    Table,
    Draw,
    Other,
}

fn namespace(value: &ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => Ns::Office,
        ResolveResult::Bound(Namespace(value)) if *value == TEXT_NS => Ns::Text,
        ResolveResult::Bound(Namespace(value)) if *value == TABLE_NS => Ns::Table,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => Ns::Draw,
        _ => Ns::Other,
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("annotation XML position overflow"))
}

fn qname(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_error| invalid_error("invalid annotation qualified name"))
}
