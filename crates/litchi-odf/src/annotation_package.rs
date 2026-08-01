//! Lossless, package-safe ODF annotation discovery and mutation.

use crate::CellAnnotation;
use crate::core::OwnedPackage;
use crate::embedded_chart::rebuild_package;
use crate::ods::annotation::AnnotationBuilder;
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DC: &str = "http://purl.org/dc/elements/1.1/";
const META: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_ANNOTATIONS: usize = 65_536;
const MAX_EVENTS: usize = 1_000_000;

/// Rich annotation value shared by ODT, ODS, and ODP.
pub type OdfAnnotation = CellAnnotation;

/// A schema location to which an annotation is attached.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfAnnotationPosition {
    TextParagraph {
        paragraph_index: usize,
    },
    SpreadsheetCell {
        sheet_index: usize,
        row: usize,
        column: usize,
    },
    PresentationPage {
        page_index: usize,
    },
    PresentationShape {
        page_index: usize,
        shape_name: String,
    },
    AnnotationBody {
        annotation_index: usize,
    },
}

/// Start position and optional named-range end position.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfAnnotationAnchor {
    pub start: OdfAnnotationPosition,
    pub end: Option<OdfAnnotationPosition>,
}

impl OdfAnnotationAnchor {
    pub fn point(start: OdfAnnotationPosition) -> Self {
        Self { start, end: None }
    }

    pub fn range(start: OdfAnnotationPosition, end: OdfAnnotationPosition) -> Self {
        Self {
            start,
            end: Some(end),
        }
    }
}

/// One annotation in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfAnnotationInfo {
    pub index: usize,
    pub annotation: OdfAnnotation,
    pub anchor: OdfAnnotationAnchor,
}

/// Partial typed metadata update. `None` retains a value; `Some(None)` clears it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfAnnotationUpdate {
    pub creator: Option<Option<String>>,
    pub date: Option<Option<String>>,
    pub date_string: Option<Option<String>>,
    pub initials: Option<Option<String>>,
    pub display: Option<Option<bool>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum AnnotationHost {
    Text,
    Spreadsheet,
    Presentation,
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
    close_start: Option<usize>,
    qname: String,
}

#[derive(Clone)]
struct Site {
    position: OdfAnnotationPosition,
    span: Span,
}

struct Record {
    span: Span,
    parent_start: usize,
    annotation: Option<OdfAnnotation>,
    start_position: OdfAnnotationPosition,
    end: Option<(Span, OdfAnnotationPosition)>,
}

struct EndMarker {
    span: Span,
    name: String,
    position: OdfAnnotationPosition,
}

struct Scan {
    records: Vec<Record>,
    sites: Vec<Site>,
}

enum FrameKind {
    Table {
        sheet: usize,
        next_row: usize,
    },
    Row {
        sheet: usize,
        row: usize,
        next_column: usize,
    },
    Cell {
        site: usize,
    },
    Page {
        site: usize,
        page: usize,
    },
    Shape {
        site: usize,
    },
    Paragraph {
        site: Option<usize>,
    },
    Annotation {
        record: usize,
    },
    Other,
}

struct Frame {
    start: usize,
    kind: FrameKind,
    namespace_changes: Vec<(String, Option<String>)>,
}

struct ActiveBuilder {
    record: usize,
    builder: AnnotationBuilder,
}

pub(crate) fn annotations(content: &str, host: AnnotationHost) -> Result<Vec<OdfAnnotationInfo>> {
    let scan = scan(content, host)?;
    scan.records
        .into_iter()
        .enumerate()
        .map(|(index, mut record)| {
            Ok(OdfAnnotationInfo {
                index,
                annotation: record
                    .annotation
                    .take()
                    .ok_or_else(|| invalid_error("unterminated annotation"))?,
                anchor: OdfAnnotationAnchor {
                    start: record.start_position,
                    end: record.end.map(|(_, position)| position),
                },
            })
        })
        .collect()
}

pub(crate) fn find_annotation(
    content: &str,
    host: AnnotationHost,
    name: &str,
) -> Result<Option<OdfAnnotationInfo>> {
    if name.is_empty() {
        return invalid("annotation name cannot be empty");
    }
    Ok(annotations(content, host)?
        .into_iter()
        .find(|item| item.annotation.name() == Some(name)))
}

pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    anchor: &OdfAnnotationAnchor,
    annotation: &OdfAnnotation,
) -> Result<(Vec<u8>, usize)> {
    let (updated, index) = add_xml(content, host, anchor, annotation)?;
    rebuild(package, &updated).map(|bytes| (bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    annotation: &OdfAnnotation,
) -> Result<Vec<u8>> {
    rebuild(package, &replace_xml(content, host, index, annotation)?)
}

pub(crate) fn update(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    update: &OdfAnnotationUpdate,
) -> Result<Vec<u8>> {
    let items = annotations(content, host)?;
    let len = items.len();
    let mut info = items
        .into_iter()
        .nth(index)
        .ok_or_else(|| bounds(index, len))?;
    if let Some(value) = &update.creator {
        info.annotation.set_creator(value.as_deref());
    }
    if let Some(value) = &update.date {
        info.annotation.set_date(value.as_deref());
    }
    if let Some(value) = &update.date_string {
        info.annotation.set_date_string(value.as_deref());
    }
    if let Some(value) = &update.initials {
        info.annotation.set_initials(value.as_deref());
    }
    if let Some(value) = update.display {
        info.annotation.set_display(value);
    }
    replace(package, content, host, index, &info.annotation)
}

pub(crate) fn remove(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
) -> Result<Vec<u8>> {
    rebuild(package, &remove_xml(content, host, index)?)
}

pub(crate) fn reorder(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    from: usize,
    to: usize,
) -> Result<Vec<u8>> {
    rebuild(package, &reorder_xml(content, host, from, to)?)
}

fn add_xml(
    content: &str,
    host: AnnotationHost,
    anchor: &OdfAnnotationAnchor,
    annotation: &OdfAnnotation,
) -> Result<(String, usize)> {
    validate_anchor_host(host, anchor)?;
    annotation.validate()?;
    let scan = scan(content, host)?;
    validate_new_name(&scan, annotation.name())?;
    if anchor.end.is_some() && annotation.name().is_none() {
        return invalid("a ranged annotation requires a non-empty office:name");
    }
    let start_site = site_for(&scan, &anchor.start)?;
    let start_at = insertion_position(start_site);
    let fragment = serialize(annotation)?;
    let updated = if let Some(end_position) = &anchor.end {
        let end_site = site_for(&scan, end_position)?;
        let end_at = insertion_position(end_site);
        if start_at > end_at {
            return invalid("annotation range end precedes its start");
        }
        let marker = end_marker(annotation.name().expect("range name validated"));
        if start_site.span.start == end_site.span.start {
            insert_child(content, start_site, &format!("{fragment}{marker}"))?
        } else {
            apply_edits(
                content,
                vec![
                    child_edit(content, start_site, &fragment)?,
                    child_edit(content, end_site, &marker)?,
                ],
            )?
        }
    } else {
        insert_child(content, start_site, &fragment)?
    };
    let index = scan
        .records
        .iter()
        .filter(|record| record.span.start < start_at)
        .count();
    self::scan(&updated, host)?;
    Ok((updated, index))
}

fn replace_xml(
    content: &str,
    host: AnnotationHost,
    index: usize,
    annotation: &OdfAnnotation,
) -> Result<String> {
    annotation.validate()?;
    let scan = scan(content, host)?;
    let record = scan
        .records
        .get(index)
        .ok_or_else(|| bounds(index, scan.records.len()))?;
    if record.end.is_some()
        && record.annotation.as_ref().and_then(CellAnnotation::name) != annotation.name()
    {
        return invalid("replacing a ranged annotation cannot change its office:name");
    }
    if let Some(name) = annotation.name() {
        for (other_index, other) in scan.records.iter().enumerate() {
            if other_index != index
                && other.annotation.as_ref().and_then(CellAnnotation::name) == Some(name)
            {
                return invalid(format!("duplicate annotation name '{name}'"));
            }
        }
    }
    let updated = apply_edits(
        content,
        vec![Edit {
            start: record.span.start,
            end: record.span.end,
            replacement: serialize(annotation)?,
        }],
    )?;
    self::scan(&updated, host)?;
    Ok(updated)
}

fn remove_xml(content: &str, host: AnnotationHost, index: usize) -> Result<String> {
    let scan = scan(content, host)?;
    let record = scan
        .records
        .get(index)
        .ok_or_else(|| bounds(index, scan.records.len()))?;
    let mut edits = vec![Edit {
        start: record.span.start,
        end: record.span.end,
        replacement: String::new(),
    }];
    if let Some((end, _)) = &record.end {
        edits.push(Edit {
            start: end.start,
            end: end.end,
            replacement: String::new(),
        });
    }
    let updated = apply_edits(content, edits)?;
    self::scan(&updated, host)?;
    Ok(updated)
}

fn reorder_xml(content: &str, host: AnnotationHost, from: usize, to: usize) -> Result<String> {
    let scan = scan(content, host)?;
    let first = scan
        .records
        .get(from)
        .ok_or_else(|| bounds(from, scan.records.len()))?;
    let second = scan
        .records
        .get(to)
        .ok_or_else(|| bounds(to, scan.records.len()))?;
    if first.end.is_some() || second.end.is_some() {
        return invalid("ranged annotations cannot be reordered independently of their text");
    }
    if first.parent_start != second.parent_start {
        return invalid("annotations can only be reordered among XML siblings");
    }
    if first.span.start == second.span.start {
        return Ok(content.to_string());
    }
    let (left, right) = if first.span.start < second.span.start {
        (first, second)
    } else {
        (second, first)
    };
    let mut updated = String::with_capacity(content.len());
    updated.push_str(&content[..left.span.start]);
    updated.push_str(&content[right.span.start..right.span.end]);
    updated.push_str(&content[left.span.end..right.span.start]);
    updated.push_str(&content[left.span.start..left.span.end]);
    updated.push_str(&content[right.span.end..]);
    self::scan(&updated, host)?;
    Ok(updated)
}

fn scan(xml: &str, host: AnnotationHost) -> Result<Scan> {
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
                        builder: AnnotationBuilder::new(
                            &element,
                            reader.decoder(),
                            namespaces.clone(),
                        )?,
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
                    let builder =
                        AnnotationBuilder::new(&element, reader.decoder(), namespaces.clone())?;
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
                    let _ = structural_kind(
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
        let position = OdfAnnotationPosition::SpreadsheetCell {
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
            OdfAnnotationPosition::PresentationPage { page_index: page },
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
            OdfAnnotationPosition::PresentationShape {
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
            Some(OdfAnnotationPosition::AnnotationBody {
                annotation_index: annotation,
            })
        } else if host == AnnotationHost::Text {
            let index = *next_paragraph;
            *next_paragraph = next_paragraph
                .checked_add(1)
                .ok_or_else(|| invalid_error("paragraph index overflow"))?;
            Some(OdfAnnotationPosition::TextParagraph {
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
    position: OdfAnnotationPosition,
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
) -> Result<OdfAnnotationPosition> {
    if let Some(annotation) = current_annotation(frames) {
        return Ok(OdfAnnotationPosition::AnnotationBody {
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
        if let Some(name) = record.annotation.as_ref().and_then(CellAnnotation::name) {
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

fn validate_new_name(scan: &Scan, name: Option<&str>) -> Result<()> {
    let Some(name) = name else { return Ok(()) };
    if name.is_empty() {
        return invalid("annotation office:name cannot be empty");
    }
    if scan
        .records
        .iter()
        .any(|record| record.annotation.as_ref().and_then(CellAnnotation::name) == Some(name))
    {
        return invalid(format!("duplicate annotation name '{name}'"));
    }
    Ok(())
}

fn validate_anchor_host(host: AnnotationHost, anchor: &OdfAnnotationAnchor) -> Result<()> {
    let valid = |position: &OdfAnnotationPosition| {
        matches!(
            (host, position),
            (_, OdfAnnotationPosition::AnnotationBody { .. })
                | (
                    AnnotationHost::Text,
                    OdfAnnotationPosition::TextParagraph { .. }
                )
                | (
                    AnnotationHost::Spreadsheet,
                    OdfAnnotationPosition::SpreadsheetCell { .. }
                )
                | (
                    AnnotationHost::Presentation,
                    OdfAnnotationPosition::PresentationPage { .. }
                )
                | (
                    AnnotationHost::Presentation,
                    OdfAnnotationPosition::PresentationShape { .. }
                )
        )
    };
    if !valid(&anchor.start) || anchor.end.as_ref().is_some_and(|end| !valid(end)) {
        return invalid("annotation anchor does not belong to this document family");
    }
    if anchor.end.is_some()
        && host != AnnotationHost::Text
        && !matches!(anchor.start, OdfAnnotationPosition::AnnotationBody { .. })
    {
        return invalid("named annotation ranges must be inserted in text paragraph content");
    }
    Ok(())
}

fn site_for<'a>(scan: &'a Scan, position: &OdfAnnotationPosition) -> Result<&'a Site> {
    let mut matches = scan.sites.iter().filter(|site| &site.position == position);
    let site = matches
        .next()
        .ok_or_else(|| invalid_error(format!("annotation anchor {position:?} was not found")))?;
    if matches.next().is_some()
        && matches!(position, OdfAnnotationPosition::PresentationShape { .. })
    {
        return invalid("presentation shape annotation anchor is ambiguous");
    }
    Ok(site)
}

fn serialize(annotation: &OdfAnnotation) -> Result<String> {
    let mut annotation = annotation.clone();
    for (prefix, uri) in [
        ("office", OFFICE),
        ("text", TEXT),
        ("table", TABLE),
        ("draw", DRAW),
        ("svg", SVG),
        ("dc", DC),
        ("meta", META),
        ("xlink", XLINK),
        ("loext", LOEXT),
    ] {
        annotation.set_namespace(prefix, uri)?;
    }
    annotation.validate()?;
    let mut output = String::new();
    annotation.write_xml(&mut output);
    Ok(output)
}

fn end_marker(name: &str) -> String {
    format!(
        "<office:annotation-end xmlns:office=\"{OFFICE}\" office:name=\"{}\"/>",
        escape_xml(name)
    )
}

fn insertion_position(site: &Site) -> usize {
    site.span
        .close_start
        .unwrap_or(site.span.end.saturating_sub(2))
}

fn insert_child(xml: &str, site: &Site, fragment: &str) -> Result<String> {
    apply_edits(xml, vec![child_edit(xml, site, fragment)?])
}

fn child_edit(xml: &str, site: &Site, fragment: &str) -> Result<Edit> {
    if let Some(close) = site.span.close_start {
        Ok(Edit {
            start: close,
            end: close,
            replacement: fragment.to_string(),
        })
    } else {
        let raw = xml
            .get(site.span.start..site.span.end)
            .ok_or_else(|| invalid_error("invalid empty annotation anchor span"))?;
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| invalid_error("invalid empty annotation anchor"))?;
        Ok(Edit {
            start: site.span.start,
            end: site.span.end,
            replacement: format!("{}>{}</{}>", &raw[..slash], fragment, site.span.qname),
        })
    }
}

struct Edit {
    start: usize,
    end: usize,
    replacement: String,
}

fn apply_edits(xml: &str, mut edits: Vec<Edit>) -> Result<String> {
    edits.sort_by(|left, right| right.start.cmp(&left.start).then(right.end.cmp(&left.end)));
    let mut previous_start = xml.len();
    let mut output = xml.to_string();
    for edit in edits {
        if edit.start > edit.end || edit.end > xml.len() || edit.end > previous_start {
            return invalid("overlapping or invalid annotation XML edit");
        }
        output.replace_range(edit.start..edit.end, &edit.replacement);
        previous_start = edit.start;
    }
    Ok(output)
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
                .map_err(|_| invalid_error("invalid repeated table count"))
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
        .map_err(|_| invalid_error("annotation XML position overflow"))
}

fn qname(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| invalid_error("invalid annotation qualified name"))
}

fn rebuild(package: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    rebuild_package(
        package,
        content,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

fn bounds(index: usize, len: usize) -> Error {
    invalid_error(format!(
        "annotation index {index} is out of bounds for {len} entries"
    ))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

macro_rules! annotation_facade_methods {
    ($host:ident) => {
        /// Inspect annotations in document order without following body links.
        pub fn annotations(&self) -> litchi_core::Result<Vec<crate::OdfAnnotationInfo>> {
            crate::annotation_package::annotations(
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
            )
        }

        /// Find a uniquely named annotation.
        pub fn find_annotation(
            &self,
            name: &str,
        ) -> litchi_core::Result<Option<crate::OdfAnnotationInfo>> {
            crate::annotation_package::find_annotation(
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                name,
            )
        }

        /// Add a point or named-range annotation atomically.
        pub fn add_annotation(
            &mut self,
            anchor: &crate::OdfAnnotationAnchor,
            annotation: &crate::OdfAnnotation,
        ) -> litchi_core::Result<usize> {
            let (bytes, index) = crate::annotation_package::add(
                &self.package,
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                anchor,
                annotation,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(index)
        }

        /// Replace an annotation body and metadata while retaining its anchor.
        pub fn replace_annotation(
            &mut self,
            index: usize,
            annotation: &crate::OdfAnnotation,
        ) -> litchi_core::Result<()> {
            let bytes = crate::annotation_package::replace(
                &self.package,
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                index,
                annotation,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Apply a partial typed annotation metadata update.
        pub fn update_annotation(
            &mut self,
            index: usize,
            update: &crate::OdfAnnotationUpdate,
        ) -> litchi_core::Result<()> {
            let bytes = crate::annotation_package::update(
                &self.package,
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                index,
                update,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Remove an annotation and its paired end marker, if any.
        pub fn remove_annotation(&mut self, index: usize) -> litchi_core::Result<()> {
            let bytes = crate::annotation_package::remove(
                &self.package,
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                index,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Reorder point annotations that are direct XML siblings.
        pub fn reorder_annotation(&mut self, from: usize, to: usize) -> litchi_core::Result<()> {
            let bytes = crate::annotation_package::reorder(
                &self.package,
                self.content.xml_content(),
                crate::annotation_package::AnnotationHost::$host,
                from,
                to,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }
    };
}

pub(crate) use annotation_facade_methods;

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = "xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0'";

    #[test]
    fn scans_rich_nested_text_ranges_and_initials() {
        let xml = format!(
            "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='outer'><dc:creator>Ada</dc:creator><dc:date>2026-07-19T00:00:00Z</dc:date><meta:creator-initials>AL</meta:creator-initials><text:p>rich <office:annotation><text:p>nested</text:p></office:annotation></text:p></office:annotation>x<office:annotation-end office:name='outer'/></text:p></office:text></office:body></office:document>"
        );
        let items = annotations(&xml, AnnotationHost::Text).unwrap();
        assert_eq!(items.len(), 2);
        assert_eq!(items[0].annotation.initials().as_deref(), Some("AL"));
        assert!(items[0].anchor.end.is_some());
        assert_eq!(
            items[1].anchor.start,
            OdfAnnotationPosition::AnnotationBody {
                annotation_index: 0
            }
        );
    }

    #[test]
    fn scans_spreadsheet_cell_and_presentation_shape_anchors() {
        let ods = format!(
            "<office:document {NS}><office:body><office:spreadsheet><table:table><table:table-row><table:table-cell><office:annotation><text:p>cell</text:p></office:annotation></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>"
        );
        let item = annotations(&ods, AnnotationHost::Spreadsheet)
            .unwrap()
            .remove(0);
        assert_eq!(
            item.anchor.start,
            OdfAnnotationPosition::SpreadsheetCell {
                sheet_index: 0,
                row: 0,
                column: 0
            }
        );

        let odp = format!(
            "<office:document {NS}><office:body><office:presentation><draw:page><draw:frame draw:name='Title'><office:annotation><text:p>shape</text:p></office:annotation></draw:frame></draw:page></office:presentation></office:body></office:document>"
        );
        let item = annotations(&odp, AnnotationHost::Presentation)
            .unwrap()
            .remove(0);
        assert_eq!(
            item.anchor.start,
            OdfAnnotationPosition::PresentationShape {
                page_index: 0,
                shape_name: "Title".to_string()
            }
        );
    }

    #[test]
    fn rejects_crossing_and_duplicate_ranges() {
        let crossing = format!(
            "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='a'/><office:annotation office:name='b'/>x<office:annotation-end office:name='a'/><office:annotation-end office:name='b'/></text:p></office:text></office:body></office:document>"
        );
        assert!(annotations(&crossing, AnnotationHost::Text).is_err());
        let duplicate = format!(
            "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='a'/><office:annotation office:name='a'/></text:p></office:text></office:body></office:document>"
        );
        assert!(annotations(&duplicate, AnnotationHost::Text).is_err());
    }

    #[test]
    fn generated_text_mutations_preserve_unknown_xml() {
        let xml = format!(
            "<office:document {NS} xmlns:v='urn:vendor'><office:body><office:text><text:p><v:keep key='1'/>text</text:p></office:text></office:body></office:document>"
        );
        let mut annotation = OdfAnnotation::new("review");
        annotation.set_name(Some("r1"));
        annotation.set_creator(Some("Ada"));
        let anchor = OdfAnnotationAnchor::range(
            OdfAnnotationPosition::TextParagraph { paragraph_index: 0 },
            OdfAnnotationPosition::TextParagraph { paragraph_index: 0 },
        );
        let (updated, index) = add_xml(&xml, AnnotationHost::Text, &anchor, &annotation).unwrap();
        assert_eq!(index, 0);
        assert!(updated.contains("<v:keep key='1'/>") && updated.contains("office:annotation-end"));
        let replaced = replace_xml(&updated, AnnotationHost::Text, 0, &annotation).unwrap();
        let removed = remove_xml(&replaced, AnnotationHost::Text, 0).unwrap();
        assert!(removed.contains("<v:keep key='1'/>") && !removed.contains("office:annotation"));
    }
}
