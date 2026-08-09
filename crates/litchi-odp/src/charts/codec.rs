//! Bounded ODP chart discovery and XML span codecs.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::items_after_statements,
    reason = "the codec keeps parsing helpers beside the state machines that use them; reordering them would make the bounded scans harder to audit"
)]

use super::model::{Chart, Limits, Location, Page, Part, Storage};
use crate::core::OwnedPackage;
use litchi_core::{Error, Result};
use litchi_odf_common::chart::authoring::{
    CachedCell, CachedRow, CachedTable, Definition, SeriesSpec, serialize_content,
    serialize_series_fragment,
};
use litchi_odf_common::chart::read;
use litchi_odf_common::constants::{ODF_CHART, ODF_CHART_TEMPLATE};
use litchi_odf_common::drawing::Part as DrawingPart;
use litchi_odf_common::embedded::{Kind, Root, Source, scan_package};
use litchi_odf_common::package::splice;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::Arc;
use xml_minifier::audit;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_HOST_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 512;

/// One direct presentation page end tag insertion point.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PageSpan {
    pub(crate) index: usize,
    pub(crate) name: Option<String>,
    pub(crate) end: usize,
}

/// Discover chart occurrences in the current ODP `content.xml` snapshot.
pub(crate) fn inventory(source: &OwnedPackage, limits: Limits) -> Result<Vec<Chart>> {
    let content = content_xml(source)?;
    if content.len() > MAX_HOST_XML_BYTES {
        return invalid("ODP content.xml exceeds the chart host byte limit");
    }
    let package = source.package()?;
    let objects = scan_package(&content, None, &package)?;
    let spans = locate_objects(&content)?;
    let mut content_objects = 0usize;
    let mut total_bytes = 0usize;
    let mut charts = Vec::new();

    for object in objects {
        if object.part != DrawingPart::Content {
            continue;
        }
        let span = spans
            .get(content_objects)
            .ok_or_else(|| invalid_error("ODP chart object and XML span scanners disagree"))?;
        content_objects = content_objects
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP chart object count overflow"))?;
        if !is_chart_object(&object) {
            continue;
        }
        if charts.len() >= limits.max_charts() {
            return invalid(format!(
                "ODP embedded chart count exceeds {}",
                limits.max_charts()
            ));
        }

        let (part, storage, content_path) = match object.source {
            Source::PackageSubdocument {
                content_path,
                manifest_media_type,
                ..
            } => {
                if !is_chart_media_type(manifest_media_type.as_deref()) {
                    continue;
                }
                let bytes = source.get_file(&content_path)?;
                let xml = String::from_utf8(bytes)
                    .map_err(|_err| invalid_error("embedded ODP chart content.xml is not UTF-8"))?;
                (
                    Part::from_xml_with_limit(xml, limits.max_part_bytes())?,
                    Storage::PackageSubdocument,
                    Some(content_path),
                )
            },
            Source::InlineXml {
                root: Root::OpenDocument,
                xml,
                ..
            } => {
                let Some(media_type) = inline_mimetype(&xml)? else {
                    continue;
                };
                if !is_chart_media_type(Some(media_type.as_str())) {
                    continue;
                }
                (
                    Part::from_inline_with_limit(xml, limits.max_part_bytes())?,
                    Storage::InlineXml,
                    None,
                )
            },
            Source::InlineXml { .. }
            | Source::InlineBinary { .. }
            | Source::PackageFile { .. }
            | Source::MissingPackagePart { .. }
            | Source::Linked { .. }
            | Source::Missing
            | _ => continue,
        };

        total_bytes = total_bytes
            .checked_add(part.xml().len())
            .ok_or_else(|| invalid_error("ODP chart byte count overflow"))?;
        if total_bytes > limits.max_total_bytes() {
            return invalid(format!(
                "ODP embedded chart content exceeds {} bytes",
                limits.max_total_bytes()
            ));
        }
        charts.push(Chart {
            frame: object.frame,
            storage,
            part,
            location: Location::Existing {
                object_start: span.start,
                object_end: span.end,
                payload: span.payload,
                content_path,
            },
        });
    }

    if content_objects != spans.len() {
        return invalid("ODP chart object scanner found an inconsistent object count");
    }
    Ok(charts)
}

impl Part {
    /// Parse an authored compact standalone chart content part.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn from_xml(source_xml: impl Into<String>) -> Result<Self> {
        let xml = source_xml.into();
        verify_authored_xml(&xml, Limits::default().max_part_bytes())?;
        Self::from_xml_with_limit(xml, Limits::default().max_part_bytes())
    }

    /// Serialize a checked common ODF chart definition into an ODP part.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn from_definition(definition: &Definition) -> Result<Self> {
        Self::from_xml(serialize_content(definition)?)
    }

    /// Return a checked part with one typed series appended to its plot area.
    ///
    /// # Errors
    ///
    /// Returns an error when the chart has no unique plot area or the series is invalid.
    pub fn with_series_added(&self, series: &SeriesSpec) -> Result<Self> {
        let fragment = serialize_series_fragment(series)?;
        let plot_areas = locate_element_spans(self.xml(), CHART_NS, b"plot-area")?;
        let plot_area = unique_span(&plot_areas, "ODP chart plot area")?;
        let closing = closing_tag_start(self.xml(), plot_area)?;
        Self::from_xml(splice(self.xml(), closing, closing, &fragment)?)
    }

    /// Return a checked part with one physical series replaced by index.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index or invalid series.
    pub fn with_series_replaced(&self, index: usize, series: &SeriesSpec) -> Result<Self> {
        let fragment = serialize_series_fragment(series)?;
        let spans = locate_element_spans(self.xml(), CHART_NS, b"series")?;
        let span = spans
            .get(index)
            .copied()
            .ok_or_else(|| invalid_error("ODP chart series index is out of bounds"))?;
        Self::from_xml(splice(self.xml(), span.start, span.end, &fragment)?)
    }

    /// Return a checked part with one physical series removed by index.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index or malformed chart XML.
    pub fn with_series_removed(&self, index: usize) -> Result<Self> {
        let spans = locate_element_spans(self.xml(), CHART_NS, b"series")?;
        let span = spans
            .get(index)
            .copied()
            .ok_or_else(|| invalid_error("ODP chart series index is out of bounds"))?;
        Self::from_xml(splice(self.xml(), span.start, span.end, "")?)
    }

    /// Return a checked part with one physical cached-table cell replaced.
    ///
    /// Row indexing includes cached header rows. Repeated row/cell runs remain
    /// physical XML entries and are not expanded implicitly.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing cached table, out-of-range coordinate, or invalid cell.
    pub fn with_cached_cell_replaced(
        &self,
        row: usize,
        column: usize,
        cell: &CachedCell,
    ) -> Result<Self> {
        let rows = locate_element_spans(self.xml(), TABLE_NS, b"table-row")?;
        let row_span = rows
            .get(row)
            .copied()
            .ok_or_else(|| invalid_error("ODP chart cached-table row is out of bounds"))?;
        let cells = locate_element_spans(self.xml(), TABLE_NS, b"table-cell")?;
        let cell_span = cells
            .iter()
            .filter(|span| span.start >= row_span.start && span.end <= row_span.end)
            .nth(column)
            .copied()
            .ok_or_else(|| invalid_error("ODP chart cached-table column is out of bounds"))?;
        let fragment = serialize_cached_cell_fragment(cell)?;
        Self::from_xml(splice(
            self.xml(),
            cell_span.start,
            cell_span.end,
            &fragment,
        )?)
    }

    /// Parse an inline `office:document` chart payload.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn from_inline_xml(source_xml: impl Into<String>) -> Result<Self> {
        let xml = source_xml.into();
        verify_authored_xml(&xml, Limits::default().max_part_bytes())?;
        Self::from_inline_with_limit(xml, Limits::default().max_part_bytes())
    }

    pub(crate) fn from_xml_with_limit(xml: String, max_bytes: usize) -> Result<Self> {
        if xml.is_empty() || xml.len() > max_bytes {
            return invalid("ODP chart content is empty or exceeds its byte limit");
        }
        let chart = read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml),
            chart: Arc::new(chart),
        })
    }

    #[allow(
        clippy::needless_pass_by_value,
        reason = "the owned XML is retained by the returned source-backed chart part without a second allocation"
    )]
    pub(crate) fn from_inline_with_limit(xml: String, max_bytes: usize) -> Result<Self> {
        let Some(media_type) = inline_mimetype(&xml)? else {
            return invalid("inline ODP chart has no office:mimetype");
        };
        if !is_chart_media_type(Some(media_type.as_str())) {
            return invalid("inline ODP object is not an OpenDocument chart");
        }
        let content = rename_document_root(&xml, "document", "document-content", None)?;
        Self::from_xml_with_limit(content, max_bytes)
    }
}

const CHART_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";

#[derive(Clone, Copy)]
struct ElementSpan {
    start: usize,
    end: usize,
}

fn locate_element_spans(xml: &str, namespace: &[u8], local: &[u8]) -> Result<Vec<ElementSpan>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active = Vec::<(usize, usize, Vec<u8>)>::new();
    let mut spans = Vec::new();
    loop {
        let start = position(&reader, utf8_bom_len(xml))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODP chart XML: {error}")))?;
        let matches =
            matches!(resolved, ResolveResult::Bound(Namespace(uri)) if *uri == *namespace);
        let end = position(&reader, utf8_bom_len(xml))?;
        match event {
            Event::Start(element) => {
                if matches && element.local_name().as_ref() == local {
                    active.push((depth, start, element.name().as_ref().to_vec()));
                }
                depth = checked_depth(depth)?;
            },
            Event::Empty(element) => {
                if matches && element.local_name().as_ref() == local {
                    spans.push(ElementSpan { start, end });
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODP chart XML depth underflow"))?;
                if active.last().is_some_and(|(active_depth, _, qualified)| {
                    *active_depth == depth && qualified.as_slice() == element.name().as_ref()
                }) {
                    let (_, active_start, _) = active
                        .pop()
                        .ok_or_else(|| invalid_error("ODP chart span state disappeared"))?;
                    spans.push(ElementSpan {
                        start: active_start,
                        end,
                    });
                }
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP chart XML"),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !active.is_empty() || depth != 0 {
        return invalid("unterminated ODP chart XML");
    }
    spans.sort_unstable_by_key(|span| span.start);
    Ok(spans)
}

fn unique_span(spans: &[ElementSpan], kind: &str) -> Result<ElementSpan> {
    match spans {
        [span] => Ok(*span),
        [] => invalid(format!("{kind} was not found")),
        _ => invalid(format!("{kind} is ambiguous")),
    }
}

fn closing_tag_start(xml: &str, span: ElementSpan) -> Result<usize> {
    xml.get(span.start..span.end)
        .and_then(|fragment| fragment.rfind("</"))
        .and_then(|relative| span.start.checked_add(relative))
        .ok_or_else(|| invalid_error("ODP chart element has no closing tag"))
}

fn serialize_cached_cell_fragment(cell: &CachedCell) -> Result<String> {
    let mut definition = Definition::new(litchi_odf_common::chart::ChartClass::line());
    let mut table = CachedTable::new("cell", 1);
    table.rows.push(CachedRow::new(vec![cell.clone()]));
    definition.cached_table = Some(table);
    let xml = serialize_content(&definition)?;
    let cells = locate_element_spans(&xml, TABLE_NS, b"table-cell")?;
    let span = unique_span(&cells, "serialized ODP chart cached cell")?;
    xml.get(span.start..span.end)
        .map(str::to_string)
        .ok_or_else(|| invalid_error("serialized ODP cached-cell span is invalid"))
}

fn verify_authored_xml(xml: &str, max_bytes: usize) -> Result<()> {
    if xml.contains("> <") {
        return Err(Error::Unsupported(
            "authored ODP chart XML contains inter-element spacing".to_string(),
        ));
    }
    let limits = audit::Limits::new(
        max_bytes, MAX_DEPTH, 1_000_000, 250_000, max_bytes, max_bytes,
    )
    .map_err(|source| invalid_error(format!("invalid ODP chart XML audit limits: {source}")))?;
    let _report = audit::verify(xml.as_bytes(), limits).map_err(|source| match source {
        audit::Error::NotCompact(_) => {
            Error::Unsupported(format!("authored ODP chart XML is not compact: {source}"))
        },
        audit::Error::Limit { .. }
        | audit::Error::Encoding { .. }
        | audit::Error::Malformed { .. }
        | audit::Error::Doctype { .. }
        | audit::Error::Allocation
        | _ => Error::InvalidFormat(format!("authored ODP chart XML failed audit: {source}")),
    })?;
    Ok(())
}

pub(crate) fn content_xml(source: &OwnedPackage) -> Result<String> {
    let bytes = source.get_file("content.xml")?;
    String::from_utf8(bytes).map_err(|_err| invalid_error("ODP content.xml is not UTF-8"))
}

fn is_chart_object(object: &litchi_odf_common::embedded::Object) -> bool {
    if object.part != DrawingPart::Content || object.kind != Kind::Object {
        return false;
    }
    if matches!(
        object.source,
        Source::InlineXml {
            root: Root::OpenDocument,
            ..
        }
    ) {
        return true;
    }
    object.class_id.is_none()
        && object.code.is_none()
        && object.archive.is_none()
        && object.may_script.is_none()
        && object.applet_name.is_none()
        && object.mime_type.is_none()
        && object.parameters.is_empty()
        && object
            .link_type
            .as_deref()
            .is_none_or(|value| value == "simple")
        && object.show.as_deref().is_none_or(|value| value == "embed")
        && object
            .actuate
            .as_deref()
            .is_none_or(|value| value == "onLoad")
}

fn is_chart_media_type(value: Option<&str>) -> bool {
    matches!(value, Some(ODF_CHART | ODF_CHART_TEMPLATE))
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct ObjectSpan {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) payload: Option<(usize, usize)>,
}

/// Locate every drawing embedded-object element in source order.
pub(crate) fn locate_objects(xml: &str) -> Result<Vec<ObjectSpan>> {
    if xml.len() > MAX_HOST_XML_BYTES {
        return invalid("ODP drawing XML exceeds the chart host byte limit");
    }
    struct Active {
        depth: usize,
        start: usize,
        payload: Option<(usize, usize, usize)>,
    }

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active = None;
    let mut spans = Vec::new();
    let position_offset = utf8_bom_len(xml);

    loop {
        let start = position(&reader, position_offset)?;
        let token = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid_error(format!("invalid ODP drawing XML: {error}")))?;
            let kind = namespace_kind(&namespace);
            match event {
                Event::Start(element) => Token::Start {
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::Empty(element) => Token::Empty {
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::End(element) => Token::End {
                    kind,
                    object: kind == NamespaceKind::Draw
                        && is_object_name(element.local_name().as_ref()),
                    document: kind == NamespaceKind::Office
                        && element.local_name().as_ref() == b"document",
                },
                Event::DocType(_) => return invalid("DTDs are not allowed in ODP chart hosts"),
                Event::Eof => Token::Eof,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => Token::Other,
            }
        };
        let end = position(&reader, position_offset)?;
        match token {
            Token::Start { object, document } => {
                if object {
                    if active.is_some() {
                        return invalid("nested ODP embedded objects are not supported");
                    }
                    active = Some(Active {
                        depth,
                        start,
                        payload: None,
                    });
                } else if let Some(current) = active.as_mut()
                    && depth == current.depth + 1
                    && document
                {
                    current.payload = Some((depth, start, 0));
                }
                depth = checked_depth(depth)?;
            },
            Token::Empty { object, document } => {
                if object {
                    if active.is_some() {
                        return invalid("nested ODP embedded objects are not supported");
                    }
                    spans.push(ObjectSpan {
                        start,
                        end,
                        payload: None,
                    });
                } else if active
                    .as_ref()
                    .is_some_and(|current| depth == current.depth + 1 && document)
                {
                    return invalid("inline ODP chart document cannot be empty");
                }
            },
            Token::End {
                kind,
                object,
                document,
            } => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODP drawing XML depth underflow"))?;
                if let Some(current) = active.as_mut()
                    && current
                        .payload
                        .is_some_and(|(payload_depth, _, _)| payload_depth == depth)
                    && document
                    && let Some((payload_depth, payload_start, _)) = current.payload
                {
                    current.payload = Some((payload_depth, payload_start, end));
                }
                if active
                    .as_ref()
                    .is_some_and(|current| current.depth == depth)
                    && kind == NamespaceKind::Draw
                    && object
                {
                    let current = active
                        .take()
                        .ok_or_else(|| invalid_error("ODP embedded object state disappeared"))?;
                    spans.push(ObjectSpan {
                        start: current.start,
                        end,
                        payload: current.payload.and_then(|(_, payload_start, payload_end)| {
                            (payload_end != 0).then_some((payload_start, payload_end))
                        }),
                    });
                }
            },
            Token::Eof => break,
            Token::Other => {},
        }
        buffer.clear();
    }
    if active.is_some() || depth != 0 {
        return invalid("unterminated ODP embedded-object XML");
    }
    Ok(spans)
}

/// Locate insertion points for all non-empty `draw:page` elements.
pub(crate) fn locate_pages(xml: &str) -> Result<Vec<PageSpan>> {
    if xml.len() > MAX_HOST_XML_BYTES {
        return invalid("ODP drawing XML exceeds the chart host byte limit");
    }
    struct Active {
        depth: usize,
        index: usize,
        name: Option<String>,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active = None;
    let mut pages = Vec::new();
    let position_offset = utf8_bom_len(xml);
    loop {
        let start = position(&reader, position_offset)?;
        let token = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid_error(format!("invalid ODP page XML: {error}")))?;
            let kind = namespace_kind(&namespace);
            match event {
                Event::Start(element)
                    if kind == NamespaceKind::Draw && element.local_name().as_ref() == b"page" =>
                {
                    TokenPage::Start(read_attribute(&reader, &element, DRAW_NS, b"name")?)
                },
                Event::Start(_) => TokenPage::StartOther,
                Event::Empty(element)
                    if kind == NamespaceKind::Draw && element.local_name().as_ref() == b"page" =>
                {
                    TokenPage::Empty
                },
                Event::Empty(_) => TokenPage::EmptyOther,
                Event::End(element)
                    if kind == NamespaceKind::Draw && element.local_name().as_ref() == b"page" =>
                {
                    TokenPage::End
                },
                Event::End(_) => TokenPage::EndOther,
                Event::DocType(_) => return invalid("DTDs are not allowed in ODP chart hosts"),
                Event::Eof => TokenPage::Eof,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => TokenPage::Other,
            }
        };
        match token {
            TokenPage::Start(name) => {
                if active.is_some() {
                    return invalid("nested ODP draw:page elements are not supported");
                }
                active = Some(Active {
                    depth,
                    index: pages.len(),
                    name,
                });
                depth = checked_depth(depth)?;
            },
            TokenPage::StartOther => {
                depth = checked_depth(depth)?;
            },
            TokenPage::Empty => return invalid("ODP chart host page cannot be empty"),
            TokenPage::EmptyOther | TokenPage::Other => {},
            TokenPage::End => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODP page XML depth underflow"))?;
                if active.as_ref().is_some_and(|page| page.depth == depth) {
                    let page = active
                        .take()
                        .ok_or_else(|| invalid_error("ODP page state disappeared"))?;
                    pages.push(PageSpan {
                        index: page.index,
                        name: page.name,
                        end: start,
                    });
                }
            },
            TokenPage::EndOther => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODP page XML depth underflow"))?;
            },
            TokenPage::Eof => break,
        }
        buffer.clear();
    }
    if active.is_some() || depth != 0 {
        return invalid("unterminated ODP page XML");
    }
    Ok(pages)
}

pub(crate) fn page_index(xml: &str, page: Page<'_>) -> Result<usize> {
    let pages = locate_pages(xml)?;
    match page {
        Page::Index(index) => pages
            .get(index)
            .map(|candidate| candidate.index)
            .ok_or_else(|| invalid_error("ODP chart page selector is out of bounds")),
        Page::Name(name) => {
            let mut found = None;
            for value in &pages {
                if value.name.as_deref() == Some(name) {
                    if found.is_some() {
                        return invalid("ODP chart page selector is ambiguous");
                    }
                    found = Some(value.index);
                }
            }
            found.ok_or_else(|| invalid_error("ODP chart page selector did not match"))
        },
    }
}

fn inline_mimetype(xml: &str) -> Result<Option<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid inline ODP chart XML: {error}")))?;
        match event {
            Event::Start(element)
                if namespace_kind(&namespace) == NamespaceKind::Office
                    && element.local_name().as_ref() == b"document" =>
            {
                return read_attribute(&reader, &element, OFFICE_NS, b"mimetype");
            },
            Event::Empty(_) => return invalid("inline ODP chart document root cannot be empty"),
            Event::DocType(_) => return invalid("DTDs are not allowed in inline ODP charts"),
            Event::Text(value) if value.iter().all(u8::is_ascii_whitespace) => {},
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("inline ODP chart has no office:document root"),
            Event::Start(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {
                return invalid("inline ODP chart has invalid content before its root");
            },
        }
        buffer.clear();
    }
}

pub(crate) fn content_inline(xml: &str) -> Result<String> {
    let mimetype = (!root_has_office_mimetype(xml)?).then_some(("office:mimetype", ODF_CHART));
    rename_document_root(xml, "document-content", "document", mimetype)
}

fn root_has_office_mimetype(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODP chart XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let value = read_attribute(&reader, &element, OFFICE_NS, b"mimetype")?;
                return Ok(value.is_some());
            },
            Event::Empty(_) => return invalid("ODP chart document root cannot be empty"),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(value) if value.iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return invalid("ODP chart XML has no document root"),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return invalid("ODP chart XML has invalid content before its root");
            },
        }
        buffer.clear();
    }
}

fn rename_document_root(
    xml: &str,
    expected_local: &str,
    replacement_local: &str,
    added_attribute: Option<(&str, &str)>,
) -> Result<String> {
    if xml.len() > MAX_HOST_XML_BYTES {
        return invalid("ODP chart XML exceeds the host byte limit");
    }
    let mut root_start = xml
        .find('<')
        .ok_or_else(|| invalid_error("ODP chart XML has no root"))?;
    if xml[root_start..].starts_with("<?xml") {
        let declaration_end = xml[root_start..]
            .find("?>")
            .ok_or_else(|| invalid_error("unterminated ODP chart XML declaration"))?
            + root_start
            + 2;
        root_start = xml[declaration_end..]
            .find('<')
            .map(|offset| declaration_end + offset)
            .ok_or_else(|| invalid_error("ODP chart XML has no document root"))?;
    }
    let name_end = xml[root_start + 1..]
        .find(|character: char| character.is_whitespace() || character == '>' || character == '/')
        .map(|offset| root_start + 1 + offset)
        .ok_or_else(|| invalid_error("invalid ODP chart root start tag"))?;
    let qname = &xml[root_start + 1..name_end];
    let (prefix, local) = qname.rsplit_once(':').unwrap_or(("", qname));
    if prefix.is_empty() || local != expected_local {
        return invalid(format!("expected office:{expected_local} ODP chart root"));
    }
    let close_start = xml
        .rfind("</")
        .ok_or_else(|| invalid_error("ODP chart root is not closed"))?;
    let close_name_end = xml[close_start + 2..]
        .find('>')
        .map(|offset| close_start + 2 + offset)
        .ok_or_else(|| invalid_error("invalid ODP chart root closing tag"))?;
    if xml[close_start + 2..close_name_end].trim() != qname {
        return invalid("ODP chart root start/end names do not match");
    }
    let replacement = format!("{prefix}:{replacement_local}");
    let mut output = String::with_capacity(xml.len() + 96);
    output.push_str(&xml[..=root_start]);
    output.push_str(&replacement);
    if let Some((name, value)) = added_attribute {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(value);
        output.push('"');
    }
    output.push_str(&xml[name_end..close_start + 2]);
    output.push_str(&replacement);
    output.push_str(&xml[close_name_end..]);
    Ok(output)
}

fn read_attribute(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|error| invalid_error(format!("invalid ODP XML attribute: {error}")))?;
        let (attribute_namespace, attribute_local) =
            reader.resolver().resolve_attribute(attribute.key);
        if matches!(attribute_namespace, ResolveResult::Bound(Namespace(uri)) if *uri == *namespace)
            && attribute_local.as_ref() == local
        {
            if value.is_some() {
                return invalid("duplicate ODP chart attribute");
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        invalid_error(format!("invalid ODP XML attribute value: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn is_object_name(local: &[u8]) -> bool {
    matches!(
        local,
        b"object" | b"object-ole" | b"applet" | b"plugin" | b"floating-frame"
    )
}

fn position(reader: &NsReader<&[u8]>, offset: usize) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid_error("ODP XML position exceeds platform limits"))?
        .checked_add(offset)
        .ok_or_else(|| invalid_error("ODP XML position offset overflow"))
}

fn utf8_bom_len(xml: &str) -> usize {
    if xml.as_bytes().starts_with(b"\xEF\xBB\xBF") {
        3
    } else {
        0
    }
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid_error("ODP XML depth overflow"))?;
    if next > MAX_DEPTH {
        return invalid("ODP chart host nesting exceeds its depth limit");
    }
    Ok(next)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Other,
}

enum Token {
    Start {
        object: bool,
        document: bool,
    },
    Empty {
        object: bool,
        document: bool,
    },
    End {
        kind: NamespaceKind,
        object: bool,
        document: bool,
    },
    Eof,
    Other,
}

enum TokenPage {
    Start(Option<String>),
    StartOther,
    Empty,
    EmptyOther,
    End,
    EndOther,
    Eof,
    Other,
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
