//! Provenance-spliced advanced spreadsheet authoring used by the unified document root.

#![deny(
    clippy::cast_possible_truncation,
    clippy::expect_used,
    clippy::panic,
    clippy::unreachable,
    clippy::unwrap_used
)]

use std::{fmt::Write as _, ops::Range};

use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::core::{
    AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication, rebuild_package_with_xml_splices,
};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::{Cell, CellValue, Row, Sheet, package::Package};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const CALCEXT: &str = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const CONTENT_PATH: &str = "content.xml";
const MAX_ELEMENTS: usize = 1_048_576;

/// One rich-text inline in a spreadsheet cell paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum RichRun {
    /// Escaped ordinary character data.
    Text(String),
    /// Character data with an optional text-style reference.
    Span { text: String, style_name: String },
    /// An inert hyperlink; the target is never followed.
    Link { text: String, href: String },
    /// A positive run of ODF spaces.
    Space(usize),
    /// One tab stop.
    Tab,
    /// One line break.
    LineBreak,
}

/// Checked rich cell text represented as paragraphs and inline runs.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RichText {
    paragraphs: Vec<Vec<RichRun>>,
}

impl RichText {
    /// Construct rich text. At least one paragraph is required.
    ///
    /// # Errors
    ///
    /// Returns an error for empty input, unsafe links, empty style names, zero spaces, or bounds.
    pub fn new(paragraphs: Vec<Vec<RichRun>>) -> Result<Self> {
        if paragraphs.is_empty() || paragraphs.len() > 65_536 {
            return invalid("ODS rich cell text requires a bounded non-empty paragraph list");
        }
        let value = Self { paragraphs };
        let _markup = value.markup()?;
        Ok(value)
    }

    /// Borrow paragraphs in document order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Vec<RichRun>] {
        &self.paragraphs
    }

    fn plain_text(&self) -> String {
        let mut output = String::new();
        for (paragraph_index, paragraph) in self.paragraphs.iter().enumerate() {
            if paragraph_index != 0 {
                output.push('\n');
            }
            for run in paragraph {
                match run {
                    RichRun::Text(text)
                    | RichRun::Span { text, .. }
                    | RichRun::Link { text, .. } => {
                        output.push_str(text);
                    },
                    RichRun::Space(count) => output.extend(std::iter::repeat_n(' ', *count)),
                    RichRun::Tab => output.push('\t'),
                    RichRun::LineBreak => output.push('\n'),
                }
            }
        }
        output
    }

    fn markup(&self) -> Result<String> {
        let mut output = String::new();
        for paragraph in &self.paragraphs {
            output.push_str("<text:p>");
            for run in paragraph {
                match run {
                    RichRun::Text(text) => output.push_str(&escape_xml(text)),
                    RichRun::Span { text, style_name } => {
                        validate_token(style_name, "rich-text style name")?;
                        output.push_str("<text:span text:style-name=\"");
                        output.push_str(&escape_xml(style_name));
                        output.push_str("\">");
                        output.push_str(&escape_xml(text));
                        output.push_str("</text:span>");
                    },
                    RichRun::Link { text, href } => {
                        validate_href(href)?;
                        output.push_str("<text:a xlink:href=\"");
                        output.push_str(&escape_xml(href));
                        output.push_str("\" xlink:type=\"simple\">");
                        output.push_str(&escape_xml(text));
                        output.push_str("</text:a>");
                    },
                    RichRun::Space(count) => {
                        if *count == 0 || *count > 1_048_576 {
                            return invalid("ODS rich-text space count is invalid");
                        }
                        output.push_str("<text:s");
                        if *count > 1 {
                            output.push_str(" text:c=\"");
                            output.push_str(&count.to_string());
                            output.push('"');
                        }
                        output.push_str("/>");
                    },
                    RichRun::Tab => output.push_str("<text:tab/>"),
                    RichRun::LineBreak => output.push_str("<text:line-break/>"),
                }
            }
            output.push_str("</text:p>");
        }
        if output.len() > 16 * 1024 * 1024 {
            return invalid("ODS rich cell text exceeds the byte limit");
        }
        Ok(output)
    }
}

/// A compact automatic table-cell style owned by `content.xml`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellStyle {
    /// Exact style name referenced by cells.
    pub name: String,
    /// Optional `#RRGGBB` background.
    pub background: Option<String>,
    /// Optional `#RRGGBB` text color.
    pub color: Option<String>,
    /// Optional bold text setting.
    pub bold: Option<bool>,
}

impl CellStyle {
    fn markup(&self) -> Result<String> {
        validate_token(&self.name, "cell style name")?;
        validate_color(self.background.as_deref())?;
        validate_color(self.color.as_deref())?;
        let mut output = format!(
            "<style:style xmlns:style=\"{STYLE}\" xmlns:fo=\"{FO}\" style:name=\"{}\" style:family=\"table-cell\">",
            escape_xml(&self.name)
        );
        if let Some(background) = &self.background {
            output.push_str("<style:table-cell-properties fo:background-color=\"");
            output.push_str(background);
            output.push_str("\"/>");
        }
        if self.color.is_some() || self.bold.is_some() {
            output.push_str("<style:text-properties");
            if let Some(color) = &self.color {
                output.push_str(" fo:color=\"");
                output.push_str(color);
                output.push('"');
            }
            if let Some(bold) = self.bold {
                output.push_str(" fo:font-weight=\"");
                output.push_str(if bold { "bold" } else { "normal" });
                output.push('"');
            }
            output.push_str("/>");
        }
        output.push_str("</style:style>");
        Ok(output)
    }
}

/// One inert form button inside the spreadsheet form container.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl {
    /// Stable form identifier.
    pub id: String,
    /// Producer-visible label.
    pub label: String,
}

/// One sheet drawing frame backed by a package resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Drawing {
    /// Exact drawing name.
    pub name: String,
    /// Package-relative dependency path.
    pub resource_path: String,
}

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    start: usize,
    tag_end: usize,
    close_start: usize,
    end: usize,
    parent: Option<usize>,
}

struct CellLocation<'a> {
    cell: &'a Cell,
    range: Range<usize>,
    tag: Range<usize>,
    empty: bool,
}

pub(crate) fn set_rich_text(
    source: &[u8],
    sheet: &str,
    row: usize,
    column: usize,
    rich: &RichText,
    max_output: usize,
) -> Result<Vec<u8>> {
    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    let sheet_model = spreadsheet
        .sheet(sheet)
        .ok_or_else(|| invalid_error(format!("ODS sheet '{sheet}' was not found")))?;
    let location = selected_cell(source, sheet_model, row, column)?;
    let mut cell = location.cell.clone();
    let plain = rich.plain_text();
    cell.value = CellValue::Text(plain.clone());
    cell.text = plain;
    let markup = rich.markup()?;
    let replacement = crate::worksheet::codec::write_cell_fragment(&cell, Some(&markup))?;
    splice_content(source, location.range, replacement.into_bytes(), max_output)
}

pub(crate) fn set_cell_formula(
    source: &[u8],
    sheet: &str,
    row: usize,
    column: usize,
    formula: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    let sheet_model = spreadsheet
        .sheet(sheet)
        .ok_or_else(|| invalid_error(format!("ODS sheet '{sheet}' was not found")))?;
    let location = selected_cell(source, sheet_model, row, column)?;
    let mut cell = location.cell.clone();
    cell.set_formula(formula)?;
    let replacement = crate::worksheet::codec::write_cell_fragment(&cell, None)?;
    splice_cell_start(source, location, replacement, max_output)
}

pub(crate) fn set_cell_style(
    source: &[u8],
    sheet: &str,
    row: usize,
    column: usize,
    style_name: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    let sheet_model = spreadsheet
        .sheet(sheet)
        .ok_or_else(|| invalid_error(format!("ODS sheet '{sheet}' was not found")))?;
    let location = selected_cell(source, sheet_model, row, column)?;
    let mut cell = location.cell.clone();
    cell.set_style_name(style_name)?;
    let replacement = crate::worksheet::codec::write_cell_fragment(&cell, None)?;
    splice_cell_start(source, location, replacement, max_output)
}

pub(crate) fn append_row(
    source: &[u8],
    sheet: &str,
    row: &Row,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let rows = children(&spans, table, TABLE, "table-row");
    let insertion = rows.last().map_or_else(
        || first_sheet_extension(&spans, table).unwrap_or(spans[table].close_start),
        |index| spans[*index].end,
    );
    let markup =
        crate::worksheet::codec::write_rows_bounded(std::slice::from_ref(row), max_output)?;
    splice_content(
        source,
        insertion..insertion,
        markup.into_bytes(),
        max_output,
    )
}

pub(crate) fn remove_row(
    source: &[u8],
    sheet: &str,
    physical_position: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let rows = children(&spans, table, TABLE, "table-row");
    let index = *rows
        .get(physical_position)
        .ok_or_else(|| invalid_error("ODS physical row position is out of bounds"))?;
    if attribute(&xml, &spans[index], b"table:number-rows-repeated")?
        .is_some_and(|value| value != "1")
    {
        return invalid("ODS row removal refuses repeated physical runs");
    }
    splice_content(
        source,
        spans[index].start..spans[index].end,
        Vec::new(),
        max_output,
    )
}

pub(crate) fn append_column(
    source: &[u8],
    sheet: &str,
    column: &crate::model::structure::Column,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let columns = children(&spans, table, TABLE, "table-column");
    let insertion = columns
        .last()
        .map_or(spans[table].tag_end, |index| spans[*index].end);
    let mut markup = String::new();
    crate::model::structure::write_columns(&mut markup, std::slice::from_ref(column));
    markup = markup.replacen(
        "<table:table-column",
        &format!("<table:table-column xmlns:table=\"{TABLE}\""),
        1,
    );
    splice_content(
        source,
        insertion..insertion,
        markup.into_bytes(),
        max_output,
    )
}

pub(crate) fn remove_column(
    source: &[u8],
    sheet: &str,
    physical_position: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let columns = children(&spans, table, TABLE, "table-column");
    let index = *columns
        .get(physical_position)
        .ok_or_else(|| invalid_error("ODS physical column position is out of bounds"))?;
    if attribute(&xml, &spans[index], b"table:number-columns-repeated")?
        .is_some_and(|value| value != "1")
    {
        return invalid("ODS column removal refuses repeated physical runs");
    }
    splice_content(
        source,
        spans[index].start..spans[index].end,
        Vec::new(),
        max_output,
    )
}

pub(crate) fn append_sheet(source: &[u8], sheet: &Sheet, max_output: usize) -> Result<Vec<u8>> {
    let (_xml, spans) = content_spans(source)?;
    let spreadsheet = one(&spans, OFFICE, "spreadsheet")?;
    let markup = crate::worksheet::codec::write_sheet(sheet)?;
    splice_content(
        source,
        spans[spreadsheet].close_start..spans[spreadsheet].close_start,
        markup.into_bytes(),
        max_output,
    )
}

pub(crate) fn remove_sheet(source: &[u8], sheet: &str, max_output: usize) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    splice_content(
        source,
        spans[table].start..spans[table].end,
        Vec::new(),
        max_output,
    )
}

pub(crate) fn put_cell_style(
    source: &[u8],
    style: &CellStyle,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let automatic = spans
        .iter()
        .enumerate()
        .find(|(_, span)| is_element(span, OFFICE, "automatic-styles"))
        .map(|(index, _)| index);
    let style_markup = style.markup()?;
    let Some(automatic) = automatic else {
        let body = one(&spans, OFFICE, "body")?;
        let markup = format!(
            "<office:automatic-styles xmlns:office=\"{OFFICE}\">{style_markup}</office:automatic-styles>"
        );
        return splice_content(
            source,
            spans[body].start..spans[body].start,
            markup.into_bytes(),
            max_output,
        );
    };
    for index in children(&spans, automatic, STYLE, "style") {
        if attribute(&xml, &spans[index], b"style:name")?.as_deref() == Some(&style.name) {
            return invalid("ODS automatic cell style already exists");
        }
    }
    if spans[automatic].start == spans[automatic].close_start
        || xml[spans[automatic].start..spans[automatic].tag_end].ends_with("/>")
    {
        let markup = format!(
            "<office:automatic-styles xmlns:office=\"{OFFICE}\">{style_markup}</office:automatic-styles>"
        );
        splice_content(
            source,
            spans[automatic].start..spans[automatic].end,
            markup.into_bytes(),
            max_output,
        )
    } else {
        splice_content(
            source,
            spans[automatic].close_start..spans[automatic].close_start,
            style_markup.into_bytes(),
            max_output,
        )
    }
}

pub(crate) fn set_conditional_formats(
    source: &[u8],
    sheet: &str,
    formats: &[crate::model::conditional_format::Format],
    max_output: usize,
) -> Result<Vec<u8>> {
    let mut markup = String::new();
    crate::model::conditional_format::write_conditional_formats(&mut markup, formats)?;
    bind_root_namespace(
        &mut markup,
        "calcext:conditional-formats",
        "calcext",
        CALCEXT,
    );
    set_sheet_container(
        source,
        sheet,
        CALCEXT,
        "conditional-formats",
        markup,
        max_output,
    )
}

pub(crate) fn set_sparkline_groups(
    source: &[u8],
    sheet: &str,
    groups: &[crate::model::sparkline::Group],
    max_output: usize,
) -> Result<Vec<u8>> {
    let mut markup = String::new();
    crate::model::sparkline::write_sparkline_groups(&mut markup, groups)?;
    bind_root_namespace(&mut markup, "calcext:sparkline-groups", "calcext", CALCEXT);
    set_sheet_container(
        source,
        sheet,
        CALCEXT,
        "sparkline-groups",
        markup,
        max_output,
    )
}

pub(crate) fn put_drawing(
    source: &[u8],
    sheet: &str,
    drawing: &Drawing,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_token(&drawing.name, "drawing name")?;
    super::document::validate_detached_resource_path(&drawing.resource_path)?;
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let shapes = children(&spans, table, TABLE, "shapes");
    if shapes.len() > 1 {
        return invalid("ODS sheet has more than one shapes container");
    }
    if let Some(shapes) = shapes.first().copied()
        && descendants(&spans, shapes, DRAW, "frame")
            .into_iter()
            .any(|index| {
                attribute(&xml, &spans[index], b"draw:name")
                    .ok()
                    .flatten()
                    .as_deref()
                    == Some(&drawing.name)
            })
    {
        return invalid("ODS drawing name already exists");
    }
    let frame = format!(
        "<draw:frame xmlns:draw=\"{DRAW}\" xmlns:xlink=\"{XLINK}\" draw:name=\"{}\"><draw:image xlink:href=\"{}\" xlink:type=\"simple\"/></draw:frame>",
        escape_xml(&drawing.name),
        escape_xml(&drawing.resource_path)
    );
    let (range, markup) = if let Some(index) = shapes.first().copied() {
        (spans[index].close_start..spans[index].close_start, frame)
    } else {
        let position = spans[table].close_start;
        (
            position..position,
            format!("<table:shapes xmlns:table=\"{TABLE}\">{frame}</table:shapes>"),
        )
    };
    splice_content(source, range, markup.into_bytes(), max_output)
}

pub(crate) fn remove_drawing(
    source: &[u8],
    sheet: &str,
    name: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let shapes = children(&spans, table, TABLE, "shapes");
    let frame = shapes
        .into_iter()
        .flat_map(|index| descendants(&spans, index, DRAW, "frame"))
        .find(|index| {
            attribute(&xml, &spans[*index], b"draw:name")
                .ok()
                .flatten()
                .as_deref()
                == Some(name)
        })
        .ok_or_else(|| invalid_error("ODS drawing was not found"))?;
    splice_content(
        source,
        spans[frame].start..spans[frame].end,
        Vec::new(),
        max_output,
    )
}

pub(crate) fn set_forms(
    source: &[u8],
    controls: &[FormControl],
    max_output: usize,
) -> Result<Vec<u8>> {
    let mut markup = String::new();
    if !controls.is_empty() {
        write!(
            markup,
            "<office:forms xmlns:office=\"{OFFICE}\" xmlns:form=\"{FORM}\"><form:form form:name=\"LitchiForm\">"
        )
        .map_err(|_error| invalid_error("ODS form markup formatting failed"))?;
        for control in controls {
            validate_token(&control.id, "form control id")?;
            if control.label.len() > 65_536 || control.label.chars().any(char::is_control) {
                return invalid("ODS form control label is invalid");
            }
            markup.push_str("<form:button form:id=\"");
            markup.push_str(&escape_xml(&control.id));
            markup.push_str("\" form:label=\"");
            markup.push_str(&escape_xml(&control.label));
            markup.push_str("\"/>");
        }
        markup.push_str("</form:form></office:forms>");
    }
    let (_xml, spans) = content_spans(source)?;
    let spreadsheet = one(&spans, OFFICE, "spreadsheet")?;
    let existing = children(&spans, spreadsheet, OFFICE, "forms");
    if existing.len() > 1 {
        return invalid("ODS spreadsheet has more than one forms container");
    }
    let range = existing.first().copied().map_or_else(
        || {
            let insertion = children(&spans, spreadsheet, TABLE, "table")
                .first()
                .map_or(spans[spreadsheet].close_start, |index| spans[*index].start);
            insertion..insertion
        },
        |index| spans[index].start..spans[index].end,
    );
    splice_content(source, range, markup.into_bytes(), max_output)
}

fn selected_cell<'a>(
    source: &[u8],
    sheet: &'a Sheet,
    row: usize,
    column: usize,
) -> Result<CellLocation<'a>> {
    let mut row_start = 0usize;
    let (physical_row, row_model) = sheet
        .rows
        .iter()
        .enumerate()
        .find_map(|(index, candidate)| {
            let end = row_start.saturating_add(candidate.repeat());
            let selected = (row < end).then_some((index, candidate));
            row_start = end;
            selected
        })
        .ok_or_else(|| invalid_error("ODS rich cell row is missing"))?;
    if row_model.repeat() != 1 {
        return invalid("ODS fine-grained cell edit refuses repeated row runs");
    }
    let mut column_start = 0usize;
    let (physical_cell, cell) = row_model
        .cells
        .iter()
        .enumerate()
        .find_map(|(index, candidate)| {
            let end = column_start.saturating_add(candidate.repeat());
            let selected = (column < end).then_some((index, candidate));
            column_start = end;
            selected
        })
        .ok_or_else(|| invalid_error("ODS rich cell column is missing"))?;
    if cell.repeat() != 1 {
        return invalid("ODS fine-grained cell edit refuses repeated cell runs");
    }
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, &sheet.name)?;
    let rows = children(&spans, table, TABLE, "table-row");
    let row_span = *rows
        .get(physical_row)
        .ok_or_else(|| invalid_error("ODS row span inventory differs from the typed graph"))?;
    let cells = children_any_local(
        &spans,
        row_span,
        TABLE,
        &["table-cell", "covered-table-cell"],
    );
    let cell_span = *cells
        .get(physical_cell)
        .ok_or_else(|| invalid_error("ODS cell span inventory differs from the typed graph"))?;
    let span = &spans[cell_span];
    let empty = xml[span.start..span.tag_end].ends_with("/>");
    Ok(CellLocation {
        cell,
        range: span.start..span.end,
        tag: span.start..span.tag_end,
        empty,
    })
}

fn set_sheet_container(
    source: &[u8],
    sheet: &str,
    namespace: &str,
    local: &str,
    markup: String,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let existing = children(&spans, table, namespace, local);
    if existing.len() > 1 {
        return invalid("ODS sheet extension container is duplicated");
    }
    let range = existing.first().copied().map_or_else(
        || {
            let insertion =
                first_sheet_extension(&spans, table).unwrap_or(spans[table].close_start);
            insertion..insertion
        },
        |index| spans[index].start..spans[index].end,
    );
    splice_content(source, range, markup.into_bytes(), max_output)
}

fn bind_root_namespace(markup: &mut String, root: &str, prefix: &str, namespace: &str) {
    if markup.is_empty() {
        return;
    }
    let opening = format!("<{root}");
    let bound = format!("<{root} xmlns:{prefix}=\"{namespace}\"");
    *markup = markup.replacen(&opening, &bound, 1);
}

fn splice_content(
    source: &[u8],
    range: Range<usize>,
    replacement: Vec<u8>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let package = Package::from_bytes(source.to_vec())?;
    let part = XmlSourcePart::load(package.package(), CONTENT_PATH)?;
    let expected = part
        .bytes()
        .get(range.clone())
        .ok_or_else(|| invalid_error("ODS content splice range is invalid"))?;
    let proof = part.checked_range(range, expected)?;
    let fragment = if replacement.is_empty() {
        AuthoredXmlFragment::deletion()
    } else {
        AuthoredXmlFragment::markup(replacement)?
    };
    let mut publication = XmlSplicePublication::new(part);
    publication.replace(proof, fragment)?;
    rebuild_package_with_xml_splices(package.package(), vec![publication], max_output)
}

fn splice_cell_start(
    source: &[u8],
    location: CellLocation<'_>,
    replacement: String,
    max_output: usize,
) -> Result<Vec<u8>> {
    if location.empty {
        return splice_content(source, location.range, replacement.into_bytes(), max_output);
    }
    let tag_end = replacement
        .find('>')
        .ok_or_else(|| invalid_error("ODS replacement cell start tag is missing"))?;
    let tag = replacement.as_bytes()[..=tag_end].to_vec();
    let package = Package::from_bytes(source.to_vec())?;
    let part = XmlSourcePart::load(package.package(), CONTENT_PATH)?;
    let expected = part
        .bytes()
        .get(location.tag.clone())
        .ok_or_else(|| invalid_error("ODS cell start-tag splice range is invalid"))?;
    let proof = part.checked_range(location.tag, expected)?;
    let fragment = AuthoredXmlFragment::start_tag(tag)?;
    let mut publication = XmlSplicePublication::new(part);
    publication.replace(proof, fragment)?;
    rebuild_package_with_xml_splices(package.package(), vec![publication], max_output)
}

fn content_spans(source: &[u8]) -> Result<(String, Vec<Span>)> {
    let package = Package::from_bytes(source.to_vec())?;
    let xml = package.content_xml().to_string();
    let spans = scan(&xml)?;
    Ok((xml, spans))
}

fn scan(xml: &str) -> Result<Vec<Span>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut spans = Vec::new();
    let mut open = Vec::<usize>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS content XML: {error}")))?;
        match event {
            Event::Start(element) => {
                if spans.len() >= MAX_ELEMENTS {
                    return invalid("ODS content element limit exceeded");
                }
                let namespace = resolve_namespace(&namespace)?;
                let element = element.into_owned();
                let parent = open.last().copied();
                let tag_end = position(&reader)?;
                let index = spans.len();
                spans.push(Span {
                    namespace,
                    local: decode(element.local_name().as_ref(), "element local name")?,
                    start: tag_start(xml, tag_end)?,
                    tag_end,
                    close_start: tag_end,
                    end: tag_end,
                    parent,
                });
                open.push(index);
            },
            Event::Empty(element) => {
                if spans.len() >= MAX_ELEMENTS {
                    return invalid("ODS content element limit exceeded");
                }
                let namespace = resolve_namespace(&namespace)?;
                let element = element.into_owned();
                let parent = open.last().copied();
                let tag_end = position(&reader)?;
                spans.push(Span {
                    namespace,
                    local: decode(element.local_name().as_ref(), "element local name")?,
                    start: tag_start(xml, tag_end)?,
                    tag_end,
                    close_start: tag_end,
                    end: tag_end,
                    parent,
                });
            },
            Event::End(_) => {
                let index = open
                    .pop()
                    .ok_or_else(|| invalid_error("ODS content element stack underflow"))?;
                let end = position(&reader)?;
                let close_start = xml.as_bytes()[..end]
                    .windows(2)
                    .rposition(|window| window == b"</")
                    .ok_or_else(|| invalid_error("ODS closing tag is missing"))?;
                spans[index].close_start = close_start;
                spans[index].end = end;
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !open.is_empty() {
        return invalid("ODS content contains unclosed elements");
    }
    Ok(spans)
}

fn select_sheet(xml: &str, spans: &[Span], name: &str) -> Result<usize> {
    let spreadsheet = one(spans, OFFICE, "spreadsheet")?;
    let mut selected = children(spans, spreadsheet, TABLE, "table")
        .into_iter()
        .filter(|index| {
            attribute(xml, &spans[*index], b"table:name")
                .ok()
                .flatten()
                .as_deref()
                == Some(name)
        });
    let result = selected
        .next()
        .ok_or_else(|| invalid_error(format!("ODS sheet '{name}' was not found")))?;
    if selected.next().is_some() {
        return invalid("ODS sheet selector is ambiguous");
    }
    Ok(result)
}

fn attribute(xml: &str, span: &Span, name: &[u8]) -> Result<Option<String>> {
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS element start-tag span is invalid"))?;
    let mut reader = quick_xml::Reader::from_str(tag);
    let event = reader
        .read_event()
        .map_err(|error| invalid_error(format!("invalid ODS start tag: {error}")))?;
    let (Event::Start(element) | Event::Empty(element)) = event else {
        return invalid("ODS element start tag is missing");
    };
    for candidate in element.attributes().with_checks(true) {
        let candidate =
            candidate.map_err(|error| invalid_error(format!("invalid ODS attribute: {error}")))?;
        if candidate.key.as_ref() == name {
            return candidate
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid_error(format!("invalid ODS attribute value: {error}")));
        }
    }
    Ok(None)
}

fn one(spans: &[Span], namespace: &str, local: &str) -> Result<usize> {
    let mut found = spans
        .iter()
        .enumerate()
        .filter(|(_, span)| is_element(span, namespace, local));
    let result = found
        .next()
        .map(|(index, _)| index)
        .ok_or_else(|| invalid_error(format!("ODS {local} element is missing")))?;
    if found.next().is_some() {
        return invalid(format!("ODS {local} element is duplicated"));
    }
    Ok(result)
}

fn children(spans: &[Span], parent: usize, namespace: &str, local: &str) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter_map(|(index, span)| {
            (span.parent == Some(parent) && is_element(span, namespace, local)).then_some(index)
        })
        .collect()
}

fn children_any_local(
    spans: &[Span],
    parent: usize,
    namespace: &str,
    locals: &[&str],
) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter_map(|(index, span)| {
            (span.parent == Some(parent)
                && span.namespace.as_deref() == Some(namespace)
                && locals.contains(&span.local.as_str()))
            .then_some(index)
        })
        .collect()
}

fn descendants(spans: &[Span], parent: usize, namespace: &str, local: &str) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter_map(|(index, span)| {
            (is_descendant(spans, index, parent) && is_element(span, namespace, local))
                .then_some(index)
        })
        .collect()
}

fn is_descendant(spans: &[Span], mut index: usize, parent: usize) -> bool {
    while let Some(candidate) = spans[index].parent {
        if candidate == parent {
            return true;
        }
        index = candidate;
    }
    false
}

fn first_sheet_extension(spans: &[Span], table: usize) -> Option<usize> {
    spans
        .iter()
        .enumerate()
        .filter(|(_, span)| span.parent == Some(table))
        .filter(|(_, span)| {
            !(span.namespace.as_deref() == Some(TABLE)
                && matches!(span.local.as_str(), "table-column" | "table-row"))
        })
        .map(|(_, span)| span.start)
        .min()
}

fn is_element(span: &Span, namespace: &str, local: &str) -> bool {
    span.namespace.as_deref() == Some(namespace) && span.local == local
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("ODS XML position overflows usize"))
}

fn tag_start(xml: &str, tag_end: usize) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.as_bytes()[..tag_end].iter().enumerate().rev() {
        match (quote, byte) {
            (Some(delimiter), current) if current == &delimiter => quote = None,
            (Some(_), _) => {},
            (None, b'\'' | b'"') => quote = Some(*byte),
            (None, b'<') => return Ok(index),
            _ => {},
        }
    }
    invalid("ODS XML element start is missing")
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(decode(uri, "element namespace")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound ODS prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn decode(bytes: &[u8], label: &str) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|_error| invalid_error(format!("ODS {label} is not UTF-8")))
}

fn validate_token(value: &str, label: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 65_536
        || value.trim() != value
        || value.chars().any(char::is_control)
    {
        invalid(format!("ODS {label} is invalid"))
    } else {
        Ok(())
    }
}

fn validate_href(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 1_048_576
        || value.chars().any(char::is_control)
        || value.starts_with("javascript:")
        || value.starts_with("data:")
    {
        invalid("ODS rich-text hyperlink is unsafe")
    } else {
        Ok(())
    }
}

fn validate_color(value: Option<&str>) -> Result<()> {
    if value.is_some_and(|color| {
        color.len() != 7
            || !color.starts_with('#')
            || !color[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    }) {
        invalid("ODS cell style color must be #RRGGBB")
    } else {
        Ok(())
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
