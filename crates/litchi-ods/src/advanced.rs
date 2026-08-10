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
const NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const DOM: &str = "http://www.w3.org/2001/xml-events";
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

/// Text properties shared by automatic text and table-cell styles.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct TextProperties {
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub font_size_pt: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
}

/// Table-cell properties retained by a deep automatic style node.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CellProperties {
    pub background: Option<String>,
    pub horizontal_align: Option<String>,
    pub vertical_align: Option<String>,
    pub wrap: Option<bool>,
    pub border: Option<String>,
}

/// One automatic table-cell style with checked parent and number-style dependencies.
#[derive(Clone, Debug, PartialEq)]
pub struct CellStyleNode {
    pub name: String,
    pub parent: Option<String>,
    pub data_style: Option<String>,
    pub cell: CellProperties,
    pub text: TextProperties,
}

/// One automatic text style referenced by rich cell runs.
#[derive(Clone, Debug, PartialEq)]
pub struct TextStyleNode {
    pub name: String,
    pub parent: Option<String>,
    pub text: TextProperties,
}

/// One bounded decimal number style referenced by a cell style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NumberStyleNode {
    pub name: String,
    pub decimal_places: u8,
    pub min_integer_digits: u8,
    pub prefix: Option<String>,
    pub suffix: Option<String>,
}

/// Common ODF data-style families supported by automatic style graphs.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum DataStyleNode {
    Date {
        name: String,
    },
    Time {
        name: String,
        decimal_places: u8,
    },
    Currency {
        name: String,
        symbol: String,
        decimal_places: u8,
        min_integer_digits: u8,
    },
    Percentage {
        name: String,
        decimal_places: u8,
        min_integer_digits: u8,
    },
    Boolean {
        name: String,
    },
}

impl DataStyleNode {
    fn name(&self) -> &str {
        match self {
            Self::Date { name }
            | Self::Time { name, .. }
            | Self::Currency { name, .. }
            | Self::Percentage { name, .. }
            | Self::Boolean { name } => name,
        }
    }
}

/// A dependency-checked automatic style graph.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct StyleGraph {
    pub cell_styles: Vec<CellStyleNode>,
    pub text_styles: Vec<TextStyleNode>,
    pub number_styles: Vec<NumberStyleNode>,
    pub data_styles: Vec<DataStyleNode>,
}

/// Fully inherited table-cell style properties resolved from one closed graph.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct EffectiveCellStyle {
    pub lineage: Vec<String>,
    pub data_style: Option<String>,
    pub cell: CellProperties,
    pub text: TextProperties,
}

impl StyleGraph {
    /// Resolve one cell style from oldest parent to selected child.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid graph or absent/cyclic style dependency.
    pub fn resolve_cell_style(&self, name: &str) -> Result<EffectiveCellStyle> {
        let _validated = style_graph_markup(self)?;
        let styles = self
            .cell_styles
            .iter()
            .map(|style| (style.name.as_str(), style))
            .collect::<std::collections::BTreeMap<_, _>>();
        let mut lineage = Vec::new();
        let mut current = Some(name);
        while let Some(selected) = current {
            let style = styles.get(selected).ok_or_else(|| {
                invalid_error(format!("ODS cell style '{selected}' was not found"))
            })?;
            lineage.push(*style);
            current = style.parent.as_deref();
        }
        lineage.reverse();
        let mut effective = EffectiveCellStyle::default();
        for style in lineage {
            effective.lineage.push(style.name.clone());
            if style.data_style.is_some() {
                effective.data_style.clone_from(&style.data_style);
            }
            merge_cell_properties(&mut effective.cell, &style.cell);
            merge_text_properties(&mut effective.text, &style.text);
        }
        Ok(effective)
    }
}

/// One bound inert form control with an optional package image dependency.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BoundFormControl {
    pub id: String,
    pub label: String,
    pub linked_cell: Option<String>,
    pub source_range: Option<String>,
    pub image_path: Option<String>,
}

/// Inert ODF form-control vocabulary supported by the unified root.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum FormControlKind {
    Button,
    CheckBox,
    Radio,
    ListBox,
    ComboBox,
    Text,
    Image,
    Date,
    Time,
    FixedText,
    FormattedText,
    Number,
    File,
    Password,
    TextArea,
    Hidden,
    ValueRange,
    Frame,
    ImageFrame,
    GenericControl,
}

/// One inert form event binding. Macro targets are retained but never executed.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormEvent {
    pub event_name: String,
    pub macro_name: String,
}

/// A typed form control with bindings, optional image data, and inert events.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RichFormControl {
    pub id: String,
    pub label: String,
    pub kind: FormControlKind,
    pub linked_cell: Option<String>,
    pub source_range: Option<String>,
    pub image_path: Option<String>,
    pub events: Vec<FormEvent>,
}

/// A positioned drawing frame with one package image dependency.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingFrame {
    pub name: String,
    pub resource_path: String,
    pub anchor_cell: Option<String>,
    pub x_cm: f64,
    pub y_cm: f64,
    pub width_cm: f64,
    pub height_cm: f64,
    pub z_index: u32,
}

/// One bounded text frame inside a drawing group.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingTextBox {
    pub name: String,
    pub text: RichText,
    pub x_cm: f64,
    pub y_cm: f64,
    pub width_cm: f64,
    pub height_cm: f64,
    pub z_index: u32,
}

/// One named drawing group containing positioned text frames.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingGroup {
    pub name: String,
    pub children: Vec<DrawingTextBox>,
}

/// Supported inert drawing geometry.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum DrawingGeometryKind {
    Rectangle,
    Ellipse,
    Line,
    Connector,
}

/// One integer point in a bounded drawing view box.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DrawingPoint {
    pub x: u32,
    pub y: u32,
}

/// A positioned polygon or polyline using checked view-box coordinates.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingPolygon {
    pub name: String,
    pub closed: bool,
    pub points: Vec<DrawingPoint>,
    pub view_box_width: u32,
    pub view_box_height: u32,
    pub text: Option<RichText>,
    pub x_cm: f64,
    pub y_cm: f64,
    pub width_cm: f64,
    pub height_cm: f64,
    pub z_index: u32,
}

/// One finite positioned geometry with optional rich text.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingGeometry {
    pub name: String,
    pub kind: DrawingGeometryKind,
    pub text: Option<RichText>,
    pub x_cm: f64,
    pub y_cm: f64,
    pub width_cm: f64,
    pub height_cm: f64,
    pub z_index: u32,
}

/// A package-backed chart object and its compact chart `content.xml` dependency.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartObject {
    pub name: String,
    pub object_path: String,
    pub content_xml: String,
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

pub(crate) fn put_style_graph(
    source: &[u8],
    graph: &StyleGraph,
    max_output: usize,
) -> Result<Vec<u8>> {
    let markup = style_graph_markup(graph)?;
    if markup.is_empty() {
        return Ok(source.to_vec());
    }
    insert_automatic_styles(source, &markup, max_output)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum StyleNodeKind {
    Number,
    Date,
    Time,
    Currency,
    Percentage,
    Boolean,
    Text,
    Cell,
}

fn style_graph_nodes(graph: &StyleGraph) -> Result<Vec<(String, StyleNodeKind, String)>> {
    let _validated = style_graph_markup(graph)?;
    let mut nodes = Vec::with_capacity(
        graph
            .number_styles
            .len()
            .saturating_add(graph.data_styles.len())
            .saturating_add(graph.text_styles.len())
            .saturating_add(graph.cell_styles.len()),
    );
    for style in &graph.number_styles {
        nodes.push((
            style.name.clone(),
            StyleNodeKind::Number,
            number_style_markup(style)?,
        ));
    }
    for style in &graph.data_styles {
        nodes.push((
            style.name().to_string(),
            data_style_kind(style),
            data_style_markup(style)?,
        ));
    }
    for style in &graph.text_styles {
        nodes.push((
            style.name.clone(),
            StyleNodeKind::Text,
            text_style_markup(style)?,
        ));
    }
    for style in &graph.cell_styles {
        nodes.push((
            style.name.clone(),
            StyleNodeKind::Cell,
            cell_style_node_markup(style)?,
        ));
    }
    Ok(nodes)
}

fn style_graph_markup(graph: &StyleGraph) -> Result<String> {
    let total = graph
        .cell_styles
        .len()
        .saturating_add(graph.text_styles.len())
        .saturating_add(graph.number_styles.len())
        .saturating_add(graph.data_styles.len());
    if total > 4_096 {
        return invalid("ODS automatic style graph exceeds the node limit");
    }
    let cell_names = graph
        .cell_styles
        .iter()
        .map(|style| style.name.as_str())
        .collect::<std::collections::BTreeSet<_>>();
    let text_names = graph
        .text_styles
        .iter()
        .map(|style| style.name.as_str())
        .collect::<std::collections::BTreeSet<_>>();
    let number_names = graph
        .number_styles
        .iter()
        .map(|style| style.name.as_str())
        .collect::<std::collections::BTreeSet<_>>();
    let data_names = graph
        .data_styles
        .iter()
        .map(DataStyleNode::name)
        .collect::<std::collections::BTreeSet<_>>();
    let all_names = cell_names
        .iter()
        .chain(text_names.iter())
        .chain(number_names.iter())
        .chain(data_names.iter())
        .copied()
        .collect::<std::collections::BTreeSet<_>>();
    if cell_names.len() != graph.cell_styles.len()
        || text_names.len() != graph.text_styles.len()
        || number_names.len() != graph.number_styles.len()
        || data_names.len() != graph.data_styles.len()
        || all_names.len() != total
    {
        return invalid("ODS automatic style graph contains duplicate names");
    }
    validate_parent_graph(
        graph
            .cell_styles
            .iter()
            .map(|style| (style.name.as_str(), style.parent.as_deref())),
    )?;
    validate_parent_graph(
        graph
            .text_styles
            .iter()
            .map(|style| (style.name.as_str(), style.parent.as_deref())),
    )?;
    for style in &graph.cell_styles {
        validate_token(&style.name, "cell style name")?;
        if style
            .parent
            .as_deref()
            .is_some_and(|parent| !cell_names.contains(parent))
            || style
                .data_style
                .as_deref()
                .is_some_and(|name| !number_names.contains(name) && !data_names.contains(name))
        {
            return invalid("ODS cell style dependency is unresolved");
        }
        validate_cell_properties(&style.cell)?;
        validate_text_properties(&style.text)?;
    }
    for style in &graph.text_styles {
        validate_token(&style.name, "text style name")?;
        if style
            .parent
            .as_deref()
            .is_some_and(|parent| !text_names.contains(parent))
        {
            return invalid("ODS text style parent dependency is unresolved");
        }
        validate_text_properties(&style.text)?;
    }
    for style in &graph.number_styles {
        validate_token(&style.name, "number style name")?;
        if style.decimal_places > 20
            || style.min_integer_digits == 0
            || style.min_integer_digits > 20
            || style.prefix.as_deref().is_some_and(invalid_style_text)
            || style.suffix.as_deref().is_some_and(invalid_style_text)
        {
            return invalid("ODS number style digits or text are invalid");
        }
    }
    for style in &graph.data_styles {
        validate_token(style.name(), "data style name")?;
        match style {
            DataStyleNode::Date { .. } | DataStyleNode::Boolean { .. } => {},
            DataStyleNode::Time { decimal_places, .. } if *decimal_places <= 9 => {},
            DataStyleNode::Currency {
                symbol,
                decimal_places,
                min_integer_digits,
                ..
            } if !symbol.is_empty()
                && symbol.len() <= 32
                && !symbol.chars().any(char::is_control)
                && *decimal_places <= 20
                && (1..=20).contains(min_integer_digits) => {},
            DataStyleNode::Percentage {
                decimal_places,
                min_integer_digits,
                ..
            } if *decimal_places <= 20 && (1..=20).contains(min_integer_digits) => {},
            DataStyleNode::Time { .. }
            | DataStyleNode::Currency { .. }
            | DataStyleNode::Percentage { .. } => {
                return invalid("ODS data style parameters are invalid");
            },
        }
    }
    let mut output = String::new();
    for style in &graph.number_styles {
        output.push_str(&number_style_markup(style)?);
    }
    for style in &graph.data_styles {
        output.push_str(&data_style_markup(style)?);
    }
    for style in &graph.text_styles {
        output.push_str(&text_style_markup(style)?);
    }
    for style in &graph.cell_styles {
        output.push_str(&cell_style_node_markup(style)?);
    }
    Ok(output)
}

fn number_style_markup(style: &NumberStyleNode) -> Result<String> {
    let mut output = String::new();
    write!(
        output,
        "<number:number-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{}\">",
        escape_xml(&style.name)
    )
    .map_err(|_error| invalid_error("ODS number style formatting failed"))?;
    if let Some(prefix) = &style.prefix {
        output.push_str("<number:text>");
        output.push_str(&escape_xml(prefix));
        output.push_str("</number:text>");
    }
    write!(
        output,
        "<number:number number:decimal-places=\"{}\" number:min-integer-digits=\"{}\"/>",
        style.decimal_places, style.min_integer_digits
    )
    .map_err(|_error| invalid_error("ODS number style formatting failed"))?;
    if let Some(suffix) = &style.suffix {
        output.push_str("<number:text>");
        output.push_str(&escape_xml(suffix));
        output.push_str("</number:text>");
    }
    output.push_str("</number:number-style>");
    Ok(output)
}

fn data_style_kind(style: &DataStyleNode) -> StyleNodeKind {
    match style {
        DataStyleNode::Date { .. } => StyleNodeKind::Date,
        DataStyleNode::Time { .. } => StyleNodeKind::Time,
        DataStyleNode::Currency { .. } => StyleNodeKind::Currency,
        DataStyleNode::Percentage { .. } => StyleNodeKind::Percentage,
        DataStyleNode::Boolean { .. } => StyleNodeKind::Boolean,
    }
}

fn data_style_markup(style: &DataStyleNode) -> Result<String> {
    let name = escape_xml(style.name());
    Ok(match style {
        DataStyleNode::Date { .. } => format!(
            "<number:date-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{name}\"><number:year number:style=\"long\"/><number:text>-</number:text><number:month number:style=\"long\"/><number:text>-</number:text><number:day number:style=\"long\"/></number:date-style>"
        ),
        DataStyleNode::Time { decimal_places, .. } => format!(
            "<number:time-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{name}\"><number:hours number:style=\"long\"/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text><number:seconds number:style=\"long\" number:decimal-places=\"{decimal_places}\"/></number:time-style>"
        ),
        DataStyleNode::Currency {
            symbol,
            decimal_places,
            min_integer_digits,
            ..
        } => format!(
            "<number:currency-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{name}\"><number:currency-symbol>{}</number:currency-symbol><number:number number:decimal-places=\"{decimal_places}\" number:min-integer-digits=\"{min_integer_digits}\"/></number:currency-style>",
            escape_xml(symbol)
        ),
        DataStyleNode::Percentage {
            decimal_places,
            min_integer_digits,
            ..
        } => format!(
            "<number:percentage-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{name}\"><number:number number:decimal-places=\"{decimal_places}\" number:min-integer-digits=\"{min_integer_digits}\"/><number:text>%</number:text></number:percentage-style>"
        ),
        DataStyleNode::Boolean { .. } => format!(
            "<number:boolean-style xmlns:number=\"{NUMBER}\" xmlns:style=\"{STYLE}\" style:name=\"{name}\"><number:boolean/></number:boolean-style>"
        ),
    })
}

fn text_style_markup(style: &TextStyleNode) -> Result<String> {
    let mut output = String::new();
    write_style_open(
        &mut output,
        &style.name,
        "text",
        style.parent.as_deref(),
        None,
    )?;
    write_text_properties(&mut output, &style.text)?;
    output.push_str("</style:style>");
    Ok(output)
}

fn cell_style_node_markup(style: &CellStyleNode) -> Result<String> {
    let mut output = String::new();
    write_style_open(
        &mut output,
        &style.name,
        "table-cell",
        style.parent.as_deref(),
        style.data_style.as_deref(),
    )?;
    write_cell_properties(&mut output, &style.cell)?;
    write_text_properties(&mut output, &style.text)?;
    output.push_str("</style:style>");
    Ok(output)
}

fn validate_parent_graph<'a>(
    nodes: impl IntoIterator<Item = (&'a str, Option<&'a str>)>,
) -> Result<()> {
    let parents = nodes
        .into_iter()
        .collect::<std::collections::BTreeMap<_, _>>();
    for name in parents.keys() {
        let mut seen = std::collections::BTreeSet::new();
        let mut current = Some(*name);
        while let Some(node) = current {
            if !seen.insert(node) {
                return invalid("ODS automatic style graph contains a parent cycle");
            }
            current = parents.get(node).copied().flatten();
        }
    }
    Ok(())
}

fn invalid_style_text(value: &str) -> bool {
    value.len() > 65_536 || value.chars().any(char::is_control)
}

fn merge_cell_properties(target: &mut CellProperties, source: &CellProperties) {
    if source.background.is_some() {
        target.background.clone_from(&source.background);
    }
    if source.horizontal_align.is_some() {
        target.horizontal_align.clone_from(&source.horizontal_align);
    }
    if source.vertical_align.is_some() {
        target.vertical_align.clone_from(&source.vertical_align);
    }
    if source.wrap.is_some() {
        target.wrap = source.wrap;
    }
    if source.border.is_some() {
        target.border.clone_from(&source.border);
    }
}

fn merge_text_properties(target: &mut TextProperties, source: &TextProperties) {
    if source.color.is_some() {
        target.color.clone_from(&source.color);
    }
    if source.font_family.is_some() {
        target.font_family.clone_from(&source.font_family);
    }
    if source.font_size_pt.is_some() {
        target.font_size_pt = source.font_size_pt;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline = source.underline;
    }
}

pub(crate) fn resolve_package_cell_style(source: &[u8], name: &str) -> Result<EffectiveCellStyle> {
    validate_token(name, "effective cell style name")?;
    let package = Package::from_bytes(source.to_vec())?;
    let mut styles = std::collections::BTreeMap::new();
    if let Some(xml) = package.styles_xml() {
        collect_package_cell_styles(xml, &mut styles)?;
    }
    collect_package_cell_styles(package.content_xml(), &mut styles)?;
    let mut lineage = Vec::new();
    let mut visited = std::collections::BTreeSet::new();
    let mut current = Some(name);
    while let Some(selected) = current {
        if !visited.insert(selected.to_string()) {
            return invalid("ODS package cell style inheritance is cyclic");
        }
        let style = styles.get(selected).ok_or_else(|| {
            invalid_error(format!("ODS package cell style '{selected}' was not found"))
        })?;
        lineage.push(style);
        current = style.parent.as_deref();
    }
    lineage.reverse();
    let mut effective = EffectiveCellStyle::default();
    for style in lineage {
        effective.lineage.push(style.name.clone());
        if style.data_style.is_some() {
            effective.data_style.clone_from(&style.data_style);
        }
        merge_cell_properties(&mut effective.cell, &style.cell);
        merge_text_properties(&mut effective.text, &style.text);
    }
    Ok(effective)
}

fn collect_package_cell_styles(
    xml: &str,
    styles: &mut std::collections::BTreeMap<String, CellStyleNode>,
) -> Result<()> {
    let spans = scan(xml)?;
    let mut part_names = std::collections::BTreeSet::new();
    for (index, span) in spans.iter().enumerate() {
        if !is_element(span, STYLE, "style")
            || attribute(xml, span, b"style:family")?.as_deref() != Some("table-cell")
        {
            continue;
        }
        let Some(name) = attribute(xml, span, b"style:name")? else {
            return invalid("ODS package table-cell style has no name");
        };
        if !part_names.insert(name.clone()) {
            return invalid(format!(
                "ODS package table-cell style '{name}' is duplicated in one XML part"
            ));
        }
        let mut cell = CellProperties::default();
        let mut text = TextProperties::default();
        for child in descendants(&spans, index, STYLE, "table-cell-properties") {
            cell.background = attribute(xml, &spans[child], b"fo:background-color")?;
            cell.horizontal_align = attribute(xml, &spans[child], b"fo:text-align")?;
            cell.vertical_align = attribute(xml, &spans[child], b"style:vertical-align")?;
            cell.border = attribute(xml, &spans[child], b"fo:border")?;
            cell.wrap = attribute(xml, &spans[child], b"fo:wrap-option")?.and_then(|value| {
                match value.as_str() {
                    "wrap" => Some(true),
                    "no-wrap" => Some(false),
                    _ => None,
                }
            });
        }
        for child in descendants(&spans, index, STYLE, "text-properties") {
            text.color = attribute(xml, &spans[child], b"fo:color")?;
            text.font_family = attribute(xml, &spans[child], b"fo:font-family")?;
            text.font_size_pt = attribute(xml, &spans[child], b"fo:font-size")?
                .and_then(|value| value.strip_suffix("pt")?.parse::<f64>().ok())
                .filter(|value| value.is_finite() && *value > 0.0);
            text.bold = parse_style_switch(
                attribute(xml, &spans[child], b"fo:font-weight")?.as_deref(),
                "bold",
                "normal",
            );
            text.italic = parse_style_switch(
                attribute(xml, &spans[child], b"fo:font-style")?.as_deref(),
                "italic",
                "normal",
            );
            text.underline = parse_style_switch(
                attribute(xml, &spans[child], b"style:text-underline-style")?.as_deref(),
                "solid",
                "none",
            );
        }
        styles.insert(
            name.clone(),
            CellStyleNode {
                name,
                parent: attribute(xml, span, b"style:parent-style-name")?,
                data_style: attribute(xml, span, b"style:data-style-name")?,
                cell,
                text,
            },
        );
    }
    Ok(())
}

fn parse_style_switch(value: Option<&str>, enabled: &str, disabled: &str) -> Option<bool> {
    value.and_then(|value| match value {
        value if value == enabled => Some(true),
        value if value == disabled => Some(false),
        _ => None,
    })
}

fn insert_automatic_styles(source: &[u8], markup: &str, max_output: usize) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let automatic = spans
        .iter()
        .enumerate()
        .find(|(_, span)| is_element(span, OFFICE, "automatic-styles"))
        .map(|(index, _)| index);
    let Some(automatic) = automatic else {
        let body = one(&spans, OFFICE, "body")?;
        let container = format!(
            "<office:automatic-styles xmlns:office=\"{OFFICE}\">{markup}</office:automatic-styles>"
        );
        return splice_content(
            source,
            spans[body].start..spans[body].start,
            container.into_bytes(),
            max_output,
        );
    };
    for name in style_names(markup)? {
        for (index, span) in spans.iter().enumerate() {
            if span.parent == Some(automatic)
                && attribute(&xml, &spans[index], b"style:name")?.as_deref() == Some(name.as_str())
            {
                return invalid("ODS automatic style name already exists");
            }
        }
    }
    if xml[spans[automatic].start..spans[automatic].tag_end].ends_with("/>") {
        let container = format!(
            "<office:automatic-styles xmlns:office=\"{OFFICE}\">{markup}</office:automatic-styles>"
        );
        splice_content(
            source,
            spans[automatic].start..spans[automatic].end,
            container.into_bytes(),
            max_output,
        )
    } else {
        splice_content(
            source,
            spans[automatic].close_start..spans[automatic].close_start,
            markup.as_bytes().to_vec(),
            max_output,
        )
    }
}

pub(crate) fn replace_style_graph(
    source: &[u8],
    graph: &StyleGraph,
    max_output: usize,
) -> Result<Vec<u8>> {
    let nodes = style_graph_nodes(graph)?;
    if nodes.is_empty() {
        return Ok(source.to_vec());
    }
    let (xml, spans) = content_spans(source)?;
    let automatic = one(&spans, OFFICE, "automatic-styles")?;
    let mut edits = Vec::with_capacity(nodes.len());
    for (name, expected_kind, markup) in nodes {
        let matches = direct_named_styles(&xml, &spans, automatic, &name)?;
        if matches.len() != 1 {
            return invalid(format!(
                "ODS automatic style '{name}' must exist exactly once for replacement"
            ));
        }
        let index = matches[0];
        if automatic_style_kind(&xml, &spans[index])? != Some(expected_kind) {
            return invalid(format!(
                "ODS automatic style '{name}' has a different family"
            ));
        }
        edits.push((spans[index].start..spans[index].end, markup.into_bytes()));
    }
    splice_content_edits(source, edits, max_output)
}

pub(crate) fn remove_automatic_styles(
    source: &[u8],
    names: &[String],
    max_output: usize,
) -> Result<Vec<u8>> {
    if names.is_empty() {
        return Ok(source.to_vec());
    }
    let requested = names
        .iter()
        .map(String::as_str)
        .collect::<std::collections::BTreeSet<_>>();
    if requested.len() != names.len() {
        return invalid("ODS automatic style removal contains duplicate names");
    }
    for name in &requested {
        validate_token(name, "automatic style removal name")?;
    }
    refuse_package_style_references(source, &requested)?;
    let (xml, spans) = content_spans(source)?;
    let automatic = one(&spans, OFFICE, "automatic-styles")?;
    let mut removals = std::collections::BTreeSet::new();
    for name in &requested {
        let matches = direct_named_styles(&xml, &spans, automatic, name)?;
        if matches.len() != 1 {
            return invalid(format!(
                "ODS automatic style '{name}' must exist exactly once for removal"
            ));
        }
        removals.insert(matches[0]);
    }
    for (index, span) in spans.iter().enumerate() {
        if removals.contains(&index)
            || removals
                .iter()
                .any(|owner| is_descendant(&spans, index, *owner))
        {
            continue;
        }
        let opening = &xml[span.start..span.tag_end];
        for name in &requested {
            let escaped = escape_xml(name);
            if opening.contains(&format!("style-name=\"{escaped}\""))
                || opening.contains(&format!("style-name='{escaped}'"))
            {
                return invalid(format!("ODS automatic style '{name}' is still referenced"));
            }
        }
    }
    let edits = removals
        .into_iter()
        .map(|index| (spans[index].start..spans[index].end, Vec::new()))
        .collect();
    splice_content_edits(source, edits, max_output)
}

fn refuse_package_style_references(
    source: &[u8],
    requested: &std::collections::BTreeSet<&str>,
) -> Result<()> {
    let package = Package::from_bytes(source.to_vec())?;
    for path in package.package().files()? {
        let is_inspectable_xml = std::path::Path::new(&path)
            .extension()
            .is_some_and(|extension| extension == "xml" || extension == "rdf");
        if path == CONTENT_PATH || !is_inspectable_xml || !package.package().has_file(&path)? {
            continue;
        }
        let bytes = package.package().get_file(&path)?;
        let xml = std::str::from_utf8(&bytes).map_err(|_error| {
            invalid_error(format!(
                "ODS automatic style removal cannot inspect non-UTF-8 part '{path}'"
            ))
        })?;
        for name in requested {
            let escaped = escape_xml(name);
            if xml.contains(&format!("style-name=\"{escaped}\""))
                || xml.contains(&format!("style-name='{escaped}'"))
            {
                return invalid(format!(
                    "ODS automatic style '{name}' is referenced by retained part '{path}'"
                ));
            }
        }
    }
    Ok(())
}

fn direct_named_styles(
    xml: &str,
    spans: &[Span],
    automatic: usize,
    name: &str,
) -> Result<Vec<usize>> {
    let mut matches = Vec::new();
    for (index, span) in spans.iter().enumerate() {
        if span.parent == Some(automatic)
            && attribute(xml, span, b"style:name")?.as_deref() == Some(name)
        {
            matches.push(index);
        }
    }
    Ok(matches)
}

fn automatic_style_kind(xml: &str, span: &Span) -> Result<Option<StyleNodeKind>> {
    if span.namespace.as_deref() == Some(NUMBER) {
        return Ok(match span.local.as_str() {
            "number-style" => Some(StyleNodeKind::Number),
            "date-style" => Some(StyleNodeKind::Date),
            "time-style" => Some(StyleNodeKind::Time),
            "currency-style" => Some(StyleNodeKind::Currency),
            "percentage-style" => Some(StyleNodeKind::Percentage),
            "boolean-style" => Some(StyleNodeKind::Boolean),
            _ => None,
        });
    }
    if span.namespace.as_deref() != Some(STYLE) || span.local != "style" {
        return Ok(None);
    }
    Ok(match attribute(xml, span, b"style:family")?.as_deref() {
        Some("text") => Some(StyleNodeKind::Text),
        Some("table-cell") => Some(StyleNodeKind::Cell),
        _ => None,
    })
}

fn style_names(markup: &str) -> Result<Vec<String>> {
    let wrapped = format!("<root>{markup}</root>");
    let mut reader = quick_xml::Reader::from_str(&wrapped);
    let mut names = Vec::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid_error(format!("invalid ODS style graph: {error}")))?
        {
            Event::Start(element) if element.name().as_ref() != b"root" => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        invalid_error(format!("invalid ODS style attribute: {error}"))
                    })?;
                    if attribute.key.as_ref() == b"style:name" {
                        names.push(
                            attribute
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| {
                                    invalid_error(format!("invalid ODS style name: {error}"))
                                })?
                                .into_owned(),
                        );
                    }
                }
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(names)
}

fn write_style_open(
    output: &mut String,
    name: &str,
    family: &str,
    parent: Option<&str>,
    data_style: Option<&str>,
) -> Result<()> {
    write!(
        output,
        "<style:style xmlns:style=\"{STYLE}\" xmlns:fo=\"{FO}\" style:name=\"{}\" style:family=\"{family}\"",
        escape_xml(name)
    )
    .map_err(|_error| invalid_error("ODS style formatting failed"))?;
    if let Some(parent) = parent {
        write!(
            output,
            " style:parent-style-name=\"{}\"",
            escape_xml(parent)
        )
        .map_err(|_error| invalid_error("ODS style formatting failed"))?;
    }
    if let Some(data_style) = data_style {
        write!(
            output,
            " style:data-style-name=\"{}\"",
            escape_xml(data_style)
        )
        .map_err(|_error| invalid_error("ODS style formatting failed"))?;
    }
    output.push('>');
    Ok(())
}

fn validate_text_properties(value: &TextProperties) -> Result<()> {
    validate_color(value.color.as_deref())?;
    if value
        .font_family
        .as_deref()
        .is_some_and(|font| validate_token(font, "font family").is_err())
        || value
            .font_size_pt
            .is_some_and(|size| !size.is_finite() || !(1.0..=512.0).contains(&size))
    {
        return invalid("ODS text style properties are invalid");
    }
    Ok(())
}

fn validate_cell_properties(value: &CellProperties) -> Result<()> {
    validate_color(value.background.as_deref())?;
    if value
        .horizontal_align
        .as_deref()
        .is_some_and(|align| !matches!(align, "start" | "center" | "end" | "justify"))
        || value
            .vertical_align
            .as_deref()
            .is_some_and(|align| !matches!(align, "top" | "middle" | "bottom"))
        || value
            .border
            .as_deref()
            .is_some_and(|border| border.is_empty() || border.len() > 256)
    {
        return invalid("ODS cell style properties are invalid");
    }
    Ok(())
}

fn write_text_properties(output: &mut String, value: &TextProperties) -> Result<()> {
    if value == &TextProperties::default() {
        return Ok(());
    }
    output.push_str("<style:text-properties");
    push_optional_attribute(output, "fo:color", value.color.as_deref());
    push_optional_attribute(output, "fo:font-family", value.font_family.as_deref());
    if let Some(size) = value.font_size_pt {
        write!(output, " fo:font-size=\"{size}pt\"")
            .map_err(|_error| invalid_error("ODS text property formatting failed"))?;
    }
    push_optional_attribute(
        output,
        "fo:font-weight",
        value.bold.map(|set| if set { "bold" } else { "normal" }),
    );
    push_optional_attribute(
        output,
        "fo:font-style",
        value
            .italic
            .map(|set| if set { "italic" } else { "normal" }),
    );
    push_optional_attribute(
        output,
        "style:text-underline-style",
        value
            .underline
            .map(|set| if set { "solid" } else { "none" }),
    );
    output.push_str("/>");
    Ok(())
}

fn write_cell_properties(output: &mut String, value: &CellProperties) -> Result<()> {
    if value == &CellProperties::default() {
        return Ok(());
    }
    output.push_str("<style:table-cell-properties");
    push_optional_attribute(output, "fo:background-color", value.background.as_deref());
    push_optional_attribute(output, "fo:text-align", value.horizontal_align.as_deref());
    push_optional_attribute(
        output,
        "style:vertical-align",
        value.vertical_align.as_deref(),
    );
    push_optional_attribute(output, "fo:border", value.border.as_deref());
    push_optional_attribute(
        output,
        "fo:wrap-option",
        value.wrap.map(|set| if set { "wrap" } else { "no-wrap" }),
    );
    output.push_str("/>");
    Ok(())
}

fn push_optional_attribute(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
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
        && [
            "frame",
            "g",
            "rect",
            "ellipse",
            "line",
            "connector",
            "polygon",
            "polyline",
        ]
        .into_iter()
        .flat_map(|local| descendants(&spans, shapes, DRAW, local))
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

pub(crate) fn put_drawing_frame(
    source: &[u8],
    sheet: &str,
    frame: &DrawingFrame,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_frame(frame)?;
    if let Some(anchor) = &frame.anchor_cell {
        let address = crate::model::data_pilot::parse_data_pilot_range(anchor)?;
        if address.start_column != address.end_column
            || address.start_row != address.end_row
            || (!address.sheet.is_empty() && address.sheet != sheet)
        {
            return invalid("ODS drawing frame anchor must be one cell on its owning sheet");
        }
    }
    let child = format!(
        "<draw:frame xmlns:draw=\"{DRAW}\" xmlns:xlink=\"{XLINK}\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:table=\"{TABLE}\" draw:name=\"{}\" draw:z-index=\"{}\" svg:x=\"{}cm\" svg:y=\"{}cm\" svg:width=\"{}cm\" svg:height=\"{}cm\"{}><draw:image xlink:href=\"{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
        escape_xml(&frame.name),
        frame.z_index,
        frame.x_cm,
        frame.y_cm,
        frame.width_cm,
        frame.height_cm,
        frame
            .anchor_cell
            .as_deref()
            .map_or_else(String::new, |anchor| format!(
                " table:end-cell-address=\"{}\"",
                escape_xml(anchor)
            )),
        escape_xml(&frame.resource_path)
    );
    insert_shape_child(source, sheet, &frame.name, child, max_output)
}

pub(crate) fn put_chart_object(
    source: &[u8],
    sheet: &str,
    chart: &ChartObject,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_token(&chart.name, "chart drawing name")?;
    super::document::validate_detached_resource_path(&format!(
        "{}/content.xml",
        chart.object_path
    ))?;
    litchi_odf_common::compact_xml::validate(chart.content_xml.as_bytes()).map_err(Error::from)?;
    let child = format!(
        "<draw:frame xmlns:draw=\"{DRAW}\" xmlns:xlink=\"{XLINK}\" draw:name=\"{}\"><draw:object xlink:href=\"./{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
        escape_xml(&chart.name),
        escape_xml(&chart.object_path)
    );
    insert_shape_child(source, sheet, &chart.name, child, max_output)
}

pub(crate) fn put_drawing_group(
    source: &[u8],
    sheet: &str,
    group: &DrawingGroup,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_token(&group.name, "drawing group name")?;
    if group.children.is_empty() || group.children.len() > 4_096 {
        return invalid("ODS drawing group requires bounded text children");
    }
    let mut names = std::collections::BTreeSet::new();
    let mut child_markup = String::new();
    for child in &group.children {
        validate_token(&child.name, "drawing text-box name")?;
        if !names.insert(child.name.as_str())
            || [child.x_cm, child.y_cm, child.width_cm, child.height_cm]
                .into_iter()
                .any(|value| !value.is_finite())
            || child.width_cm <= 0.0
            || child.height_cm <= 0.0
        {
            return invalid("ODS drawing text-box identity or geometry is invalid");
        }
        write!(
            child_markup,
            "<draw:frame draw:name=\"{}\" draw:z-index=\"{}\" svg:x=\"{}cm\" svg:y=\"{}cm\" svg:width=\"{}cm\" svg:height=\"{}cm\"><draw:text-box>{}</draw:text-box></draw:frame>",
            escape_xml(&child.name),
            child.z_index,
            child.x_cm,
            child.y_cm,
            child.width_cm,
            child.height_cm,
            child.text.markup()?
        )
        .map_err(|_error| invalid_error("ODS drawing text formatting failed"))?;
    }
    let markup = format!(
        "<draw:g xmlns:draw=\"{DRAW}\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"{XLINK}\" draw:name=\"{}\">{child_markup}</draw:g>",
        escape_xml(&group.name)
    );
    insert_shape_child(source, sheet, &group.name, markup, max_output)
}

pub(crate) fn put_drawing_geometry(
    source: &[u8],
    sheet: &str,
    geometry: &DrawingGeometry,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_token(&geometry.name, "drawing geometry name")?;
    if [
        geometry.x_cm,
        geometry.y_cm,
        geometry.width_cm,
        geometry.height_cm,
    ]
    .into_iter()
    .any(|value| !value.is_finite())
        || geometry.width_cm <= 0.0
        || geometry.height_cm <= 0.0
    {
        return invalid("ODS drawing geometry is invalid");
    }
    let text = geometry
        .text
        .as_ref()
        .map(RichText::markup)
        .transpose()?
        .unwrap_or_default();
    let common = format!(
        "xmlns:draw=\"{DRAW}\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"{XLINK}\" draw:name=\"{}\" draw:z-index=\"{}\"",
        escape_xml(&geometry.name),
        geometry.z_index
    );
    let markup = match geometry.kind {
        DrawingGeometryKind::Rectangle => format!(
            "<draw:rect {common} svg:x=\"{}cm\" svg:y=\"{}cm\" svg:width=\"{}cm\" svg:height=\"{}cm\">{text}</draw:rect>",
            geometry.x_cm, geometry.y_cm, geometry.width_cm, geometry.height_cm
        ),
        DrawingGeometryKind::Ellipse => format!(
            "<draw:ellipse {common} svg:x=\"{}cm\" svg:y=\"{}cm\" svg:width=\"{}cm\" svg:height=\"{}cm\">{text}</draw:ellipse>",
            geometry.x_cm, geometry.y_cm, geometry.width_cm, geometry.height_cm
        ),
        DrawingGeometryKind::Line => {
            let x2 = geometry.x_cm + geometry.width_cm;
            let y2 = geometry.y_cm + geometry.height_cm;
            if !x2.is_finite() || !y2.is_finite() {
                return invalid("ODS drawing line endpoint is invalid");
            }
            format!(
                "<draw:line {common} svg:x1=\"{}cm\" svg:y1=\"{}cm\" svg:x2=\"{x2}cm\" svg:y2=\"{y2}cm\">{text}</draw:line>",
                geometry.x_cm, geometry.y_cm
            )
        },
        DrawingGeometryKind::Connector => {
            let x2 = geometry.x_cm + geometry.width_cm;
            let y2 = geometry.y_cm + geometry.height_cm;
            if !x2.is_finite() || !y2.is_finite() {
                return invalid("ODS drawing connector endpoint is invalid");
            }
            format!(
                "<draw:connector {common} draw:type=\"standard\" svg:x1=\"{}cm\" svg:y1=\"{}cm\" svg:x2=\"{x2}cm\" svg:y2=\"{y2}cm\">{text}</draw:connector>",
                geometry.x_cm, geometry.y_cm
            )
        },
    };
    insert_shape_child(source, sheet, &geometry.name, markup, max_output)
}

pub(crate) fn put_drawing_polygon(
    source: &[u8],
    sheet: &str,
    polygon: &DrawingPolygon,
    max_output: usize,
) -> Result<Vec<u8>> {
    validate_token(&polygon.name, "drawing polygon name")?;
    let minimum = if polygon.closed { 3 } else { 2 };
    if polygon.points.len() < minimum
        || polygon.points.len() > 4_096
        || polygon.view_box_width == 0
        || polygon.view_box_height == 0
        || polygon
            .points
            .iter()
            .any(|point| point.x > polygon.view_box_width || point.y > polygon.view_box_height)
        || [
            polygon.x_cm,
            polygon.y_cm,
            polygon.width_cm,
            polygon.height_cm,
        ]
        .into_iter()
        .any(|value| !value.is_finite())
        || polygon.width_cm <= 0.0
        || polygon.height_cm <= 0.0
    {
        return invalid("ODS drawing polygon geometry is invalid");
    }
    let mut points = String::new();
    for (index, point) in polygon.points.iter().enumerate() {
        if index != 0 {
            points.push(' ');
        }
        write!(points, "{},{}", point.x, point.y)
            .map_err(|_error| invalid_error("ODS drawing polygon formatting failed"))?;
    }
    let text = polygon
        .text
        .as_ref()
        .map(RichText::markup)
        .transpose()?
        .unwrap_or_default();
    let element = if polygon.closed {
        "polygon"
    } else {
        "polyline"
    };
    let markup = format!(
        "<draw:{element} xmlns:draw=\"{DRAW}\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"{XLINK}\" draw:name=\"{}\" draw:z-index=\"{}\" svg:x=\"{}cm\" svg:y=\"{}cm\" svg:width=\"{}cm\" svg:height=\"{}cm\" svg:viewBox=\"0 0 {} {}\" draw:points=\"{points}\">{text}</draw:{element}>",
        escape_xml(&polygon.name),
        polygon.z_index,
        polygon.x_cm,
        polygon.y_cm,
        polygon.width_cm,
        polygon.height_cm,
        polygon.view_box_width,
        polygon.view_box_height
    );
    insert_shape_child(source, sheet, &polygon.name, markup, max_output)
}

fn insert_shape_child(
    source: &[u8],
    sheet: &str,
    name: &str,
    child: String,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content_spans(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let shapes = children(&spans, table, TABLE, "shapes");
    if shapes.len() > 1 {
        return invalid("ODS sheet has more than one shapes container");
    }
    if let Some(shapes) = shapes.first().copied()
        && [
            "frame",
            "g",
            "rect",
            "ellipse",
            "line",
            "connector",
            "polygon",
            "polyline",
        ]
        .into_iter()
        .flat_map(|local| descendants(&spans, shapes, DRAW, local))
        .any(|index| {
            attribute(&xml, &spans[index], b"draw:name")
                .ok()
                .flatten()
                .as_deref()
                == Some(name)
        })
    {
        return invalid("ODS drawing name already exists");
    }
    let (range, markup) = if let Some(index) = shapes.first().copied() {
        (spans[index].close_start..spans[index].close_start, child)
    } else {
        let position = spans[table].close_start;
        (
            position..position,
            format!("<table:shapes xmlns:table=\"{TABLE}\">{child}</table:shapes>"),
        )
    };
    splice_content(source, range, markup.into_bytes(), max_output)
}

fn validate_frame(frame: &DrawingFrame) -> Result<()> {
    validate_token(&frame.name, "drawing frame name")?;
    super::document::validate_detached_resource_path(&frame.resource_path)?;
    if [frame.x_cm, frame.y_cm, frame.width_cm, frame.height_cm]
        .into_iter()
        .any(|value| !value.is_finite())
        || frame.width_cm <= 0.0
        || frame.height_cm <= 0.0
        || frame
            .anchor_cell
            .as_deref()
            .is_some_and(|anchor| anchor.is_empty() || anchor.len() > 65_536)
    {
        return invalid("ODS drawing frame geometry is invalid");
    }
    Ok(())
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

pub(crate) fn set_bound_forms(
    source: &[u8],
    controls: &[BoundFormControl],
    max_output: usize,
) -> Result<Vec<u8>> {
    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    let mut markup = String::new();
    if !controls.is_empty() {
        write!(
            markup,
            "<office:forms xmlns:office=\"{OFFICE}\" xmlns:form=\"{FORM}\" xmlns:xlink=\"{XLINK}\"><form:form form:name=\"LitchiForm\">"
        )
        .map_err(|_error| invalid_error("ODS form markup formatting failed"))?;
        let mut ids = std::collections::BTreeSet::new();
        for control in controls {
            validate_token(&control.id, "form control id")?;
            if !ids.insert(control.id.as_str())
                || control.label.len() > 65_536
                || control.label.chars().any(char::is_control)
            {
                return invalid("ODS bound form control is invalid or duplicated");
            }
            if let Some(path) = &control.image_path {
                super::document::validate_detached_resource_path(path)?;
            }
            if let Some(address) = &control.linked_cell {
                validate_form_address(&spreadsheet, address, true)?;
            }
            if let Some(address) = &control.source_range {
                validate_form_address(&spreadsheet, address, false)?;
            }
            markup.push_str("<form:button form:id=\"");
            markup.push_str(&escape_xml(&control.id));
            markup.push_str("\" form:label=\"");
            markup.push_str(&escape_xml(&control.label));
            markup.push('"');
            push_optional_attribute(
                &mut markup,
                "form:linked-cell",
                control.linked_cell.as_deref(),
            );
            push_optional_attribute(
                &mut markup,
                "form:source-cell-range",
                control.source_range.as_deref(),
            );
            push_optional_attribute(
                &mut markup,
                "form:image-data",
                control.image_path.as_deref(),
            );
            markup.push_str("/>");
        }
        markup.push_str("</form:form></office:forms>");
    }
    set_forms_markup(source, markup, max_output)
}

pub(crate) fn set_rich_forms(
    source: &[u8],
    controls: &[RichFormControl],
    max_output: usize,
) -> Result<Vec<u8>> {
    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    if controls.len() > 65_536 {
        return invalid("ODS rich form control count exceeds the limit");
    }
    let mut markup = String::new();
    if !controls.is_empty() {
        write!(
            markup,
            "<office:forms xmlns:office=\"{OFFICE}\" xmlns:form=\"{FORM}\" xmlns:script=\"{SCRIPT}\" xmlns:dom=\"{DOM}\" xmlns:xlink=\"{XLINK}\"><form:form form:name=\"LitchiForm\">"
        )
        .map_err(|_error| invalid_error("ODS rich form markup formatting failed"))?;
        let mut ids = std::collections::BTreeSet::new();
        for control in controls {
            validate_token(&control.id, "form control id")?;
            if !ids.insert(control.id.as_str())
                || control.label.len() > 65_536
                || control.label.chars().any(char::is_control)
                || control.events.len() > 1_024
            {
                return invalid("ODS rich form control is invalid or duplicated");
            }
            if let Some(path) = &control.image_path {
                super::document::validate_detached_resource_path(path)?;
            }
            if let Some(address) = &control.linked_cell {
                validate_form_address(&spreadsheet, address, true)?;
            }
            if let Some(address) = &control.source_range {
                validate_form_address(&spreadsheet, address, false)?;
            }
            let element = match control.kind {
                FormControlKind::Button => "button",
                FormControlKind::CheckBox => "checkbox",
                FormControlKind::Radio => "radio",
                FormControlKind::ListBox => "listbox",
                FormControlKind::ComboBox => "combobox",
                FormControlKind::Text => "text",
                FormControlKind::Image => "image",
                FormControlKind::Date => "date",
                FormControlKind::Time => "time",
                FormControlKind::FixedText => "fixed-text",
                FormControlKind::FormattedText => "formatted-text",
                FormControlKind::Number => "number",
                FormControlKind::File => "file",
                FormControlKind::Password => "password",
                FormControlKind::TextArea => "textarea",
                FormControlKind::Hidden => "hidden",
                FormControlKind::ValueRange => "value-range",
                FormControlKind::Frame => "frame",
                FormControlKind::ImageFrame => "image-frame",
                FormControlKind::GenericControl => "generic-control",
            };
            write!(
                markup,
                "<form:{element} form:id=\"{}\" form:label=\"{}\"",
                escape_xml(&control.id),
                escape_xml(&control.label)
            )
            .map_err(|_error| invalid_error("ODS rich form formatting failed"))?;
            push_optional_attribute(
                &mut markup,
                "form:linked-cell",
                control.linked_cell.as_deref(),
            );
            push_optional_attribute(
                &mut markup,
                "form:source-cell-range",
                control.source_range.as_deref(),
            );
            push_optional_attribute(
                &mut markup,
                "form:image-data",
                control.image_path.as_deref(),
            );
            if control.events.is_empty() {
                markup.push_str("/>");
                continue;
            }
            markup.push_str("><office:event-listeners>");
            let mut events = std::collections::BTreeSet::new();
            for event in &control.events {
                validate_token(&event.event_name, "form event name")?;
                validate_token(&event.macro_name, "form macro name")?;
                if !events.insert(event.event_name.as_str()) {
                    return invalid("ODS form control contains a duplicate event");
                }
                write!(
                    markup,
                    "<script:event-listener script:event-name=\"{}\" script:macro-name=\"{}\"/>",
                    escape_xml(&event.event_name),
                    escape_xml(&event.macro_name)
                )
                .map_err(|_error| invalid_error("ODS form event formatting failed"))?;
            }
            write!(markup, "</office:event-listeners></form:{element}>")
                .map_err(|_error| invalid_error("ODS rich form formatting failed"))?;
        }
        markup.push_str("</form:form></office:forms>");
    }
    set_forms_markup(source, markup, max_output)
}

fn validate_form_address(
    spreadsheet: &crate::Spreadsheet,
    value: &str,
    require_single_cell: bool,
) -> Result<()> {
    let address = crate::model::data_pilot::parse_data_pilot_range(value)?;
    if require_single_cell
        && (address.start_column != address.end_column || address.start_row != address.end_row)
    {
        return invalid("ODS form linked-cell binding must identify one cell");
    }
    if !address.sheet.is_empty() && spreadsheet.sheet(&address.sheet).is_none() {
        return invalid("ODS form binding references an unknown sheet");
    }
    Ok(())
}

fn set_forms_markup(source: &[u8], markup: String, max_output: usize) -> Result<Vec<u8>> {
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

fn splice_content_edits(
    source: &[u8],
    edits: Vec<(Range<usize>, Vec<u8>)>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let package = Package::from_bytes(source.to_vec())?;
    let part = XmlSourcePart::load(package.package(), CONTENT_PATH)?;
    let mut publication = XmlSplicePublication::new(part.clone());
    for (range, replacement) in edits {
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
        publication.replace(proof, fragment)?;
    }
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
