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
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
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
const MAX_LOGICAL_ROW_EDITS: usize = 4_096;
const MAX_SHEET_MOVE_SHEETS: usize = 1_024;
const MAX_SHEET_COPY_NAME_BYTES: usize = 1_024;
const MAX_SHEET_COPY_EVENTS: usize = 4_194_304;
const MAX_SHEET_COPY_DEPTH: usize = 256;
const MAX_SHEET_TRANSFER_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_SHEET_TRANSFER_ROWS: usize = 262_144;
const MAX_SHEET_TRANSFER_CELLS: usize = 4_194_304;
const MAX_SHEET_TRANSFER_REPETITION: usize = 1_048_576;
const MAX_SHEET_TRANSFER_NAMESPACES: usize = 256;
const MAX_SHEET_TRANSFER_DEPTH: usize = 256;

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

/// One logical-row structural operation evaluated against the result of the
/// preceding operation in the same batch.
///
/// Row positions are zero based. A move destination is expressed in the grid
/// after the moved range is removed, which makes `at == to` an exact no-op.
#[derive(Clone, Debug, PartialEq)]
#[non_exhaustive]
pub enum LogicalRowEdit {
    /// Insert the supplied physical row runs before one logical row position.
    Insert {
        /// Logical insertion position, including one-past-the-end.
        at: usize,
        /// Bounded physical row runs to insert atomically.
        rows: Vec<Row>,
    },
    /// Remove a non-empty half-open logical row range.
    Remove {
        /// First logical row to remove.
        at: usize,
        /// Number of logical rows to remove.
        count: usize,
    },
    /// Move a non-empty logical row range before another logical position.
    Move {
        /// First logical row to move.
        at: usize,
        /// Number of logical rows to move.
        count: usize,
        /// Destination in the grid after removal, including one-past-the-end.
        to: usize,
    },
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
    if resolved_attribute(&xml, &spans, index, TABLE, "number-rows-repeated")?
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

#[derive(Clone, Debug, PartialEq)]
struct PlannedRow {
    row: Row,
    origin: Option<usize>,
}

#[derive(Clone, Debug)]
struct SourceRowLexical {
    span: Range<usize>,
    original_repeat: usize,
    repeat_qname: Option<String>,
    namespaces: std::collections::BTreeMap<String, String>,
    declared_prefixes: std::collections::BTreeSet<String>,
}

/// Apply one bounded atomic batch of logical row edits to an ordinary,
/// dependency-free worksheet.
///
/// This deliberately narrow closure refuses any coordinate-bearing or opaque
/// spreadsheet owner rather than leaving formulas, ranges, merges, annotations,
/// drawings, validation bindings, or tracked changes stale. Repeated row runs
/// are split and compacted without expanding them into logical-row objects.
pub(crate) fn edit_logical_rows(
    source: &[u8],
    sheet: &str,
    edits: &[LogicalRowEdit],
    max_output: usize,
) -> Result<Vec<u8>> {
    if edits.len() > MAX_LOGICAL_ROW_EDITS {
        return invalid(format!(
            "ODS logical row edit count exceeds the {MAX_LOGICAL_ROW_EDITS} limit"
        ));
    }
    if edits.is_empty() {
        return Ok(source.to_vec());
    }
    if source.len() > max_output {
        return invalid(format!(
            "ODS source exceeds the {max_output} byte logical-row output limit"
        ));
    }

    let spreadsheet = crate::Spreadsheet::from_bytes(source.to_vec())?;
    let sheet_model = spreadsheet
        .sheet(sheet)
        .ok_or_else(|| invalid_error(format!("ODS sheet '{sheet}' was not found")))?;
    if validate_trivial_logical_row_edits(edits, sheet_model.logical_row_count())? {
        return Ok(source.to_vec());
    }

    let package = Package::from_bytes(source.to_vec())?;
    audit_logical_row_package(&package)?;
    let xml = package.content_xml().to_string();
    let spans = scan(&xml)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    audit_ordinary_row_content(&xml, sheet)?;
    audit_table_row_layout(&spans, table)?;
    let row_spans = children(&spans, table, TABLE, "table-row");
    if row_spans.len() != sheet_model.rows.len() {
        return invalid("ODS ordinary-row physical source inventory is inconsistent");
    }

    let mut source_rows = Vec::new();
    source_rows
        .try_reserve_exact(row_spans.len())
        .map_err(|_error| invalid_error("ODS source-row inventory allocation failed"))?;
    let mut planned = Vec::new();
    planned
        .try_reserve_exact(row_spans.len())
        .map_err(|_error| invalid_error("ODS logical-row plan allocation failed"))?;
    for (origin, (span_index, row)) in row_spans.iter().copied().zip(&sheet_model.rows).enumerate()
    {
        audit_inserted_row(row)?;
        source_rows.push(source_row_lexical(&xml, &spans, span_index, row.repeat())?);
        planned.push(PlannedRow {
            row: row.clone(),
            origin: Some(origin),
        });
    }
    let original = planned.clone();

    for edit in edits {
        apply_logical_row_edit(&mut planned, edit)?;
    }
    compact_planned_rows(&mut planned)?;
    validate_planned_rows(&planned)?;
    if planned == original {
        return Ok(source.to_vec());
    }

    let prefix = original
        .iter()
        .zip(&planned)
        .take_while(|(left, right)| left == right)
        .count();
    let suffix = original[prefix..]
        .iter()
        .rev()
        .zip(planned[prefix..].iter().rev())
        .take_while(|(left, right)| left == right)
        .count();
    let old_end = original.len() - suffix;
    let new_end = planned.len() - suffix;

    if row_spans.is_empty() {
        return replace_empty_table(
            source,
            &xml,
            &spans[table],
            sheet,
            &planned,
            &source_rows,
            max_output,
        );
    }

    let range = if prefix == old_end {
        let position = row_spans.get(prefix).map_or(
            spans[*row_spans.last().ok_or_else(|| {
                invalid_error("ODS source row inventory unexpectedly became empty")
            })?]
            .end,
            |index| spans[*index].start,
        );
        position..position
    } else {
        spans[row_spans[prefix]].start..spans[row_spans[old_end - 1]].end
    };
    let replacement =
        render_planned_rows(&xml, &planned[prefix..new_end], &source_rows, max_output)?;
    splice_content(source, range, replacement, max_output)
}

fn validate_trivial_logical_row_edits(
    edits: &[LogicalRowEdit],
    logical_count: usize,
) -> Result<bool> {
    for edit in edits {
        match edit {
            LogicalRowEdit::Insert { at, rows } if rows.is_empty() => {
                if *at > logical_count {
                    return invalid("ODS logical row insertion position is out of bounds");
                }
            },
            LogicalRowEdit::Remove { at, count: 0 } => {
                if *at > logical_count {
                    return invalid("ODS logical row removal position is out of bounds");
                }
            },
            LogicalRowEdit::Move { at, count: 0, to } => {
                if *at > logical_count || *to > logical_count {
                    return invalid("ODS logical row move position is out of bounds");
                }
            },
            LogicalRowEdit::Move { at, count, to } if at == to => {
                let end = at
                    .checked_add(*count)
                    .ok_or_else(|| invalid_error("ODS logical row move range overflows"))?;
                if end > logical_count || *to > logical_count - count {
                    return invalid("ODS logical row move range is out of bounds");
                }
            },
            _ => return Ok(false),
        }
    }
    Ok(true)
}

fn audit_table_row_layout(spans: &[Span], table: usize) -> Result<()> {
    let mut seen_row = false;
    for span in spans.iter().filter(|span| span.parent == Some(table)) {
        if is_element(span, TABLE, "table-row") {
            seen_row = true;
        } else if is_element(span, TABLE, "table-column") {
            if seen_row {
                return invalid(
                    "ODS ordinary logical-row edits require columns to precede every row",
                );
            }
        } else {
            return invalid("ODS ordinary logical-row table contains an unsupported direct child");
        }
    }
    Ok(())
}

fn apply_logical_row_edit(rows: &mut Vec<PlannedRow>, edit: &LogicalRowEdit) -> Result<()> {
    let logical_count = planned_logical_count(rows)?;
    match edit {
        LogicalRowEdit::Insert { at, rows: inserted } => {
            if *at > logical_count {
                return invalid("ODS logical row insertion position is out of bounds");
            }
            if inserted.is_empty() {
                return Ok(());
            }
            for row in inserted {
                audit_inserted_row(row)?;
            }
            let inserted_count = inserted.iter().try_fold(0usize, |total, row| {
                total
                    .checked_add(row.repeat())
                    .ok_or_else(|| invalid_error("ODS inserted logical row count overflows"))
            })?;
            let final_count = logical_count
                .checked_add(inserted_count)
                .ok_or_else(|| invalid_error("ODS logical row count overflows"))?;
            if final_count > crate::worksheet::validation::MAX_LOGICAL_ROWS {
                return invalid("ODS logical row insertion exceeds the worksheet row limit");
            }
            let position = split_planned_boundary(rows, *at)?;
            rows.try_reserve(inserted.len())
                .map_err(|_error| invalid_error("ODS inserted-row plan allocation failed"))?;
            rows.splice(
                position..position,
                inserted
                    .iter()
                    .cloned()
                    .map(|row| PlannedRow { row, origin: None }),
            );
        },
        LogicalRowEdit::Remove { at, count } => {
            if *count == 0 {
                if *at > logical_count {
                    return invalid("ODS logical row removal position is out of bounds");
                }
                return Ok(());
            }
            let end = at
                .checked_add(*count)
                .ok_or_else(|| invalid_error("ODS logical row removal range overflows"))?;
            if end > logical_count {
                return invalid("ODS logical row removal range is out of bounds");
            }
            let start_index = split_planned_boundary(rows, *at)?;
            let end_index = split_planned_boundary(rows, end)?;
            rows.drain(start_index..end_index);
        },
        LogicalRowEdit::Move { at, count, to } => {
            if *count == 0 {
                if *at > logical_count || *to > logical_count {
                    return invalid("ODS logical row move position is out of bounds");
                }
                return Ok(());
            }
            let end = at
                .checked_add(*count)
                .ok_or_else(|| invalid_error("ODS logical row move range overflows"))?;
            if end > logical_count {
                return invalid("ODS logical row move range is out of bounds");
            }
            let remaining = logical_count - count;
            if *to > remaining {
                return invalid("ODS logical row move destination is out of bounds");
            }
            if *at == *to {
                return Ok(());
            }
            let start_index = split_planned_boundary(rows, *at)?;
            let end_index = split_planned_boundary(rows, end)?;
            let moved = rows.drain(start_index..end_index).collect::<Vec<_>>();
            let destination = split_planned_boundary(rows, *to)?;
            rows.try_reserve(moved.len())
                .map_err(|_error| invalid_error("ODS moved-row plan allocation failed"))?;
            rows.splice(destination..destination, moved);
        },
    }
    compact_planned_rows(rows)
}

fn planned_logical_count(rows: &[PlannedRow]) -> Result<usize> {
    rows.iter().try_fold(0usize, |total, row| {
        total
            .checked_add(row.row.repeat())
            .ok_or_else(|| invalid_error("ODS logical row count overflows"))
    })
}

fn split_planned_boundary(rows: &mut Vec<PlannedRow>, target: usize) -> Result<usize> {
    let mut start = 0usize;
    for index in 0..rows.len() {
        let count = rows[index].row.repeat();
        let end = start
            .checked_add(count)
            .ok_or_else(|| invalid_error("ODS logical row address overflows"))?;
        if target == start {
            return Ok(index);
        }
        if target < end {
            let offset = target - start;
            let suffix = count - offset;
            let origin = rows[index].origin;
            let right = PlannedRow {
                row: rows[index].row.with_repeat(suffix)?,
                origin,
            };
            rows[index].row = rows[index].row.with_repeat(offset)?;
            rows.try_reserve(1)
                .map_err(|_error| invalid_error("ODS repeated-row split allocation failed"))?;
            rows.insert(index + 1, right);
            return Ok(index + 1);
        }
        start = end;
    }
    if target == start {
        Ok(rows.len())
    } else {
        invalid("ODS logical row boundary is out of bounds")
    }
}

fn compact_planned_rows(rows: &mut Vec<PlannedRow>) -> Result<()> {
    let mut compacted = Vec::<PlannedRow>::new();
    compacted
        .try_reserve_exact(rows.len())
        .map_err(|_error| invalid_error("ODS logical-row compaction allocation failed"))?;
    for row in rows.drain(..) {
        if let Some(previous) = compacted.last_mut()
            && previous.origin == row.origin
            && previous.row.equivalent_run(&row.row)
        {
            let repeat = previous
                .row
                .repeat()
                .checked_add(row.row.repeat())
                .ok_or_else(|| invalid_error("ODS row repetition overflows"))?;
            previous.row = previous.row.with_repeat(repeat)?;
        } else {
            compacted.push(row);
        }
    }
    *rows = compacted;
    Ok(())
}

fn validate_planned_rows(rows: &[PlannedRow]) -> Result<()> {
    if rows.len() > crate::worksheet::validation::MAX_PHYSICAL_RUNS {
        return invalid("ODS logical row edit exceeds the physical row-run limit");
    }
    let mut planned_rows = Vec::new();
    planned_rows
        .try_reserve_exact(rows.len())
        .map_err(|_error| invalid_error("ODS planned-row validation allocation failed"))?;
    planned_rows.extend(rows.iter().map(|row| row.row.clone()));
    let sheet = Sheet {
        name: "litchi-row-plan".to_string(),
        rows: planned_rows,
        style_name: None,
    };
    crate::worksheet::validation::validate_sheet(&sheet)
}

fn audit_inserted_row(row: &Row) -> Result<()> {
    let sheet = Sheet {
        name: "litchi-inserted-row".to_string(),
        rows: vec![row.clone()],
        style_name: None,
    };
    crate::worksheet::validation::validate_sheet(&sheet)?;
    if row.style_name.is_some()
        || row.default_cell_style_name.is_some()
        || row.cells.iter().any(|cell| {
            cell.style_name.is_some()
                || cell.formula.is_some()
                || !matches!(cell.merge, crate::Merge::None)
                || matches!(cell.value, CellValue::Unknown { .. })
        })
    {
        return invalid(
            "ODS ordinary logical-row edits refuse style references, formulas, merges, and unknown cell values",
        );
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OrdinaryElement {
    Root,
    Body,
    Spreadsheet,
    Table,
    Column,
    Row,
    Cell,
    Paragraph,
}

fn audit_logical_row_package(package: &Package) -> Result<()> {
    for path in package.package().files()? {
        let lower = path.to_ascii_lowercase();
        if matches!(
            path.as_str(),
            "mimetype"
                | "content.xml"
                | "meta.xml"
                | "manifest.rdf"
                | "META-INF/manifest.xml"
                | "META-INF/documentsignatures.xml"
                | "META-INF/macrosignatures.xml"
        ) || path.ends_with('/')
        {
            continue;
        }
        if lower == "styles.xml"
            || lower == "settings.xml"
            || lower == "scripts.xml"
            || lower.ends_with(".xml")
            || lower.ends_with(".rdf")
            || lower.starts_with("basic/")
            || lower.starts_with("scripts/")
            || lower.starts_with("configurations2/")
            || lower.starts_with("object")
        {
            return invalid(format!(
                "ODS ordinary logical-row edits refuse dependent package member '{path}'"
            ));
        }
    }
    Ok(())
}

fn audit_ordinary_row_content(xml: &str, selected_sheet: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<(OrdinaryElement, usize)>::new();
    let mut selected = 0usize;
    let mut roots = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS content XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let kind = ordinary_element(&namespace, element.local_name().as_ref())?;
                audit_ordinary_parent(stack.last().map(|entry| entry.0), kind)?;
                let name = audit_ordinary_attributes(&element, &reader, kind)?;
                if kind == OrdinaryElement::Root {
                    roots += 1;
                }
                if kind == OrdinaryElement::Table && name.as_deref() == Some(selected_sheet) {
                    selected += 1;
                }
                if kind == OrdinaryElement::Paragraph
                    && let Some((OrdinaryElement::Cell, paragraphs)) = stack.last_mut()
                {
                    *paragraphs += 1;
                    if *paragraphs > 1 {
                        return invalid(
                            "ODS ordinary logical-row edits refuse multi-paragraph cells",
                        );
                    }
                }
                stack.push((kind, 0));
                if stack.len() > 1_024 {
                    return invalid("ODS ordinary logical-row XML depth exceeds the limit");
                }
            },
            Event::Empty(element) => {
                let kind = ordinary_element(&namespace, element.local_name().as_ref())?;
                audit_ordinary_parent(stack.last().map(|entry| entry.0), kind)?;
                let name = audit_ordinary_attributes(&element, &reader, kind)?;
                if kind == OrdinaryElement::Root {
                    roots += 1;
                }
                if kind == OrdinaryElement::Table && name.as_deref() == Some(selected_sheet) {
                    selected += 1;
                }
                if kind == OrdinaryElement::Paragraph
                    && let Some((OrdinaryElement::Cell, paragraphs)) = stack.last_mut()
                {
                    *paragraphs += 1;
                    if *paragraphs > 1 {
                        return invalid(
                            "ODS ordinary logical-row edits refuse multi-paragraph cells",
                        );
                    }
                }
            },
            Event::End(_) => {
                stack.pop().ok_or_else(|| {
                    invalid_error("ODS ordinary logical-row element stack underflow")
                })?;
            },
            Event::Text(text) => {
                let text_bytes: &[u8] = text.as_ref();
                if stack.last().map(|entry| entry.0) != Some(OrdinaryElement::Paragraph)
                    && !text_bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return invalid(
                        "ODS ordinary logical-row edits refuse text outside plain paragraphs",
                    );
                }
            },
            Event::GeneralRef(reference) => {
                let reference_bytes: &[u8] = reference.as_ref();
                let predefined =
                    [b"amp".as_slice(), b"lt", b"gt", b"apos", b"quot"].contains(&reference_bytes);
                if !predefined
                    || stack.last().map(|entry| entry.0) != Some(OrdinaryElement::Paragraph)
                {
                    return invalid(
                        "ODS ordinary logical-row edits refuse general entity references",
                    );
                }
            },
            Event::Decl(_) => {},
            Event::Eof => break,
            Event::DocType(_) => {
                return invalid("ODS ordinary logical-row edits refuse document type declarations");
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) => {
                return invalid(
                    "ODS ordinary logical-row edits refuse comments, processing instructions, and CDATA",
                );
            },
        }
        buffer.clear();
    }
    if !stack.is_empty() || roots != 1 {
        return invalid("ODS ordinary logical-row content hierarchy is incomplete");
    }
    if selected != 1 {
        return invalid("ODS ordinary logical-row sheet selector is missing or ambiguous");
    }
    Ok(())
}

fn ordinary_element(namespace: &ResolveResult<'_>, local: &[u8]) -> Result<OrdinaryElement> {
    let uri = match namespace {
        ResolveResult::Bound(Namespace(uri)) => uri,
        ResolveResult::Unbound => {
            return invalid("ODS ordinary logical-row edits refuse unqualified elements");
        },
        ResolveResult::Unknown(prefix) => {
            return invalid(format!(
                "ODS ordinary logical-row edits refuse unbound prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ));
        },
    };
    match (uri, local) {
        (value, b"document-content") if *value == OFFICE.as_bytes() => Ok(OrdinaryElement::Root),
        (value, b"body") if *value == OFFICE.as_bytes() => Ok(OrdinaryElement::Body),
        (value, b"spreadsheet") if *value == OFFICE.as_bytes() => Ok(OrdinaryElement::Spreadsheet),
        (value, b"table") if *value == TABLE.as_bytes() => Ok(OrdinaryElement::Table),
        (value, b"table-column") if *value == TABLE.as_bytes() => Ok(OrdinaryElement::Column),
        (value, b"table-row") if *value == TABLE.as_bytes() => Ok(OrdinaryElement::Row),
        (value, b"table-cell") if *value == TABLE.as_bytes() => Ok(OrdinaryElement::Cell),
        (value, b"p") if *value == crate::worksheet::codec::TEXT_NAMESPACE.as_bytes() => {
            Ok(OrdinaryElement::Paragraph)
        },
        _ => invalid(format!(
            "ODS ordinary logical-row edits refuse element '{}':{}",
            String::from_utf8_lossy(uri),
            String::from_utf8_lossy(local)
        )),
    }
}

fn audit_ordinary_parent(parent: Option<OrdinaryElement>, child: OrdinaryElement) -> Result<()> {
    let valid = matches!(
        (parent, child),
        (None, OrdinaryElement::Root)
            | (Some(OrdinaryElement::Root), OrdinaryElement::Body)
            | (Some(OrdinaryElement::Body), OrdinaryElement::Spreadsheet)
            | (Some(OrdinaryElement::Spreadsheet), OrdinaryElement::Table)
            | (Some(OrdinaryElement::Table), OrdinaryElement::Column)
            | (Some(OrdinaryElement::Table), OrdinaryElement::Row)
            | (Some(OrdinaryElement::Row), OrdinaryElement::Cell)
            | (Some(OrdinaryElement::Cell), OrdinaryElement::Paragraph)
    );
    if valid {
        Ok(())
    } else {
        invalid("ODS ordinary logical-row edits refuse this element hierarchy")
    }
}

fn audit_ordinary_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    kind: OrdinaryElement,
) -> Result<Option<String>> {
    let mut table_name = None;
    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|error| invalid_error(format!("invalid ODS attribute: {error}")))?;
        let qname = raw.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let uri = match namespace {
            ResolveResult::Bound(Namespace(uri)) => uri,
            ResolveResult::Unbound => b"".as_slice(),
            ResolveResult::Unknown(prefix) => {
                return invalid(format!(
                    "ODS ordinary logical-row edits refuse unbound attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                ));
            },
        };
        let local = local.as_ref();
        let allowed = match kind {
            OrdinaryElement::Root => uri == OFFICE.as_bytes() && local == b"version",
            OrdinaryElement::Body | OrdinaryElement::Spreadsheet => false,
            OrdinaryElement::Table => uri == TABLE.as_bytes() && local == b"name",
            OrdinaryElement::Column => {
                uri == TABLE.as_bytes()
                    && matches!(local, b"number-columns-repeated" | b"visibility")
            },
            OrdinaryElement::Row => {
                uri == TABLE.as_bytes() && matches!(local, b"number-rows-repeated" | b"visibility")
            },
            OrdinaryElement::Cell => {
                uri == TABLE.as_bytes() && local == b"number-columns-repeated"
                    || uri == OFFICE.as_bytes()
                        && matches!(
                            local,
                            b"value-type"
                                | b"value"
                                | b"date-value"
                                | b"time-value"
                                | b"boolean-value"
                                | b"currency"
                        )
            },
            OrdinaryElement::Paragraph => {
                uri == b"http://www.w3.org/XML/1998/namespace" && local == b"space"
            },
        };
        if !allowed {
            return invalid(format!(
                "ODS ordinary logical-row edits refuse attribute '{}'",
                String::from_utf8_lossy(qname)
            ));
        }
        if kind == OrdinaryElement::Table && uri == TABLE.as_bytes() && local == b"name" {
            let value = raw
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid ODS table name: {error}")))?
                .into_owned();
            if table_name.replace(value).is_some() {
                return invalid("ODS ordinary logical-row table name is duplicated");
            }
        }
        if matches!(kind, OrdinaryElement::Row | OrdinaryElement::Column)
            && uri == TABLE.as_bytes()
            && local == b"visibility"
        {
            let value = raw
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid ODS visibility: {error}")))?;
            let _visibility = crate::model::structure::Visibility::parse(&value)?;
        }
        if kind == OrdinaryElement::Paragraph
            && uri == b"http://www.w3.org/XML/1998/namespace"
            && local == b"space"
        {
            let value = raw
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid ODS xml:space: {error}")))?;
            if !matches!(value.as_ref(), "default" | "preserve") {
                return invalid("ODS ordinary logical-row paragraph xml:space is invalid");
            }
        }
    }
    if kind == OrdinaryElement::Table && table_name.is_none() {
        return invalid("ODS ordinary logical-row table name is missing");
    }
    Ok(table_name)
}

fn source_row_lexical(
    xml: &str,
    spans: &[Span],
    row_index: usize,
    original_repeat: usize,
) -> Result<SourceRowLexical> {
    let span = spans
        .get(row_index)
        .ok_or_else(|| invalid_error("ODS source-row span is missing"))?;
    let mut ancestors = Vec::new();
    let mut parent = span.parent;
    while let Some(index) = parent {
        let ancestor = spans
            .get(index)
            .ok_or_else(|| invalid_error("ODS source-row ancestor span is missing"))?;
        ancestors.push(ancestor);
        parent = ancestor.parent;
    }
    let mut wrapped = String::new();
    for ancestor in ancestors.iter().rev() {
        bounded_append(
            &mut wrapped,
            xml.get(ancestor.start..ancestor.tag_end)
                .ok_or_else(|| invalid_error("ODS source-row ancestor tag is invalid"))?,
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
    }
    bounded_append(
        &mut wrapped,
        xml.get(span.start..span.end)
            .ok_or_else(|| invalid_error("ODS source-row range is invalid"))?,
        crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
    )?;
    for ancestor in &ancestors {
        bounded_append(
            &mut wrapped,
            "</",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            start_qname(xml, ancestor)?,
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            ">",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
    }

    let mut reader = NsReader::from_str(&wrapped);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut inside = false;
    let mut depth = 0usize;
    let mut namespaces = std::collections::BTreeMap::new();
    let mut declared_prefixes = std::collections::BTreeSet::new();
    let mut repeat_qname = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS source-row XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let empty = false;
                let is_row = matches!(
                    &namespace,
                    ResolveResult::Bound(Namespace(uri)) if *uri == TABLE.as_bytes()
                ) && element.local_name().as_ref() == b"table-row";
                if !inside && is_row {
                    inside = true;
                    depth = 0;
                }
                if inside {
                    register_element_namespace(
                        &mut namespaces,
                        element.name().as_ref(),
                        &namespace,
                    )?;
                    for raw in element.attributes().with_checks(true) {
                        let raw = raw.map_err(|error| {
                            invalid_error(format!("invalid ODS source-row attribute: {error}"))
                        })?;
                        let qname = raw.key.as_ref();
                        if depth == 0 && (qname == b"xmlns" || qname.starts_with(b"xmlns:")) {
                            let prefix = qname
                                .strip_prefix(b"xmlns:")
                                .map_or(b"".as_slice(), |value| value);
                            declared_prefixes.insert(decode(prefix, "namespace prefix")?);
                            continue;
                        }
                        let (attribute_namespace, local) =
                            reader.resolver().resolve_attribute(raw.key);
                        register_attribute_namespace(&mut namespaces, qname, &attribute_namespace)?;
                        if depth == 0
                            && matches!(
                                &attribute_namespace,
                                ResolveResult::Bound(Namespace(uri)) if *uri == TABLE.as_bytes()
                            )
                            && local.as_ref() == b"number-rows-repeated"
                        {
                            repeat_qname = Some(decode(qname, "row repeat attribute name")?);
                        }
                    }
                    if !empty {
                        depth = depth.checked_add(1).ok_or_else(|| {
                            invalid_error("ODS source-row namespace depth overflows")
                        })?;
                    }
                }
            },
            Event::Empty(element) => {
                let empty = true;
                let is_row = matches!(
                    &namespace,
                    ResolveResult::Bound(Namespace(uri)) if *uri == TABLE.as_bytes()
                ) && element.local_name().as_ref() == b"table-row";
                if !inside && is_row {
                    inside = true;
                    depth = 0;
                }
                if inside {
                    register_element_namespace(
                        &mut namespaces,
                        element.name().as_ref(),
                        &namespace,
                    )?;
                    for raw in element.attributes().with_checks(true) {
                        let raw = raw.map_err(|error| {
                            invalid_error(format!("invalid ODS source-row attribute: {error}"))
                        })?;
                        let qname = raw.key.as_ref();
                        if depth == 0 && (qname == b"xmlns" || qname.starts_with(b"xmlns:")) {
                            let prefix = qname
                                .strip_prefix(b"xmlns:")
                                .map_or(b"".as_slice(), |value| value);
                            declared_prefixes.insert(decode(prefix, "namespace prefix")?);
                            continue;
                        }
                        let (attribute_namespace, local) =
                            reader.resolver().resolve_attribute(raw.key);
                        register_attribute_namespace(&mut namespaces, qname, &attribute_namespace)?;
                        if depth == 0
                            && matches!(
                                &attribute_namespace,
                                ResolveResult::Bound(Namespace(uri)) if *uri == TABLE.as_bytes()
                            )
                            && local.as_ref() == b"number-rows-repeated"
                        {
                            repeat_qname = Some(decode(qname, "row repeat attribute name")?);
                        }
                    }
                    if empty && depth == 0 {
                        break;
                    }
                    if !empty {
                        depth = depth.checked_add(1).ok_or_else(|| {
                            invalid_error("ODS source-row namespace depth overflows")
                        })?;
                    }
                }
            },
            Event::End(_) if inside => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODS source-row namespace depth underflows"))?;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => {
                return invalid("ODS source-row namespace inventory ended early");
            },
            _ => {},
        }
        buffer.clear();
    }
    Ok(SourceRowLexical {
        span: span.start..span.end,
        original_repeat,
        repeat_qname,
        namespaces,
        declared_prefixes,
    })
}

fn register_element_namespace(
    namespaces: &mut std::collections::BTreeMap<String, String>,
    qname: &[u8],
    namespace: &ResolveResult<'_>,
) -> Result<()> {
    if let ResolveResult::Bound(Namespace(uri)) = namespace {
        register_namespace(namespaces, qname, uri)?;
    }
    Ok(())
}

fn register_attribute_namespace(
    namespaces: &mut std::collections::BTreeMap<String, String>,
    qname: &[u8],
    namespace: &ResolveResult<'_>,
) -> Result<()> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => register_namespace(namespaces, qname, uri),
        ResolveResult::Unbound => Ok(()),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "ODS source-row attribute prefix '{}' is unbound",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn register_namespace(
    namespaces: &mut std::collections::BTreeMap<String, String>,
    qname: &[u8],
    uri: &[u8],
) -> Result<()> {
    let prefix = qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(b"".as_slice(), |index| &qname[..index]);
    if prefix == b"xml" {
        return Ok(());
    }
    let prefix = decode(prefix, "source-row namespace prefix")?;
    let uri = decode(uri, "source-row namespace URI")?;
    if let Some(previous) = namespaces.insert(prefix.clone(), uri.clone())
        && previous != uri
    {
        return invalid(format!(
            "ODS source-row prefix '{prefix}' resolves to multiple namespaces"
        ));
    }
    Ok(())
}

fn start_qname<'a>(xml: &'a str, span: &Span) -> Result<&'a str> {
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS element start tag is invalid"))?;
    let bytes = tag.as_bytes();
    let start = usize::from(bytes.first() == Some(&b'<'));
    let end = bytes[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))
        .map_or(bytes.len(), |offset| start + offset);
    tag.get(start..end)
        .ok_or_else(|| invalid_error("ODS element qualified name is invalid"))
}

fn render_planned_rows(
    xml: &str,
    rows: &[PlannedRow],
    source_rows: &[SourceRowLexical],
    max_output: usize,
) -> Result<Vec<u8>> {
    let mut output = String::new();
    for planned in rows {
        let markup = if let Some(origin) = planned.origin {
            render_source_row(
                xml,
                source_rows
                    .get(origin)
                    .ok_or_else(|| invalid_error("ODS source-row origin is missing"))?,
                planned.row.repeat(),
            )?
        } else {
            crate::worksheet::codec::write_rows_bounded(
                std::slice::from_ref(&planned.row),
                max_output,
            )?
        };
        bounded_append(&mut output, &markup, max_output)?;
    }
    Ok(output.into_bytes())
}

fn render_source_row(xml: &str, source: &SourceRowLexical, repeat: usize) -> Result<String> {
    let raw = xml
        .get(source.span.clone())
        .ok_or_else(|| invalid_error("ODS source-row bytes are missing"))?;
    let repeated = if repeat == source.original_repeat {
        raw.to_string()
    } else {
        let qname = source.repeat_qname.as_deref().ok_or_else(|| {
            invalid_error("ODS repeated source row has no resolved repetition attribute")
        })?;
        replace_attribute_value(raw, qname, &repeat.to_string())?
    };
    bind_fragment_namespaces(&repeated, &source.namespaces, &source.declared_prefixes)
}

fn bind_fragment_namespaces(
    markup: &str,
    namespaces: &std::collections::BTreeMap<String, String>,
    declared: &std::collections::BTreeSet<String>,
) -> Result<String> {
    let bytes = markup.as_bytes();
    if bytes.first() != Some(&b'<') {
        return invalid("ODS source-row fragment has no start tag");
    }
    let insertion = bytes[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))
        .map_or(bytes.len(), |offset| offset + 1);
    let extra = namespaces
        .iter()
        .filter(|(prefix, _)| !declared.contains(*prefix))
        .try_fold(0usize, |total, (prefix, uri)| {
            total
                .checked_add(10)
                .and_then(|value| value.checked_add(prefix.len()))
                .and_then(|value| value.checked_add(uri.len()))
                .ok_or_else(|| invalid_error("ODS namespace binding size overflows"))
        })?;
    let capacity = markup
        .len()
        .checked_add(extra)
        .ok_or_else(|| invalid_error("ODS source-row fragment size overflows"))?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_error| invalid_error("ODS source-row fragment allocation failed"))?;
    output.push_str(&markup[..insertion]);
    for (prefix, uri) in namespaces {
        if declared.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            output.push_str(" xmlns=\"");
        } else {
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
        }
        output.push_str(&escape_xml(uri));
        output.push('"');
    }
    output.push_str(&markup[insertion..]);
    Ok(output)
}

fn replace_attribute_value(markup: &str, qname: &str, replacement: &str) -> Result<String> {
    let bytes = markup.as_bytes();
    let tag_end = quote_aware_tag_end(bytes)?;
    let mut cursor = bytes[1..tag_end]
        .iter()
        .position(u8::is_ascii_whitespace)
        .map_or(tag_end, |offset| offset + 1);
    while cursor < tag_end {
        while cursor < tag_end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || bytes[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < tag_end
            && !bytes[cursor].is_ascii_whitespace()
            && !matches!(bytes[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag_end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || bytes[cursor] != b'=' {
            return invalid("ODS source-row attribute has no equals sign");
        }
        cursor += 1;
        while cursor < tag_end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid_error("ODS source-row attribute value is missing"))?;
        if !matches!(quote, b'\'' | b'"') {
            return invalid("ODS source-row attribute value is not quoted");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag_end && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag_end {
            return invalid("ODS source-row attribute quote is unclosed");
        }
        let value_end = cursor;
        cursor += 1;
        if &bytes[name_start..name_end] == qname.as_bytes() {
            let capacity = markup
                .len()
                .checked_sub(value_end - value_start)
                .and_then(|value| value.checked_add(replacement.len()))
                .ok_or_else(|| invalid_error("ODS source-row attribute size overflows"))?;
            let mut output = String::new();
            output
                .try_reserve_exact(capacity)
                .map_err(|_error| invalid_error("ODS source-row attribute allocation failed"))?;
            output.push_str(&markup[..value_start]);
            output.push_str(replacement);
            output.push_str(&markup[value_end..]);
            return Ok(output);
        }
    }
    invalid(format!("ODS source-row attribute '{qname}' was not found"))
}

fn quote_aware_tag_end(bytes: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in bytes.iter().copied().enumerate() {
        match (quote, byte) {
            (Some(delimiter), current) if current == delimiter => quote = None,
            (Some(_), _) => {},
            (None, b'\'' | b'"') => quote = Some(byte),
            (None, b'>') => return Ok(index),
            _ => {},
        }
    }
    invalid("ODS source-row start tag is incomplete")
}

fn bounded_append(output: &mut String, value: &str, max_output: usize) -> Result<()> {
    let next = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| invalid_error("ODS logical-row rendered size overflows"))?;
    if next > max_output {
        return invalid(format!(
            "ODS logical-row rendering exceeds the {max_output} byte limit"
        ));
    }
    output
        .try_reserve(value.len())
        .map_err(|_error| invalid_error("ODS logical-row rendering allocation failed"))?;
    output.push_str(value);
    Ok(())
}

fn replace_empty_table(
    source: &[u8],
    xml: &str,
    table: &Span,
    _sheet: &str,
    rows: &[PlannedRow],
    source_rows: &[SourceRowLexical],
    max_output: usize,
) -> Result<Vec<u8>> {
    let raw = xml
        .get(table.start..table.end)
        .ok_or_else(|| invalid_error("ODS empty table span is invalid"))?;
    let rendered = render_planned_rows(xml, rows, source_rows, max_output)?;
    let qname = start_qname(xml, table)?;
    let mut replacement = String::new();
    if let Some(opening) = raw.strip_suffix("/>") {
        bounded_append(&mut replacement, opening, max_output)?;
        bounded_append(&mut replacement, ">", max_output)?;
    } else {
        let opening = xml
            .get(table.start..table.tag_end)
            .ok_or_else(|| invalid_error("ODS empty table opening tag is invalid"))?;
        bounded_append(&mut replacement, opening, max_output)?;
        let retained_children = xml
            .get(table.tag_end..table.close_start)
            .ok_or_else(|| invalid_error("ODS empty table child span is invalid"))?;
        bounded_append(&mut replacement, retained_children, max_output)?;
    }
    bounded_append(
        &mut replacement,
        std::str::from_utf8(&rendered)
            .map_err(|_error| invalid_error("ODS rendered rows are not UTF-8"))?,
        max_output,
    )?;
    bounded_append(&mut replacement, "</", max_output)?;
    bounded_append(&mut replacement, qname, max_output)?;
    bounded_append(&mut replacement, ">", max_output)?;
    splice_content(
        source,
        table.start..table.end,
        replacement.into_bytes(),
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

/// Reorder one complete worksheet owner while retaining every source fragment byte.
///
/// This closure is deliberately narrower than general spreadsheet editing.  It
/// accepts only ordinary, independent sheets and refuses every known owner of
/// sheet names or ordinal coordinates before constructing the replacement.
pub(crate) fn move_sheet(
    source: &[u8],
    sheet: &str,
    final_position: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    if source.len() > max_output {
        return invalid(format!(
            "ODS source exceeds the {max_output} byte sheet-move output limit"
        ));
    }
    let package = Package::from_bytes(source.to_vec())?;
    let xml = package.content_xml().to_string();
    let spans = scan(&xml)?;
    let spreadsheet = one(&spans, OFFICE, "spreadsheet")?;
    let tables = children(&spans, spreadsheet, TABLE, "table");
    if tables.is_empty() || tables.len() > MAX_SHEET_MOVE_SHEETS {
        return invalid(format!(
            "ODS sheet move requires 1..={MAX_SHEET_MOVE_SHEETS} worksheets"
        ));
    }
    if final_position >= tables.len() {
        return invalid(format!(
            "ODS sheet move destination {final_position} exceeds sheet count {}",
            tables.len()
        ));
    }
    let selected = select_sheet(&xml, &spans, sheet)?;
    let from = tables
        .iter()
        .position(|candidate| *candidate == selected)
        .ok_or_else(|| invalid_error("ODS selected sheet is outside the spreadsheet owner"))?;
    if from == final_position {
        return Ok(source.to_vec());
    }

    audit_sheet_move_package(&package, &xml, &spans, spreadsheet)?;

    // The table slots remain at their exact lexical positions.  Only the
    // complete table fragments occupying those slots are permuted; whitespace
    // and all bytes between slots remain in their original positions.
    let first = spans[tables[0]].start;
    let last = spans[*tables
        .last()
        .ok_or_else(|| invalid_error("ODS sheet inventory unexpectedly became empty"))?]
    .end;
    let mut order = Vec::new();
    order
        .try_reserve_exact(tables.len())
        .map_err(|_error| invalid_error("ODS sheet-move order allocation failed"))?;
    order.extend(0..tables.len());
    let moved = order.remove(from);
    order.insert(final_position, moved);
    let replaced_len = last
        .checked_sub(first)
        .ok_or_else(|| invalid_error("ODS sheet-move span is invalid"))?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(replaced_len)
        .map_err(|_error| invalid_error("ODS sheet-move replacement allocation failed"))?;
    for (slot, origin) in order.into_iter().enumerate() {
        let origin_span = &spans[tables[origin]];
        replacement.extend_from_slice(
            xml.as_bytes()
                .get(origin_span.start..origin_span.end)
                .ok_or_else(|| invalid_error("ODS sheet source fragment is invalid"))?,
        );
        if slot + 1 < tables.len() {
            let gap_start = spans[tables[slot]].end;
            let gap_end = spans[tables[slot + 1]].start;
            replacement.extend_from_slice(
                xml.as_bytes()
                    .get(gap_start..gap_end)
                    .ok_or_else(|| invalid_error("ODS sheet interstitial span is invalid"))?,
            );
        }
    }
    if replacement.len() != replaced_len {
        return invalid("ODS sheet move changed the checked content span length");
    }
    splice_content(source, first..last, replacement, max_output)
}

/// Clone one dependency-free worksheet owner and rename only its copied name attribute.
///
/// The source table fragment remains byte-exact apart from the required `table:name`
/// value replacement. All semantic audits run before the replacement allocation or
/// package publication.
pub(crate) fn copy_sheet(
    source: &[u8],
    source_sheet: &str,
    destination_name: &str,
    final_position: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    if source.len() > max_output {
        return invalid(format!(
            "ODS source exceeds the {max_output} byte sheet-copy output limit"
        ));
    }
    if destination_name.len() > MAX_SHEET_COPY_NAME_BYTES {
        return invalid(format!(
            "ODS copied sheet name exceeds the {MAX_SHEET_COPY_NAME_BYTES} byte limit"
        ));
    }
    let _validated_name = Sheet::new(destination_name)?;

    let package = Package::from_bytes(source.to_vec())?;
    let xml = package.content_xml().to_string();
    let spans = scan(&xml)?;
    let spreadsheet = one(&spans, OFFICE, "spreadsheet")?;
    let tables = children(&spans, spreadsheet, TABLE, "table");
    if tables.is_empty() || tables.len() >= MAX_SHEET_MOVE_SHEETS {
        return invalid(format!(
            "ODS sheet copy requires 1..{} source worksheets",
            MAX_SHEET_MOVE_SHEETS - 1
        ));
    }
    if final_position > tables.len() {
        return invalid(format!(
            "ODS sheet copy destination {final_position} exceeds final sheet count {}",
            tables.len() + 1
        ));
    }

    let selected = select_sheet(&xml, &spans, source_sheet)?;
    for table in &tables {
        if resolved_attribute(&xml, &spans, *table, TABLE, "name")?.as_deref()
            == Some(destination_name)
        {
            return invalid(format!(
                "ODS copied sheet name '{destination_name}' already exists"
            ));
        }
    }

    let name_range = audit_sheet_copy(&package, &xml, &spans, spreadsheet, selected)?;
    let selected_span = spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS selected sheet span is missing"))?;
    let fragment = xml
        .as_bytes()
        .get(selected_span.start..selected_span.end)
        .ok_or_else(|| invalid_error("ODS copied sheet fragment is invalid"))?;
    let relative_name = name_range
        .start
        .checked_sub(selected_span.start)
        .and_then(|start| {
            name_range
                .end
                .checked_sub(selected_span.start)
                .map(|end| start..end)
        })
        .filter(|range| range.start <= range.end && range.end <= fragment.len())
        .ok_or_else(|| invalid_error("ODS copied sheet name span is invalid"))?;
    let escaped_name = escape_sheet_name_attribute(destination_name);
    let replacement_len = fragment
        .len()
        .checked_sub(relative_name.len())
        .and_then(|length| length.checked_add(escaped_name.len()))
        .ok_or_else(|| invalid_error("ODS copied sheet fragment length overflows"))?;
    let final_content_len = xml
        .len()
        .checked_add(replacement_len)
        .ok_or_else(|| invalid_error("ODS copied content length overflows"))?;
    if final_content_len > max_output {
        return invalid(format!(
            "ODS sheet copy exceeds the {max_output} byte content/output limit"
        ));
    }

    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(replacement_len)
        .map_err(|_error| invalid_error("ODS sheet-copy fragment allocation failed"))?;
    replacement.extend_from_slice(&fragment[..relative_name.start]);
    replacement.extend_from_slice(escaped_name.as_bytes());
    replacement.extend_from_slice(&fragment[relative_name.end..]);

    let insertion = if final_position == tables.len() {
        spans[*tables
            .last()
            .ok_or_else(|| invalid_error("ODS sheet inventory unexpectedly became empty"))?]
        .end
    } else {
        spans[tables[final_position]].start
    };
    splice_content(source, insertion..insertion, replacement, max_output)
}

/// Transfer one worksheet whose complete selected owner is a bounded plain-scalar grid.
///
/// The source table is copied lexically, apart from its `table:name` value.  The
/// closure intentionally excludes every style, formula, merge, range, drawing,
/// extension, and package dependency so that the destination does not need any
/// source-local owner remapping.
pub(crate) fn transfer_plain_scalar_sheet(
    destination: &[u8],
    source: &[u8],
    source_sheet: &str,
    destination_name: &str,
    position: crate::document::SheetPosition,
    max_output: usize,
) -> Result<Vec<u8>> {
    if destination.len() > max_output {
        return invalid(format!(
            "ODS destination exceeds the {max_output} byte sheet-transfer output limit"
        ));
    }
    if source.len() > max_output {
        return invalid(format!(
            "ODS source exceeds the {max_output} byte sheet-transfer work limit"
        ));
    }
    if destination_name.len() > MAX_SHEET_COPY_NAME_BYTES {
        return invalid(format!(
            "ODS transferred sheet name exceeds the {MAX_SHEET_COPY_NAME_BYTES} byte limit"
        ));
    }
    validate_transfer_sheet_name(destination_name)?;

    let source_package = Package::from_bytes(copy_package_bytes(source, max_output, "source")?)?;
    audit_plain_scalar_transfer_package(&source_package, "source")?;
    let source_xml = source_package.content_xml();
    let source_spans = scan(source_xml)?;
    let source_spreadsheet = one(&source_spans, OFFICE, "spreadsheet")?;
    let source_tables = bounded_children(
        &source_spans,
        source_spreadsheet,
        TABLE,
        "table",
        MAX_SHEET_MOVE_SHEETS,
    )?;
    if source_tables.is_empty() || source_tables.len() >= MAX_SHEET_MOVE_SHEETS {
        return invalid(format!(
            "ODS sheet transfer requires 1..{} source worksheets",
            MAX_SHEET_MOVE_SHEETS - 1
        ));
    }
    let selected = select_sheet_bounded(source_xml, &source_spans, source_sheet)?;
    let name_range =
        audit_plain_scalar_sheet(source_xml, &source_spans, selected, source_spreadsheet)?;

    let destination_package =
        Package::from_bytes(copy_package_bytes(destination, max_output, "destination")?)?;
    audit_plain_scalar_transfer_package(&destination_package, "destination")?;
    let destination_xml = destination_package.content_xml();
    let destination_spans = scan(destination_xml)?;
    let destination_spreadsheet = one(&destination_spans, OFFICE, "spreadsheet")?;
    if is_self_closing_span(destination_xml, &destination_spans[destination_spreadsheet])? {
        return invalid(
            "ODS scalar sheet transfer refuses a self-closing destination spreadsheet owner",
        );
    }
    let destination_tables = bounded_children(
        &destination_spans,
        destination_spreadsheet,
        TABLE,
        "table",
        MAX_SHEET_MOVE_SHEETS,
    )?;
    if destination_tables.len() >= MAX_SHEET_MOVE_SHEETS {
        return invalid(format!(
            "ODS sheet transfer destination has too many worksheets (limit {})",
            MAX_SHEET_MOVE_SHEETS - 1
        ));
    }
    let final_position = match position {
        crate::document::SheetPosition::First => 0,
        crate::document::SheetPosition::Last => destination_tables.len(),
        crate::document::SheetPosition::Index(index) if index <= destination_tables.len() => index,
        crate::document::SheetPosition::Index(index) => {
            return invalid(format!(
                "ODS sheet transfer destination {index} exceeds final sheet count {}",
                destination_tables.len() + 1
            ));
        },
    };
    if final_position > destination_tables.len() {
        return invalid(format!(
            "ODS sheet transfer destination {final_position} exceeds final sheet count {}",
            destination_tables.len() + 1
        ));
    }
    let mut destination_names = Vec::new();
    destination_names
        .try_reserve_exact(destination_tables.len())
        .map_err(|_error| invalid_error("ODS sheet transfer name inventory allocation failed"))?;
    for table in &destination_tables {
        let name = resolved_attribute(destination_xml, &destination_spans, *table, TABLE, "name")?
            .ok_or_else(|| invalid_error("ODS sheet transfer destination has an unnamed sheet"))?;
        if destination_names.iter().any(|candidate| candidate == &name) {
            return invalid(format!(
                "ODS sheet transfer destination sheet name '{name}' is ambiguous"
            ));
        }
        validate_transfer_sheet_name(&name)?;
        if name == destination_name {
            return invalid(format!(
                "ODS transferred sheet name '{destination_name}' already exists"
            ));
        }
        destination_names.push(name);
    }

    let source_root = one(&source_spans, OFFICE, "document-content")?;
    let destination_root = one(&destination_spans, OFFICE, "document-content")?;
    ensure_transfer_namespace_compatibility(
        source_xml,
        &source_spans,
        source_root,
        selected,
        destination_xml,
        &destination_spans,
        destination_root,
        destination_spreadsheet,
    )?;

    let selected_span = source_spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS transferred sheet span is missing"))?;
    let fragment = source_xml
        .as_bytes()
        .get(selected_span.start..selected_span.end)
        .ok_or_else(|| invalid_error("ODS transferred sheet fragment is invalid"))?;
    if fragment.len() > MAX_SHEET_TRANSFER_FRAGMENT_BYTES {
        return invalid(format!(
            "ODS transferred sheet fragment exceeds the {MAX_SHEET_TRANSFER_FRAGMENT_BYTES} byte limit"
        ));
    }
    let relative_name = name_range
        .start
        .checked_sub(selected_span.start)
        .and_then(|start| {
            name_range
                .end
                .checked_sub(selected_span.start)
                .map(|end| start..end)
        })
        .filter(|range| range.start <= range.end && range.end <= fragment.len())
        .ok_or_else(|| invalid_error("ODS transferred sheet name span is invalid"))?;
    let escaped_name = escape_sheet_name_attribute_bounded(destination_name)?;
    let replacement_len = fragment
        .len()
        .checked_sub(relative_name.len())
        .and_then(|length| length.checked_add(escaped_name.len()))
        .ok_or_else(|| invalid_error("ODS transferred sheet fragment length overflows"))?;
    let final_content_len = destination_xml
        .len()
        .checked_add(replacement_len)
        .ok_or_else(|| invalid_error("ODS transferred content length overflows"))?;
    if final_content_len > max_output {
        return invalid(format!(
            "ODS sheet transfer exceeds the {max_output} byte content/output limit"
        ));
    }

    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(replacement_len)
        .map_err(|_error| invalid_error("ODS sheet-transfer fragment allocation failed"))?;
    replacement.extend_from_slice(&fragment[..relative_name.start]);
    replacement.extend_from_slice(escaped_name.as_bytes());
    replacement.extend_from_slice(&fragment[relative_name.end..]);

    let insertion = if final_position == destination_tables.len() {
        destination_tables.last().map_or(
            destination_spans[destination_spreadsheet].close_start,
            |table| destination_spans[*table].end,
        )
    } else {
        destination_spans[destination_tables[final_position]].start
    };
    splice_content(destination, insertion..insertion, replacement, max_output)
}

#[derive(Clone, Debug, Default)]
struct ScalarAttributes {
    name_key: Option<Vec<u8>>,
    rows_repeated: usize,
    rows_repeated_seen: bool,
    columns_repeated: usize,
    columns_repeated_seen: bool,
    value_type: Option<String>,
    value: Option<String>,
    boolean_value: Option<String>,
    date_value: Option<String>,
    time_value: Option<String>,
}

#[derive(Clone, Debug)]
struct ScalarCellState {
    attributes: ScalarAttributes,
    text_nonempty: bool,
    text_bytes: usize,
}

fn audit_plain_scalar_transfer_package(package: &Package, label: &str) -> Result<()> {
    let reader = package.package().package()?;
    if reader.manifest().has_encrypted_entries() {
        return invalid(format!(
            "ODS {label} sheet transfer refuses encrypted package members"
        ));
    }
    for path in reader.files()? {
        let mut lower = copy_bounded_string(&path, "package member path")?;
        lower.make_ascii_lowercase();
        if lower.contains("signature")
            || lower == "settings.xml"
            || lower == "manifest.rdf"
            || lower == "scripts.xml"
            || lower.starts_with("scripts/")
            || lower.starts_with("basic/")
            || lower.starts_with("configurations2/")
            || lower.starts_with("pictures/")
            || lower.starts_with("object")
            || lower.starts_with("links/")
            || lower.ends_with("/content.xml")
            || lower.contains("tracked")
        {
            return invalid(format!(
                "ODS {label} sheet transfer refuses dependent package member '{path}'"
            ));
        }
    }
    let protection =
        crate::protection::Snapshot::parse(package.content_xml(), package.styles_xml())?;
    if protection.document().structure_protected == Some(true)
        || protection
            .sheets()
            .iter()
            .any(|sheet| sheet.is_protected() == Some(true))
    {
        return invalid(format!(
            "ODS {label} sheet transfer refuses protected worksheets"
        ));
    }
    Ok(())
}

fn copy_package_bytes(bytes: &[u8], limit: usize, label: &str) -> Result<Vec<u8>> {
    if bytes.len() > limit {
        return invalid(format!(
            "ODS {label} package exceeds the {limit} byte transfer limit"
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_error| invalid_error(format!("ODS {label} package allocation failed")))?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn copy_bounded_string(value: &str, label: &str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_error| invalid_error(format!("ODS {label} allocation failed")))?;
    output.push_str(value);
    Ok(output)
}

fn validate_transfer_sheet_name(name: &str) -> Result<()> {
    crate::worksheet::validation::validate_text(name, "ODS transferred sheet name")?;
    if name.is_empty() {
        return invalid("ODS transferred sheet name must be non-empty");
    }
    Ok(())
}

fn audit_plain_scalar_sheet(
    xml: &str,
    spans: &[Span],
    selected: usize,
    spreadsheet: usize,
) -> Result<Range<usize>> {
    let selected_span = spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS scalar sheet span is missing"))?;
    if selected_span.parent != Some(spreadsheet) {
        return invalid("ODS scalar sheet is not a direct spreadsheet child");
    }
    let fragment = xml
        .as_bytes()
        .get(selected_span.start..selected_span.end)
        .ok_or_else(|| invalid_error("ODS scalar sheet fragment is invalid"))?;
    if fragment.len() > MAX_SHEET_TRANSFER_FRAGMENT_BYTES {
        return invalid(format!(
            "ODS scalar sheet fragment exceeds the {MAX_SHEET_TRANSFER_FRAGMENT_BYTES} byte limit"
        ));
    }
    let wrapped = wrap_selected_fragment(xml, spans, selected)?;
    let mut reader = NsReader::from_str(&wrapped);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(wrapped.len())
        .map_err(|_error| invalid_error("ODS scalar sheet XML buffer allocation failed"))?;
    let mut stack = Vec::<OrdinaryElement>::new();
    stack
        .try_reserve_exact(MAX_SHEET_TRANSFER_DEPTH)
        .map_err(|_error| invalid_error("ODS scalar sheet stack allocation failed"))?;
    let mut events = 0usize;
    let mut seen_table = false;
    let mut table_name_key = None::<Vec<u8>>;
    let mut current_row_repeat = None::<usize>;
    let mut current_row_cells = 0usize;
    let mut current_cell = None::<ScalarCellState>;
    let mut logical_rows = 0usize;
    let mut logical_cells = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS scalar sheet event count overflows"))?;
        if events > MAX_SHEET_COPY_EVENTS {
            return invalid(format!(
                "ODS scalar sheet exceeds the {MAX_SHEET_COPY_EVENTS} XML event limit"
            ));
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS scalar sheet XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let kind = ordinary_element(&namespace, element.local_name().as_ref())?;
                audit_ordinary_parent(stack.last().copied(), kind)?;
                let attributes = audit_scalar_attributes(&element, &reader, kind)?;
                match kind {
                    OrdinaryElement::Table => {
                        if seen_table {
                            return invalid(
                                "ODS scalar sheet contains a nested or duplicate table",
                            );
                        }
                        seen_table = true;
                        table_name_key = attributes.name_key;
                    },
                    OrdinaryElement::Row => {
                        if current_row_repeat.is_some() {
                            return invalid("ODS scalar sheet contains a nested row");
                        }
                        current_row_repeat = Some(attributes.rows_repeated);
                        current_row_cells = 0;
                    },
                    OrdinaryElement::Cell => {
                        if current_cell.is_some() {
                            return invalid("ODS scalar sheet contains a nested cell");
                        }
                        current_row_cells = current_row_cells
                            .checked_add(attributes.columns_repeated)
                            .ok_or_else(|| {
                                invalid_error("ODS scalar sheet column count overflows")
                            })?;
                        if current_row_cells > MAX_SHEET_TRANSFER_CELLS {
                            return invalid(format!(
                                "ODS scalar sheet row exceeds the {MAX_SHEET_TRANSFER_CELLS} cell limit"
                            ));
                        }
                        current_cell = Some(ScalarCellState {
                            attributes,
                            text_nonempty: false,
                            text_bytes: 0,
                        });
                    },
                    OrdinaryElement::Paragraph => {
                        if current_cell.is_none() {
                            return invalid("ODS scalar sheet paragraph is outside a cell");
                        }
                    },
                    _ => {},
                }
                if stack.len() >= MAX_SHEET_TRANSFER_DEPTH {
                    return invalid(format!(
                        "ODS scalar sheet exceeds the {MAX_SHEET_TRANSFER_DEPTH} XML depth limit"
                    ));
                }
                stack.push(kind);
            },
            Event::Empty(element) => {
                let kind = ordinary_element(&namespace, element.local_name().as_ref())?;
                audit_ordinary_parent(stack.last().copied(), kind)?;
                let attributes = audit_scalar_attributes(&element, &reader, kind)?;
                match kind {
                    OrdinaryElement::Table => {
                        if seen_table {
                            return invalid(
                                "ODS scalar sheet contains a nested or duplicate table",
                            );
                        }
                        seen_table = true;
                        table_name_key = attributes.name_key;
                    },
                    OrdinaryElement::Row => {
                        if current_row_repeat.is_some() {
                            return invalid("ODS scalar sheet contains a nested row");
                        }
                        logical_rows = logical_rows
                            .checked_add(attributes.rows_repeated)
                            .ok_or_else(|| invalid_error("ODS scalar sheet row count overflows"))?;
                        if logical_rows > MAX_SHEET_TRANSFER_ROWS {
                            return invalid(format!(
                                "ODS scalar sheet exceeds the {MAX_SHEET_TRANSFER_ROWS} row limit"
                            ));
                        }
                    },
                    OrdinaryElement::Cell => {
                        current_row_cells = current_row_cells
                            .checked_add(attributes.columns_repeated)
                            .ok_or_else(|| {
                                invalid_error("ODS scalar sheet column count overflows")
                            })?;
                        if current_row_cells > MAX_SHEET_TRANSFER_CELLS {
                            return invalid(format!(
                                "ODS scalar sheet row exceeds the {MAX_SHEET_TRANSFER_CELLS} cell limit"
                            ));
                        }
                        validate_scalar_cell(&attributes, false)?;
                    },
                    OrdinaryElement::Paragraph => {
                        if current_cell.is_none() {
                            return invalid("ODS scalar sheet paragraph is outside a cell");
                        }
                    },
                    _ => {},
                }
            },
            Event::End(_) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid_error("ODS scalar sheet element stack underflow"))?;
                match kind {
                    OrdinaryElement::Cell => {
                        let cell = current_cell.take().ok_or_else(|| {
                            invalid_error("ODS scalar sheet cell state is missing")
                        })?;
                        validate_scalar_cell(&cell.attributes, cell.text_nonempty)?;
                    },
                    OrdinaryElement::Row => {
                        let repeat = current_row_repeat.take().ok_or_else(|| {
                            invalid_error("ODS scalar sheet row state is missing")
                        })?;
                        logical_rows = logical_rows
                            .checked_add(repeat)
                            .ok_or_else(|| invalid_error("ODS scalar sheet row count overflows"))?;
                        if logical_rows > MAX_SHEET_TRANSFER_ROWS {
                            return invalid(format!(
                                "ODS scalar sheet exceeds the {MAX_SHEET_TRANSFER_ROWS} row limit"
                            ));
                        }
                        let row_cells = current_row_cells;
                        let row_footprint = row_cells.checked_mul(repeat).ok_or_else(|| {
                            invalid_error("ODS scalar sheet logical cell count overflows")
                        })?;
                        logical_cells =
                            logical_cells.checked_add(row_footprint).ok_or_else(|| {
                                invalid_error("ODS scalar sheet logical cell count overflows")
                            })?;
                        if logical_cells > MAX_SHEET_TRANSFER_CELLS {
                            return invalid(format!(
                                "ODS scalar sheet exceeds the {MAX_SHEET_TRANSFER_CELLS} cell limit"
                            ));
                        }
                        current_row_cells = 0;
                    },
                    _ => {},
                }
            },
            Event::Text(text) => {
                let raw_text: &[u8] = text.as_ref();
                if stack.last().copied() == Some(OrdinaryElement::Paragraph) {
                    let cell = current_cell.as_mut().ok_or_else(|| {
                        invalid_error("ODS scalar sheet paragraph cell is missing")
                    })?;
                    cell.text_nonempty |= !raw_text.is_empty();
                    cell.text_bytes = cell
                        .text_bytes
                        .checked_add(raw_text.len())
                        .ok_or_else(|| invalid_error("ODS scalar sheet text length overflows"))?;
                    if cell.text_bytes > MAX_SHEET_TRANSFER_FRAGMENT_BYTES {
                        return invalid(format!(
                            "ODS scalar sheet text exceeds the {MAX_SHEET_TRANSFER_FRAGMENT_BYTES} byte limit"
                        ));
                    }
                } else if !raw_text.iter().all(u8::is_ascii_whitespace) {
                    return invalid("ODS scalar sheet refuses text outside plain cell paragraphs");
                }
            },
            Event::Decl(_) => {},
            Event::Eof => break,
            Event::DocType(_)
            | Event::PI(_)
            | Event::Comment(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {
                return invalid("ODS scalar sheet refuses opaque XML events");
            },
        }
        buffer.clear();
    }
    if !stack.is_empty() || current_cell.is_some() || current_row_repeat.is_some() {
        return invalid("ODS scalar sheet content hierarchy is incomplete");
    }
    if !seen_table || table_name_key.is_none() {
        return invalid("ODS scalar sheet name attribute is missing");
    }
    let name_key =
        table_name_key.ok_or_else(|| invalid_error("ODS scalar sheet name key is missing"))?;
    lexical_attribute_value_range(xml, selected_span.start..selected_span.tag_end, &name_key)
}

fn audit_scalar_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    kind: OrdinaryElement,
) -> Result<ScalarAttributes> {
    let mut result = ScalarAttributes {
        rows_repeated: 1,
        columns_repeated: 1,
        ..ScalarAttributes::default()
    };
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| invalid_error(format!("invalid ODS scalar attribute: {error}")))?;
        let qname = raw.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let uri = match namespace {
            ResolveResult::Bound(Namespace(uri)) => uri,
            ResolveResult::Unbound => {
                return invalid("ODS scalar sheet refuses unqualified attribute ownership");
            },
            ResolveResult::Unknown(prefix) => {
                return invalid(format!(
                    "ODS scalar sheet refuses unbound attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                ));
            },
        };
        let decoded = raw
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                invalid_error(format!("invalid ODS scalar attribute value: {error}"))
            })?;
        let mut value = String::new();
        value
            .try_reserve_exact(decoded.len())
            .map_err(|_error| invalid_error("ODS scalar attribute allocation failed"))?;
        value.push_str(decoded.as_ref());
        let local = local.as_ref();
        let allowed = match kind {
            OrdinaryElement::Root => uri == OFFICE.as_bytes() && local == b"version",
            OrdinaryElement::Body | OrdinaryElement::Spreadsheet => false,
            OrdinaryElement::Table => uri == TABLE.as_bytes() && local == b"name",
            OrdinaryElement::Column => {
                uri == TABLE.as_bytes() && local == b"number-columns-repeated"
            },
            OrdinaryElement::Row => uri == TABLE.as_bytes() && local == b"number-rows-repeated",
            OrdinaryElement::Cell => {
                uri == TABLE.as_bytes() && local == b"number-columns-repeated"
                    || uri == OFFICE.as_bytes()
                        && matches!(
                            local,
                            b"value-type"
                                | b"value"
                                | b"date-value"
                                | b"time-value"
                                | b"boolean-value"
                        )
            },
            OrdinaryElement::Paragraph => false,
        };
        if !allowed {
            return invalid(format!(
                "ODS scalar sheet refuses attribute '{}'",
                String::from_utf8_lossy(qname)
            ));
        }
        if uri == TABLE.as_bytes() && local == b"name" {
            if result.name_key.is_some() {
                return invalid("ODS scalar sheet name attribute is duplicated");
            }
            let mut name_key = Vec::new();
            name_key
                .try_reserve_exact(qname.len())
                .map_err(|_error| invalid_error("ODS scalar sheet name allocation failed"))?;
            name_key.extend_from_slice(qname);
            result.name_key = Some(name_key);
        } else if uri == TABLE.as_bytes() && local == b"number-rows-repeated" {
            if result.rows_repeated_seen {
                return invalid("ODS scalar sheet row repetition is duplicated");
            }
            result.rows_repeated_seen = true;
            result.rows_repeated = scalar_positive(&value, "number-rows-repeated")?;
        } else if uri == TABLE.as_bytes() && local == b"number-columns-repeated" {
            if result.columns_repeated_seen {
                return invalid("ODS scalar sheet column repetition is duplicated");
            }
            result.columns_repeated_seen = true;
            result.columns_repeated = scalar_positive(&value, "number-columns-repeated")?;
        } else if uri == OFFICE.as_bytes() && local == b"value-type" {
            if result.value_type.replace(value).is_some() {
                return invalid("ODS scalar sheet value type is duplicated");
            }
        } else if uri == OFFICE.as_bytes() && local == b"value" {
            if result.value.replace(value).is_some() {
                return invalid("ODS scalar sheet value is duplicated");
            }
        } else if uri == OFFICE.as_bytes() && local == b"boolean-value" {
            if result.boolean_value.replace(value).is_some() {
                return invalid("ODS scalar sheet boolean value is duplicated");
            }
        } else if uri == OFFICE.as_bytes() && local == b"date-value" {
            if result.date_value.replace(value).is_some() {
                return invalid("ODS scalar sheet date value is duplicated");
            }
        } else if uri == OFFICE.as_bytes()
            && local == b"time-value"
            && result.time_value.replace(value).is_some()
        {
            return invalid("ODS scalar sheet time value is duplicated");
        }
    }
    if kind == OrdinaryElement::Table && result.name_key.is_none() {
        return invalid("ODS scalar sheet name attribute is missing");
    }
    Ok(result)
}

fn scalar_positive(value: &str, label: &str) -> Result<usize> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_error| invalid_error(format!("ODS scalar {label} is not positive")))?;
    if parsed == 0 || parsed > MAX_SHEET_TRANSFER_REPETITION {
        return invalid(format!(
            "ODS scalar {label} exceeds the {MAX_SHEET_TRANSFER_REPETITION} bound"
        ));
    }
    Ok(parsed)
}

fn validate_scalar_cell(attributes: &ScalarAttributes, _text_nonempty: bool) -> Result<()> {
    let has_value = attributes.value.is_some();
    let has_boolean = attributes.boolean_value.is_some();
    let has_date = attributes.date_value.is_some();
    let has_time = attributes.time_value.is_some();
    match attributes.value_type.as_deref() {
        None => {
            if has_value || has_boolean || has_date || has_time {
                return invalid("ODS scalar untyped cells cannot carry typed value attributes");
            }
        },
        Some("string") => {
            if has_value || has_boolean || has_date || has_time {
                return invalid("ODS scalar string cells cannot carry non-string value attributes");
            }
        },
        Some("float" | "double" | "decimal") => {
            let value = attributes
                .value
                .as_deref()
                .ok_or_else(|| invalid_error("ODS scalar number cell requires office:value"))?;
            let valid_number = match value.parse::<f64>() {
                Ok(number) => number.is_finite(),
                Err(_error) => false,
            };
            if !valid_number || has_boolean || has_date || has_time {
                return invalid("ODS scalar number cell value is invalid");
            }
        },
        Some("boolean") => {
            let value = attributes
                .boolean_value
                .as_deref()
                .or(attributes.value.as_deref())
                .ok_or_else(|| invalid_error("ODS scalar boolean cell requires a boolean value"))?;
            if !matches!(value, "true" | "false")
                || (has_boolean && has_value)
                || has_date
                || has_time
            {
                return invalid("ODS scalar boolean cell value is invalid");
            }
        },
        Some("date") => {
            let value = attributes
                .date_value
                .as_deref()
                .or(attributes.value.as_deref())
                .ok_or_else(|| invalid_error("ODS scalar date cell requires a date value"))?;
            if value.is_empty() || has_boolean || has_time {
                return invalid("ODS scalar date cell value is invalid");
            }
            litchi_odf_common::datatype::Date::decode(value)
                .map_err(|_error| invalid_error("ODS scalar date cell lexical value is invalid"))?;
        },
        Some("time") => {
            let value = attributes
                .time_value
                .as_deref()
                .or(attributes.value.as_deref())
                .ok_or_else(|| invalid_error("ODS scalar time cell requires a time value"))?;
            if value.is_empty() || has_boolean || has_date {
                return invalid("ODS scalar time cell value is invalid");
            }
            litchi_odf_common::datatype::Duration::decode(value)
                .map_err(|_error| invalid_error("ODS scalar time cell lexical value is invalid"))?;
        },
        Some(kind) => {
            return invalid(format!("ODS scalar sheet refuses cell value type '{kind}'"));
        },
    }
    Ok(())
}

fn wrap_selected_fragment(xml: &str, spans: &[Span], selected: usize) -> Result<String> {
    let selected_span = spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS scalar sheet span is missing"))?;
    let mut ancestors = Vec::new();
    let mut parent = selected_span.parent;
    while let Some(index) = parent {
        if ancestors.len() >= MAX_SHEET_TRANSFER_DEPTH {
            return invalid(format!(
                "ODS scalar sheet ancestor depth exceeds the {MAX_SHEET_TRANSFER_DEPTH} limit"
            ));
        }
        ancestors
            .try_reserve(1)
            .map_err(|_error| invalid_error("ODS scalar sheet ancestor allocation failed"))?;
        let ancestor = spans
            .get(index)
            .ok_or_else(|| invalid_error("ODS scalar sheet ancestor span is missing"))?;
        ancestors.push(ancestor);
        parent = ancestor.parent;
    }
    let mut output = String::new();
    let capacity = selected_span
        .end
        .checked_sub(selected_span.start)
        .and_then(|length| length.checked_add(ancestors.len().saturating_mul(128)))
        .ok_or_else(|| invalid_error("ODS scalar sheet wrapper size overflows"))?;
    if capacity > MAX_SHEET_TRANSFER_FRAGMENT_BYTES {
        return invalid(format!(
            "ODS scalar sheet wrapper exceeds the {MAX_SHEET_TRANSFER_FRAGMENT_BYTES} byte limit"
        ));
    }
    output
        .try_reserve_exact(capacity)
        .map_err(|_error| invalid_error("ODS scalar sheet wrapper allocation failed"))?;
    for ancestor in ancestors.iter().rev() {
        bounded_append(
            &mut output,
            xml.get(ancestor.start..ancestor.tag_end)
                .ok_or_else(|| invalid_error("ODS scalar sheet ancestor tag is invalid"))?,
            MAX_SHEET_TRANSFER_FRAGMENT_BYTES,
        )?;
    }
    bounded_append(
        &mut output,
        xml.get(selected_span.start..selected_span.end)
            .ok_or_else(|| invalid_error("ODS scalar sheet fragment is invalid"))?,
        MAX_SHEET_TRANSFER_FRAGMENT_BYTES,
    )?;
    for ancestor in &ancestors {
        bounded_append(&mut output, "</", MAX_SHEET_TRANSFER_FRAGMENT_BYTES)?;
        bounded_append(
            &mut output,
            start_qname(xml, ancestor)?,
            MAX_SHEET_TRANSFER_FRAGMENT_BYTES,
        )?;
        bounded_append(&mut output, ">", MAX_SHEET_TRANSFER_FRAGMENT_BYTES)?;
    }
    Ok(output)
}

fn ensure_transfer_namespace_compatibility(
    source_xml: &str,
    source_spans: &[Span],
    source_root: usize,
    selected: usize,
    destination_xml: &str,
    destination_spans: &[Span],
    destination_root: usize,
    destination_owner: usize,
) -> Result<()> {
    let source_bindings =
        effective_namespace_bindings(source_xml, source_spans, source_root, selected, "source")?;
    let destination_bindings = effective_namespace_bindings(
        destination_xml,
        destination_spans,
        destination_root,
        destination_owner,
        "destination",
    )?;
    let selected_span = source_spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS scalar sheet namespace span is missing"))?;
    let fragment = source_xml
        .get(selected_span.start..selected_span.end)
        .ok_or_else(|| invalid_error("ODS scalar sheet namespace fragment is invalid"))?;
    let prefixes = fragment_prefixes(fragment)?;
    for prefix in prefixes {
        let source_uri = source_bindings
            .iter()
            .find_map(|(candidate, uri)| (candidate == &prefix).then_some(uri))
            .ok_or_else(|| invalid_error("ODS scalar sheet namespace binding is missing"))?;
        if !destination_bindings
            .iter()
            .any(|(candidate, uri)| candidate == &prefix && uri == source_uri)
        {
            return invalid(format!(
                "ODS scalar sheet transfer refuses destination namespace collision for prefix '{prefix}'"
            ));
        }
    }
    Ok(())
}

fn effective_namespace_bindings(
    xml: &str,
    spans: &[Span],
    root: usize,
    owner: usize,
    label: &str,
) -> Result<Vec<(String, String)>> {
    let mut chain = Vec::new();
    let mut current = Some(owner);
    while let Some(index) = current {
        if chain.len() >= MAX_SHEET_TRANSFER_DEPTH {
            return invalid(format!(
                "ODS {label} namespace owner depth exceeds the {MAX_SHEET_TRANSFER_DEPTH} limit"
            ));
        }
        chain.try_reserve(1).map_err(|_error| {
            invalid_error(format!("ODS {label} namespace chain allocation failed"))
        })?;
        chain.push(index);
        if index == root {
            break;
        }
        current = spans
            .get(index)
            .ok_or_else(|| invalid_error(format!("ODS {label} namespace owner span is missing")))?
            .parent;
    }
    if chain.last().copied() != Some(root) {
        return invalid(format!(
            "ODS {label} namespace owner is outside the document-content root"
        ));
    }
    let bindings = root_namespace_bindings(xml, spans, root)?;
    for index in chain.iter().rev().skip(1) {
        let span = spans.get(*index).ok_or_else(|| {
            invalid_error(format!("ODS {label} namespace ancestor span is missing"))
        })?;
        if has_namespace_declaration(xml, span)? {
            return invalid(format!(
                "ODS scalar sheet transfer refuses {label} namespace redeclaration on '{}:{}'",
                span.namespace.as_deref().unwrap_or(""),
                span.local
            ));
        }
    }
    Ok(bindings)
}

fn has_namespace_declaration(xml: &str, span: &Span) -> Result<bool> {
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS namespace ancestor tag is invalid"))?;
    let mut reader = quick_xml::Reader::from_str(tag);
    let event = reader
        .read_event()
        .map_err(|error| invalid_error(format!("invalid ODS namespace ancestor: {error}")))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return invalid("ODS namespace ancestor start tag is missing"),
    };
    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|error| {
            invalid_error(format!("invalid ODS namespace ancestor attribute: {error}"))
        })?;
        let key = raw.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

fn root_namespace_bindings(
    xml: &str,
    spans: &[Span],
    root: usize,
) -> Result<Vec<(String, String)>> {
    let span = spans
        .get(root)
        .ok_or_else(|| invalid_error("ODS namespace root span is missing"))?;
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS namespace root tag is invalid"))?;
    let mut reader = quick_xml::Reader::from_str(tag);
    let event = reader
        .read_event()
        .map_err(|error| invalid_error(format!("invalid ODS namespace root: {error}")))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return invalid("ODS namespace root start tag is missing"),
    };
    let mut bindings = Vec::new();
    bindings
        .try_reserve(1)
        .map_err(|_error| invalid_error("ODS namespace binding allocation failed"))?;
    bindings.push((
        copy_bounded_string("xml", "namespace prefix")?,
        copy_bounded_string("http://www.w3.org/XML/1998/namespace", "namespace URI")?,
    ));
    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|error| {
            invalid_error(format!("invalid ODS namespace declaration: {error}"))
        })?;
        let key = raw.key.as_ref();
        let prefix = if key == b"xmlns" {
            copy_bounded_string("", "namespace prefix")?
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            decode(prefix, "namespace prefix")?
        } else {
            continue;
        };
        let decoded = raw
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid_error(format!("invalid ODS namespace URI: {error}")))?;
        let mut value = String::new();
        value
            .try_reserve_exact(decoded.len())
            .map_err(|_error| invalid_error("ODS namespace URI allocation failed"))?;
        value.push_str(decoded.as_ref());
        if bindings.iter().any(|(candidate, _)| candidate == &prefix) {
            return invalid(format!("ODS namespace prefix '{prefix}' is duplicated"));
        }
        if bindings.len() >= MAX_SHEET_TRANSFER_NAMESPACES {
            return invalid("ODS namespace binding limit exceeded");
        }
        bindings
            .try_reserve(1)
            .map_err(|_error| invalid_error("ODS namespace binding allocation failed"))?;
        bindings.push((prefix, value));
    }
    Ok(bindings)
}

fn fragment_prefixes(fragment: &str) -> Result<Vec<String>> {
    let mut reader = quick_xml::Reader::from_str(fragment);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(fragment.len())
        .map_err(|_error| invalid_error("ODS transfer fragment buffer allocation failed"))?;
    let mut prefixes = Vec::new();
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS transfer namespace event count overflows"))?;
        if events > MAX_SHEET_COPY_EVENTS {
            return invalid(format!(
                "ODS transfer fragment exceeds the {MAX_SHEET_COPY_EVENTS} XML event limit"
            ));
        }
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS transfer fragment: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                add_fragment_prefix(&mut prefixes, element.name().as_ref())?;
                for raw in element.attributes().with_checks(true) {
                    let raw = raw.map_err(|error| {
                        invalid_error(format!("invalid ODS transfer fragment attribute: {error}"))
                    })?;
                    let key = raw.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        return invalid(
                            "ODS scalar sheet transfer refuses local namespace declarations",
                        );
                    }
                    add_fragment_prefix(&mut prefixes, key)?;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Text(_) => {},
            Event::End(_) => {},
            Event::DocType(_)
            | Event::PI(_)
            | Event::Comment(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {
                return invalid("ODS scalar sheet transfer refuses opaque fragment events");
            },
        }
        buffer.clear();
    }
    Ok(prefixes)
}

fn add_fragment_prefix(prefixes: &mut Vec<String>, qname: &[u8]) -> Result<()> {
    let Some(index) = qname.iter().position(|byte| *byte == b':') else {
        if !prefixes.iter().any(String::is_empty) {
            if prefixes.len() >= MAX_SHEET_TRANSFER_NAMESPACES {
                return invalid("ODS transfer namespace prefix limit exceeded");
            }
            prefixes
                .try_reserve(1)
                .map_err(|_error| invalid_error("ODS transfer namespace allocation failed"))?;
            prefixes.push(String::new());
        }
        return Ok(());
    };
    let prefix = decode(&qname[..index], "fragment namespace prefix")?;
    if prefix != "xml" && !prefixes.iter().any(|candidate| candidate == &prefix) {
        if prefixes.len() >= MAX_SHEET_TRANSFER_NAMESPACES {
            return invalid("ODS transfer namespace prefix limit exceeded");
        }
        prefixes
            .try_reserve(1)
            .map_err(|_error| invalid_error("ODS transfer namespace allocation failed"))?;
        prefixes.push(prefix);
    }
    Ok(())
}

fn escape_sheet_name_attribute(value: &str) -> String {
    escape_xml(value)
        .replace('\t', "&#x9;")
        .replace('\n', "&#xA;")
        .replace('\r', "&#xD;")
}

fn escape_sheet_name_attribute_bounded(value: &str) -> Result<String> {
    let capacity = value
        .len()
        .checked_mul(6)
        .ok_or_else(|| invalid_error("ODS sheet name escape size overflows"))?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_error| invalid_error("ODS sheet name escape allocation failed"))?;
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            character => output.push(character),
        }
    }
    Ok(output)
}

fn audit_sheet_copy(
    package: &Package,
    xml: &str,
    spans: &[Span],
    spreadsheet: usize,
    selected: usize,
) -> Result<Range<usize>> {
    let reader = package.package().package()?;
    if reader.manifest().has_encrypted_entries() {
        return invalid("ODS sheet copy refuses encrypted package members");
    }
    for path in reader.files()? {
        let lower = path.to_ascii_lowercase();
        if lower == "settings.xml"
            || lower == "manifest.rdf"
            || lower == "scripts.xml"
            || lower.starts_with("scripts/")
            || lower.starts_with("basic/")
            || lower.starts_with("configurations2/")
            || lower.starts_with("pictures/")
            || lower.starts_with("object")
            || lower.starts_with("links/")
            || lower.ends_with("/content.xml")
            || lower.contains("signature")
            || lower.contains("tracked")
        {
            return invalid(format!(
                "ODS sheet copy refuses dependent package member '{path}'"
            ));
        }
    }

    if spans
        .iter()
        .any(|span| !sheet_move_namespace_supported(span.namespace.as_deref()))
    {
        return invalid("ODS sheet copy refuses MCE or unknown namespace owners");
    }
    if spans
        .iter()
        .any(|span| span.parent == Some(spreadsheet) && !is_element(span, TABLE, "table"))
    {
        return invalid("ODS sheet copy refuses non-sheet spreadsheet owners");
    }
    let selected_span = spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS selected sheet span is missing"))?;
    if let Some(span) = spans.iter().find(|span| {
        span.start >= selected_span.start
            && span.end <= selected_span.end
            && !sheet_copy_element_supported(span)
    }) {
        return invalid(format!(
            "ODS sheet copy refuses unsupported selected-sheet element '{}'",
            span.local
        ));
    }

    const FORBIDDEN_TABLE_ELEMENTS: [&str; 25] = [
        "calculation-settings",
        "cell-range-source",
        "change-deletion",
        "change-track-table-cell",
        "content-validation",
        "content-validations",
        "database-range",
        "database-ranges",
        "dde-link",
        "dde-links",
        "detective",
        "label-range",
        "label-ranges",
        "named-expression",
        "named-expressions",
        "named-range",
        "scenario",
        "shapes",
        "table-source",
        "tracked-changes",
        "insertion",
        "deletion",
        "movement",
        "cut-offs",
        "dependencies",
    ];
    const FORBIDDEN_OFFICE_ELEMENTS: [&str; 4] =
        ["event-listeners", "forms", "scripts", "tracked-changes"];
    if let Some(span) = spans.iter().find(|span| {
        (span.namespace.as_deref() == Some(TABLE)
            && FORBIDDEN_TABLE_ELEMENTS.contains(&span.local.as_str()))
            || (span.namespace.as_deref() == Some(OFFICE)
                && FORBIDDEN_OFFICE_ELEMENTS.contains(&span.local.as_str()))
    }) {
        return invalid(format!(
            "ODS sheet copy refuses dependency owner '{}:{}'",
            if span.namespace.as_deref() == Some(TABLE) {
                "table"
            } else {
                "office"
            },
            span.local
        ));
    }

    audit_sheet_copy_attributes(xml, spans, selected)
}

fn sheet_copy_element_supported(span: &Span) -> bool {
    match span.namespace.as_deref() {
        Some(TABLE) => matches!(
            span.local.as_str(),
            "table"
                | "table-column"
                | "table-columns"
                | "table-header-columns"
                | "table-column-group"
                | "table-row"
                | "table-rows"
                | "table-header-rows"
                | "table-row-group"
                | "table-cell"
                | "covered-table-cell"
        ),
        Some(TEXT) => matches!(
            span.local.as_str(),
            "p" | "span" | "s" | "tab" | "line-break" | "soft-page-break"
        ),
        Some(OFFICE | STYLE | NUMBER) | None => false,
        Some(_) => false,
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SheetCopyNamespace {
    Office,
    Table,
    Text,
    Style,
    Number,
    Xml,
    Other,
}

fn sheet_copy_namespace(uri: &[u8]) -> SheetCopyNamespace {
    match uri {
        value if value == OFFICE.as_bytes() => SheetCopyNamespace::Office,
        value if value == TABLE.as_bytes() => SheetCopyNamespace::Table,
        value if value == TEXT.as_bytes() => SheetCopyNamespace::Text,
        value if value == STYLE.as_bytes() => SheetCopyNamespace::Style,
        value if value == NUMBER.as_bytes() => SheetCopyNamespace::Number,
        b"http://www.w3.org/XML/1998/namespace" => SheetCopyNamespace::Xml,
        _ => SheetCopyNamespace::Other,
    }
}

fn sheet_copy_attribute_supported(
    element_namespace: SheetCopyNamespace,
    element_local: &[u8],
    attribute_namespace: SheetCopyNamespace,
    attribute_local: &[u8],
) -> bool {
    match (element_namespace, element_local, attribute_namespace) {
        (SheetCopyNamespace::Table, b"table", SheetCopyNamespace::Table) => matches!(
            attribute_local,
            b"name" | b"style-name" | b"display" | b"print"
        ),
        (SheetCopyNamespace::Table, b"table-column", SheetCopyNamespace::Table) => matches!(
            attribute_local,
            b"style-name" | b"default-cell-style-name" | b"number-columns-repeated" | b"visibility"
        ),
        (SheetCopyNamespace::Table, b"table-row", SheetCopyNamespace::Table) => matches!(
            attribute_local,
            b"style-name" | b"default-cell-style-name" | b"number-rows-repeated" | b"visibility"
        ),
        (
            SheetCopyNamespace::Table,
            b"table-column-group" | b"table-row-group",
            SheetCopyNamespace::Table,
        ) => attribute_local == b"display",
        (
            SheetCopyNamespace::Table,
            b"table-cell" | b"covered-table-cell",
            SheetCopyNamespace::Table,
        ) => matches!(
            attribute_local,
            b"style-name"
                | b"number-columns-repeated"
                | b"number-columns-spanned"
                | b"number-rows-spanned"
        ),
        (SheetCopyNamespace::Table, b"table-cell", SheetCopyNamespace::Office) => matches!(
            attribute_local,
            b"value-type"
                | b"value"
                | b"boolean-value"
                | b"date-value"
                | b"time-value"
                | b"string-value"
                | b"currency"
        ),
        (SheetCopyNamespace::Text, b"p" | b"span", SheetCopyNamespace::Text) => {
            attribute_local == b"style-name"
        },
        (SheetCopyNamespace::Text, b"s", SheetCopyNamespace::Text) => attribute_local == b"c",
        (SheetCopyNamespace::Text, b"p" | b"span", SheetCopyNamespace::Xml) => {
            matches!(attribute_local, b"lang" | b"space")
        },
        _ => false,
    }
}

fn audit_sheet_copy_attributes(xml: &str, spans: &[Span], selected: usize) -> Result<Range<usize>> {
    let selected_span = spans
        .get(selected)
        .ok_or_else(|| invalid_error("ODS selected sheet span is missing"))?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut name_key = None::<Vec<u8>>;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS sheet-copy XML event count overflows"))?;
        if events > MAX_SHEET_COPY_EVENTS {
            return invalid(format!(
                "ODS sheet copy exceeds the {MAX_SHEET_COPY_EVENTS} XML event limit"
            ));
        }
        let (element_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS sheet-copy XML: {error}")))?;
        let element_namespace = if matches!(&event, Event::Start(_) | Event::Empty(_)) {
            Some(match element_namespace {
                ResolveResult::Bound(Namespace(uri)) => sheet_copy_namespace(uri),
                ResolveResult::Unbound => {
                    return invalid("ODS sheet copy refuses unbound element ownership");
                },
                ResolveResult::Unknown(_) => {
                    return invalid("ODS sheet copy refuses unbound element prefixes");
                },
            })
        } else {
            None
        };
        let opens_depth = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let element_namespace = element_namespace.ok_or_else(|| {
                    invalid_error("ODS sheet-copy element namespace classification is missing")
                })?;
                let tag_end = position(&reader)?;
                let tag_start = tag_start(xml, tag_end)?;
                let selected_start = tag_start == selected_span.start;
                let within_selected =
                    tag_start >= selected_span.start && tag_end <= selected_span.end;
                let element_local = element.local_name();
                for raw in element.attributes().with_checks(true) {
                    let raw = raw.map_err(|error| {
                        invalid_error(format!("invalid ODS sheet-copy attribute: {error}"))
                    })?;
                    if raw.key.as_ref() == b"xmlns" || raw.key.as_ref().starts_with(b"xmlns:") {
                        continue;
                    }
                    let local = raw.key.local_name();
                    let resolved = reader.resolver().resolve_attribute(raw.key).0;
                    match resolved {
                        ResolveResult::Unbound => {
                            return invalid("ODS sheet copy refuses unbound attribute ownership");
                        },
                        ResolveResult::Bound(Namespace(uri)) => {
                            let namespace = sheet_copy_namespace(uri);
                            if namespace == SheetCopyNamespace::Other {
                                return invalid(
                                    "ODS sheet copy refuses MCE or unknown attribute namespace owners",
                                );
                            }
                            if namespace == SheetCopyNamespace::Table
                                && matches!(
                                    local.as_ref(),
                                    b"formula"
                                        | b"cell-range-address"
                                        | b"base-cell-address"
                                        | b"target-range-address"
                                        | b"source-cell-range-address"
                                        | b"source-range-address"
                                        | b"print-ranges"
                                        | b"content-validation-name"
                                        | b"protected"
                                        | b"structure-protected"
                                        | b"protection-key"
                                        | b"protection-key-digest-algorithm"
                                        | b"protection-key-digest-algorithm-2"
                                )
                            {
                                return invalid(
                                    "ODS sheet copy refuses name/range/protection attributes",
                                );
                            }
                            if namespace == SheetCopyNamespace::Xml && local.as_ref() == b"id" {
                                return invalid("ODS sheet copy refuses duplicated xml:id owners");
                            }
                            if within_selected
                                && !sheet_copy_attribute_supported(
                                    element_namespace,
                                    element_local.as_ref(),
                                    namespace,
                                    local.as_ref(),
                                )
                            {
                                return invalid(
                                    "ODS sheet copy refuses unsupported selected-sheet attributes",
                                );
                            }
                            if selected_start
                                && namespace == SheetCopyNamespace::Table
                                && local.as_ref() == b"name"
                            {
                                if name_key.replace(raw.key.as_ref().to_vec()).is_some() {
                                    return invalid(
                                        "ODS copied sheet has duplicate name attributes",
                                    );
                                }
                            }
                        },
                        ResolveResult::Unknown(_) => {
                            return invalid(
                                "ODS sheet copy refuses MCE or unknown attribute namespace owners",
                            );
                        },
                    }
                }
                if opens_depth {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid_error("ODS sheet-copy XML depth overflows"))?;
                    if depth > MAX_SHEET_COPY_DEPTH {
                        return invalid(format!(
                            "ODS sheet copy exceeds the {MAX_SHEET_COPY_DEPTH} XML depth limit"
                        ));
                    }
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODS sheet-copy XML depth underflows"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "ODS sheet copy refuses document types and processing instructions",
                );
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return invalid("ODS sheet-copy XML has an unclosed element");
    }
    let key =
        name_key.ok_or_else(|| invalid_error("ODS copied sheet name attribute is missing"))?;
    lexical_attribute_value_range(xml, selected_span.start..selected_span.tag_end, &key)
}

fn lexical_attribute_value_range(xml: &str, tag: Range<usize>, key: &[u8]) -> Result<Range<usize>> {
    let bytes = xml
        .as_bytes()
        .get(tag.clone())
        .ok_or_else(|| invalid_error("ODS copied sheet start-tag span is invalid"))?;
    let mut cursor = 1usize;
    while cursor < bytes.len() && !bytes[cursor].is_ascii_whitespace() && bytes[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < bytes.len() {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || matches!(bytes[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len()
            && !bytes[cursor].is_ascii_whitespace()
            && !matches!(bytes[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return invalid("ODS copied sheet attribute has no equals sign");
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid_error("ODS copied sheet attribute is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            return invalid("ODS copied sheet attribute value is unterminated");
        }
        let value_end = cursor;
        cursor += 1;
        if bytes.get(name_start..name_end) == Some(key) {
            return Ok((tag.start + value_start)..(tag.start + value_end));
        }
    }
    invalid("ODS copied sheet name attribute lexical span was not found")
}

fn audit_sheet_move_package(
    package: &Package,
    xml: &str,
    spans: &[Span],
    spreadsheet: usize,
) -> Result<()> {
    for path in package.package().files()? {
        if path.eq_ignore_ascii_case("settings.xml")
            || path.eq_ignore_ascii_case("manifest.rdf")
            || path.eq_ignore_ascii_case("scripts.xml")
            || starts_with_ascii_case_insensitive(&path, "scripts/")
            || starts_with_ascii_case_insensitive(&path, "basic/")
            || starts_with_ascii_case_insensitive(&path, "configurations2/")
            || contains_ascii_case_insensitive(&path, "tracked")
            || starts_with_ascii_case_insensitive(&path, "object")
            || ends_with_ascii_case_insensitive(&path, "/content.xml")
        {
            return invalid(format!(
                "ODS sheet move refuses dependent package member '{path}'"
            ));
        }
    }

    if spans
        .iter()
        .any(|span| !sheet_move_namespace_supported(span.namespace.as_deref()))
    {
        return invalid("ODS sheet move refuses MCE or unknown namespace owners");
    }
    audit_sheet_move_attributes(xml)?;
    if spans
        .iter()
        .any(|span| span.parent == Some(spreadsheet) && !is_element(span, TABLE, "table"))
    {
        return invalid("ODS sheet move refuses non-sheet spreadsheet owners");
    }

    const BLOCKERS: [&str; 28] = [
        "markup-compatibility",
        "alternatecontent",
        "office:scripts",
        "script:",
        "table:formula",
        ":formula=",
        "cell-range-address",
        "base-cell-address",
        "target-range-address",
        "print-ranges",
        "named-expressions",
        "named-range",
        "database-ranges",
        "database-range",
        "content-validations",
        "content-validation",
        "validation-name",
        "tracked-changes",
        "change-track",
        "label-ranges",
        "label-range",
        "dde-links",
        "table:scenario",
        "table:detective",
        "table:protected",
        ":protected=",
        "protection-key",
        "table:calculation-settings",
    ];
    if let Some(blocker) = BLOCKERS
        .iter()
        .find(|blocker| contains_ascii_case_insensitive(xml, blocker))
    {
        return invalid(format!(
            "ODS sheet move refuses sheet-order dependency owner '{blocker}'"
        ));
    }
    Ok(())
}

fn sheet_move_namespace_supported(namespace: Option<&str>) -> bool {
    matches!(namespace, Some(OFFICE | TABLE | TEXT | STYLE | NUMBER))
}

fn audit_sheet_move_attributes(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let (_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS sheet-move XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for raw in element.attributes().with_checks(true) {
                    let raw = raw.map_err(|error| {
                        invalid_error(format!("invalid ODS sheet-move attribute: {error}"))
                    })?;
                    if raw.key.as_ref() == b"xmlns" || raw.key.as_ref().starts_with(b"xmlns:") {
                        continue;
                    }
                    match reader.resolver().resolve_attribute(raw.key).0 {
                        ResolveResult::Unbound => {},
                        ResolveResult::Bound(Namespace(uri))
                            if std::str::from_utf8(uri).is_ok_and(|namespace| {
                                sheet_move_namespace_supported(Some(namespace))
                                    || namespace == "http://www.w3.org/XML/1998/namespace"
                            }) => {},
                        ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
                            return invalid(
                                "ODS sheet move refuses MCE or unknown attribute namespace owners",
                            );
                        },
                    }
                }
            },
            Event::DocType(_) => {
                return invalid("ODS sheet move refuses document type declarations");
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn starts_with_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .as_bytes()
        .get(..prefix.len())
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix.as_bytes()))
}

fn ends_with_ascii_case_insensitive(value: &str, suffix: &str) -> bool {
    value
        .as_bytes()
        .get(value.len().saturating_sub(suffix.len())..)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(suffix.as_bytes()))
}

fn contains_ascii_case_insensitive(value: &str, needle: &str) -> bool {
    value
        .as_bytes()
        .windows(needle.len())
        .any(|candidate| candidate.eq_ignore_ascii_case(needle.as_bytes()))
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
    let package = Package::from_bytes(copy_package_bytes(source, max_output, "content splice")?)?;
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
    buffer
        .try_reserve_exact(xml.len())
        .map_err(|_error| invalid_error("ODS content XML buffer allocation failed"))?;
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
                spans
                    .try_reserve(1)
                    .map_err(|_error| invalid_error("ODS content span allocation failed"))?;
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
                open.try_reserve(1)
                    .map_err(|_error| invalid_error("ODS content stack allocation failed"))?;
                open.push(index);
            },
            Event::Empty(element) => {
                if spans.len() >= MAX_ELEMENTS {
                    return invalid("ODS content element limit exceeded");
                }
                spans
                    .try_reserve(1)
                    .map_err(|_error| invalid_error("ODS content span allocation failed"))?;
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
            resolved_attribute(xml, spans, *index, TABLE, "name")
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

fn select_sheet_bounded(xml: &str, spans: &[Span], name: &str) -> Result<usize> {
    let spreadsheet = one(spans, OFFICE, "spreadsheet")?;
    let mut selected = None;
    let mut count = 0usize;
    for (index, span) in spans.iter().enumerate() {
        if span.parent != Some(spreadsheet) || !is_element(span, TABLE, "table") {
            continue;
        }
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS sheet count overflows"))?;
        if count >= MAX_SHEET_MOVE_SHEETS {
            return invalid(format!(
                "ODS sheet transfer requires fewer than {MAX_SHEET_MOVE_SHEETS} worksheets"
            ));
        }
        if resolved_attribute(xml, spans, index, TABLE, "name")?.as_deref() == Some(name) {
            if selected.replace(index).is_some() {
                return invalid("ODS sheet selector is ambiguous");
            }
        }
    }
    selected.ok_or_else(|| invalid_error(format!("ODS sheet '{name}' was not found")))
}

fn resolved_attribute(
    xml: &str,
    spans: &[Span],
    index: usize,
    namespace: &str,
    local: &str,
) -> Result<Option<String>> {
    let span = spans
        .get(index)
        .ok_or_else(|| invalid_error("ODS resolved attribute element is missing"))?;
    let mut ancestors = Vec::new();
    let mut parent = span.parent;
    while let Some(parent_index) = parent {
        if ancestors.len() >= MAX_SHEET_COPY_DEPTH {
            return invalid(format!(
                "ODS resolved attribute ancestor depth exceeds the {MAX_SHEET_COPY_DEPTH} limit"
            ));
        }
        ancestors
            .try_reserve(1)
            .map_err(|_error| invalid_error("ODS resolved attribute ancestor allocation failed"))?;
        let ancestor = spans
            .get(parent_index)
            .ok_or_else(|| invalid_error("ODS resolved attribute ancestor is missing"))?;
        ancestors.push(ancestor);
        parent = ancestor.parent;
    }
    let mut wrapped = String::new();
    for ancestor in ancestors.iter().rev() {
        bounded_append(
            &mut wrapped,
            xml.get(ancestor.start..ancestor.tag_end)
                .ok_or_else(|| invalid_error("ODS resolved attribute ancestor tag is invalid"))?,
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
    }
    bounded_append(
        &mut wrapped,
        xml.get(span.start..span.tag_end)
            .ok_or_else(|| invalid_error("ODS resolved attribute start tag is invalid"))?,
        crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
    )?;
    if !wrapped.ends_with("/>") {
        bounded_append(
            &mut wrapped,
            "</",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            start_qname(xml, span)?,
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            ">",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
    }
    for ancestor in &ancestors {
        bounded_append(
            &mut wrapped,
            "</",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            start_qname(xml, ancestor)?,
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
        bounded_append(
            &mut wrapped,
            ">",
            crate::worksheet::validation::MAX_CONTENT_XML_BYTES,
        )?;
    }
    let mut reader = NsReader::from_str(&wrapped);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let (element_namespace, event) =
            reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| {
                    invalid_error(format!("invalid ODS resolved attribute XML: {error}"))
                })?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if matches!(
                    &element_namespace,
                    ResolveResult::Bound(Namespace(uri))
                        if span.namespace.as_deref() == std::str::from_utf8(uri).ok()
                ) && element.local_name().as_ref() == span.local.as_bytes() =>
            {
                for raw in element.attributes().with_checks(true) {
                    let raw = raw.map_err(|error| {
                        invalid_error(format!("invalid ODS resolved attribute: {error}"))
                    })?;
                    let (resolved, attribute_local) = reader.resolver().resolve_attribute(raw.key);
                    if matches!(
                        resolved,
                        ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes()
                    ) && attribute_local.as_ref() == local.as_bytes()
                    {
                        return raw
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| {
                                invalid_error(format!(
                                    "invalid ODS resolved attribute value: {error}"
                                ))
                            })
                            .and_then(|value| {
                                copy_bounded_string(value.as_ref(), "resolved ODS attribute value")
                                    .map(Some)
                            });
                    }
                }
                return Ok(None);
            },
            Event::Eof => return invalid("ODS resolved attribute element was not found"),
            _ => {},
        }
        buffer.clear();
    }
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

fn bounded_children(
    spans: &[Span],
    parent: usize,
    namespace: &str,
    local: &str,
    limit: usize,
) -> Result<Vec<usize>> {
    let mut result = Vec::new();
    for (index, span) in spans.iter().enumerate() {
        if span.parent != Some(parent) || !is_element(span, namespace, local) {
            continue;
        }
        if result.len() >= limit {
            return invalid(format!("ODS {local} child count exceeds the {limit} limit"));
        }
        result
            .try_reserve(1)
            .map_err(|_error| invalid_error(format!("ODS {local} child allocation failed")))?;
        result.push(index);
    }
    Ok(result)
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

fn is_self_closing_span(xml: &str, span: &Span) -> Result<bool> {
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS element tag span is invalid"))?;
    Ok(tag.trim_end().ends_with("/>"))
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
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|_error| invalid_error(format!("ODS {label} allocation failed")))?;
    owned.extend_from_slice(bytes);
    String::from_utf8(owned).map_err(|_error| invalid_error(format!("ODS {label} is not UTF-8")))
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

#[cfg(test)]
mod scalar_transfer_tests {
    use super::{TABLE, audit_plain_scalar_sheet, children, one, scan};
    use litchi_core::Result;

    const CONTENT_PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet>"#;
    const CONTENT_SUFFIX: &str = r#"</office:spreadsheet></office:body></office:document-content>"#;

    fn scalar_sheet_is_refused(table: &str) -> Result<()> {
        let xml = format!("{CONTENT_PREFIX}{table}{CONTENT_SUFFIX}");
        let spans = scan(&xml)?;
        let spreadsheet = one(&spans, super::OFFICE, "spreadsheet")?;
        let table = children(&spans, spreadsheet, TABLE, "table")
            .into_iter()
            .next()
            .ok_or_else(|| super::invalid_error("test table is missing"))?;
        assert!(audit_plain_scalar_sheet(&xml, &spans, table, spreadsheet).is_err());
        Ok(())
    }

    #[test]
    fn first_value_one_duplicate_row_repetition_is_refused() -> Result<()> {
        scalar_sheet_is_refused(
            r#"<t:table t:name="Source"><t:table-row t:number-rows-repeated="1" t:number-rows-repeated="1"><t:table-cell/></t:table-row></t:table>"#,
        )
    }

    #[test]
    fn first_value_one_duplicate_cell_repetition_is_refused() -> Result<()> {
        scalar_sheet_is_refused(
            r#"<t:table t:name="Source"><t:table-row><t:table-cell t:number-columns-repeated="1" t:number-columns-repeated="1"/></t:table-row></t:table>"#,
        )
    }

    #[test]
    fn malformed_date_and_time_lexicals_are_refused() -> Result<()> {
        for table in [
            r#"<t:table t:name="Source"><t:table-row><t:table-cell office:value-type="date" office:date-value="2026-02-30"/></t:table-row></t:table>"#,
            r#"<t:table t:name="Source"><t:table-row><t:table-cell office:value-type="time" office:time-value="not-a-duration"/></t:table-row></t:table>"#,
        ] {
            scalar_sheet_is_refused(table)?;
        }
        Ok(())
    }
}
