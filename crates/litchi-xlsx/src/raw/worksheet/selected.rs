//! Internal one-cell `SpreadsheetML` streaming selection.
//!
//! This is a source-reader foundation, not a public `SourceWorksheet` route.
//! The scanner consumes the committed MCE stream together with the x14ac raw
//! observer in one pass and publishes an eligible result only after XML EOF.
//! A [`ScanOutcome::NotEligible`] result asserts only that streaming
//! XML/MCE/raw validation reached EOF successfully; it is not full worksheet
//! semantic validation. Callers must discard it and fall back to the existing
//! materialized worksheet parser.
//!
//! The implementation retains at most the requested cell, one active-cell
//! lexical scratch area, and bounded parser state. It makes no fixed-memory,
//! RSS, or OOM-safety claim. quick-xml, the MCE stream, and observer callback
//! allocations are outside this scanner's bounded-state description.

use std::io::BufRead;

use litchi_ooxml_common::mce::{Capabilities, SemanticElement, SemanticEvent, StreamLimits};
use litchi_sheet::Cell as Address;

use super::super::formula::Range as FormulaRange;
use super::super::namespace::{SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE};
use super::super::strings::decode_spreadsheet_text;
use super::model::{MAX_CELL_CHARACTERS, MAX_ENCODED_CELL_BYTES, MAX_FORMULA_CHARACTERS};
use super::{parse_a1, parse_one_based_row, x14ac};
use crate::cell::{Cell, ErrorValue, Number, Text, Value};
use crate::error::{Result, allocation, invalid};
use crate::formula::{Cache, Formula, Kind};

const MAX_XML_DEPTH: usize = 256;
const X14AC_NAMESPACE: &[u8] = x14ac::NAMESPACE;

/// One physical cell selection result.
///
/// `None` means that the requested coordinate has no stored `<c>` record;
/// `Some(Cell::Empty)` is an explicitly stored empty cell.
#[derive(Debug)]
pub struct SelectedCell {
    /// Requested coordinate, retained so later source readers can bind the
    /// result without reconstructing the selector.
    pub address: Address,
    /// Stored semantic cell, or `None` when the coordinate has no `<c>`
    /// record. An explicit empty record is represented by `Cell::Empty`.
    pub cell: Option<Cell>,
}

/// Why a worksheet was structurally valid but outside the scanner's first
/// parity slice. The caller must use the eager materialized parser for these
/// cases.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum NotEligibleReason {
    /// The worksheet contains a structure outside the first streaming slice.
    UnsupportedStructure,
    /// Merge anchors and covered-cell semantics require worksheet-wide state.
    MergeSemantics,
    /// Resolving a shared-string index requires the workbook string table.
    SharedStrings,
    /// Styles or inherited row/column formatting are not selected-cell safe.
    Styles,
    /// Shared, array, data-table, or otherwise deferred formula semantics.
    FormulaSemantics,
    /// Rich inline runs are not flattened by this scalar-only scanner.
    RichInlineText,
    /// The cell type is not part of the supported scalar value slice.
    UnsupportedCellType,
    /// The source order is not strictly ascending for eligible selection.
    Ordering,
    /// A text payload contains an XML reference unavailable to this observer.
    GeneralReference,
}

/// Result of one bounded selected-cell stream.
#[derive(Debug)]
pub enum ScanOutcome {
    /// The selected coordinate was scanned with the first-slice semantics.
    Eligible(SelectedCell),
    /// Streaming XML/MCE/raw validation succeeded, but full worksheet
    /// semantics were deferred; the caller must use the materialized parser.
    NotEligible(NotEligibleReason),
}

/// Result returned by the selected worksheet stream.
///
/// MCE, XML, raw x14ac, input, and allocation failures remain separate typed
/// stream errors; `NotEligible` is carried inside the successful result.
pub type StreamResult<T> =
    std::result::Result<T, litchi_ooxml_common::mce::StreamError<crate::Error, crate::Error>>;

/// Scan one requested coordinate through a source worksheet stream.
///
/// The returned eligible value is published only after the shared MCE/XML
/// stream reaches EOF. A successful [`ScanOutcome::NotEligible`] outcome is
/// not a semantic success: it asserts only successful streaming XML/MCE/raw
/// validation, and callers are required to fall back to the existing
/// materialized parser. Input, XML/MCE, raw x14ac, and allocation failures
/// remain typed stream errors and are never converted to `NotEligible`.
#[expect(
    clippy::result_large_err,
    reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
)]
pub fn scan(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    requested: Address,
) -> StreamResult<ScanOutcome> {
    let mut scanner = Scanner::new(requested);
    let result = x14ac::capture_stream_with_active(
        input,
        capabilities,
        limits,
        x14ac::RowMode::ValidateOnly,
        |event| scanner.event(event),
    );
    match result {
        Ok(_) => {
            scanner
                .finish()
                .map_err(|error| litchi_ooxml_common::mce::StreamError::Callback {
                    raw_error: None,
                    active_error: Some(error),
                })
        },
        Err(error) => Err(error),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    Worksheet,
    SheetData,
    Row,
    Cell,
    Formula,
    Value,
    Inline,
    InlineText,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellKind {
    Untyped,
    Numeric,
    Boolean,
    Error,
    String,
    InlineString,
    SharedString,
    Date,
    Unknown,
}

impl CellKind {
    fn from_attribute(value: Option<&str>) -> Self {
        match value {
            None => Self::Untyped,
            Some("n") => Self::Numeric,
            Some("b") => Self::Boolean,
            Some("e") => Self::Error,
            Some("str") => Self::String,
            Some("inlineStr") => Self::InlineString,
            Some("s") => Self::SharedString,
            Some("d") => Self::Date,
            Some(_) => Self::Unknown,
        }
    }
}

#[derive(Debug)]
struct PendingCell {
    selected: bool,
    kind: CellKind,
    saw_value: bool,
    value: String,
    value_bytes: usize,
    value_characters: usize,
    formula: Option<String>,
    formula_characters: usize,
    saw_inline: bool,
    inline: String,
    inline_bytes: usize,
    inline_characters: usize,
    saw_inline_simple: bool,
    saw_inline_run: bool,
}

impl PendingCell {
    fn new(selected: bool, kind: CellKind) -> Self {
        Self {
            selected,
            kind,
            saw_value: false,
            value: String::new(),
            value_bytes: 0,
            value_characters: 0,
            formula: None,
            formula_characters: 0,
            saw_inline: false,
            inline: String::new(),
            inline_bytes: 0,
            inline_characters: 0,
            saw_inline_simple: false,
            saw_inline_run: false,
        }
    }

    fn append_value(&mut self, text: &str) -> Result<()> {
        append_bounded(
            &mut self.value,
            &mut self.value_bytes,
            &mut self.value_characters,
            text,
            MAX_ENCODED_CELL_BYTES,
            MAX_CELL_CHARACTERS,
            "selected worksheet cell value",
        )
    }

    fn append_inline(&mut self, text: &str) -> Result<()> {
        append_bounded(
            &mut self.inline,
            &mut self.inline_bytes,
            &mut self.inline_characters,
            text,
            MAX_ENCODED_CELL_BYTES,
            MAX_CELL_CHARACTERS,
            "selected worksheet inline text",
        )
    }

    fn append_formula(&mut self, text: &str) -> Result<()> {
        let characters = text.chars().count();
        self.formula_characters = self
            .formula_characters
            .checked_add(characters)
            .filter(|length| *length <= MAX_FORMULA_CHARACTERS)
            .ok_or_else(|| {
                invalid(format!(
                    "selected worksheet formula exceeds {MAX_FORMULA_CHARACTERS} characters"
                ))
            })?;
        let formula = self.formula.get_or_insert_with(String::new);
        formula
            .try_reserve(text.len())
            .map_err(|source| allocation("selected worksheet formula", source))?;
        formula.push_str(text);
        Ok(())
    }
}

fn append_bounded(
    target: &mut String,
    bytes: &mut usize,
    characters: &mut usize,
    text: &str,
    max_bytes: usize,
    max_characters: usize,
    resource: &'static str,
) -> Result<()> {
    *bytes = bytes
        .checked_add(text.len())
        .filter(|length| *length <= max_bytes)
        .ok_or_else(|| invalid(format!("{resource} is too large")))?;
    *characters = characters
        .checked_add(text.chars().count())
        .filter(|length| *length <= max_characters)
        .ok_or_else(|| invalid(format!("{resource} exceeds {max_characters} characters")))?;
    target
        .try_reserve(text.len())
        .map_err(|source| allocation(resource, source))?;
    target.push_str(text);
    Ok(())
}

#[derive(Debug)]
struct Scanner {
    requested: Address,
    target_row: u32,
    target_column: u32,
    stack: Vec<Frame>,
    row: Option<(u32, u32)>,
    previous_row: u32,
    root_seen: bool,
    root_closed: bool,
    seen_sheet_data: bool,
    pending_cell: Option<PendingCell>,
    selected: Option<Cell>,
    not_eligible: Option<NotEligibleReason>,
}

impl Scanner {
    fn new(requested: Address) -> Self {
        Self {
            target_row: requested.row().get() + 1,
            target_column: requested.column().get() + 1,
            requested,
            stack: Vec::new(),
            row: None,
            previous_row: 0,
            root_seen: false,
            root_closed: false,
            seen_sheet_data: false,
            pending_cell: None,
            selected: None,
            not_eligible: None,
        }
    }

    fn finish(self) -> Result<ScanOutcome> {
        if let Some(reason) = self.not_eligible {
            return Ok(ScanOutcome::NotEligible(reason));
        }
        if !self.root_seen || !self.root_closed || !self.stack.is_empty() {
            return Err(invalid(
                "worksheet XML has an incomplete SpreadsheetML worksheet root",
            ));
        }
        if !self.seen_sheet_data {
            return Err(invalid("worksheet XML is missing required sheetData"));
        }
        Ok(ScanOutcome::Eligible(SelectedCell {
            address: self.requested,
            cell: self.selected,
        }))
    }

    fn event(&mut self, event: &SemanticEvent<'_>) -> Result<()> {
        if self.not_eligible.is_some() {
            return Ok(());
        }
        match event {
            SemanticEvent::Start(element) => self.start(element, false),
            SemanticEvent::Empty(element) => self.start(element, true),
            SemanticEvent::End(element) => self.end(element),
            SemanticEvent::Text(text) | SemanticEvent::CData(text) => self.text(text.text()),
            SemanticEvent::GeneralRef(_) => {
                if self.in_payload() {
                    self.mark(NotEligibleReason::GeneralReference);
                }
                Ok(())
            },
            SemanticEvent::Comment(_) | SemanticEvent::Decl(_) => Ok(()),
            _ => Ok(()),
        }
    }

    fn start(&mut self, element: &SemanticElement<'_>, empty: bool) -> Result<()> {
        if self.stack.is_empty() {
            if self.root_seen || self.root_closed {
                return Err(invalid("worksheet XML has multiple worksheet roots"));
            }
            if !is_spreadsheetml_element(element, "worksheet") {
                return Err(invalid(
                    "worksheet XML must have one SpreadsheetML worksheet root",
                ));
            }
            self.root_seen = true;
            self.validate_attributes(element, &[], false)?;
            if empty {
                self.root_closed = true;
            } else if self.not_eligible.is_none() {
                self.push(Frame::Worksheet)?;
            }
            return Ok(());
        }

        let parent = *self
            .stack
            .last()
            .ok_or_else(|| invalid("selected worksheet parser lost its parent context"))?;
        let frame = match parent {
            Frame::Worksheet => self.start_worksheet_child(element)?,
            Frame::SheetData => self.start_sheet_data_child(element)?,
            Frame::Row => self.start_row_child(element)?,
            Frame::Cell => self.start_cell_child(element)?,
            Frame::Formula | Frame::Value | Frame::InlineText => {
                self.mark(NotEligibleReason::UnsupportedStructure);
                None
            },
            Frame::Inline => self.start_inline_child(element)?,
        };
        if let Some(frame) = frame {
            if empty {
                self.finish_empty(frame)
            } else if self.not_eligible.is_none() {
                self.push(frame)
            } else {
                Ok(())
            }
        } else {
            Ok(())
        }
    }

    fn start_worksheet_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "sheetData") {
            if self.seen_sheet_data {
                return Err(invalid("worksheet has duplicate sheetData"));
            }
            self.validate_attributes(element, &[], false)?;
            self.seen_sheet_data = true;
            return Ok(Some(Frame::SheetData));
        }
        if is_spreadsheetml_element(element, "mergeCells")
            || is_spreadsheetml_element(element, "mergeCell")
        {
            self.mark(NotEligibleReason::MergeSemantics);
        } else if is_spreadsheetml_element(element, "cols")
            || is_spreadsheetml_element(element, "col")
        {
            self.mark(NotEligibleReason::Styles);
        } else {
            self.mark(NotEligibleReason::UnsupportedStructure);
        }
        Ok(None)
    }

    fn start_sheet_data_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "row") {
            self.start_row(element)
        } else {
            if is_spreadsheetml_element(element, "mergeCells")
                || is_spreadsheetml_element(element, "mergeCell")
            {
                self.mark(NotEligibleReason::MergeSemantics);
            } else {
                self.mark(NotEligibleReason::UnsupportedStructure);
            }
            Ok(None)
        }
    }

    fn start_row_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "c") {
            return self.start_cell(element);
        }
        self.mark(NotEligibleReason::UnsupportedStructure);
        Ok(None)
    }

    fn start_cell_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "f") {
            self.start_formula(element)?;
            return Ok(Some(Frame::Formula));
        }
        if is_spreadsheetml_element(element, "v") {
            self.validate_attributes(element, &[], false)?;
            let cell = self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("worksheet value outside a cell"))?;
            if cell.saw_value {
                return Err(invalid("duplicate worksheet cell value"));
            }
            cell.saw_value = true;
            return Ok(Some(Frame::Value));
        }
        if is_spreadsheetml_element(element, "is") {
            self.validate_attributes(element, &[], false)?;
            let cell = self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("inline string outside a worksheet cell"))?;
            if cell.saw_inline {
                return Err(invalid("duplicate worksheet inline string"));
            }
            cell.saw_inline = true;
            return Ok(Some(Frame::Inline));
        }
        self.mark(NotEligibleReason::UnsupportedStructure);
        Ok(None)
    }

    fn start_inline_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "t") {
            self.validate_attributes(element, &[], false)?;
            let cell = self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("inline text outside a worksheet cell"))?;
            if cell.saw_inline_simple || cell.saw_inline_run {
                return Err(invalid("inline string mixes or duplicates text"));
            }
            cell.saw_inline_simple = true;
            return Ok(Some(Frame::InlineText));
        }
        if is_spreadsheetml_element(element, "r") {
            self.mark(NotEligibleReason::RichInlineText);
        } else {
            self.mark(NotEligibleReason::UnsupportedStructure);
        }
        Ok(None)
    }

    fn start_row(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.validate_attributes(element, &["r"], true)?;
        let number = match unqualified_attribute(element, "r") {
            Some(value) => parse_one_based_row(value)?,
            None => self
                .previous_row
                .checked_add(1)
                .filter(|value| *value <= litchi_sheet::ROWS)
                .ok_or_else(|| invalid("inferred worksheet row exceeds the spreadsheet grid"))?,
        };
        if self.previous_row != 0 {
            if number < self.previous_row {
                self.mark(NotEligibleReason::Ordering);
                return Ok(None);
            }
            if number == self.previous_row {
                return Err(invalid(format!("duplicate worksheet row {number}")));
            }
        }
        self.previous_row = number;
        self.row = Some((number, 0));
        if self.not_eligible.is_some() {
            Ok(None)
        } else {
            Ok(Some(Frame::Row))
        }
    }

    fn start_cell(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.validate_attributes(element, &["r", "t"], false)?;
        let (row, last_column) = self
            .row
            .ok_or_else(|| invalid("worksheet cell outside a row"))?;
        let column = match unqualified_attribute(element, "r") {
            Some(reference) => {
                let (reference_row, column) = parse_a1(reference)?;
                if reference_row != row {
                    return Err(invalid(format!(
                        "cell reference '{reference}' does not belong to row {row}"
                    )));
                }
                column
            },
            None => last_column
                .checked_add(1)
                .filter(|value| *value <= litchi_sheet::COLUMNS)
                .ok_or_else(|| invalid("inferred worksheet column exceeds the grid"))?,
        };
        if last_column != 0 {
            if column < last_column {
                self.mark(NotEligibleReason::Ordering);
                return Ok(None);
            }
            if column == last_column {
                return Err(invalid(format!(
                    "duplicate worksheet cell at row {row}, column {column}"
                )));
            }
        }
        self.row = Some((row, column));

        let kind = CellKind::from_attribute(unqualified_attribute(element, "t"));
        match kind {
            CellKind::SharedString => self.mark(NotEligibleReason::SharedStrings),
            CellKind::Date | CellKind::Unknown => self.mark(NotEligibleReason::UnsupportedCellType),
            CellKind::Untyped
            | CellKind::Numeric
            | CellKind::Boolean
            | CellKind::Error
            | CellKind::String
            | CellKind::InlineString => {},
        }
        if self.not_eligible.is_some() {
            return Ok(None);
        }
        let selected = row == self.target_row && column == self.target_column;
        self.pending_cell = Some(PendingCell::new(selected, kind));
        Ok(Some(Frame::Cell))
    }

    fn start_formula(&mut self, element: &SemanticElement<'_>) -> Result<()> {
        self.validate_attributes(element, &["t", "ref", "si", "bx"], false)?;
        let formula_type = unqualified_attribute(element, "t").unwrap_or("normal");
        if let Some(reference) = unqualified_attribute(element, "ref") {
            FormulaRange::parse(reference)?;
        }
        if let Some(index) = unqualified_attribute(element, "si") {
            index
                .parse::<u32>()
                .map_err(|_source| invalid(format!("invalid shared formula index '{index}'")))?;
            self.mark(NotEligibleReason::FormulaSemantics);
        }
        if let Some(bx) = unqualified_attribute(element, "bx") {
            if !matches!(bx, "0" | "1" | "true" | "false") {
                return Err(invalid(format!("invalid formula bx '{bx}'")));
            }
            if matches!(bx, "1" | "true") {
                return Err(invalid("Office requires formula bx to be false"));
            }
        }
        match formula_type {
            "normal" => {},
            "array" | "dataTable" | "shared" => self.mark(NotEligibleReason::FormulaSemantics),
            _ => self.mark(NotEligibleReason::FormulaSemantics),
        }
        if self.not_eligible.is_some() {
            return Ok(());
        }
        let cell = self
            .pending_cell
            .as_mut()
            .ok_or_else(|| invalid("worksheet formula outside a cell"))?;
        if cell.formula.is_some() {
            return Err(invalid("duplicate worksheet formula"));
        }
        cell.formula = Some(String::new());
        Ok(())
    }

    fn end(&mut self, element: &litchi_ooxml_common::mce::SemanticEnd<'_>) -> Result<()> {
        let frame = self
            .stack
            .pop()
            .ok_or_else(|| invalid("worksheet XML has an unexpected end"))?;
        if !matches_frame(frame, element) {
            return Err(invalid("worksheet XML has an unexpected semantic end"));
        }
        match frame {
            Frame::Worksheet => self.root_closed = true,
            Frame::SheetData => {},
            Frame::Row => self.finish_row()?,
            Frame::Cell => self.finish_cell()?,
            Frame::Formula | Frame::Value | Frame::InlineText | Frame::Inline => {},
        }
        Ok(())
    }

    fn finish_empty(&mut self, frame: Frame) -> Result<()> {
        match frame {
            Frame::Worksheet => self.root_closed = true,
            Frame::SheetData => {},
            Frame::Row => self.finish_row()?,
            Frame::Cell => self.finish_cell()?,
            Frame::Formula | Frame::Value | Frame::InlineText | Frame::Inline => {},
        }
        Ok(())
    }

    fn finish_row(&mut self) -> Result<()> {
        if self.pending_cell.is_some() {
            return Err(invalid("unterminated worksheet cell"));
        }
        self.row = None;
        Ok(())
    }

    fn finish_cell(&mut self) -> Result<()> {
        let mut cell = self
            .pending_cell
            .take()
            .ok_or_else(|| invalid("missing worksheet cell"))?;
        if cell.saw_inline && cell.saw_value {
            return Err(invalid(
                "worksheet cell contains both inline text and a value",
            ));
        }
        if cell.saw_inline && !matches!(cell.kind, CellKind::Untyped | CellKind::InlineString) {
            return Err(invalid("inline string has a non-inline cell type"));
        }
        if matches!(cell.kind, CellKind::InlineString) {
            cell.saw_inline = true;
        }
        let formula = if let Some(formula) = cell.formula.take() {
            if formula.trim().is_empty() || formula.trim_start().starts_with('=') {
                return Err(invalid(
                    "formula expression must be non-empty and omit the leading '='",
                ));
            }
            let formula = Formula::new(formula.clone())?;
            Some((formula.text().to_owned(), formula))
        } else {
            None
        };
        let inline = if cell.saw_inline {
            if formula.is_some() {
                return Err(invalid("formula cell cannot contain an inline string"));
            }
            let value = decode_spreadsheet_text(&cell.inline)?;
            if value.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "inline string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
            Some(value)
        } else {
            None
        };
        let value = if cell.saw_value {
            Some(parse_value(cell.kind, &cell.value)?)
        } else {
            None
        };
        let semantic = if let Some((formula_text, _formula)) = formula {
            let cached = value.flatten().map(Cache::stored);
            Cell::Formula(Formula::parsed(formula_text, Kind::Scalar, cached))
        } else if let Some(inline) = inline {
            Cell::Value(Value::Text(inline.into()))
        } else if let Some(value) = value.flatten() {
            Cell::Value(value)
        } else if matches!(cell.kind, CellKind::String) {
            Cell::Value(Value::Text(Text::from("")))
        } else {
            Cell::Empty
        };
        if cell.selected {
            self.selected = Some(semantic);
        }
        Ok(())
    }

    fn text(&mut self, text: &str) -> Result<()> {
        let Some(frame) = self.stack.last().copied() else {
            return Err(invalid("worksheet text appears outside its root"));
        };
        match frame {
            Frame::Formula => self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("worksheet formula text outside a cell"))?
                .append_formula(text),
            Frame::Value => self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("worksheet value text outside a cell"))?
                .append_value(text),
            Frame::InlineText => self
                .pending_cell
                .as_mut()
                .ok_or_else(|| invalid("worksheet inline text outside a cell"))?
                .append_inline(text),
            Frame::Worksheet | Frame::SheetData | Frame::Row | Frame::Cell | Frame::Inline => {
                if text.trim().is_empty() {
                    Ok(())
                } else {
                    self.mark(NotEligibleReason::UnsupportedStructure);
                    Ok(())
                }
            },
        }
    }

    fn in_payload(&self) -> bool {
        matches!(
            self.stack.last(),
            Some(Frame::Formula | Frame::Value | Frame::InlineText)
        )
    }

    fn validate_attributes(
        &mut self,
        element: &SemanticElement<'_>,
        allowed: &[&str],
        allow_x14ac: bool,
    ) -> Result<()> {
        let mut seen = [false; 4];
        for attribute in &element.attributes {
            let namespace = attribute.expanded_name.namespace.as_bytes();
            let local = attribute.expanded_name.local_name.as_str();
            if namespace.is_empty() {
                if let Some(index) = allowed.iter().position(|name| *name == local) {
                    if seen[index] {
                        return Err(invalid(format!("duplicate worksheet attribute '{local}'")));
                    }
                    seen[index] = true;
                } else {
                    self.mark(NotEligibleReason::UnsupportedStructure);
                }
            } else if !(allow_x14ac && namespace == X14AC_NAMESPACE && local == "dyDescent") {
                self.mark(NotEligibleReason::UnsupportedStructure);
            }
        }
        Ok(())
    }

    fn push(&mut self, frame: Frame) -> Result<()> {
        if self.stack.len() >= MAX_XML_DEPTH {
            return Err(invalid(format!(
                "worksheet XML exceeds {MAX_XML_DEPTH} levels"
            )));
        }
        self.stack
            .try_reserve(1)
            .map_err(|source| allocation("selected worksheet element stack", source))?;
        self.stack.push(frame);
        Ok(())
    }

    fn mark(&mut self, reason: NotEligibleReason) {
        if self.not_eligible.is_none() {
            self.not_eligible = Some(reason);
        }
    }
}

fn parse_value(kind: CellKind, value: &str) -> Result<Option<Value>> {
    match kind {
        CellKind::Untyped | CellKind::Numeric => {
            if value.trim().is_empty() {
                Ok(None)
            } else {
                Number::new(value.to_owned().into_boxed_str())
                    .map(Value::Number)
                    .map(Some)
            }
        },
        CellKind::String => {
            let value = decode_spreadsheet_text(value)?;
            if value.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "worksheet string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
            Ok(Some(Value::Text(value.into())))
        },
        CellKind::Boolean => match value.trim() {
            "1" | "true" => Ok(Some(Value::Bool(true))),
            "0" | "false" => Ok(Some(Value::Bool(false))),
            other => Err(invalid(format!("invalid worksheet boolean '{other}'"))),
        },
        CellKind::Error => Ok(Some(Value::Error(ErrorValue::parse(value)))),
        CellKind::InlineString => Err(invalid(
            "inline-string cell stores text in an is element, not v",
        )),
        CellKind::SharedString | CellKind::Date | CellKind::Unknown => {
            Err(invalid("unsupported worksheet cell value type"))
        },
    }
}

fn unqualified_attribute<'a>(element: &'a SemanticElement<'_>, local: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| {
            attribute.expanded_name.namespace.is_empty()
                && attribute.expanded_name.local_name == local
        })
        .map(|attribute| attribute.decoded_value.as_ref())
}

fn is_spreadsheetml_element(element: &SemanticElement<'_>, local: &str) -> bool {
    (element.expanded_name.namespace.as_bytes() == SPREADSHEETML_NAMESPACE
        || element.expanded_name.namespace.as_bytes() == STRICT_SPREADSHEETML_NAMESPACE)
        && element.expanded_name.local_name == local
}

fn matches_frame(frame: Frame, element: &litchi_ooxml_common::mce::SemanticEnd<'_>) -> bool {
    let local = element.expanded_name.local_name.as_str();
    let namespace = element.expanded_name.namespace.as_bytes();
    let expected = match frame {
        Frame::Worksheet => "worksheet",
        Frame::SheetData => "sheetData",
        Frame::Row => "row",
        Frame::Cell => "c",
        Frame::Formula => "f",
        Frame::Value => "v",
        Frame::Inline => "is",
        Frame::InlineText => "t",
    };
    (namespace == SPREADSHEETML_NAMESPACE || namespace == STRICT_SPREADSHEETML_NAMESPACE)
        && local == expected
}
