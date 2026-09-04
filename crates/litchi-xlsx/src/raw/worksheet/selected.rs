//! Internal sparse `SpreadsheetML` worksheet-range streaming selection.
//!
//! This is a source-reader foundation, not a public `SourceWorksheet` route.
//! The scanner consumes the committed MCE stream together with the x14ac raw
//! observer in one pass and publishes an eligible result only after XML EOF.
//! A [`ScanOutcome::NotEligible`] result asserts only that streaming
//! XML/MCE/raw validation reached EOF successfully; it is not full worksheet
//! semantic validation. Callers must discard it and fall back to the existing
//! materialized worksheet parser.
//!
//! The implementation retains only physical records inside the requested
//! rectangle, one active-cell lexical scratch area, and bounded parser state.
//! It makes no fixed-memory, RSS, or OOM-safety claim. quick-xml, the MCE
//! stream, and observer callback allocations are outside this scanner's
//! bounded-state description; event and input limits provide the aggregate
//! bound for the stream.

use std::io::BufRead;

use litchi_ooxml_common::mce::{Capabilities, SemanticElement, SemanticEvent, StreamLimits};
use litchi_sheet::{Cell as Address, Rect};
use quick_xml::events::BytesRef;

use super::super::formula::Range as FormulaRange;
use super::super::namespace::{SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE};
use super::super::strings::decode_spreadsheet_text;
use super::model::{
    MAX_CELL_CHARACTERS, MAX_CELL_STYLE, MAX_ENCODED_CELL_BYTES, MAX_FORMULA_CHARACTERS,
    merge_successor,
};
use super::{parse_a1, parse_one_based_row, x14ac};
use crate::cell::{Cell, ErrorValue, Number, Text, Value};
use crate::error::{Result, allocation, invalid};
use crate::formula::{Cache, Formula, Kind};

const MAX_XML_DEPTH: usize = 256;
const MAX_GENERAL_REFERENCE_TOKEN_BYTES: usize = 12;
const MAX_MERGED_RANGES: usize = 16_384;
// The 12-byte cap keeps 8192 * 12 formula bytes and 32767 * 12 value bytes
// below MAX_ENCODED_CELL_BYTES (458738); inline numeric refs stay ineligible
// because `_xHHHH_` decoding could otherwise compose around the encoded bound.
const X14AC_NAMESPACE: &[u8] = x14ac::NAMESPACE;

/// One-coordinate compatibility result from the worksheet range scanner.
///
/// `cell: None` means that the requested coordinate has no stored `<c>` record
/// or that a valid non-formula shared-string record was deferred. The latter
/// is identified by [`SelectedCell::dependencies`] retaining its shared-string
/// index. `Some(Cell::Empty)` is an explicitly stored empty cell. The range
/// API exposes the same distinction directly through [`SelectedRecord`].
#[derive(Debug)]
pub struct SelectedCell {
    /// Requested coordinate, retained so later source readers can bind the
    /// result without reconstructing the selector.
    pub address: Address,
    /// Stored semantic cell. `None` means either that the coordinate has no
    /// `<c>` record or that a valid non-formula shared-string record was
    /// deferred; the dependency metadata identifies the latter. An explicit
    /// empty record is represented by `Cell::Empty`.
    pub cell: Option<Cell>,
    /// Merged range covering the requested coordinate when it is not the
    /// merge anchor. The range is absent for anchors and unmerged cells.
    pub covering_merge: Option<Rect>,
    /// Bounded dependency metadata observed while scanning the complete
    /// worksheet stream.
    pub dependencies: SelectedDependencies,
}

/// One physical cell record retained by a rectangular worksheet selection.
///
/// The range scanner is sparse: absent coordinates do not produce records.
/// An explicit empty `<c>` record is represented by `cell: Some(Cell::Empty)`.
/// A valid non-formula shared-string record is deferred instead, with its
/// index in [`Self::shared_string_index`] and `cell` set to `None`.
#[derive(Debug)]
pub struct SelectedRecord {
    /// Physical worksheet coordinate of the stored `<c>` record.
    pub address: Address,
    /// Stored semantic cell, or `None` for a deferred shared-string record.
    pub cell: Option<Cell>,
    /// Zero-based shared-string index for a deferred `t="s"` record.
    pub shared_string_index: Option<u32>,
}

/// Sparse physical records retained by an eligible rectangular scan.
#[derive(Debug)]
pub struct SelectedCells {
    /// Stored `<c>` records inside the requested rectangle, in source order.
    /// Missing coordinates are intentionally omitted.
    pub cells: Vec<SelectedRecord>,
    /// Dependency metadata observed across the complete worksheet stream.
    pub dependencies: SelectedDependencies,
}

/// Dependency metadata retained by an eligible selected worksheet scan.
///
/// All indexes are zero-based and are retained without resolving workbook
/// parts. `None` means that no matching reference was observed. The maxima
/// include unselected cells, so callers can use them to bound the dependency
/// tables needed by a later materialized read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SelectedDependencies {
    /// Largest shared-string index referenced by any non-formula `t="s"`
    /// cell with a valid `<v>` value.
    pub max_shared_string_index: Option<u32>,
    /// Largest direct cell style index present in any valid `c@s` attribute.
    /// Row and column style inheritance is not represented; it makes the scan
    /// [`NotEligibleReason::Styles`] instead.
    pub max_direct_style_index: Option<u32>,
    /// Shared-string index referenced by the requested cell, when that cell is
    /// a valid non-formula `t="s"` cell. `Some` identifies a deferred selected
    /// cell and is distinct from a missing coordinate or an explicit empty
    /// cell. For a range scan this compatibility field is populated only when
    /// exactly one physical record was selected; use
    /// [`SelectedRecord::shared_string_index`] for per-record indexes.
    pub target_shared_string_index: Option<u32>,
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
    /// A shared-string reference has lexical content outside the bounded
    /// index form retained by the scanner.
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

/// Result of one bounded rectangular worksheet stream.
#[derive(Debug)]
pub enum RangeScanOutcome {
    /// The selected physical records were scanned with the first-slice
    /// semantics.
    Eligible(SelectedCells),
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
pub fn scan_range(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    requested: Rect,
) -> StreamResult<RangeScanOutcome> {
    scan_stream(input, capabilities, limits, requested).map(|(outcome, _)| outcome)
}

#[expect(
    clippy::result_large_err,
    reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
)]
fn scan_stream(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    requested: Rect,
) -> StreamResult<(RangeScanOutcome, Option<Rect>)> {
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

/// Scan one requested coordinate through a source worksheet stream.
///
/// This compatibility wrapper uses the rectangular scanner with a one-cell
/// rectangle. A successful eligible result contains `None` for a missing
/// coordinate, `Some(Cell::Empty)` for an explicit empty record, or the
/// deferred shared-string representation used by the original API. The
/// result is published only after the shared MCE/XML stream reaches EOF.
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
    let (outcome, covering_merge) =
        scan_stream(input, capabilities, limits, Rect::single(requested))?;
    match outcome {
        RangeScanOutcome::NotEligible(reason) => Ok(ScanOutcome::NotEligible(reason)),
        RangeScanOutcome::Eligible(selected) => {
            let mut records = selected.cells;
            let record = match records.len() {
                0 => None,
                1 => records.pop(),
                _ => {
                    return Err(litchi_ooxml_common::mce::StreamError::Callback {
                        raw_error: None,
                        active_error: Some(invalid(
                            "one-cell worksheet selection produced multiple records",
                        )),
                    });
                },
            };
            let mut dependencies = selected.dependencies;
            dependencies.target_shared_string_index = record
                .as_ref()
                .and_then(|record| record.shared_string_index);
            Ok(ScanOutcome::Eligible(SelectedCell {
                address: requested,
                cell: record.and_then(|record| record.cell),
                covering_merge,
                dependencies,
            }))
        },
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    Worksheet,
    Dimension,
    SheetData,
    MergeCells,
    MergeCell,
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

    fn uses_numeric_value_scratch(self) -> bool {
        matches!(self, Self::Untyped | Self::Numeric)
    }
}

#[derive(Debug)]
struct PendingCell {
    address: Address,
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
    saw_inline_simple: bool,
    saw_inline_run: bool,
}

impl PendingCell {
    fn new(address: Address, kind: CellKind, value: String) -> Self {
        Self {
            address,
            kind,
            saw_value: false,
            value,
            value_bytes: 0,
            value_characters: 0,
            formula: None,
            formula_characters: 0,
            saw_inline: false,
            inline: String::new(),
            inline_bytes: 0,
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
        self.inline_bytes = self
            .inline_bytes
            .checked_add(text.len())
            .filter(|length| *length <= MAX_ENCODED_CELL_BYTES)
            .ok_or_else(|| invalid("selected worksheet inline text is too large"))?;
        self.inline
            .try_reserve(text.len())
            .map_err(|source| allocation("selected worksheet inline text", source))?;
        self.inline.push_str(text);
        Ok(())
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
    requested: Rect,
    stack: Vec<Frame>,
    row: Option<(u32, u32)>,
    previous_row: u32,
    root_seen: bool,
    root_closed: bool,
    seen_dimension: bool,
    seen_sheet_data: bool,
    seen_merge_cells: bool,
    merge_window_closed: bool,
    merge_count: Option<u32>,
    merges: Vec<Rect>,
    pending_cell: Option<PendingCell>,
    numeric_value_scratch: String,
    selected: Vec<SelectedRecord>,
    selected_count: usize,
    dependencies: SelectedDependencies,
    not_eligible: Option<NotEligibleReason>,
}

impl Scanner {
    fn new(requested: Rect) -> Self {
        Self {
            requested,
            stack: Vec::new(),
            row: None,
            previous_row: 0,
            root_seen: false,
            root_closed: false,
            seen_dimension: false,
            seen_sheet_data: false,
            seen_merge_cells: false,
            merge_window_closed: false,
            merge_count: None,
            merges: Vec::new(),
            pending_cell: None,
            numeric_value_scratch: String::new(),
            selected: Vec::new(),
            selected_count: 0,
            dependencies: SelectedDependencies::default(),
            not_eligible: None,
        }
    }

    fn finish(self) -> Result<(RangeScanOutcome, Option<Rect>)> {
        if let Some(reason) = self.not_eligible {
            return Ok((RangeScanOutcome::NotEligible(reason), None));
        }
        if !self.root_seen || !self.root_closed || !self.stack.is_empty() {
            return Err(invalid(
                "worksheet XML has an incomplete SpreadsheetML worksheet root",
            ));
        }
        if !self.seen_sheet_data {
            return Err(invalid("worksheet XML is missing required sheetData"));
        }
        if self.selected_count != self.selected.len() {
            return Err(invalid(
                "selected worksheet record count lost synchronization",
            ));
        }
        if self.seen_merge_cells && self.merges.is_empty() {
            return Err(invalid(
                "worksheet mergeCells contains no mergeCell records",
            ));
        }
        if let Some(count) = self.merge_count
            && count
                != u32::try_from(self.merges.len())
                    .map_err(|_source| invalid("worksheet merged-range count does not fit u32"))?
        {
            return Err(invalid(format!(
                "worksheet merged-range count differs from {} records",
                self.merges.len()
            )));
        }
        let index = crate::merge::Index::new(self.merges)?;
        let covering_merge = if self.requested.rows() == 1 && self.requested.columns() == 1 {
            let address = self.requested.start();
            index
                .containing(address)
                .filter(|range| range.start() != address)
        } else {
            None
        };
        let target_shared_string_index = (self.selected.len() == 1)
            .then(|| self.selected[0].shared_string_index)
            .flatten();
        let mut dependencies = self.dependencies;
        dependencies.target_shared_string_index = target_shared_string_index;
        Ok((
            RangeScanOutcome::Eligible(SelectedCells {
                cells: self.selected,
                dependencies,
            }),
            covering_merge,
        ))
    }

    fn event(&mut self, event: &SemanticEvent<'_>) -> Result<()> {
        if self.not_eligible.is_some() {
            if let SemanticEvent::Start(element) | SemanticEvent::Empty(element) = event
                && self.stack.last() == Some(&Frame::Worksheet)
            {
                if is_spreadsheetml_element(element, "dimension") {
                    self.validate_dimension_placement()?;
                } else if is_spreadsheetml_element(element, "mergeCells") {
                    self.validate_merge_cells_placement()?;
                }
            }
            return Ok(());
        }
        match event {
            SemanticEvent::Start(element) => self.start(element, false),
            SemanticEvent::Empty(element) => self.start(element, true),
            SemanticEvent::End(element) => self.end(element),
            SemanticEvent::Text(text) | SemanticEvent::CData(text) => self.text(text.text()),
            SemanticEvent::GeneralRef(reference) if self.in_payload() => {
                self.general_reference(reference)
            },
            SemanticEvent::GeneralRef(reference)
                if self.stack.last() == Some(&Frame::Dimension) =>
            {
                self.general_reference(reference)
            },
            SemanticEvent::GeneralRef(_) if self.in_merge() => {
                self.mark(NotEligibleReason::MergeSemantics);
                Ok(())
            },
            SemanticEvent::GeneralRef(_) => Ok(()),
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
            Frame::Dimension => self.start_dimension_child(element)?,
            Frame::SheetData => self.start_sheet_data_child(element)?,
            Frame::MergeCells => self.start_merge_cells_child(element)?,
            Frame::MergeCell => self.start_merge_cell_child(element)?,
            Frame::Row => self.start_row_child(element)?,
            Frame::Cell => self.start_cell_child(element)?,
            Frame::Formula | Frame::Value | Frame::InlineText => {
                if is_merge_markup(element) {
                    return Err(invalid(
                        "worksheet merge markup appears outside its schema context",
                    ));
                }
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
        if is_spreadsheetml_element(element, "dimension") {
            return self.start_dimension(element);
        }
        if is_spreadsheetml_element(element, "sheetData") {
            if self.seen_sheet_data {
                return Err(invalid("worksheet has duplicate sheetData"));
            }
            self.validate_attributes(element, &[], false)?;
            self.seen_sheet_data = true;
            return Ok(Some(Frame::SheetData));
        }
        if is_spreadsheetml_element(element, "mergeCells") {
            self.validate_merge_cells_placement()?;
            self.seen_merge_cells = true;
            return self.start_merge_cells(element);
        }
        if is_spreadsheetml_element(element, "mergeCell") {
            return Err(invalid(
                "worksheet merge markup appears outside its schema context",
            ));
        }
        let local = element.expanded_name.local_name.as_str();
        if self.seen_sheet_data
            && is_spreadsheetml_element(element, local)
            && merge_successor(local.as_bytes())
        {
            self.merge_window_closed = true;
        }
        if is_spreadsheetml_element(element, "cols") || is_spreadsheetml_element(element, "col") {
            self.mark(NotEligibleReason::Styles);
        } else {
            self.mark(NotEligibleReason::UnsupportedStructure);
        }
        Ok(None)
    }

    fn start_dimension(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.validate_dimension_placement()?;
        self.seen_dimension = true;
        self.validate_attributes(element, &["ref"], false)?;
        let reference = unqualified_attribute(element, "ref")
            .ok_or_else(|| invalid("worksheet dimension is missing ref"))?;
        Rect::from_a1(reference).map_err(|error| {
            invalid(format!(
                "invalid worksheet dimension '{reference}': {error}"
            ))
        })?;
        if self.not_eligible.is_some() {
            Ok(None)
        } else {
            Ok(Some(Frame::Dimension))
        }
    }

    fn validate_dimension_placement(&self) -> Result<()> {
        if self.seen_sheet_data {
            return Err(invalid(
                "worksheet dimension appears after column or cell data",
            ));
        }
        if self.seen_dimension {
            return Err(invalid("worksheet has duplicate dimension elements"));
        }
        Ok(())
    }

    fn start_dimension_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_merge_markup(element) {
            return Err(invalid(
                "worksheet merge markup appears outside its schema context",
            ));
        }
        self.mark(NotEligibleReason::UnsupportedStructure);
        Ok(None)
    }

    fn start_sheet_data_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "row") {
            self.start_row(element)
        } else {
            if is_merge_markup(element) {
                return Err(invalid(
                    "worksheet merge markup appears outside its schema context",
                ));
            }
            self.mark(NotEligibleReason::UnsupportedStructure);
            Ok(None)
        }
    }

    fn start_row_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "c") {
            return self.start_cell(element);
        }
        if is_merge_markup(element) {
            return Err(invalid(
                "worksheet merge markup appears outside its schema context",
            ));
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
        if is_merge_markup(element) {
            return Err(invalid(
                "worksheet merge markup appears outside its schema context",
            ));
        }
        self.mark(NotEligibleReason::UnsupportedStructure);
        Ok(None)
    }

    fn start_merge_cells(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.validate_attributes(element, &["count"], false)?;
        self.merge_count = unqualified_attribute(element, "count")
            .map(|value| {
                value.parse::<u32>().map_err(|_source| {
                    invalid(format!("invalid worksheet merged-range count '{value}'"))
                })
            })
            .transpose()?;
        if self.not_eligible.is_some() {
            return Ok(None);
        }
        Ok(Some(Frame::MergeCells))
    }

    fn validate_merge_cells_placement(&self) -> Result<()> {
        if !self.seen_sheet_data {
            return Err(invalid("worksheet mergeCells appears before sheetData"));
        }
        if self.seen_merge_cells {
            return Err(invalid("worksheet has duplicate mergeCells elements"));
        }
        if self.merge_window_closed {
            return Err(invalid(
                "worksheet mergeCells appears after a schema successor",
            ));
        }
        Ok(())
    }

    fn start_merge_cells_child(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if is_spreadsheetml_element(element, "mergeCell") {
            return self.start_merge_cell(element);
        }
        if is_merge_markup(element) {
            return Err(invalid("worksheet mergeCells has an unmodeled child"));
        }
        self.mark(NotEligibleReason::MergeSemantics);
        Ok(None)
    }

    fn start_merge_cell(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.validate_attributes(element, &["ref"], false)?;
        if self.merges.len() >= MAX_MERGED_RANGES {
            self.merges.clear();
            self.mark(NotEligibleReason::MergeSemantics);
            return Ok(None);
        }
        let reference = unqualified_attribute(element, "ref")
            .ok_or_else(|| invalid("worksheet mergeCell is missing ref"))?;
        let range = Rect::from_a1(reference)
            .map_err(|error| invalid(format!("invalid merged range '{reference}': {error}")))?;
        if range.rows() == 1 && range.columns() == 1 {
            return Err(invalid(format!(
                "worksheet merged range '{reference}' contains only one cell"
            )));
        }
        if self.not_eligible.is_some() {
            return Ok(None);
        }
        self.merges
            .try_reserve(1)
            .map_err(|source| allocation("selected worksheet merged ranges", source))?;
        self.merges.push(range);
        Ok(Some(Frame::MergeCell))
    }

    fn start_merge_cell_child(&mut self, _element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        self.mark(NotEligibleReason::MergeSemantics);
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
            if is_merge_markup(element) {
                return Err(invalid(
                    "worksheet merge markup appears outside its schema context",
                ));
            }
            self.mark(NotEligibleReason::UnsupportedStructure);
        }
        Ok(None)
    }

    fn start_row(&mut self, element: &SemanticElement<'_>) -> Result<Option<Frame>> {
        if unqualified_attribute(element, "s").is_some()
            || unqualified_attribute(element, "customFormat").is_some()
        {
            self.mark(NotEligibleReason::Styles);
        }
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
        self.validate_attributes(element, &["r", "t", "s"], false)?;
        if let Some(style) = unqualified_attribute(element, "s") {
            let Some(style) = style.parse::<u32>().ok() else {
                self.mark(NotEligibleReason::Styles);
                return Ok(None);
            };
            if style > MAX_CELL_STYLE {
                self.mark(NotEligibleReason::Styles);
                return Ok(None);
            }
            self.dependencies.max_direct_style_index = Some(
                self.dependencies
                    .max_direct_style_index
                    .map_or(style, |current| current.max(style)),
            );
        }
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
            CellKind::Date | CellKind::Unknown => self.mark(NotEligibleReason::UnsupportedCellType),
            CellKind::Untyped
            | CellKind::Numeric
            | CellKind::Boolean
            | CellKind::Error
            | CellKind::String
            | CellKind::InlineString
            | CellKind::SharedString => {},
        }
        if self.not_eligible.is_some() {
            return Ok(None);
        }
        let address = Address::at(row - 1, column - 1)
            .map_err(|_error| invalid("worksheet cell address exceeds the grid"))?;
        let value = if kind.uses_numeric_value_scratch() {
            let mut value = std::mem::take(&mut self.numeric_value_scratch);
            value.clear();
            value
        } else {
            String::new()
        };
        self.pending_cell = Some(PendingCell::new(address, kind, value));
        Ok(Some(Frame::Cell))
    }

    fn start_formula(&mut self, element: &SemanticElement<'_>) -> Result<()> {
        self.validate_attributes(element, &["t", "ref", "si", "bx"], false)?;
        if matches!(
            self.pending_cell.as_ref().map(|cell| cell.kind),
            Some(CellKind::SharedString)
        ) {
            self.mark(NotEligibleReason::FormulaSemantics);
        }
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
            Frame::Dimension => {},
            Frame::SheetData => {},
            Frame::MergeCells => self.finish_merge_cells()?,
            Frame::MergeCell => {},
            Frame::Row => self.finish_row()?,
            Frame::Cell => self.finish_cell()?,
            Frame::Formula | Frame::Value | Frame::InlineText | Frame::Inline => {},
        }
        Ok(())
    }

    fn finish_empty(&mut self, frame: Frame) -> Result<()> {
        match frame {
            Frame::Worksheet => self.root_closed = true,
            Frame::Dimension => {},
            Frame::SheetData => {},
            Frame::MergeCells => self.finish_merge_cells()?,
            Frame::MergeCell => {},
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

    fn finish_merge_cells(&self) -> Result<()> {
        if self.merges.is_empty() {
            return Err(invalid(
                "worksheet mergeCells contains no mergeCell records",
            ));
        }
        if self
            .merge_count
            .is_some_and(|count| usize::try_from(count).ok() != Some(self.merges.len()))
        {
            return Err(invalid(format!(
                "worksheet merged-range count differs from {} records",
                self.merges.len()
            )));
        }
        Ok(())
    }

    fn finish_cell(&mut self) -> Result<()> {
        let mut cell = self
            .pending_cell
            .take()
            .ok_or_else(|| invalid("missing worksheet cell"))?;
        let address = cell.address;
        let effective_inline = cell.saw_inline || matches!(cell.kind, CellKind::InlineString);
        if effective_inline && cell.saw_value {
            return Err(invalid(
                "worksheet cell contains both inline text and a value",
            ));
        }
        if effective_inline && !matches!(cell.kind, CellKind::Untyped | CellKind::InlineString) {
            return Err(invalid("inline string has a non-inline cell type"));
        }
        let formula = if let Some(formula) = cell.formula.take() {
            if effective_inline {
                return Err(invalid("formula cell cannot contain an inline string"));
            }
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
        let shared_string_index =
            if formula.is_none() && matches!(cell.kind, CellKind::SharedString) && cell.saw_value {
                let Some(index) = cell.value.trim().parse::<u32>().ok() else {
                    self.mark(NotEligibleReason::SharedStrings);
                    return Ok(());
                };
                self.dependencies.max_shared_string_index = Some(
                    self.dependencies
                        .max_shared_string_index
                        .map_or(index, |current| current.max(index)),
                );
                Some(index)
            } else {
                None
            };
        if let Some(index) = shared_string_index {
            if self.requested.contains(address) {
                self.retain_selected(SelectedRecord {
                    address,
                    cell: None,
                    shared_string_index: Some(index),
                })?;
            }
            return Ok(());
        }
        let inline = if effective_inline {
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
        if self.requested.contains(address) {
            self.retain_selected(SelectedRecord {
                address,
                cell: Some(semantic),
                shared_string_index: None,
            })?;
        }
        self.recycle_numeric_value(cell);
        Ok(())
    }

    fn recycle_numeric_value(&mut self, cell: PendingCell) {
        if cell.kind.uses_numeric_value_scratch() && cell.value.capacity() <= MAX_CELL_CHARACTERS {
            self.numeric_value_scratch = cell.value;
        }
    }

    fn retain_selected(&mut self, record: SelectedRecord) -> Result<()> {
        self.selected_count = self
            .selected_count
            .checked_add(1)
            .ok_or_else(|| invalid("selected worksheet record count overflow"))?;
        self.selected
            .try_reserve(1)
            .map_err(|source| allocation("selected worksheet range records", source))?;
        self.selected.push(record);
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
            Frame::MergeCells | Frame::MergeCell => {
                if text.trim().is_empty() {
                    Ok(())
                } else {
                    self.mark(NotEligibleReason::MergeSemantics);
                    Ok(())
                }
            },
            Frame::Worksheet
            | Frame::Dimension
            | Frame::SheetData
            | Frame::Row
            | Frame::Cell
            | Frame::Inline => {
                if text.trim().is_empty() {
                    Ok(())
                } else {
                    self.mark(NotEligibleReason::UnsupportedStructure);
                    Ok(())
                }
            },
        }
    }

    fn general_reference(
        &mut self,
        reference: &litchi_ooxml_common::mce::SemanticGeneralRef<'_>,
    ) -> Result<()> {
        let name = reference.name.as_ref();
        let Some(token_bytes) = name.len().checked_add(2) else {
            self.mark(NotEligibleReason::GeneralReference);
            return Ok(());
        };
        if !name.is_ascii() || token_bytes > MAX_GENERAL_REFERENCE_TOKEN_BYTES {
            self.mark(NotEligibleReason::GeneralReference);
            return Ok(());
        }
        if matches!(self.stack.last(), Some(Frame::InlineText)) && name.starts_with(b"#") {
            self.mark(NotEligibleReason::GeneralReference);
            return Ok(());
        }
        let Ok(name) = std::str::from_utf8(name) else {
            self.mark(NotEligibleReason::GeneralReference);
            return Ok(());
        };
        let reference = BytesRef::new(name);
        let decoded = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
        if name.starts_with('#') && decoded.chars().any(|character| !is_xml_10_char(character)) {
            self.mark(NotEligibleReason::GeneralReference);
            return Ok(());
        }
        self.text(&decoded)
    }

    fn in_payload(&self) -> bool {
        matches!(
            self.stack.last(),
            Some(Frame::Formula | Frame::Value | Frame::InlineText)
        )
    }

    fn in_merge(&self) -> bool {
        matches!(
            self.stack.last(),
            Some(Frame::MergeCells | Frame::MergeCell)
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

fn is_xml_10_char(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
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

fn is_merge_markup(element: &SemanticElement<'_>) -> bool {
    is_spreadsheetml_element(element, "mergeCells")
        || is_spreadsheetml_element(element, "mergeCell")
}

fn matches_frame(frame: Frame, element: &litchi_ooxml_common::mce::SemanticEnd<'_>) -> bool {
    let local = element.expanded_name.local_name.as_str();
    let namespace = element.expanded_name.namespace.as_bytes();
    let expected = match frame {
        Frame::Worksheet => "worksheet",
        Frame::Dimension => "dimension",
        Frame::SheetData => "sheetData",
        Frame::MergeCells => "mergeCells",
        Frame::MergeCell => "mergeCell",
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
