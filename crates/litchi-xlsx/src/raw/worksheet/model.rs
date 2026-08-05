//! Lossless parser state and raw worksheet records.

use std::collections::HashSet;

use litchi_sheet::{Cell as Address, Rect};

use super::super::formula::Range as FormulaRange;
use super::x14ac;
use crate::column::{self, Assignments};
use crate::layout::Defaults;
use crate::row;

pub(super) const MAX_CELL_CHARACTERS: usize = 32_767;
pub(super) const MAX_FORMULA_CHARACTERS: usize = 8_192;
// A supplementary Unicode scalar can occupy two seven-byte `_xHHHH_`
// SpreadsheetML escapes before decoding.
pub(super) const MAX_ENCODED_CELL_BYTES: usize = MAX_CELL_CHARACTERS * 14;
pub(super) const MAX_CELL_STYLE: u32 = 65_490;
pub(super) const MAX_COLUMN_STYLE: u32 = 65_429;
pub(super) const MAX_METADATA_INDEX: u32 = 2_147_483_647;
pub(super) const MAX_XML_DEPTH: usize = 256;

pub(crate) fn merge_successor(local: &[u8]) -> bool {
    matches!(
        local,
        b"phoneticPr"
            | b"conditionalFormatting"
            | b"dataValidations"
            | b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Context {
    Worksheet,
    SheetFormat,
    Columns,
    SheetData,
    MergeCells,
    Merge,
    Row,
    Cell,
    Formula,
    Value,
    Inline,
    Run,
    Text(TextTarget),
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum TextTarget {
    Formula,
    Value,
    Inline,
}

#[derive(Debug)]
pub(super) struct PendingRow {
    pub(super) number: u32,
    pub(super) last_column: u32,
    pub(super) properties: row::Properties,
}

#[derive(Debug)]
pub(super) struct PendingCell {
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) style: Option<u32>,
    pub(super) cell_metadata: Option<u32>,
    pub(super) value_metadata: Option<u32>,
    pub(super) cell_type: Option<String>,
    pub(super) value: String,
    pub(super) value_bytes: usize,
    pub(super) saw_value: bool,
    pub(super) formula: String,
    pub(super) formula_characters: usize,
    pub(super) formula_kind: Option<RawFormulaKind>,
    pub(super) inline: String,
    pub(super) inline_bytes: usize,
    pub(super) saw_inline: bool,
    pub(super) saw_inline_simple: bool,
    pub(super) saw_inline_run: bool,
    pub(super) run_has_text: bool,
}

#[derive(Debug)]
pub(super) enum RawFormulaKind {
    Scalar,
    Array(Option<String>),
    DataTable(Option<String>),
    Shared { index: u32, range: Option<String> },
    Unknown(String),
}

#[derive(Debug)]
pub(super) struct RawCell {
    pub(super) address: Address,
    pub(super) style: Option<u32>,
    pub(super) cell_metadata: Option<u32>,
    pub(super) value_metadata: Option<u32>,
    pub(super) cell_type: Option<String>,
    pub(super) value: Option<String>,
    pub(super) inline: Option<String>,
    pub(super) formula: Option<RawFormula>,
}

#[derive(Debug)]
pub(super) struct RawFormula {
    pub(super) text: String,
    pub(super) kind: RawFormulaKind,
}

#[derive(Debug)]
pub(super) struct SharedMember {
    pub(super) cell_index: usize,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) index: u32,
    pub(super) range: Option<String>,
    pub(super) text: String,
}

#[derive(Debug)]
pub(super) struct SharedMaster {
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) range: FormulaRange,
    pub(super) text: String,
}

#[derive(Debug)]
pub(super) struct Parser {
    pub(super) cells: Vec<RawCell>,
    pub(super) rows: Vec<row::Stored>,
    pub(super) columns: Option<Assignments<column::Properties>>,
    pub(super) defaults: Option<Defaults>,
    pub(super) extensions: x14ac::Values,
    pub(super) declared_extent: Option<Rect>,
    pub(super) row: Option<PendingRow>,
    pub(super) cell: Option<PendingCell>,
    pub(super) seen_rows: HashSet<u32>,
    pub(super) previous_row: u32,
    pub(super) seen_dimension: bool,
    pub(super) seen_defaults: bool,
    pub(super) seen_columns: bool,
    pub(super) column_records: usize,
    pub(super) seen_sheet_data: bool,
    pub(super) merges: Vec<Rect>,
    pub(super) merge_count: Option<usize>,
    pub(super) seen_merges: bool,
    pub(super) merge_window_closed: bool,
}
