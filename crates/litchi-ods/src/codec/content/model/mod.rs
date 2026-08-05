//! Semantic assembly builders for ODS `content.xml`.
//!
//! The builders own only the in-memory spreadsheet state assembled by the
//! streaming codec. XML traversal stays in the sibling `codec` owner, while
//! the narrow `pub(crate)` methods below are the boundary between the two
//! layers.

use super::super::{
    Annotation, Cell, CellDetective, CellMatrixSpan, CellMerge, CellRangeSource, CellTextContent,
    CellValue, Column, ConditionalFormat, Row, Sheet, SheetPrintSettings, SheetProtection,
    SheetScenario, SheetStyle, SheetTableSource, SparklineGroup, TableGroup, TableRange,
    TableStructure, TableVisibility,
    conditional_format::MAX_CONDITIONAL_FORMATS_PER_SHEET,
    dde::DdeSource,
    sparkline::MAX_SPARKLINE_GROUPS_PER_SHEET,
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, MAX_TABLE_STRUCTURE_DEPTH,
    },
};
use crate::model::hyperlink::Link;
use litchi_core::{Error, Result};

const MAX_EXPANDED_CELLS_PER_ROW: usize = 1_048_576;
const MAX_EXPANDED_CELLS_PER_SHEET: usize = 4_194_304;
/// Interleaved runs of empty rows kept unmaterialised before the parser gives
/// up and expands them, so deferral cannot grow without bound.
const MAX_DEFERRED_BLANK_ROW_RUNS: usize = 4_096;
/// Longest run of cell-less rows still kept at the end of a table. Anything
/// longer is the full-height grid padding every ODF producer writes.
const MAX_TRAILING_EMPTY_ROWS: usize = 4_096;

mod cell;
mod row;
mod sheet;
mod structure;

use structure::StructureStack;

pub(crate) use cell::CellBuilder;
pub(crate) use row::RowBuilder;
pub(crate) use sheet::SheetBuilder;
