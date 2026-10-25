//! Numbers Spreadsheet Support
//!
//! This module provides Numbers creation and editing support. Use
//! [`litchi_numbers::Workbook`] for archive-free spreadsheet reading.
//!
//! ## Features
//!
//! - Sheet extraction
//! - Table parsing with cell data
//! - Formula extraction
//! - CSV export
//! - Cell formatting information
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi_numbers::Workbook;
//!
//! let doc = Workbook::open("spreadsheet.numbers")?;
//! let sheets = doc.sheets();
//!
//! for sheet in sheets.iter() {
//!     println!("Sheet: {}", sheet.name());
//!     for table in sheet.tables() {
//!         println!("  Table: {}", table.name());
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub(crate) mod cell;
pub mod creation;
pub(crate) mod editor;
pub(crate) mod formula;

pub(crate) use litchi_numbers::table::dimension::{
    Dimension as NumbersTableDimension, Points as NumbersTablePoints,
    Size as NumbersTableDimensionSize,
};
pub(crate) use litchi_numbers_wire as bnc;

/// Select the bounded raw-cell profile for one native source payload.
pub(crate) fn bnc_limits(source: &[u8]) -> bnc::Limits {
    if source.is_empty() {
        return bnc::Limits::DEFAULT;
    }
    bnc::Limits::DEFAULT
        .with_input_bytes(source.len())
        .and_then(|limits| limits.with_tail_bytes(source.len()))
        .expect("source-derived Numbers BNC limits stay within hard ceilings")
}

/// Encode an owned native cell under the Numbers editor's bounded profile.
pub(crate) fn bnc_encode(
    cell: &bnc::BncCell,
    limits: &bnc::Limits,
) -> std::result::Result<Vec<u8>, bnc::Error> {
    cell.encode(limits)
}
mod formula_owner;
mod function_map;
mod table_uid_map;

impl From<bnc::Error> for crate::Error {
    fn from(error: bnc::Error) -> Self {
        match error {
            bnc::Error::InvalidFormat(message) => Self::InvalidFormat(message),
            bnc::Error::ParseError(message) => Self::ParseError(message),
            bnc::Error::LimitExceeded { .. } | bnc::Error::Allocation { .. } => {
                Self::ParseError(error.to_string())
            },
            _ => Self::ParseError("Numbers cell wire error".to_string()),
        }
    }
}

impl From<litchi_numbers::table::headers::Error> for crate::Error {
    fn from(error: litchi_numbers::table::headers::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

impl From<litchi_numbers::table::merge::Error> for crate::Error {
    fn from(error: litchi_numbers::table::merge::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

pub use creation::Builder;
pub use editor::{
    Direction, Draft, PivotCategoryInfo, RemovedSheetAudio, RemovedSheetChart, RemovedSheetImage,
    RemovedSheetMovie, RemovedSheetShape, RemovedTextBox, SheetAudioInfo, SheetAudioOptions,
    SheetChartInfo, SheetImageInfo, SheetImageOptions, SheetInfo, SheetMovieInfo,
    SheetMovieOptions, SheetShapeInfo, TableCellConditionalHighlightInfo,
    TableCellParagraphIndents, TableCellParagraphLineSpacing, TableCellParagraphList,
    TableCellParagraphListBullet, TableCellParagraphListBulletGeometry,
    TableCellParagraphListIndentation, TableCellParagraphListLabelColor,
    TableCellParagraphListLevel, TableCellParagraphListLevelPlacement,
    TableCellParagraphListNumberFormat, TableCellParagraphListNumberScale,
    TableCellParagraphListNumberTiering, TableCellParagraphListNumbering,
    TableCellParagraphListPlacement, TableCellParagraphSpacing, TableCellParagraphTabStops,
    TableCellTextAlignment, TableCellTextBackground, TableCellTextBaselineShift,
    TableCellTextCapitalization, TableCellTextCharacterSpacing, TableCellTextColor,
    TableCellTextDecorations, TableCellTextFont, TableCellTextLigatures, TableCellTextOutline,
    TableCellTextScript, TableCellTextShadow, TableCellTextStyle, TableInfo, TableSortColumnIndex,
    TableSortDirection, TableSortOrder, TableSortRowRange, TableSortRule, TableSortScope,
    TextBoxInfo, UnverifiedAxisChange,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use litchi_numbers::cell::{
    APPLE_EPOCH_UNIX_OFFSET_SECONDS, FiniteF64, FiniteF64Error, Type, Update, Value,
};
#[cfg(test)]
pub(crate) trait SemanticTableCellAssertions {
    fn get_cell(&self, row: usize, column: usize) -> Option<&litchi_numbers::cell::Value>;
}

#[cfg(test)]
impl SemanticTableCellAssertions for litchi_numbers::Table {
    fn get_cell(&self, row: usize, column: usize) -> Option<&litchi_numbers::cell::Value> {
        litchi_numbers::At::try_from_usize(row, column)
            .ok()
            .and_then(|position| self.get(position))
    }
}
