//! Numbers Spreadsheet Support
//!
//! This module provides Numbers creation and editing support. Use
//! [`litchi_numbers::Document`] for archive-free spreadsheet reading.
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
//! use litchi_numbers::Document;
//!
//! let doc = Document::open("spreadsheet.numbers")?;
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
pub mod editor;
pub(crate) mod formula;
pub mod table;
pub mod table_extractor;

pub(crate) use litchi_numbers::table::dimension::{
    Dimension as NumbersTableDimension, Points as NumbersTablePoints,
    Size as NumbersTableDimensionSize,
};
pub(crate) use litchi_numbers_wire as bnc;
mod formula_owner;
mod function_map;
mod table_uid_map;

impl From<bnc::Error> for crate::Error {
    fn from(error: bnc::Error) -> Self {
        match error {
            bnc::Error::InvalidFormat(message) => Self::InvalidFormat(message),
            bnc::Error::ParseError(message) => Self::ParseError(message),
            bnc::Error::OutputLimitExceeded { .. } | bnc::Error::Allocation { .. } => {
                Self::ParseError(error.to_string())
            },
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

pub use creation::NumbersDocumentBuilder;
pub use editor::{
    Direction, NumbersEditor, NumbersPivotCategoryInfo, NumbersSheetAudioInfo,
    NumbersSheetAudioOptions, NumbersSheetChartInfo, NumbersSheetImageInfo,
    NumbersSheetImageOptions, NumbersSheetInfo, NumbersSheetMovieInfo, NumbersSheetMovieOptions,
    NumbersSheetShapeInfo, NumbersTableCellParagraphList, NumbersTableCellParagraphListBullet,
    NumbersTableCellParagraphListBulletGeometry, NumbersTableCellParagraphListIndentation,
    NumbersTableCellParagraphListLabelColor, NumbersTableCellParagraphListLevel,
    NumbersTableCellParagraphListLevelPlacement, NumbersTableCellParagraphListNumberFormat,
    NumbersTableCellParagraphListNumberScale, NumbersTableCellParagraphListNumberTiering,
    NumbersTableCellParagraphListNumbering, NumbersTableCellParagraphListPlacement,
    NumbersTableCellParagraphTabStops, NumbersTableCellTextBackground,
    NumbersTableCellTextBaselineShift, NumbersTableCellTextCapitalization,
    NumbersTableCellTextCharacterSpacing, NumbersTableCellTextColor,
    NumbersTableCellTextDecorations, NumbersTableCellTextFont, NumbersTableCellTextLigatures,
    NumbersTableCellTextOutline, NumbersTableCellTextScript, NumbersTableCellTextShadow,
    NumbersTableCellTextStyle, NumbersTableInfo, NumbersTableSortColumnIndex,
    NumbersTableSortDirection, NumbersTableSortOrder, NumbersTableSortRowRange,
    NumbersTableSortRule, NumbersTableSortScope, NumbersTextBoxInfo, RemovedNumbersSheetAudio,
    RemovedNumbersSheetChart, RemovedNumbersSheetImage, RemovedNumbersSheetMovie,
    RemovedNumbersSheetShape, RemovedNumbersTextBox, Settings, TableCellConditionalHighlightInfo,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use litchi_numbers::cell::{APPLE_EPOCH_UNIX_OFFSET_SECONDS, Type, Update, Value};
pub use table::NumbersTable;
pub use table_extractor::TableDataExtractor;

#[cfg(test)]
pub(crate) trait SemanticTableCellAssertions {
    fn get_cell(&self, row: usize, column: usize) -> Option<&litchi_numbers::cell::Value>;
}

#[cfg(test)]
impl SemanticTableCellAssertions for litchi_numbers::Table {
    fn get_cell(&self, row: usize, column: usize) -> Option<&litchi_numbers::cell::Value> {
        litchi_numbers::Position::try_from_usize(row, column)
            .ok()
            .and_then(|position| self.get(position))
    }
}
