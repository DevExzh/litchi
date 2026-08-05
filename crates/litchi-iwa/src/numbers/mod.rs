//! Numbers Spreadsheet Support
//!
//! This module provides comprehensive support for parsing Apple Numbers spreadsheets,
//! including table extraction, cell data parsing, and formula support.
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
//! use litchi_iwa::numbers::NumbersDocument;
//!
//! let doc = NumbersDocument::open("spreadsheet.numbers")?;
//! let sheets = doc.sheets()?;
//!
//! for sheet in sheets {
//!     println!("Sheet: {}", sheet.name);
//!     for table in &sheet.tables {
//!         println!("  Table: {}", table.name());
//!         println!("{}", table.to_csv());
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub(crate) mod cell;
pub mod creation;
pub mod document;
pub mod editor;
pub(crate) mod formula;
pub mod sheet;
pub mod table;
pub mod table_extractor;

pub(crate) use litchi_numbers::cell::wire as bnc;
mod formula_owner;
mod function_map;
mod table_uid_map;

impl From<litchi_numbers::cell::wire::Error> for crate::Error {
    fn from(error: litchi_numbers::cell::wire::Error) -> Self {
        match error {
            litchi_numbers::cell::wire::Error::InvalidFormat(message) => {
                Self::InvalidFormat(message)
            },
            litchi_numbers::cell::wire::Error::ParseError(message) => Self::ParseError(message),
        }
    }
}

pub(crate) use cell::CellValue;
pub use creation::NumbersDocumentBuilder;
pub use document::NumbersDocument;
pub use editor::{
    Direction, IWorkTableCellRegion, NumbersCellCommentInfo, NumbersCellCommentReplyInfo,
    NumbersEditor, NumbersPivotCategoryInfo, NumbersSheetAudioInfo, NumbersSheetAudioOptions,
    NumbersSheetChartInfo, NumbersSheetImageInfo, NumbersSheetImageOptions, NumbersSheetInfo,
    NumbersSheetMovieInfo, NumbersSheetMovieOptions, NumbersSheetShapeInfo, NumbersSheetShapeKind,
    NumbersTableCellParagraphIndents, NumbersTableCellParagraphLineSpacing,
    NumbersTableCellParagraphList, NumbersTableCellParagraphListBullet,
    NumbersTableCellParagraphListBulletGeometry, NumbersTableCellParagraphListIndentation,
    NumbersTableCellParagraphListLabelColor, NumbersTableCellParagraphListLevel,
    NumbersTableCellParagraphListLevelPlacement, NumbersTableCellParagraphListNumberFormat,
    NumbersTableCellParagraphListNumberScale, NumbersTableCellParagraphListNumberTiering,
    NumbersTableCellParagraphListNumbering, NumbersTableCellParagraphListPlacement,
    NumbersTableCellParagraphSpacing, NumbersTableCellParagraphTabStops,
    NumbersTableCellTextAlignment, NumbersTableCellTextBackground,
    NumbersTableCellTextBaselineShift, NumbersTableCellTextCapitalization,
    NumbersTableCellTextCharacterSpacing, NumbersTableCellTextColor,
    NumbersTableCellTextDecorations, NumbersTableCellTextFont, NumbersTableCellTextLigatures,
    NumbersTableCellTextOutline, NumbersTableCellTextScript, NumbersTableCellTextShadow,
    NumbersTableCellTextStyle, NumbersTableDimension, NumbersTableDimensionSize,
    NumbersTableHeaderCount, NumbersTableHeaderSettings, NumbersTableInfo, NumbersTablePoints,
    NumbersTableSortColumnIndex, NumbersTableSortDirection, NumbersTableSortOrder,
    NumbersTableSortRowRange, NumbersTableSortRule, NumbersTableSortScope,
    NumbersTableTitleSettings, NumbersTextBoxInfo, RemovedNumbersSheetAudio,
    RemovedNumbersSheetChart, RemovedNumbersSheetImage, RemovedNumbersSheetMovie,
    RemovedNumbersSheetShape, RemovedNumbersTextBox, TableCellConditionalHighlightInfo,
    TableColumnDeletion, TableColumnInsertion, TableRowDeletion, TableRowInsertion,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use litchi_numbers::cell::{APPLE_EPOCH_UNIX_OFFSET_SECONDS, Type, Update, Value};
pub use sheet::NumbersSheet;
pub use table::{NumbersCellComment, NumbersCommentUuid, NumbersTable};
pub use table_extractor::TableDataExtractor;
