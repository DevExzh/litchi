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
//!         println!("  Table: {}", table.name);
//!         println!("{}", table.to_csv());
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod cell;
pub mod creation;
pub mod document;
pub mod editor;
pub mod formula;
pub mod sheet;
pub mod table;
pub mod table_extractor;

mod bnc;
mod formula_owner;
mod function_map;
mod table_uid_map;

pub use cell::{APPLE_EPOCH_UNIX_OFFSET_SECONDS, CellType, CellValue, TableCellUpdate};
pub use creation::NumbersDocumentBuilder;
pub use document::NumbersDocument;
pub use editor::{
    ChartSeriesDirection, IWorkTableCellRegion, NumbersCellCommentInfo,
    NumbersCellCommentReplyInfo, NumbersEditor, NumbersPivotCategoryInfo, NumbersSheetAudioInfo,
    NumbersSheetAudioOptions, NumbersSheetChartInfo, NumbersSheetImageInfo,
    NumbersSheetImageOptions, NumbersSheetInfo, NumbersSheetMovieInfo, NumbersSheetMovieOptions,
    NumbersSheetShapeInfo, NumbersSheetShapeKind, NumbersTableCellParagraphIndents,
    NumbersTableCellParagraphLineSpacing, NumbersTableCellParagraphList,
    NumbersTableCellParagraphListPlacement, NumbersTableCellParagraphSpacing,
    NumbersTableCellParagraphTabStops, NumbersTableCellTextAlignment,
    NumbersTableCellTextBackground, NumbersTableCellTextBaselineShift,
    NumbersTableCellTextCapitalization, NumbersTableCellTextCharacterSpacing,
    NumbersTableCellTextColor, NumbersTableCellTextDecorations, NumbersTableCellTextFont,
    NumbersTableCellTextLigatures, NumbersTableCellTextOutline, NumbersTableCellTextScript,
    NumbersTableCellTextShadow, NumbersTableCellTextStyle, NumbersTableDimension,
    NumbersTableDimensionSize, NumbersTableHeaderCount, NumbersTableHeaderSettings,
    NumbersTableInfo, NumbersTablePoints, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRowRange, NumbersTableSortRule, NumbersTableSortScope,
    NumbersTableTitleSettings, NumbersTextBoxInfo, RemovedNumbersSheetAudio,
    RemovedNumbersSheetChart, RemovedNumbersSheetImage, RemovedNumbersSheetMovie,
    RemovedNumbersSheetShape, RemovedNumbersTextBox, TableCellConditionalHighlightInfo,
    TableColumnDeletion, TableColumnInsertion, TableRowDeletion, TableRowInsertion,
};
pub use formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
pub use sheet::NumbersSheet;
pub use table::{NumbersCellComment, NumbersCommentUuid, NumbersTable};
pub use table_extractor::TableDataExtractor;
