//! Pages Document Support
//!
//! This module provides comprehensive support for parsing Apple Pages documents,
//! including text extraction, section management, and document structure analysis.
//!
//! ## Features
//!
//! - Document metadata extraction
//! - Section and paragraph parsing
//! - Text style information
//! - Floating drawables (images, shapes)
//! - Header and footer extraction
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi_iwa::pages::PagesDocument;
//!
//! let doc = PagesDocument::open("document.pages")?;
//! let text = doc.text()?;
//! let sections = doc.sections()?;
//!
//! for section in sections {
//!     println!("Section: {:?}", section.heading);
//!     for para in &section.paragraphs {
//!         println!("  {}", para);
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod creation;
pub mod document;
pub mod editor;
pub mod section;

pub use creation::PagesDocumentBuilder;
pub use document::PagesDocument;
pub use editor::{
    PagesAudioInfo, PagesAudioOptions, PagesBodyChartInfo, PagesBodyShapeInfo, PagesBodyShapeKind,
    PagesCellValue, PagesDocumentOptions, PagesDrawableTextInfo, PagesEditor, PagesFootnote,
    PagesFootnoteFormat, PagesFootnoteGap, PagesFootnoteId, PagesFootnoteKind,
    PagesFootnoteNumbering, PagesFootnoteSettings, PagesHeaderFooterInfo, PagesHeaderFooterKind,
    PagesImageInfo, PagesImageOptions, PagesMovieInfo, PagesMovieOptions, PagesPageLayout,
    PagesPageNumber, PagesPageOrientation, PagesRgbColorSpace, PagesRgbaColor,
    PagesSectionBackground, PagesSectionInfo, PagesSectionPageNumbering, PagesSectionSettings,
    PagesSectionStart, PagesTable, PagesTableAxisIndex, PagesTableCellCheckboxFormat,
    PagesTableCellCurrencyFormat, PagesTableCellDataFormat, PagesTableCellDateTimeFormat,
    PagesTableCellDecimalPlaces, PagesTableCellDurationFormat, PagesTableCellDurationStyle,
    PagesTableCellDurationUnit, PagesTableCellDurationUnitRange, PagesTableCellDurationUnits,
    PagesTableCellFixedDecimalPlaces, PagesTableCellFractionFormat, PagesTableCellInset,
    PagesTableCellInsets, PagesTableCellLayout, PagesTableCellNegativeNumberStyle,
    PagesTableCellNumberFormat, PagesTableCellNumeralSystemFormat, PagesTableCellPercentageFormat,
    PagesTableCellRegion, PagesTableCellScientificFormat, PagesTableCellStarRatingFormat,
    PagesTableCellTextWrap, PagesTableCellThousandsSeparator, PagesTableCellUpdate,
    PagesTableCellVerticalAlignment, PagesTableColumnDeletion, PagesTableColumnInsertion,
    PagesTableDimension, PagesTableDimensionSize, PagesTableFormulaAxisReference,
    PagesTableFormulaBinaryOperator, PagesTableFormulaCachedValue, PagesTableFormulaCellReference,
    PagesTableFormulaExpression, PagesTableHeaderCount, PagesTableHeaderSettings,
    PagesTableHiddenAxes, PagesTableInfo, PagesTablePoints, PagesTableRowDeletion,
    PagesTableRowInsertion, PagesTableSortColumnIndex, PagesTableSortDirection,
    PagesTableSortOrder, PagesTableSortRowRange, PagesTableSortRule, PagesTableSortScope,
    PagesTableTitleSettings, PagesTemplateKind, RemovedPagesAudio, RemovedPagesBodyChart,
    RemovedPagesBodyShape, RemovedPagesImage, RemovedPagesMovie, RemovedPagesTextBox,
};
pub use section::{PagesSection, PagesSectionType};
