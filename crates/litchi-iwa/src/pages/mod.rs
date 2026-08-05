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

pub use creation::PagesDocumentBuilder;
pub use document::PagesDocument;
pub use editor::{
    Kind, PagesAudioInfo, PagesAudioOptions, PagesBodyChartInfo, PagesBodyShapeInfo,
    PagesBodyShapeKind, PagesCellValue, PagesDrawableTextInfo, PagesEditor, PagesFootnote,
    PagesFootnoteId, PagesHeaderFooterInfo, PagesImageInfo, PagesImageOptions, PagesMovieInfo,
    PagesMovieOptions, PagesSectionInfo, PagesTable,
    PagesTableCellCheckboxFormat, PagesTableCellCurrencyFormat,
    PagesTableCellDataFormat, PagesTableCellDateTimeFormat, PagesTableCellDecimalPlaces,
    PagesTableCellDurationFormat, PagesTableCellDurationStyle, PagesTableCellDurationUnit,
    PagesTableCellDurationUnitRange, PagesTableCellDurationUnits, PagesTableCellFixedDecimalPlaces,
    PagesTableCellFractionFormat, PagesTableCellInset, PagesTableCellInsets, PagesTableCellLayout,
    PagesTableCellNegativeNumberStyle, PagesTableCellNumberFormat,
    PagesTableCellNumeralSystemFormat, PagesTableCellParagraphIndents,
    PagesTableCellParagraphLineSpacing, PagesTableCellParagraphList,
    PagesTableCellParagraphListBullet, PagesTableCellParagraphListBulletGeometry,
    PagesTableCellParagraphListIndentation, PagesTableCellParagraphListLabelColor,
    PagesTableCellParagraphListLevel, PagesTableCellParagraphListLevelPlacement,
    PagesTableCellParagraphListNumberFormat, PagesTableCellParagraphListNumberScale,
    PagesTableCellParagraphListNumberTiering, PagesTableCellParagraphListNumbering,
    PagesTableCellParagraphListPlacement, PagesTableCellParagraphSpacing,
    PagesTableCellParagraphTabStops, PagesTableCellPercentageFormat, PagesTableCellPopUpMenuFormat,
    PagesTableCellPopUpMenuInitialSelection, PagesTableCellPopUpMenuItem, PagesTableCellRegion,
    PagesTableCellScientificFormat, PagesTableCellSliderDisplayFormat, PagesTableCellSliderFormat,
    PagesTableCellSliderRange, PagesTableCellStarRatingFormat, PagesTableCellStepperDisplayFormat,
    PagesTableCellStepperFormat, PagesTableCellStepperRange, PagesTableCellTextAlignment,
    PagesTableCellTextBackground, PagesTableCellTextBaselineShift,
    PagesTableCellTextCapitalization, PagesTableCellTextCharacterSpacing, PagesTableCellTextColor,
    PagesTableCellTextDecorations, PagesTableCellTextFont, PagesTableCellTextFormat,
    PagesTableCellTextLigatures, PagesTableCellTextOutline, PagesTableCellTextScript,
    PagesTableCellTextShadow, PagesTableCellTextStyle, PagesTableCellTextWrap,
    PagesTableCellThousandsSeparator, PagesTableCellUpdate, PagesTableCellVerticalAlignment,
    PagesTableColumnDeletion, PagesTableColumnInsertion, PagesTableDimension,
    PagesTableDimensionSize, PagesTableFormulaAxisReference, PagesTableFormulaBinaryOperator,
    PagesTableFormulaCachedValue, PagesTableFormulaCellReference, PagesTableFormulaExpression,
    PagesTableHeaderCount, PagesTableHeaderSettings, PagesTableInfo, PagesTablePoints,
    PagesTableRowDeletion, PagesTableRowInsertion, PagesTableSortColumnIndex,
    PagesTableSortDirection, PagesTableSortOrder, PagesTableSortRowRange, PagesTableSortRule,
    PagesTableSortScope, PagesTableTitleSettings, RemovedPagesAudio, RemovedPagesBodyChart,
    RemovedPagesBodyShape, RemovedPagesImage, RemovedPagesMovie, RemovedPagesTextBox, Template,
};
pub use litchi_pages::{Section, SectionType};
