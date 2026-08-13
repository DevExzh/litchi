//! Pages Document Support
//!
//! This module provides comprehensive support for creating and editing Apple
//! Pages documents. Read-only semantic parsing lives in [`litchi_pages`].
//!
//! ## Features
//!
//! - Floating drawables (images, shapes)
//! - Header and footer extraction
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi_pages::Document;
//!
//! let doc = Document::open("document.pages")?;
//! let text = doc.plain_text();
//! let sections = doc.sections();
//!
//! for section in sections {
//!     println!("Section: {:?}", section.heading());
//!     for para in section.paragraphs() {
//!         println!("  {}", para);
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod creation;
pub mod editor;

pub use creation::PagesDocumentBuilder;
pub use editor::{
    PagesAudioInfo, PagesBodyChartInfo, PagesBodyShapeInfo, PagesCellValue, PagesDrawableTextInfo,
    PagesEditor, PagesHeaderFooterInfo, PagesImageInfo, PagesMovieInfo, PagesSectionInfo,
    PagesTable, PagesTableCellParagraphList, PagesTableCellParagraphListBullet,
    PagesTableCellParagraphListBulletGeometry, PagesTableCellParagraphListIndentation,
    PagesTableCellParagraphListLabelColor, PagesTableCellParagraphListLevel,
    PagesTableCellParagraphListLevelPlacement, PagesTableCellParagraphListNumberFormat,
    PagesTableCellParagraphListNumberScale, PagesTableCellParagraphListNumberTiering,
    PagesTableCellParagraphListNumbering, PagesTableCellParagraphListPlacement,
    PagesTableCellParagraphTabStops, PagesTableCellTextBackground, PagesTableCellTextBaselineShift,
    PagesTableCellTextCapitalization, PagesTableCellTextCharacterSpacing, PagesTableCellTextColor,
    PagesTableCellTextDecorations, PagesTableCellTextFont, PagesTableCellTextLigatures,
    PagesTableCellTextOutline, PagesTableCellTextScript, PagesTableCellTextShadow,
    PagesTableCellTextStyle, PagesTableCellUpdate, PagesTableDimension, PagesTableDimensionSize,
    PagesTableFormulaAxisReference, PagesTableFormulaBinaryOperator, PagesTableFormulaCachedValue,
    PagesTableFormulaCellReference, PagesTableFormulaExpression, PagesTableInfo, PagesTablePoints,
    PagesTableSortColumnIndex, PagesTableSortDirection, PagesTableSortOrder,
    PagesTableSortRowRange, PagesTableSortRule, PagesTableSortScope, PagesTableTitleSettings,
    RemovedPagesAudio, RemovedPagesBodyChart, RemovedPagesBodyShape, RemovedPagesImage,
    RemovedPagesMovie, RemovedPagesTextBox,
};
pub use litchi_pages::footnote::body::{Footnote, Position, Selector};
pub use litchi_pages::{Section, SectionType};
