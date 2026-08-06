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
pub mod document;
pub mod editor;

pub use creation::PagesDocumentBuilder;
pub use document::PagesDocument;
pub use editor::{
    PagesAudioInfo, PagesBodyChartInfo, PagesBodyShapeInfo, PagesCellValue, PagesDrawableTextInfo,
    PagesEditor, PagesHeaderFooterInfo, PagesImageInfo, PagesMovieInfo, PagesSectionInfo,
    PagesTable, PagesTableCellParagraphIndents, PagesTableCellParagraphLineSpacing,
    PagesTableCellParagraphList, PagesTableCellParagraphListBullet,
    PagesTableCellParagraphListBulletGeometry, PagesTableCellParagraphListIndentation,
    PagesTableCellParagraphListLabelColor, PagesTableCellParagraphListLevel,
    PagesTableCellParagraphListLevelPlacement, PagesTableCellParagraphListNumberFormat,
    PagesTableCellParagraphListNumberScale, PagesTableCellParagraphListNumberTiering,
    PagesTableCellParagraphListNumbering, PagesTableCellParagraphListPlacement,
    PagesTableCellParagraphSpacing, PagesTableCellParagraphTabStops, PagesTableCellTextAlignment,
    PagesTableCellTextBackground, PagesTableCellTextBaselineShift,
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
