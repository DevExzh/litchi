//! Pages Document Support
//!
//! This module provides comprehensive support for creating and editing Apple
//! Pages documents. Read-only semantic parsing lives in [`litchi_pages`].
//!
//! ## Features
//!
//! - Floating drawables (images, shapes)
//! - Header/footer number-attachment compatibility operations
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi_pages::Package;
//!
//! let doc = Package::open("document.pages")?.document();
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
pub(crate) mod editor;

pub use creation::Builder;
pub use editor::{
    AudioInfo, BodyChartInfo, BodyShapeInfo, CellValue, Draft, DrawableTextInfo, ImageInfo,
    MovieInfo, RemovedAudio, RemovedBodyChart, RemovedBodyShape, RemovedImage, RemovedMovie,
    RemovedTextBox, SectionInfo, Table, TableCellConditionalHighlightInfo, TableCellInset,
    TableCellInsets, TableCellLayout, TableCellParagraphList, TableCellParagraphListBullet,
    TableCellParagraphListBulletGeometry, TableCellParagraphListIndentation,
    TableCellParagraphListLabelColor, TableCellParagraphListLevel,
    TableCellParagraphListLevelPlacement, TableCellParagraphListNumberFormat,
    TableCellParagraphListNumberScale, TableCellParagraphListNumberTiering,
    TableCellParagraphListNumbering, TableCellParagraphListPlacement, TableCellParagraphTabStops,
    TableCellTextBackground, TableCellTextBaselineShift, TableCellTextCapitalization,
    TableCellTextCharacterSpacing, TableCellTextColor, TableCellTextDecorations, TableCellTextFont,
    TableCellTextLigatures, TableCellTextOutline, TableCellTextScript, TableCellTextShadow,
    TableCellTextStyle, TableCellTextWrap, TableCellUpdate, TableCellVerticalAlignment,
    TableFormulaAxisReference, TableFormulaBinaryOperator, TableFormulaCachedValue,
    TableFormulaCellReference, TableFormulaExpression, TableInfo, TableSortRowRange,
};
pub use litchi_pages::footnote::body::{Footnote, Position, Selector};
pub use litchi_pages::{Section, SectionType};
