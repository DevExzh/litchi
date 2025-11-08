//! OpenDocument Format (ODF) support.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating OpenDocument
//! files conforming to ISO/IEC 26300 (ODF 1.2), including text documents (.odt), spreadsheets (.ods),
//! and presentations (.odp).
//!
//! # Implementation Progress
//!
//! This implementation is inspired by ODF Toolkit (Java) and odfpy (Python), aiming for a complete,
//! production-ready ODF reader/writer in Rust with high performance and memory efficiency.
//!
//! ## ✅ Core Infrastructure (COMPLETE)
//!
//! - **Package System** (`core/package.rs`)
//!   - ✅ ZIP archive reading with `Package<R>`
//!   - ✅ Manifest parsing and MIME type detection
//!   - ✅ File extraction and existence checking
//!   - ✅ Optimized zero-copy package with buffer pooling
//!   - ✅ PackageWriter for creating ODF files
//!
//! - **XML Processing** (`core/xml.rs`)
//!   - ✅ Content.xml parsing (main document content)
//!   - ✅ Styles.xml parsing (document styles)
//!   - ✅ Meta.xml parsing (document metadata)
//!   - ✅ Settings.xml support (document settings)
//!   - ✅ High-performance quick-xml based parsing
//!
//! - **Element Model** (`elements/`)
//!   - ✅ Text elements (paragraphs, spans, lists)
//!   - ✅ Table elements (tables, rows, cells)
//!   - ✅ Style elements (paragraph, text, table styles)
//!   - ✅ Draw elements (shapes, frames, images)
//!   - ✅ Field elements (date, time, page number)
//!   - ✅ Bookmark and reference support
//!   - ✅ Namespace handling for ODF XML
//!
//! - **Constants & Utilities** (`constants.rs`, `datatype.rs`)
//!   - ✅ MIME type constants and mappings
//!   - ✅ File extension detection
//!   - ✅ Standard ODF part paths
//!   - ✅ Data type conversions (Boolean, Date, DateTime, Duration)
//!   - ✅ A1 notation coordinate conversion
//!
//! ## ✅ ODT - Text Documents (COMPLETE for reading/writing)
//!
//! ### Reading (`odt/document.rs`, `odt/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Full text extraction
//! - ✅ Paragraph and span parsing
//! - ✅ Table parsing (nested tables supported)
//! - ✅ List parsing (ordered and unordered)
//! - ✅ Heading hierarchy extraction
//! - ✅ Style registry and style resolution
//! - ✅ Metadata extraction
//! - ✅ Hyperlink extraction
//! - ✅ Footnote and endnote support
//! - ✅ Bookmark and reference tracking
//! - ✅ Comment and change tracking parsing
//! - ✅ Section parsing
//!
//! ### Writing (`odt/builder.rs`, `odt/mutable.rs`)
//! - ✅ DocumentBuilder for creating new ODT files
//! - ✅ Add paragraphs with text and styling
//! - ✅ Add tables with rows and cells
//! - ✅ Add lists (ordered/unordered)
//! - ✅ Add headings with levels
//! - ✅ MutableDocument for modifying existing documents
//! - ✅ Set metadata (title, author, description, etc.)
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Table of contents generation
//! - ⚠️ Index generation
//! - ⚠️ Mail merge and field insertion
//! - ⚠️ Drawing object support (beyond basic shapes)
//! - ⚠️ Form controls
//! - ⚠️ Master page manipulation
//!
//! ## ✅ ODS - Spreadsheets (COMPLETE for reading/writing)
//!
//! ### Reading (`ods/spreadsheet.rs`, `ods/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Sheet parsing (multiple sheets)
//! - ✅ Cell value extraction (String, Number, Boolean, Date, DateTime, Duration, Percentage, Currency)
//! - ✅ Formula representation
//! - ✅ Row and column operations
//! - ✅ Cell coordinate conversion (A1 notation)
//! - ✅ CSV export
//! - ✅ Metadata extraction
//! - ✅ Style parsing
//! - ✅ Repeated cells/rows expansion
//! - ✅ Merged cell handling
//!
//! ### Writing (`ods/builder.rs`, `ods/mutable.rs`)
//! - ✅ SpreadsheetBuilder for creating new ODS files
//! - ✅ Add sheets with names
//! - ✅ Set cell values (all types)
//! - ✅ Set cell formulas
//! - ✅ Set cell styles
//! - ✅ MutableSpreadsheet for modifying existing spreadsheets
//! - ✅ Insert/delete rows and columns
//! - ✅ Set metadata
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Chart creation and parsing
//! - ⚠️ Data validation rules
//! - ⚠️ Conditional formatting
//! - ⚠️ Pivot tables
//! - ⚠️ Named ranges
//! - ⚠️ Cell comments (notes)
//! - ⚠️ Sheet protection
//! - ⚠️ Filter and sort criteria
//!
//! ## ✅ ODP - Presentations (COMPLETE for reading/writing)
//!
//! ### Reading (`odp/presentation.rs`, `odp/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Slide parsing
//! - ✅ Shape extraction (text boxes, images, etc.)
//! - ✅ Slide layouts
//! - ✅ Master page parsing
//! - ✅ Text extraction from slides
//! - ✅ Metadata extraction
//!
//! ### Writing (`odp/builder.rs`, `odp/mutable.rs`)
//! - ✅ PresentationBuilder for creating new ODP files
//! - ✅ Add slides
//! - ✅ Add shapes (text boxes, rectangles, etc.)
//! - ✅ Set slide layouts
//! - ✅ MutablePresentation for modifying existing presentations
//! - ✅ Set metadata
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Slide transitions
//! - ⚠️ Animations
//! - ⚠️ Speaker notes
//! - ⚠️ Multimedia embedding (audio, video)
//! - ⚠️ Custom slide layouts
//! - ⚠️ Advanced shape properties
//! - ⚠️ Connector lines
//! - ⚠️ Slide master manipulation
//!
//! ## 🚧 Additional ODF Formats (NOT IMPLEMENTED)
//!
//! These formats are recognized but not fully supported yet:
//! - 🔲 ODG - OpenDocument Drawing (.odg)
//! - 🔲 ODC - OpenDocument Chart (.odc) - standalone charts
//! - 🔲 ODF - OpenDocument Formula (.odf) - mathematical formulas
//! - 🔲 ODI - OpenDocument Image (.odi)
//! - 🔲 ODM - OpenDocument Master (.odm) - master documents
//! - 🔲 Template variants (.ott, .ots, .otp, .otg, etc.)
//!
//! ## 🚧 Advanced Features (PLANNED)
//!
//! ### Embedded Objects
//! - 🔲 Embedded images (basic support exists, advanced features needed)
//! - 🔲 Embedded charts in documents/spreadsheets
//! - 🔲 Embedded objects (OLE)
//! - 🔲 Embedded videos and audio
//!
//! ### Collaboration Features
//! - 🔲 Change tracking (basic parsing exists, manipulation needed)
//! - 🔲 Comments and annotations (basic parsing exists)
//! - 🔲 Version control
//! - 🔲 Document comparison
//!
//! ### Performance Optimizations
//! - 🔲 Parallel sheet parsing for large ODS files
//! - 🔲 Streaming API for memory-constrained environments
//! - 🔲 Incremental parsing for large documents
//! - 🔲 Background saving
//!
//! ### Advanced Styling
//! - 🔲 Custom style creation
//! - 🔲 Style inheritance and cascading
//! - 🔲 Page layout manipulation
//! - 🔲 Header and footer customization
//!
//! ## References
//!
//! - **ODF Toolkit** (Java): `3rdparty/odftoolkit/` - ODFDOM framework, validation tools
//! - **odfpy** (Python): `3rdparty/odfpy/` - Pure Python ODF manipulation library
//! - **ODF Specification**: ISO/IEC 26300:2015 (ODF 1.2)
//! - **calamine** (Rust): Spreadsheet parsing patterns
//!
//! # Examples
//!
//! ```no_run
//! use litchi::odf::{Document, Spreadsheet, Presentation};
//!
//! # fn main() -> litchi::Result<()> {
//! // Open a text document
//! let mut doc = Document::open("document.odt")?;
//! let text = doc.text()?;
//!
//! // Open a spreadsheet
//! let mut sheet = Spreadsheet::open("data.ods")?;
//! let csv = sheet.to_csv()?;
//!
//! // Open a presentation
//! let mut pres = Presentation::open("slides.odp")?;
//! let slide_count = pres.slide_count()?;
//!
//! # Ok(())
//! # }
//! ```

/// ODF constants, MIME types, and XML tags
pub mod constants;
/// Cell coordinate conversion utilities (A1 notation)
pub mod coordinates;
/// Core ODF parsing functionality
mod core;
/// ODF data type conversions (Boolean, Date, DateTime, Duration)
pub mod datatype;
/// ODF XML element classes
pub mod elements;
/// ODF presentation (.odp) support
mod odp;
/// ODF spreadsheet (.ods) support
mod ods;
/// ODF text document (.odt) support
mod odt;

// Re-export common utilities for convenience
// These are used across all Office formats, not ODF-specific
pub use crate::common::RGBColor as Color;
pub use crate::common::unit::{Length, LengthUnit};

// Re-export main types for convenience
pub use odp::{MutablePresentation, Presentation, PresentationBuilder};
pub use ods::{
    Cell as SCell, CellValue, MutableSpreadsheet, Row as SRow, Sheet, Spreadsheet,
    SpreadsheetBuilder,
};
pub use odt::{Document, DocumentBuilder, MutableDocument};

// Re-export shapes for presentations
pub use odp::{Shape, Slide};

// Re-export document element types for unified API (for ODT tables)
pub use elements::table::{Table, TableCell as Cell, TableRow as Row};
pub use elements::text::Paragraph;
pub use elements::text::Span as Run; // Span is equivalent to Run in ODF

// Re-export parser types for document element iteration
pub use elements::parser::{DocumentOrderElement, DocumentParser};
