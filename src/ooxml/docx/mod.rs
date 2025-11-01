/// Word (.docx) document support.
///
/// This module provides parsing and manipulation of Microsoft Word documents
/// in the Office Open XML (OOXML) format (.docx files).
///
/// # Architecture
///
/// The module is organized around these key types:
/// - `Package`: The overall .docx file package
/// - `Document`: The main document content and API
/// - `Paragraph`: A paragraph with runs
/// - `Run`: A text run with formatting
/// - `Table`: A table with rows and cells
/// - `Section`: A document section with page properties
/// - `Styles`: Collection of document styles
/// - `DocumentPart`: The core document.xml part
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// // Open a document
/// let package = Package::open("document.docx")?;
/// let doc = package.document()?;
///
/// // Access paragraphs and runs
/// for para in doc.paragraphs()? {
///     println!("Paragraph: {}", para.text()?);
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
///     }
/// }
///
/// // Access tables
/// for table in doc.tables()? {
///     for row in table.rows()? {
///         for cell in row.cells()? {
///             println!("Cell: {}", cell.text()?);
///         }
///     }
/// }
///
/// // Access sections
/// let mut sections = doc.sections()?;
/// for section in sections.iter_mut() {
///     println!("Page width: {:?}", section.page_width());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod document;
pub mod enums;
pub mod format;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod section;
pub mod styles;
pub mod table;
pub mod template;
pub mod writer;

pub use document::Document;
pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use package::Package;
pub use paragraph::{Paragraph, Run, RunProperties};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use styles::{Style, Styles};
pub use table::{Cell, Row, Table, VMergeState};
// Re-export shared formatting types
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
// Re-export writer types
pub use writer::{
    CellProperties, ListType, MutableDocument, MutableHyperlink, MutableInlineImage,
    MutableParagraph, MutableRun, MutableTable, Note, PageNumberFormat, PageOrientation,
    RunContent, SectionProperties, TableBorder, TableBorders,
};
