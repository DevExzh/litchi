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
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod document;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod table;

pub use document::Document;
pub use package::Package;
pub use paragraph::{Paragraph, Run};
pub use table::{Cell, Row, Table};
