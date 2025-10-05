/// Word (.doc) document support.
///
/// This module provides parsing of Microsoft Word documents in the legacy
/// binary format (.doc files), which uses OLE2 structured storage.
///
/// # Architecture
///
/// The module is organized around these key types:
/// - `Package`: The overall .doc file package (OLE container)
/// - `Document`: The main document content and API
/// - `Paragraph`: A paragraph with runs (formatted text)
/// - `Run`: A text run with formatting
/// - `Table`: A table with rows and cells
///
/// # DOC File Structure
///
/// A .doc file is an OLE2 structured storage containing several streams:
/// - **WordDocument**: Main document stream containing the FIB and text
/// - **1Table** or **0Table**: Contains formatting and structure information
/// - **Data**: Contains embedded objects and images
/// - **\x05SummaryInformation**: Document metadata
///
/// # Example
///
/// ```rust,no_run
/// use litchi::doc::Package;
///
/// // Open a document
/// let package = Package::open("document.doc")?;
/// let doc = package.document()?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Access paragraphs
/// for para in doc.paragraphs()? {
///     println!("Paragraph: {}", para.text()?);
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

