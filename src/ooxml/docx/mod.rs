pub mod document;
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
/// - `DocumentPart`: The core document.xml part
/// - Various part types: `StylesPart`, `NumberingPart`, etc.
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
/// // Access paragraphs
/// for para in doc.paragraphs() {
///     println!("Paragraph text: {}", para.text());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod package;
pub mod parts;

pub use document::Document;
pub use package::Package;
