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
/// - `Comment`: A comment with author metadata and body paragraphs
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
/// use litchi_ole::doc::Package;
///
/// // Open a document
/// let mut package = Package::open("document.doc")?;
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
pub mod bookmark;
pub mod comment;
pub mod document;
pub mod footnote;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod revision;
pub mod shapes;
pub mod table;

/// DOC file writing
pub mod writer;

pub use bookmark::Bookmark;
pub use comment::{Comment, CommentDateTime, CommentExtendedMetadata};
pub use document::Document;
pub use footnote::{Endnote, Footnote};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::{Image, ImageError};
pub use package::Package;
pub use paragraph::{Paragraph, Run};
pub use parts::numbering::{ListLevel, ListTables, NumberFormat};
pub use parts::styles::{
    StyleDefinition, StyleFlags, StyleKind, StylePost2000, StyleSheet, StyleSheetHeader,
};
pub use revision::{
    DisplayFieldRevisionMark, NumberingRevisionMark, RevisionKind, RevisionMark, RevisionReason,
    SectionRevisionMark,
};
pub use shapes::DocShape;
pub use table::{Cell, Row, Table};
pub use writer::{
    BookmarkEntry, CharacterFormatting, CommentEntry, DisplayFieldRevision, DocWriteError,
    DocWriter, FormattingRevision, LineSpacing, NumberingRevision, ParagraphBorder,
    ParagraphBorderStyle, ParagraphBorders, ParagraphFormatting, TextBoxTightWrap, TextRevision,
};

/// Crate-native ordered document element returned by [`Document::elements`].
///
/// The umbrella `litchi` crate maps this into its public `DocumentElement`
/// variants. Keeping it crate-local avoids a reverse dependency from
/// `litchi-ole` back to the umbrella's `document` types.
#[derive(Debug, Clone)]
pub enum DocElement {
    /// A paragraph element.
    Paragraph(Box<Paragraph>),
    /// A table element.
    Table(Box<Table>),
}
