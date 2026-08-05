//! DOC file writer implementation
//!
//! This module provides functionality to create and modify Microsoft Word documents
//! in the legacy binary format (.doc files) using OLE2 structured storage.
//!
//! # Architecture
//!
//! The writer generates the complex DOC file structure including:
//! - FIB (File Information Block) - contains file metadata and stream offsets
//! - Text stream - contains the actual document text
//! - Table stream (0Table/1Table) - contains formatting and structure
//! - Data stream - contains embedded objects
//!
//! # DOC File Format Overview
//!
//! DOC files use a "piece table" architecture where:
//! 1. Text is stored in one or more pieces (continuous runs)
//! 2. Character formatting (CHP) is stored separately
//! 3. Paragraph formatting (PAP) is stored separately
//! 4. All formatting uses SPRMs (Single Property Modifiers)
//!
//! # Critical Implementation Details
//!
//! ## Stream Creation Order
//!
//! Microsoft Word requires `WordDocument` to be allocated at **sector 0** of the
//! OLE file. This is achieved by creating the `WordDocument` stream BEFORE any
//! other streams. The stream creation order in `save()` method is:
//!
//! 1. `WordDocument` → sector 0 (REQUIRED by Microsoft Word)
//! 2. `1Table` → next available sector
//!
//! ## Directory Entry Ordering
//!
//! Directory entries are sorted using Apache POI's PropertyComparator rules:
//! - Sort by name length first (shorter names before longer names)
//! - Then alphabetically (case-insensitive) for same-length names
//!
//! For DOC files, this results in the tree structure:
//! ```text
//! Root Entry
//!     └─ WordDocument (midpoint of sorted list)
//!          └─ 1Table (left child, shorter name)
//! ```
//!
//! **Note**: Stream ALLOCATION order (sector assignment) is DIFFERENT from
//! directory ENTRY order (tree structure). See `OleWriter` documentation for details.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_doc::Writer;
//!
//! let mut writer = Writer::new();
//!
//! // Add paragraphs
//! writer.add_paragraph("Hello, World!")?;
//! writer.add_paragraph("This is a second paragraph.")?;
//!
//! // Save the document
//! writer.save("output.doc")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Compatibility
//!
//! Generated DOC files are compatible with:
//! - Microsoft Word 97-2003
//! - Microsoft Word 2007+ (compatibility mode)
//! - LibreOffice Writer
//! - Apache POI (HWPF)
//! - Other OLE2-based Word readers

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    CharacterFormatting, HeaderFooterParagraph, HeaderKind, LineSpacing, ParagraphFormatting,
    WriteError, Writer,
};

pub(super) use model::pack_dttm;
