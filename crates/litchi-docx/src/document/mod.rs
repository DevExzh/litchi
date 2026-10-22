//! Semantic `WordprocessingML` main-document facade.
//!
//! The document owner keeps its public model, document-XML codec, and
//! package-bound orchestration in separate layers while retaining the
//! historical `crate::document` entry point.
//!
//! Both eager and source-backed `Package::edit()` retain the complete source package. Committing
//! validates and returns the immutable package through [`Commit::package`].
//! The owner type is the generic parameter of [`Edit`] and [`Commit`];
//! it defaults to the eager [`crate::Package`].
//!
//! ```
//! use litchi_docx::{Package, document, source_backed};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let bytes = include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
//!     "/../../test-data/ooxml/docx/documentProperties.docx"));
//! let package = Package::from_reader(bytes.as_slice())?;
//! let commit: document::Result<document::Commit> = package.edit()?.commit();
//! assert!(!commit?.diagnostics().changed());
//! let document: litchi_docx::Document = package.document()?;
//! drop(package);
//! assert!(document.paragraph_count().is_ok());
//! let source = source_backed::Package::from_reader(bytes.as_slice())?;
//! assert!(!source.edit()?.commit()?.diagnostics().changed());
//! # Ok(())
//! # }
//! ```

mod blocks;
mod codec;
pub mod composition;
mod limit;
mod model;
mod owned;
mod package;
pub mod raw;
#[cfg(test)]
mod tests;
mod transaction;

/// Formatting-preserving text substitution over the text of a story: the
/// main document, a header, a footer, the notes or the comments
/// (records 0274 and 0285). It is the one leaf primitive behind template
/// merge and find/replace.
///
/// [`Edit::story`], [`Edit::story_within`] and [`Edit::part_story`] read a
/// [`substitute::Story`]; [`Edit::substitute`], [`Edit::substitute_part`]
/// and their `_formatted` forms apply [`substitute::Replacement`]s planned
/// over it.
pub mod substitute {
    pub use super::transaction::substitute::{
        Field, FieldKind, OBJECT, Outcome, Paragraph, Replacement, Segment, Story, StoryKind,
        StoryPart,
    };
}

/// Field results of the main story, indexed and rewritten for field refresh
/// (record 0279).
///
/// [`Edit::fields`] reads a [`fields::Fields`] index; [`Edit::refresh_fields`]
/// applies [`fields::Update`]s planned over it. Nothing is evaluated here.
pub mod fields {
    pub use super::transaction::fields::{
        Bookmark, Field, Fields, Outcome, Paragraph, TocEntry, Update,
    };
}

pub use limit::LimitKind;
pub(crate) use owned::story_parts_of;
pub use owned::{
    CommentChange, Commit, Edit, HyperlinkTarget, NewChart, NewComment, NewNote, NewPicture,
    NewSmartArt, NewStory, NoteKind, PictureChange, SmartArtParts, StoryTarget, chart_workbook,
};
pub(crate) use transaction::MAX_DOCUMENT_XML_BYTES;
pub(crate) use transaction::validate_paragraph_text;
pub(crate) use transaction::{ManagedAdmission, ManagedNamespaceAdmission, ManagedParserAdmission};
pub(crate) use transaction::{TransferPart, TransferRelationship, durable_transfer_operations};

pub use blocks::{BlockContentControl, BlockKind, Blocks, IndexedBlock};
pub(crate) use blocks::{BlockIndex, ContainerRoot, index as index_blocks};
pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use transaction::{
    BodyAnchor, BodyBlock, CompactionPolicy, Diagnostics, ElementAnchor, ElementPath, ElementStep,
    Error, GraphChange, HyperlinkTextReplacement, MAX_NEW_TABLE_DIMENSION, NewCell, NewParagraph,
    NewRun, NewTable, Operation, ParagraphHyperlinkAddress, ParagraphTextReplacement,
    ParagraphTransfer, Patch, Refusal, Result, RevisionAction, RevisionKind, RevisionSelector,
    RunAnchor, Snapshot, TableCellAddress, TransferGraph, TransferRefusal,
};
