//! Semantic `WordprocessingML` main-document facade.
//!
//! The document owner keeps its public model, document-XML codec, and
//! package-bound orchestration in separate layers while retaining the
//! historical `crate::document` entry point.
//!
//! Both package owners start the primary transaction with `edit()`. An edit
//! is isolated from its source; committing it produces a document commit for
//! the package publication boundary.
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
//! let source = source_backed::Package::from_reader(bytes.as_slice())?;
//! assert!(!source.edit()?.commit()?.diagnostics().changed());
//! # Ok(())
//! # }
//! ```

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;
mod transaction;

pub(crate) use transaction::MAX_DOCUMENT_XML_BYTES;
pub(crate) use transaction::validate_paragraph_text;
pub(crate) use transaction::{ManagedAdmission, ManagedNamespaceAdmission, ManagedParserAdmission};
pub(crate) use transaction::{TransferPart, TransferRelationship, durable_transfer_operations};

pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use transaction::{
    Commit, CompactionPolicy, Composition, CompositionLimits, Diagnostics, Edit, Error, History,
    HistoryLimits, HyperlinkTextReplacement, JoinError, MergeChoice, Operation,
    ParagraphHyperlinkAddress, ParagraphTextReplacement, ParagraphTransfer, Patch, PreparedEdit,
    Refusal, Result, RevisionAction, RevisionKind, RevisionSelector, Snapshot, SubEditConflict,
    SubEditJoinFailure, TableCellAddress, ThreeWayError, ThreeWayMergeFailure, ThreeWayPlan,
    TransferGraph, TransferRefusal,
};
