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

mod codec;
mod limit;
mod model;
mod owned;
mod package;
#[cfg(test)]
mod tests;
mod transaction;

/// Transitional standalone XML transactions used by source-backed publication.
///
/// These values retain main-document XML, not an editable package. Ordinary
/// callers use [`crate::Package::edit`] or [`crate::source_backed::Package::edit`]
/// and [`crate::document::Edit`] instead.
pub mod snapshot {
    pub use super::transaction::{Commit, Edit};
}

pub use limit::LimitKind;
pub use owned::{Commit, Edit};

pub(crate) use transaction::MAX_DOCUMENT_XML_BYTES;
pub(crate) use transaction::validate_paragraph_text;
pub(crate) use transaction::{ManagedAdmission, ManagedNamespaceAdmission, ManagedParserAdmission};
pub(crate) use transaction::{TransferPart, TransferRelationship, durable_transfer_operations};

pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use transaction::{
    CompactionPolicy, Composition, CompositionLimits, Diagnostics, Error, History, HistoryLimits,
    HyperlinkTextReplacement, JoinError, MergeChoice, Operation, ParagraphHyperlinkAddress,
    ParagraphTextReplacement, ParagraphTransfer, Patch, PreparedEdit, Refusal, Result,
    RevisionAction, RevisionKind, RevisionSelector, Snapshot, SubEditConflict, SubEditJoinFailure,
    TableCellAddress, ThreeWayError, ThreeWayMergeFailure, ThreeWayPlan, TransferGraph,
    TransferRefusal,
};
