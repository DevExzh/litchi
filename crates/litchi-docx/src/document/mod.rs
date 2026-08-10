//! Semantic `WordprocessingML` main-document facade.
//!
//! The document owner keeps its public model, document-XML codec, and
//! package-bound orchestration in separate layers while retaining the
//! historical `crate::document` entry point.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;
mod transaction;

pub(crate) use transaction::{TransferPart, TransferRelationship, durable_transfer_operations};

pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use transaction::{
    Commit, Composition, CompositionLimits, Diagnostics, Edit, History, HistoryLimits,
    HyperlinkTextReplacement, JoinError, MergeChoice, Operation, ParagraphHyperlinkAddress,
    ParagraphTextReplacement, ParagraphTransfer, Patch, PreparedEdit, Refusal, RevisionKind,
    Snapshot, SubEditConflict, SubEditJoinFailure, TableCellAddress, ThreeWayError,
    ThreeWayMergeFailure, ThreeWayPlan, TransactionError, TransactionResult, TransferGraph,
    TransferRefusal,
};
