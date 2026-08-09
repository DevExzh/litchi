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

pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use transaction::{
    Commit, Composition, CompositionLimits, Diagnostics, Edit, History, HistoryLimits, JoinError,
    Operation, Patch, PreparedEdit, Refusal, Snapshot, SubEditJoinFailure, TransactionError,
    TransactionResult,
};
