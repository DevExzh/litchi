//! Contextual ODS document metadata.
//!
//! The Dublin Core/ODF value model and retained-source patcher stay owned by
//! `litchi-odf-common`.  This module only binds those owners to an ODS
//! package snapshot and exposes the transaction boundary used by the facade.

mod model;
mod transaction;

#[cfg(test)]
mod tests;

pub use litchi_odf_common::core::metadata::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Metadata, TemplateMetadata,
    UserDefinedMetadata, UserDefinedValueType,
};
pub use model::Snapshot;
pub use transaction::{Commit, Editor, Transaction};
