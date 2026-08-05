//! Chart future records (FRT) owned by the legacy BIFF chart stream.
//!
//! The public surface is grouped by the semantic record family instead of by
//! the historical flat source files.  All records remain inert: unknown,
//! reserved, and unsupported bytes are retained by their owning record.
//!
//! The `info`, `label`, and `blocks` facades expose the records formerly
//! grouped in the implementation-only `records` module.  Keeping that
//! implementation module private gives callers short names in a chart/FRT
//! context without duplicating models.

mod records;

/// Chart future-record version and CFrtId range metadata.
pub mod info {
    pub use super::records::{Info, RecordRange, Version};
}

/// Axis-label future-record metadata.
pub mod label {
    pub use super::records::{Alignment, CatLab};
}

/// StartBlock and EndBlock chart future-record scopes.
pub mod blocks {
    pub use super::records::{BlockKind, EndBlock, StartBlock};
}

/// StartObject, EndObject, and FrtWrapper records.
pub mod wrapper;

/// CrtMlFrt and its continuation-chain assembly.
pub mod continuation;

pub use blocks::{BlockKind, EndBlock, StartBlock};
pub use continuation::CrtMlFrt;
pub use info::{Info, RecordRange, Version};
pub use label::{Alignment, CatLab};
pub use wrapper::{EndObject, ObjectKind, StartObject, Wrapper};
