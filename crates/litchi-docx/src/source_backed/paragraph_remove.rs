//! Exact source-backed removal of deliberately plain main-story paragraphs.
//!
//! The implementation shares the closed plain-paragraph grammar and exact
//! package publication machinery with [`super::paragraph_copy`]. This module
//! gives removal callers a focused public namespace without duplicating that
//! security and preservation closure.

pub use super::paragraph_copy::{
    Error, Limits, Refusal, RemovalCommit as Commit, RemovalEdit as Edit,
    RemovalEffectReport as EffectReport, RemovalPatch as Patch, RemovalPublication as Publication,
    Result, Snapshot,
};
