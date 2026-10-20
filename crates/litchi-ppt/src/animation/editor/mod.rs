//! Snapshot editor for persisted `PowerPoint` animation and timing records.
//!
//! The record-level editor behind [`super::Edit`] stays crate-private; the
//! facade keeps the public read types compact while the implementation
//! is split into semantic state, transactional snapshot rewrites, validation,
//! and focused tests. Actions, hyperlinks, sounds, commands, and media
//! references remain inert record metadata; this owner never resolves or
//! executes them.

mod semantic;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use semantic::{Editor, EditorLimits};
pub use semantic::{LegacyShapeAnimation, Scope, Timeline};
