//! Shared diagnostics of the scoped PPTX edit families (record 0252).

/// The outcome of a scoped edit or commit.
///
/// `changed()` is the canonical no-op introspection (ADR 0034 §3); patches
/// answer the same question through `is_noop()`.
pub use litchi_core::edit::Diagnostics;
