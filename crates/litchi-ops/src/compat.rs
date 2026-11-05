//! Element wrapper enums carried over from the `litchi` facade so the family
//! views can return them (record 0003 §2.4).
//!
//! These are the facade's `document::{Paragraph, Run, Table, Row, Cell,
//! CellMerge, DocumentElement}` and `presentation::Slide` types, moved
//! unchanged except that their reads return the unified, typed
//! [`Error`](crate::Error) instead of a string-flattened leaf error. The
//! facade's Markdown adapter and the Python bindings consume them while the
//! unified API grows its node view.
//!
//! **Every item in this module is scheduled for deletion in M1**, when the
//! generic node view and the per-format ADR 0007 vocabulary replace these
//! wrapper enums (design §16.1). New code should not build on them.

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub mod word;

#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
pub mod slide;
