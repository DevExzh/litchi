//! Canonical PresentationML limit dimensions.

// The definitions live at the crate error boundary so every PPTX family and
// the root error use one type. This focused path is a re-export, not a second
// enum.
pub use crate::error::{CodecLimitKind, LimitKind};
