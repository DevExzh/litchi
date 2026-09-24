//! Shared owners for exact Microsoft chart extension vocabulary.
//!
//! These owners deliberately stop at the shared DrawingML fragment boundary.
//! Package hosts decide where a fragment is attached, and an extension is only
//! exposed here when its specification defines enough grammar to edit it
//! without guessing a containing chart type.

pub mod formatcode2;
