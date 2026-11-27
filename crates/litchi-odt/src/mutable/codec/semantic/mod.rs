//! Layered semantic facade for lossless mutable ODT snapshot edits.
//!
//! The inherent `MutableDocument` API is partitioned by document concern while
//! keeping the public facade at the original codec boundary.

mod content;
mod forms;
mod metadata;
mod references;
mod structure;
mod styles;
mod tracking;
mod xforms;
