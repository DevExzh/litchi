//! Inert PowerPoint 10 document-comparison metadata.
//!
//! The owner is split by responsibility: typed records live in [`model`],
//! binary record I/O in [`codec`], structural and resource checks in
//! [`validation`], and focused regression coverage in [`tests`]. Parsing
//! never compares presentations, opens external data, or executes embedded
//! content.

mod codec;
mod model;
mod validation;

pub use model::{
    DiffFlags, DiffNode, DiffRecordHeaders, DiffTree10, DiffType, DocDiffFlags, ElementType,
    MainMasterDiffFlags, POWERPOINT_DIFF_MAX_DEPTH, POWERPOINT_DIFF_MAX_RECORDS,
    ReviewingToolbarStates, ShapeDiffFlags, SlideCreationEntry, SlideDiffFlags, SlideListTable10,
    TableDiffFlags, TextDiffFlags,
};

#[cfg(test)]
mod tests;
