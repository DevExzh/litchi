//! Semantic `WordprocessingML` section and page-layout APIs.
//!
//! The facade intentionally keeps the section owner small and contextual. The
//! semantic values live in `model`, while `codec` owns the bounded,
//! lossless `w:sectPr` reader and writer. The mutable document writer has a
//! separate, richer section editor under `writer::section`; this module does
//! not duplicate or depend on that editor's orchestration state.

mod codec;
pub mod footnote_columns;
mod inventory;
pub mod layout;
mod model;

#[cfg(test)]
mod tests;

pub use inventory::{
    Descriptor, Inventory, Limits, Ownership, ParagraphRange, Property, PropertyValue, Selector,
    Snapshot,
};
pub use model::{
    Column, Columns, Emu, Margins, Orientation, PageSize, Reference, Section, Sections, Start,
};

pub use layout::{
    Commit as LayoutCommit, Edit as LayoutEdit, Patch as LayoutPatch,
    Publication as LayoutPublication, Snapshot as LayoutSnapshot,
};
