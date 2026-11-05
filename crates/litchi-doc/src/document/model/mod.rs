//! Layered document state and semantic query facade.
//!
//! `Document` owns the parsed OLE2-backed state in [`state`], exposes the
//! contextual reading API from [`semantic`], and keeps table-shape invariants
//! in [`validation`]. The crate root re-exports the facade directly, so
//! callers never need to know this internal topology.

mod semantic;
mod state;
mod validation;

pub(in crate::document) use semantic::stories::StoryIndexCache;
pub use semantic::stories::{StoryBlock, StoryBlockKind, StoryBlocks, StoryIndexDiagnostics};
pub use semantic::{BodyParagraphs, HeaderFooterSlot};
pub use state::Document;

pub(in crate::document) use state::DocumentState;
