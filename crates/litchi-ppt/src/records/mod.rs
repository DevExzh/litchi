//! PPT record types and parsing.
//!
//! This module provides structures and functions for parsing `PowerPoint` binary records.

pub mod document_info;
pub mod record;
pub mod slide_atoms_set;
pub mod slide_info;

// Re-export commonly used types
pub use document_info::DocumentInfo;
pub use record::Record;
#[allow(
    unused_imports,
    reason = "re-export consumed by codec modules in configurations where it is currently unused"
)]
pub(crate) use record::RecordParseSession;
pub use slide_atoms_set::SlideAtomsSet;
pub use slide_info::SlideInfo;
