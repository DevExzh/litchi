/// PPT record types and parsing.
///
/// This module provides structures and functions for parsing PowerPoint binary records.

pub mod record;
pub mod document_info;
pub mod slide_info;
pub mod slide_atoms_set;

// Re-export commonly used types
pub use record::PptRecord;
pub use document_info::DocumentInfo;
pub use slide_info::SlideInfo;
pub use slide_atoms_set::SlideAtomsSet;

