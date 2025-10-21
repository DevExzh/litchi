//! PPT record types and parsing.
//!
//! This module provides structures and functions for parsing PowerPoint binary records.

pub mod document_info;
pub mod record;
pub mod slide_atoms_set;
pub mod slide_info;

// Re-export commonly used types
pub use document_info::DocumentInfo;
pub use record::PptRecord;
pub use slide_atoms_set::SlideAtomsSet;
pub use slide_info::SlideInfo;
