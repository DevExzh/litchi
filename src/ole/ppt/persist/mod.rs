//! Persist pointer system for mapping slide IDs to byte offsets.
//!
//! Based on Apache POI's PersistPtrHolder and related infrastructure.

pub mod mapping;
pub mod ptr_holder;

// Re-export main types
pub use mapping::PersistMapping;
pub use ptr_holder::PersistPtrHolder;
