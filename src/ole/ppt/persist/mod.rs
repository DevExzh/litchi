/// Persist pointer system for mapping slide IDs to byte offsets.
///
/// Based on Apache POI's PersistPtrHolder and related infrastructure.

pub mod ptr_holder;
pub mod mapping;

// Re-export main types
pub use ptr_holder::PersistPtrHolder;
pub use mapping::PersistMapping;

