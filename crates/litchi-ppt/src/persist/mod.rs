//! Persist pointer system for mapping slide IDs to byte offsets.
//!
//! Based on Apache POI's `PersistPtrHolder` and related infrastructure.

pub mod mapping;
pub mod ptr_holder;

// Re-export main types
#[allow(
    clippy::module_name_repetitions,
    reason = "`PersistMapping` is the established public API name re-exported from the crate root; renaming it would break downstream crates"
)]
pub use mapping::PersistMapping;
#[allow(
    clippy::module_name_repetitions,
    reason = "`PersistPtrHolder` is the established public API name re-exported from the crate root; renaming it would break downstream crates"
)]
pub use ptr_holder::PersistPtrHolder;
