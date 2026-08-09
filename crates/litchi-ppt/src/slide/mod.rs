//! Slide parsing and management with high-performance zero-copy design.

pub mod directory;
pub mod factory;
pub mod notes;
pub mod types;

// Re-export main types
#[allow(
    clippy::module_name_repetitions,
    reason = "`SlideDirectory` and `SlideDirectoryEntry` are the established public API names under `slide::`; renaming them would break downstream crates"
)]
pub use directory::{SlideDirectory, SlideDirectoryEntry};
#[allow(
    clippy::module_name_repetitions,
    reason = "`SlideData` and `SlideFactory` are the established public API names under `slide::`; renaming them would break downstream crates"
)]
pub use factory::{SlideData, SlideFactory};
pub use notes::SpeakerNotes;
pub use types::{ParsedComment, ParsedSlideTiming, Slide};
