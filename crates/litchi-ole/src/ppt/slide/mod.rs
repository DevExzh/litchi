//! Slide parsing and management with high-performance zero-copy design.

pub mod directory;
pub mod factory;
pub mod notes;
pub mod types;

// Re-export main types
pub use directory::{SlideDirectory, SlideDirectoryEntry};
pub use factory::{SlideData, SlideFactory};
pub use notes::SpeakerNotes;
pub use types::{ParsedComment, ParsedSlideTiming, Slide};
