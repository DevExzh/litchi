//! Slide parsing and management with high-performance zero-copy design.

pub mod factory;
pub mod types;

// Re-export main types
pub use factory::{SlideData, SlideFactory};
pub use types::Slide;
