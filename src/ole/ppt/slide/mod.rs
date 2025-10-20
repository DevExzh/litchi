//! Slide parsing and management with high-performance zero-copy design.

pub mod factory;
pub mod types;

// Re-export main types
pub use factory::{SlideFactory, SlideData};
pub use types::Slide;

