/// Common types, traits, and utilities shared across formats.
///
/// This module provides unified types and traits used by both OLE2 (legacy)
/// and OOXML (modern) implementations, ensuring a consistent API for users.
pub mod error;
pub mod metadata;
pub mod shapes;
pub mod style;

// Re-export commonly used types
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use shapes::{ShapeType, PlaceholderType};
pub use style::{RGBColor, Length};

