/// EMF Record Definitions
///
/// Complete set of EMF record structures based on [MS-EMF] specification
///
/// This module provides zero-copy parsing of EMF records using the `zerocopy` crate
/// for maximum performance and minimal memory allocation.
pub mod bitmap;
pub mod drawing;
pub mod objects;
pub mod path;
pub mod state;
pub mod text;
pub mod types;

// Re-export commonly used types
pub use bitmap::*;
pub use drawing::*;
pub use objects::*;
pub use path::*;
pub use state::*;
pub use text::*;
pub use types::*;
