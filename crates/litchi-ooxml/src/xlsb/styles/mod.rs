//! XLSB styles submodules
//!
//! This module contains specialized parsers for different style components.

pub mod alignment_parser;
pub mod border_parser;

// Re-export main types for public API
pub use alignment_parser::Alignment;
pub use border_parser::Border;

// Re-export detailed types for internal use
#[allow(unused_imports)]
pub(crate) use alignment_parser::{HorizontalAlignment, VerticalAlignment};
#[allow(unused_imports)]
pub(crate) use border_parser::{BorderSide, BorderStyle};
