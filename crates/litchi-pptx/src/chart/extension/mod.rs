//! Layered Microsoft `ChartEx` (`cx:chartSpace`) semantics.

pub mod codec;
pub mod model;
pub mod package;

pub use codec::CONTENT_TYPE;
pub use model::*;
pub use package::{load, read, related};

/// Borrowed `ChartEx` package part.
pub use package::Part;
