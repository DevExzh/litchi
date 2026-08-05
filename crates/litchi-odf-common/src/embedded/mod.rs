//! Inert embedded-resource inventory for OpenDocument XML parts.

mod model;
mod reader;

pub use model::{Kind, Object, Parameter, Root, Source};
pub use reader::{scan_flat, scan_package};
