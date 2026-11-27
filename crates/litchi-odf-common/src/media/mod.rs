//! Inert image inventory and authoring resources for `OpenDocument` XML parts.

pub mod authoring;
mod model;
mod reader;

#[cfg(test)]
mod tests;

pub use model::{Image, Source};
pub use reader::{scan_content, scan_flat, scan_package};
