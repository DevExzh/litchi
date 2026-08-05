//! Bounded, inert Ribbon customization storage shared by every OOXML host.
//!
//! Ribbon callback names and image payloads remain opaque. This module only
//! validates the declared OPC graph and the Custom UI document boundary; it
//! never invokes callbacks, resolves commands, decodes images, or contacts an
//! external resource. XML is consumed without transcoding and therefore must
//! be UTF-8; an encoding declaration must say `UTF-8` case-insensitively.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Family, Limits, Set, Ui, Version};
pub use package::{load, load_with, put, put_with, remove};
