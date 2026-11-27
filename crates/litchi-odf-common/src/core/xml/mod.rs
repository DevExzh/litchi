//! Shared ODF XML part models.
//!
//! Each standard XML part has a focused module so callers can distinguish
//! content, styles, metadata, and the common UTF-8 storage primitive without
//! putting all part models in one implementation file.

pub mod content;
pub mod meta;
pub mod part;
pub mod styles;

#[cfg(test)]
mod tests;

pub use content::Content;
pub use meta::Meta;
pub use part::XmlPart;
pub use styles::Styles;
