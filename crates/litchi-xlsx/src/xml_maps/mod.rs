//! Layered SpreadsheetML Custom XML Maps ownership.
//!
//! Typed contextual values live in [`model`], bounded XML conversion lives in
//! [`codec`], workbook relationship traversal lives in [`package`], and focused
//! regression coverage lives in [`tests`]. Schema and binding payloads remain
//! inert, bounded XML rather than being resolved or executed.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapSchema};
pub use package::{
    load_from_package, load_from_package_with_conformance, remove_from_package, store_in_package,
};

fn invalid(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
}
