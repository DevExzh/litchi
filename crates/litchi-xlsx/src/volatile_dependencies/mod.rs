//! Layered, inert `SpreadsheetML` volatile-dependency records.
//!
//! Typed semantic values live in `model`, bounded XML conversion in
//! `codec`, and workbook relationship traversal in `package`. The
//! historical `litchi_xlsx::volatile_dependencies::*` facade remains exposed
//! through the re-exports below.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::*;
pub use package::{
    load_from_package, load_from_package_with_conformance, remove_from_package, store_in_package,
};

pub(super) fn invalid(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
}
