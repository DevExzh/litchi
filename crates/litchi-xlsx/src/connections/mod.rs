//! Layered SpreadsheetML external data connection owner.
//!
//! Typed connection models live in the model module, bounded XML conversion in the codec
//! module, and workbook relationship integration in the package module. External targets
//! remain inert data and are never dereferenced.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::*;
pub use package::{
    load_from_package, remove_from_package, store_in_package,
    store_in_package_with_query_table_validator,
};

fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
