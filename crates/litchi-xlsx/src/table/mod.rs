//! Layered SpreadsheetML table ownership.
//!
//! Typed table models and semantic validation live in `model`, bounded XML
//! conversion in `codec`, package serialization in `package`, and focused
//! regression coverage in `tests`.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_table_xml;
pub use model::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction, validate_table,
};
pub use package::{serialize_table, write_table_xml};

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MAX_COLUMNS: usize = 16_384;
const MAX_SORT_CONDITIONS: usize = 64;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_EXCEL_COLUMN: u32 = 16_384;
const MAX_EXCEL_ROW: u32 = 1_048_576;

fn xml_error(error: impl std::fmt::Display) -> crate::error::Error {
    crate::error::Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

fn limit(resource: &str) -> crate::error::Error {
    crate::error::invalid(format!("table {resource} limit exceeded"))
}
