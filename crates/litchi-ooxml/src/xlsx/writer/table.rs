//! Compatibility entry point for XLSX table XML serialization.

use crate::xlsx::table::Table;
use litchi_core::sheet::Result as SheetResult;

/// Serialize a table through the canonical `litchi-xlsx` codec.
pub fn serialize_table(table: &Table) -> SheetResult<String> {
    crate::xlsx::table::serialize_table(table)
}
