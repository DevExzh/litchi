//! Numbers-specific structured extraction.

use crate::Result;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use litchi_numbers::Table;

/// Extract tables directly into the canonical Numbers semantic model.
pub(super) fn extract(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    if !TableDataExtractor::has_table_models(object_index) {
        return Ok(Vec::new());
    }
    TableDataExtractor::new(bundle, object_index).extract_all_semantic_tables()
}
