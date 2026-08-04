//! Compatibility exports for the canonical XLSX-owned query-table codec.
//!
//! The semantic model, bounded XML codec, and package graph operations live in
//! the canonical `litchi_xlsx::query_table` module. This module retains the
//! historical host path.

pub use litchi_xlsx::query_table::*;

#[cfg(test)]
mod tests {
    #[test]
    fn loads_real_libreoffice_query_table_parts_through_workbook() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let workbook = crate::xlsx::Workbook::open(
            root.join("test-data/poi/test-data/spreadsheet/StructuredRefs-lots-with-lookups.xlsx"),
        )
        .unwrap();
        let tables = workbook.query_tables_on_sheet("Query").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].query_table().connection_id(), 2);
        assert_eq!(tables[0].query_table().name(), "Query from RDS");
    }
}
