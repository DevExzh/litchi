//! Compatibility adapter for the canonical XLSX table model and XML codec.
//!
//! The owner crate implements the bounded SpreadsheetML table grammar and
//! semantic validation.  This module keeps the historical host path and
//! translates owner failures at the OOXML facade boundary.

use litchi_core::sheet::Result as SheetResult;

use crate::error::OoxmlError;

pub use litchi_xlsx::table::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction, validate_table,
    write_table_xml,
};

/// Parse one table part while retaining the historical host result type.
pub fn parse_table_xml(xml: &str) -> SheetResult<Option<Table>> {
    litchi_xlsx::table::parse_table_xml(xml.as_bytes())
        .map_err(|error| boxed_owner_error(map_owner_error(error)))
}

/// Serialize one table through the canonical owner codec.
pub fn serialize_table(table: &Table) -> SheetResult<String> {
    litchi_xlsx::table::serialize_table(table)
        .map_err(|error| boxed_owner_error(map_owner_error(error)))
}

fn boxed_owner_error(error: OoxmlError) -> Box<dyn std::error::Error + Send + Sync> {
    Box::new(error)
}

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}
