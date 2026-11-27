//! Typed Custom XML Maps and inert BIFF12 map bindings.
//!
//! The SpreadsheetML catalog is shared with XLSX through
//! `litchi-ooxml-common`. This module owns only XLSB table-column and
//! single-cell binding records. It validates XPath lexically but never
//! resolves schemas, evaluates XPath, imports XML, or performs external I/O.

mod codec;
mod host_api;
mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;
mod validation;
mod workbook;

#[cfg(test)]
mod tests;

pub use codec::{
    SingleCellsSource, TableBindingsSource, apply_table_bindings, parse_single_cells,
    parse_table_bindings, patch_single_cells, patch_table_bindings, serialize_column_binding,
    serialize_single_cells, serialize_table_bindings,
};
pub use litchi_core::edit::Diagnostics;
pub use model::{
    CellReference, ColumnBinding, Limits, MappedTable, SingleCellBinding, SingleCellTable, XPath,
    XmlDataType,
};
pub use patch::{Commit, Patch};
pub use snapshot::{SemanticLimits, Snapshot};
pub use transaction::{Edit, EditOptions};
pub use validation::{validate_binding_map_ids, validate_catalog};

pub use host_api::{
    parse_xml_map_info, parse_xml_map_info_with_conformance, patch_xml_map_info_source,
    serialize_xml_map_info, validate_xml_map_info,
};
pub use litchi_ooxml_common::spreadsheet_xml_maps::{
    DataBinding, XmlMap, XmlMapConformance, XmlMapInfo, XmlMapLimits, XmlSchema,
};

pub use crate::package::error::{Error, Result};

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of the MapInfo part.
    MapInfoBytes,
    /// Mapped tables of the workbook.
    MappedTables,
    /// Single-cell bindings of one worksheet.
    SingleCellBindings,
}
