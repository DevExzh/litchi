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
pub use model::{
    CellReference, ColumnBinding, Limits, MappedTable, SingleCellBinding, SingleCellTable, XPath,
    XmlDataType,
};
pub use patch::{Commit, Patch};
pub use snapshot::{ReadLimits, Snapshot};
pub use transaction::Transaction;
pub use validation::{validate_binding_map_ids, validate_catalog};

pub use host_api::{
    parse_xml_map_info, parse_xml_map_info_with_conformance,
    parse_xml_map_info_with_conformance_and_limits, parse_xml_map_info_with_limits,
    patch_xml_map_info_source, patch_xml_map_info_source_with_limits, serialize_xml_map_info,
    serialize_xml_map_info_with_limits, validate_xml_map_info, validate_xml_map_info_with_limits,
};
pub use litchi_ooxml_common::spreadsheet_xml_maps::{
    DataBinding, XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapLimits,
    XmlMapSchema, XmlSchema,
};

pub use crate::package::error::{Error, Result};
