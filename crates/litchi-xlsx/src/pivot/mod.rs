//! Typed `SpreadsheetML` pivot vocabulary and part codecs.
//!
//! The compact types at this facade are the semantic API. XML-shaped
//! cache/table records live in the nested modules so callers can opt into
//! package-level fidelity without duplicating the semantic models.

use std::collections::HashMap;

pub mod cache;
pub mod chart;
pub mod fields;
pub mod filters;
#[path = "server_formats/mod.rs"]
pub(crate) mod formats;
pub mod reader;
pub mod styles;
#[cfg(test)]
mod tests;
pub mod writer;

pub use formats::{cached_unique_names, table_data};

pub mod server_formats {
    //! Source-bound support for the MS-XLSX `pivotTableServerFormats`
    //! extension of a non-worksheet PivotTable.
    //!
    //! [`crate::Workbook::pivot_table_server_formats`] reads the typed values;
    //! [`crate::Package::edit_pivot_table`] starts an [`Edit`] that stages
    //! [`ServerFormat`] values or explicit [`Attribute`] operations through
    //! [`Format`]; [`Edit::commit`] returns a [`Commit`] holding the new
    //! package, the exact reversible [`Patch`] and the [`Diagnostics`], and
    //! [`crate::Package::apply_pivot_table`] replays a patch. The source
    //! package never changes.

    pub use super::formats::{
        Attribute, Commit, Diagnostics, Edit, Error, Format, LimitKind, Patch, PivotTableSelector,
        PivotTableServerFormats, PivotTableView, ServerFormat, Snapshot, load,
    };
    #[cfg(test)]
    pub(crate) use super::formats::{apply_patch, edit};
    pub(crate) use super::formats::{edit_workbook, view_workbook};
}

pub use cached_unique_names::{
    CachedUniqueName, CachedUniqueNames, PivotCacheFieldHandle, PivotCacheFieldSelector,
    PivotCacheHandle, PivotCacheId, PivotCacheSelector,
};
pub use reader::{
    read_pivot_cache_definition, read_pivot_cache_records, read_pivot_table_definition,
    read_pivot_tables,
};
pub use server_formats::{
    PivotTableSelector, PivotTableServerFormats, PivotTableView, ServerFormat,
};
pub use table_data::{
    PivotCellAddress, PivotCellType, PivotRowView, PivotTableDataDiagnostic, PivotTableDataView,
    PivotValueCellExtraView, PivotValueCellView,
};
pub use writer::{write_pivot_cache_definition, write_pivot_cache_records, write_pivot_table};

/// The worksheet area occupied by a pivot field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotAxis {
    Row,
    Column,
    Filter,
    Data,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotValueFunction {
    Sum,
    Count,
    Average,
    Min,
    Max,
    Custom,
}

#[derive(Debug, Clone)]
pub struct PivotFieldRole {
    pub field_name: String,
    pub axis: PivotAxis,
    pub position: u32,
}

#[derive(Debug, Clone)]
pub struct PivotDataField {
    pub field_name: String,
    pub function: PivotValueFunction,
    pub display_name: Option<String>,
}

#[derive(Debug, Clone)]
pub struct PivotCacheField {
    pub name: String,
    pub shared_items: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct PivotCacheDefinition {
    pub id: u32,
    pub source_ref: Option<String>,
    pub fields: Vec<PivotCacheField>,
}

#[derive(Debug, Clone)]
pub struct PivotTable {
    pub name: String,
    pub source_sheet: Option<String>,
    pub source_ref: Option<String>,
    pub field_names: Vec<String>,
    pub sheet_name: String,
    pub cache_id: u32,
    pub location_ref: String,
    pub row_fields: Vec<PivotFieldRole>,
    pub column_fields: Vec<PivotFieldRole>,
    pub filter_fields: Vec<PivotFieldRole>,
    pub data_fields: Vec<PivotDataField>,
}

impl PivotTable {
    #[must_use]
    pub fn fields_by_axis(&self, axis: PivotAxis) -> &[PivotFieldRole] {
        match axis {
            PivotAxis::Row => &self.row_fields,
            PivotAxis::Column => &self.column_fields,
            PivotAxis::Filter => &self.filter_fields,
            PivotAxis::Data => &[],
        }
    }

    #[must_use]
    pub fn data_fields_map(&self) -> HashMap<&str, &PivotDataField> {
        let mut map = HashMap::with_capacity(self.data_fields.len());
        for field in &self.data_fields {
            map.insert(field.field_name.as_str(), field);
        }
        map
    }
}

// These are XML vocabulary enums used by the lossless nested codecs. They
// remain below the pivot facade rather than being duplicated in each codec.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ItemType {
    Data,
    Default,
    Sum,
    CountA,
    Avg,
    Max,
    Min,
    Product,
    Count,
    StdDev,
    StdDevP,
    Var,
    VarP,
    Grand,
    Blank,
}

impl ItemType {
    #[must_use]
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Data => "data",
            Self::Default => "default",
            Self::Sum => "sum",
            Self::CountA => "countA",
            Self::Avg => "avg",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::Count => "count",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevP",
            Self::Var => "var",
            Self::VarP => "varP",
            Self::Grand => "grand",
            Self::Blank => "blank",
        }
    }

    #[must_use]
    pub fn parse_str(value: &str) -> Self {
        match value {
            "default" => Self::Default,
            "sum" => Self::Sum,
            "countA" => Self::CountA,
            "avg" => Self::Avg,
            "max" => Self::Max,
            "min" => Self::Min,
            "product" => Self::Product,
            "count" => Self::Count,
            "stdDev" => Self::StdDev,
            "stdDevP" => Self::StdDevP,
            "var" => Self::Var,
            "varP" => Self::VarP,
            "grand" => Self::Grand,
            "blank" => Self::Blank,
            _ => Self::Data,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum SortType {
    Manual,
    Ascending,
    Descending,
}

impl SortType {
    #[must_use]
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Manual => "manual",
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }

    #[must_use]
    pub fn parse_str(value: &str) -> Self {
        match value {
            "ascending" => Self::Ascending,
            "descending" => Self::Descending,
            _ => Self::Manual,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum AxisType {
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}

impl AxisType {
    #[must_use]
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
        }
    }

    #[must_use]
    pub fn parse_str(value: &str) -> Self {
        match value {
            "axisCol" => Self::AxisCol,
            "axisPage" => Self::AxisPage,
            "axisValues" => Self::AxisValues,
            _ => Self::AxisRow,
        }
    }
}
