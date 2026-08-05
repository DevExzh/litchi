//! Typed SpreadsheetML pivot vocabulary and part codecs.
//!
//! The compact types at this facade are the semantic API. XML-shaped
//! cache/table records live in the nested modules so callers can opt into
//! package-level fidelity without duplicating the semantic models.

use std::collections::HashMap;

pub mod cache;
pub mod fields;
pub mod filters;
pub mod reader;
pub mod styles;
pub mod writer;

pub use reader::{
    read_pivot_cache_definition, read_pivot_cache_records, read_pivot_table_definition,
    read_pivot_tables,
};
pub use writer::{
    write_pivot_cache_definition, write_pivot_cache_records, write_pivot_table,
};

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
    pub fn fields_by_axis(&self, axis: PivotAxis) -> &[PivotFieldRole] {
        match axis {
            PivotAxis::Row => &self.row_fields,
            PivotAxis::Column => &self.column_fields,
            PivotAxis::Filter => &self.filter_fields,
            PivotAxis::Data => &[],
        }
    }

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
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Manual => "manual",
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }

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
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
        }
    }

    pub fn parse_str(value: &str) -> Self {
        match value {
            "axisCol" => Self::AxisCol,
            "axisPage" => Self::AxisPage,
            "axisValues" => Self::AxisValues,
            _ => Self::AxisRow,
        }
    }
}
