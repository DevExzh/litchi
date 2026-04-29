pub mod cache;
pub mod fields;
pub mod filters;
pub mod reader;
pub mod styles;
pub mod writer;

pub use cache::{PivotCacheDefinition, PivotCacheField, PivotCacheRecords, SharedItem};
pub use fields::{DataField, FieldItem, PageField, PivotField, RowColField, RowColItem, Subtotal};
pub use filters::{PivotArea, PivotFilter, Reference};
pub use reader::{read_pivot_cache_definition, read_pivot_table_definition, read_pivot_tables};
pub use styles::{Location, PivotTableStyle};
pub use writer::{
    PivotTableDefinition, write_pivot_cache_definition, write_pivot_cache_records,
    write_pivot_table,
};

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

    pub fn parse_str(s: &str) -> Self {
        match s {
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

    pub fn parse_str(s: &str) -> Self {
        match s {
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

    pub fn parse_str(s: &str) -> Self {
        match s {
            "axisCol" => Self::AxisCol,
            "axisPage" => Self::AxisPage,
            "axisValues" => Self::AxisValues,
            _ => Self::AxisRow,
        }
    }
}
