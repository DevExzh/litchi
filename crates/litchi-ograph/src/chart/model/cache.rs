use super::super::{Kind, cache as chart_cache};
use super::series::RowCol;

/// One cached chart value.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Blank,
}
/// Excel cached value, including the producer-specific `BoolErr` union.
#[derive(Debug, Clone, PartialEq)]
pub enum XlValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(chart_cache::Fault),
    Blank,
}

impl From<Value> for XlValue {
    fn from(value: Value) -> Self {
        match value {
            Value::Number(value) => Self::Number(value),
            Value::Text(value) => Self::Text(value),
            Value::Blank => Self::Blank,
        }
    }
}

/// Borrowed producer-neutral view of a cached value.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ValueRef<'a> {
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Error(chart_cache::Fault),
    Blank,
}

/// Producer-typed cached chart cell.
///
/// Each variant owns its producer-specific coordinate, section, and format;
/// Graph/Excel mixtures are therefore unrepresentable.
#[derive(Debug, PartialEq)]
pub enum Cache {
    /// Excel SERIESDATA cell.
    Excel {
        section: chart_cache::Index,
        row: u16,
        col: u8,
        xf: chart_cache::Xf,
        value: XlValue,
    },
    /// Standalone Graph datasheet cell.
    Graph {
        row: RowCol,
        col: RowCol,
        ifmt: chart_cache::Ifmt,
        value: Value,
    },
}

impl Cache {
    /// Creates an Excel cache cell.
    pub fn excel(
        section: chart_cache::Index,
        row: u16,
        col: u8,
        xf: chart_cache::Xf,
        value: impl Into<XlValue>,
    ) -> Self {
        Self::Excel {
            section,
            row,
            col,
            xf,
            value: value.into(),
        }
    }

    /// Creates a standalone Graph cache cell.
    pub const fn graph(row: RowCol, col: RowCol, ifmt: chart_cache::Ifmt, value: Value) -> Self {
        Self::Graph {
            row,
            col,
            ifmt,
            value,
        }
    }

    /// Producer grammar owned by this cell.
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Excel { .. } => Kind::Excel,
            Self::Graph { .. } => Kind::Graph,
        }
    }

    /// Cached value.
    pub fn value(&self) -> ValueRef<'_> {
        match self {
            Self::Excel { value, .. } => match value {
                XlValue::Number(value) => ValueRef::Number(*value),
                XlValue::Text(value) => ValueRef::Text(value),
                XlValue::Bool(value) => ValueRef::Bool(*value),
                XlValue::Error(value) => ValueRef::Error(*value),
                XlValue::Blank => ValueRef::Blank,
            },
            Self::Graph { value, .. } => match value {
                Value::Number(value) => ValueRef::Number(*value),
                Value::Text(value) => ValueRef::Text(value),
                Value::Blank => ValueRef::Blank,
            },
        }
    }
}
