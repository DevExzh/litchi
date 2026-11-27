//! Cached table data emitted with a chart definition.

use super::extensions::Extensions;

#[derive(Debug, Clone, PartialEq, Default)]
pub enum CachedValue {
    #[default]
    Empty,
    Float(f64),
    Percentage(f64),
    Currency {
        value: f64,
        currency: String,
    },
    Boolean(bool),
    Date(String),
    Time(String),
    String(String),
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct CachedCell {
    pub value: CachedValue,
    /// An `OpenDocument` formula stored as inert text; this crate never evaluates it.
    pub formula: Option<String>,
    pub repeated: u32,
}

impl CachedCell {
    #[must_use]
    pub fn new(value: CachedValue) -> Self {
        Self {
            value,
            formula: None,
            repeated: 1,
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct CachedRow {
    pub cells: Vec<CachedCell>,
    pub repeated: u32,
}

impl CachedRow {
    #[must_use]
    pub fn new(cells: Vec<CachedCell>) -> Self {
        Self { cells, repeated: 1 }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CachedTable {
    pub name: String,
    pub columns: u32,
    pub header_columns: u32,
    pub header_rows: Vec<CachedRow>,
    pub rows: Vec<CachedRow>,
    pub extensions: Extensions,
}

impl CachedTable {
    pub fn new(name: impl Into<String>, columns: u32) -> Self {
        Self {
            name: name.into(),
            columns,
            header_columns: 0,
            header_rows: Vec::new(),
            rows: Vec::new(),
            extensions: Extensions::default(),
        }
    }
}
