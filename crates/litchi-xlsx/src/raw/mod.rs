//! Validated low-level workbook catalog records.
//!
//! Relationship IDs and native sheet IDs are deliberately isolated here.

mod catalog;
mod formula;
pub mod namespace;
pub(crate) mod strings;
pub(crate) mod worksheet;

pub use catalog::{parse_catalog, parse_sheet};

/// Physical visibility value retained from `sheet/@state`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Visibility {
    Visible,
    Hidden,
    VeryHidden,
    Unknown(Box<str>),
}

/// One validated `workbook/sheets/sheet` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sheet {
    pub name: String,
    pub relationship_id: String,
    pub sheet_id: u32,
    pub visibility: Visibility,
}

/// One inert SpreadsheetML defined-name record.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DefinedName {
    pub name: String,
    pub reference: String,
    pub comment: Option<String>,
    pub local_sheet_id: Option<u32>,
    pub custom_menu: Option<String>,
    pub description: Option<String>,
    pub help: Option<String>,
    pub status_bar: Option<String>,
    pub shortcut_key: Option<String>,
    pub hidden: bool,
    pub function: bool,
    pub vb_procedure: bool,
    pub xlm: bool,
    pub function_group_id: Option<u32>,
    pub publish_to_server: bool,
    pub workbook_parameter: bool,
}

/// Workbook-level pivot-cache relationship record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCache {
    pub cache_id: u32,
    pub relationship_id: String,
}

/// Validated catalog extracted from `workbook.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Catalog {
    pub sheets: Vec<Sheet>,
    pub active_sheet_index: usize,
    pub uses_1904_date_system: bool,
    pub defined_names: Vec<DefinedName>,
    pub pivot_caches: Vec<PivotCache>,
    pub external_reference_ids: Vec<String>,
}
