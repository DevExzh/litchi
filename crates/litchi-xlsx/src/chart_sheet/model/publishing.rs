//! Inert chartsheet web-publishing children.

use crate::error::{Error, Result};

/// Schema-complete `ST_WebSourceType` values for inert web-publishing metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebSourceType {
    Sheet,
    PrintArea,
    AutoFilter,
    Range,
    Chart,
    PivotTable,
    Query,
    Label,
}

impl WebSourceType {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Sheet => "sheet",
            Self::PrintArea => "printArea",
            Self::AutoFilter => "autoFilter",
            Self::Range => "range",
            Self::Chart => "chart",
            Self::PivotTable => "pivotTable",
            Self::Query => "query",
            Self::Label => "label",
        }
    }

    pub(in crate::chart_sheet) fn parse(value: &str) -> Result<Self> {
        match value {
            "sheet" => Ok(Self::Sheet),
            "printArea" => Ok(Self::PrintArea),
            "autoFilter" => Ok(Self::AutoFilter),
            "range" => Ok(Self::Range),
            "chart" => Ok(Self::Chart),
            "pivotTable" => Ok(Self::PivotTable),
            "query" => Ok(Self::Query),
            "label" => Ok(Self::Label),
            _ => Err(Error::Invalid(format!(
                "invalid web publish sourceType '{value}'"
            ))),
        }
    }
}

/// One `webPublishItem`; all paths, names, and references are opaque inert strings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPublishItem {
    pub id: u32,
    pub div_id: String,
    pub source_type: WebSourceType,
    pub source_ref: Option<String>,
    pub source_object: Option<String>,
    pub destination_file: String,
    pub title: Option<String>,
    /// `None` preserves the schema default; no publishing action is ever performed.
    pub auto_republish: Option<bool>,
}

/// A present `webPublishItems` collection is schema-required to be non-empty.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPublishItems {
    /// Preserves explicit `count`; when present it must equal `items.len()`.
    pub count: Option<u32>,
    pub items: Vec<WebPublishItem>,
}
