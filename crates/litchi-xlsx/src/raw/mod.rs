//! Validated low-level workbook catalog records.
//!
//! Relationship IDs and native sheet IDs are deliberately isolated here.

mod catalog;
pub(crate) mod catalog_edit;
pub(crate) mod compact;
mod formula;
pub mod namespace;
pub(crate) mod properties_edit;
pub(crate) mod recalc;
pub(crate) mod reference_edit;
pub(crate) mod reference_scan;
pub mod shared_strings;
pub(crate) mod sheet_view_edit;
pub(crate) mod strings;
pub(crate) mod styles;
pub(crate) mod styles_edit;
pub mod web;
pub(crate) mod worksheet;
pub(crate) mod worksheet_property;

/// Bounded source selection of one worksheet coordinate or rectangular range.
///
/// This surface is an internal format-owner API. `NotEligible` is a
/// capability result, not full worksheet validation; callers must fall back
/// to the materialized worksheet parser whenever it is returned. Eligible
/// range results are sparse physical records and include
/// [`selected_worksheet::SelectedDependencies`] for the zero-based
/// shared-string and direct-cell-style indexes observed across the complete
/// stream. The one-cell [`selected_worksheet::scan`] function is retained as
/// a compatibility wrapper around [`selected_worksheet::scan_range`].
pub mod selected_worksheet {
    pub use super::worksheet::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCell, SelectedCells,
        SelectedDependencies, SelectedPayload, SelectedRecord, StreamResult, scan, scan_range,
    };
}

pub(crate) use catalog::parse_catalog_repairing;
pub use catalog::{parse_catalog, parse_catalog_with_mce, parse_sheet};

/// Check the caller's source budget before a raw parser materializes XML.
pub(crate) fn check_source_bytes(
    actual: usize,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    let maximum = limits.max_part_bytes().min(usize::MAX as u64);
    let maximum = usize::try_from(maximum).unwrap_or(usize::MAX);
    if actual > maximum {
        return Err(crate::Error::Package(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::PartBytes,
            observed: actual as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

/// Check the materialized OPC graph against the caller's raw-operation
/// policy after a staged graph operation.
pub(crate) fn check_package_limits(
    package: &litchi_opc::OpcPackage,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    let parts = package.iter_parts().count();
    if parts > limits.max_parts() {
        return Err(crate::Error::Package(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::Parts,
            observed: parts as u64,
            maximum: limits.max_parts() as u64,
        }));
    }

    let mut total = 0_u64;
    for part in package.iter_parts() {
        let bytes = package.get_part(part.partname())?.blob().len();
        check_source_bytes(bytes, limits)?;
        total = total.checked_add(bytes as u64).ok_or_else(|| {
            crate::Error::Package(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::TotalPartBytes,
                observed: u64::MAX,
                maximum: limits.max_total_part_bytes(),
            })
        })?;
    }
    if total > limits.max_total_part_bytes() {
        return Err(crate::Error::Package(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::TotalPartBytes,
            observed: total,
            maximum: limits.max_total_part_bytes(),
        }));
    }
    Ok(())
}

/// Parse only the direct worksheet-grid defaults from one `WorksheetPart`.
///
/// This focused read path shares the canonical [`crate::layout::Defaults`]
/// model and validation with the full worksheet parser without materializing
/// cells, rows, columns, or merges. Markup compatibility and the supported
/// `x14ac:dyDescent` extension are processed using the same bounded rules as
/// the full parser.
pub fn parse_worksheet_defaults(
    content: &[u8],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<Option<crate::layout::Defaults>> {
    check_source_bytes(content.len(), limits)?;
    worksheet::parse_defaults(content)
}

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

/// One inert `SpreadsheetML` defined-name record.
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
