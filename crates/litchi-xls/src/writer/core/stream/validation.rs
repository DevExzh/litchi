//! Workbook-stream input validation and identity staging.

use crate::writer::formatting::FormattingManager;
use crate::{Error, Result};

use super::super::named_range::DefinedName as InternalDefinedName;
use super::super::worksheet::WritableWorksheet;
use super::super::{CustomTableStyles, DefinedNameRecordOptions, WorkbookWindowOptions};
use super::semantic::{PivotCacheIdentity, stage_pivot_cache_identities};

/// Validate the cross-record invariants required before BIFF bytes are
/// emitted, then stage the workbook-global pivot-cache identity map.
pub(super) fn validate_workbook_inputs(
    fmt: &FormattingManager,
    custom_table_styles: Option<&CustomTableStyles>,
    defined_names: &[InternalDefinedName],
    defined_name_records: &[(DefinedNameRecordOptions, crate::DefinedNameFutureRecords)],
    workbook_window: &WorkbookWindowOptions,
    worksheets: &[WritableWorksheet],
) -> Result<Vec<Vec<PivotCacheIdentity>>> {
    if let Some(styles) = custom_table_styles {
        styles.validate(fmt)?;
    }
    super::super::validate_list_object_relationships(
        worksheets,
        custom_table_styles,
        defined_names,
        defined_name_records,
    )?;
    workbook_window.validate_for_sheet_count(worksheets.len())?;
    let active_sheet = usize::from(workbook_window.active_sheet_index);
    let active_worksheet = worksheets.get(active_sheet).ok_or_else(|| {
        Error::InvalidData(format!(
            "active worksheet index {active_sheet} is outside the sheet collection"
        ))
    })?;
    if !active_worksheet.view.is_selected() {
        return Err(Error::InvalidData(format!(
            "active worksheet {active_sheet} must be selected in Window2"
        )));
    }
    let selected_sheet_count = worksheets
        .iter()
        .filter(|sheet| sheet.view.is_selected())
        .count();
    if selected_sheet_count != usize::from(workbook_window.selected_sheet_count) {
        return Err(Error::InvalidData(format!(
            "Window1 selected sheet count {} disagrees with Window2 selected state ({selected_sheet_count})",
            workbook_window.selected_sheet_count
        )));
    }

    stage_pivot_cache_identities(worksheets)
}
