//! Typed state owned by the XLSB worksheet cell reader.

use crate::conditional_formatting::Formatting;
use crate::hyperlinks::Hyperlink;
use crate::merged_cells::MergedCell;
use crate::package::cell::CellHeader;
use crate::package::data_validation::{Settings, Validation};
use crate::package::formula::{Context, Group, ParsedFormula};
use crate::package::records::Stream;
use crate::package::shared_strings::SharedString;
use crate::package::web_extension_bindings::Binding;
use crate::sheet::{AutoFilter, ColumnInfo, RowInfo, SheetProtection, StrongProtection};
use litchi_core::sheet::CellValue;
use std::io::{Read, Seek};
use std::sync::Arc;

pub(super) struct ParsedFormulaCell {
    pub(super) header: CellHeader,
    pub(super) cached_value: CellValue,
    pub(super) formula: ParsedFormula,
    pub(super) flags: u16,
}

/// Dimensions of a worksheet
#[derive(Debug, Clone, Copy)]
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct Dimensions {
    pub start: (u32, u32),
    pub end: (u32, u32),
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
impl Dimensions {
    pub fn len(&self) -> usize {
        ((self.end.0 - self.start.0 + 1) * (self.end.1 - self.start.1 + 1)) as usize
    }
}

/// XLSB cells reader
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct CellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub(super) iter: Stream<RS>,
    pub(super) shared_strings: &'a [SharedString],
    pub(super) formula_context: &'a Context,
    pub(super) cell_xf_count: usize,
    pub(super) dimensions: Dimensions,
    pub(super) current_row: u32,
    pub(super) last_row: Option<u32>,
    pub(super) buf: Vec<u8>,
    pub(super) pending_record: Option<(crate::raw::Kind, Vec<u8>)>,
    pub(super) formula_groups: Vec<Arc<Group>>,
    /// Merged cells found in the worksheet
    pub merged_cells: Vec<MergedCell>,
    /// Hyperlinks found in the worksheet
    pub hyperlinks: Vec<Hyperlink>,
    /// Column formatting records found before sheet data.
    pub column_infos: Vec<ColumnInfo>,
    /// Row header metadata found within sheet data.
    pub row_infos: Vec<RowInfo>,
    /// Worksheet AutoFilter range.
    pub auto_filter: Option<AutoFilter>,
    /// Worksheet protection settings.
    pub sheet_protection: Option<SheetProtection>,
    /// ISO strong password-verifier metadata.
    pub strong_sheet_protection: Option<StrongProtection>,
    /// Classic worksheet data-validation rules.
    pub data_validations: Vec<Validation>,
    /// UI settings from the classic validation collection.
    pub data_validation_settings: Option<Settings>,
    /// UI settings from the Office 2013 validation collection.
    pub data_validation14_settings: Option<Settings>,
    /// Classic and Office 2013 conditional-formatting blocks in stream order.
    pub conditional_formattings: Vec<Formatting>,
    /// Inert Office Add-in bindings from the worksheet WEBEXTENSIONS collection.
    pub web_extension_bindings: Vec<Binding>,
    /// Sheet views from the worksheet WSVIEWS collection.
    pub views: Vec<litchi_sheet::view::View>,
    pub(super) saw_web_extension_collection: bool,
}
