use super::super::formatting::FormattingManager;
use super::data_validation::{DataValidation, DataValidationBiffPayload};
use super::named_range::DefinedNameRecordOptions;
use super::worksheet::WritableWorksheet;
use crate::encryption::WriterEncryption;
use crate::error::{Error, Result};
use crate::page_setup::{PrintComments, PrintErrors, PrintOrder, PrintOrientation};
use crate::{DifferentialFormat, TableStyle, TableStyles, XfProperty};
use std::collections::{HashMap, HashSet};
/// Public configuration for adding a pivot table via [`Writer::add_pivot_table`].
#[derive(Debug, Clone)]
pub struct PivotTableConfig {
    /// Pivot table name.
    pub name: String,
    /// Source type (0x0001 = Worksheet, 0x0002 = External).
    pub source_type: u16,

    // -- Source data range --
    /// Name of the worksheet that holds the source data.
    pub source_sheet_name: String,
    /// First row of the source data range (0-based, **including** the header row).
    pub source_first_row: u16,
    /// Last row of the source data range (0-based, inclusive).
    pub source_last_row: u16,
    /// First column of the source data range (0-based).
    pub source_first_col: u16,
    /// Last column of the source data range (0-based).
    pub source_last_col: u16,

    // -- Output range --
    /// First row of the pivot table output.
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    /// First header row in the output.
    pub first_header_row: u16,
    /// First data row in the output.
    pub first_data_row: u16,
    /// First data column in the output.
    pub first_data_col: u16,
    /// Data field header name (e.g. "Values").
    pub data_field_name: String,
    /// Axis for the data field header (0=none, 1=row, 2=col, 4=page, 8=data).
    pub data_axis: u16,
    /// Position of data label within the axis.
    pub data_position: u16,
    /// Field definitions.
    pub fields: Vec<PivotFieldConfig>,
    /// Data item (value field) definitions.
    pub data_items: Vec<PivotDataItemConfig>,
    /// Page field entries: `(item_index, field_index, object_id)`.
    pub page_entries: Vec<(u16, u16, u16)>,
    /// Source data rows for the pivot cache (fSaveData).
    ///
    /// Each inner `Vec` has one entry per field in the same order as `fields`.
    /// String fields use [`PivotCacheValue::StringIndex`] (index into that
    /// field's `cache_items`), numeric fields use [`PivotCacheValue::Number`].
    ///
    /// When non-empty, SXDBB + SXNUM records are written to the cache stream
    /// and the SXDB `fSaveData` flag is set.
    pub source_data: Vec<Vec<PivotCacheValue>>,
}

/// A single pivot field definition.
#[derive(Debug, Clone)]
pub struct PivotFieldConfig {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    /// Number of subtotals.
    pub subtotal_count: u16,
    /// Subtotal function bitmask.
    pub subtotal_flags: u16,
    /// Items belonging to this field.
    pub items: Vec<PivotItemConfig>,
    /// Optional SXVD display name override (`None` → use cache name, i.e. cch=0xFFFF).
    pub name: Option<String>,
    /// Source column name used in the pivot cache SXFDB record.
    /// This is the actual header text from the source data range.
    pub cache_name: String,
    /// Unique source data values for this field's cache items.
    /// These become SXSTRING records in the pivot cache stream.
    /// For data-axis (numeric) fields, leave this empty.
    pub cache_items: Vec<crate::PivotCacheItem>,
    /// Whether this is a numeric (data-axis) field.
    ///
    /// Numeric fields use SXFDB flags `0x0560` and contribute SXNUM records
    /// (instead of SXDBB indices) in the cache source data.
    pub is_numeric: bool,
    /// Optional numeric, calendar, or discrete grouping definition.
    pub grouping: Option<crate::PivotCacheGrouping>,
}

/// A single cell value in the pivot cache source data.
#[derive(Debug, Clone, Copy)]
pub enum PivotCacheValue {
    /// Index into the field's `cache_items` (for string fields).
    StringIndex(u8),
    /// Index into the field's typed shared-item list.
    SharedItemIndex(u16),
    /// Raw numeric value (for numeric/data-axis fields).
    Number(f64),
    Boolean(bool),
    Error(crate::PivotCacheError),
    DateTime(crate::PivotCacheDateTime),
    Empty,
}

impl PivotCacheValue {
    pub(super) fn shared_item(&self) -> Option<crate::PivotCacheItem> {
        match self {
            Self::Boolean(value) => Some(crate::PivotCacheItem::Boolean(*value)),
            Self::Error(value) => Some(crate::PivotCacheItem::Error(*value)),
            Self::DateTime(value) => Some(crate::PivotCacheItem::DateTime(*value)),
            Self::Empty => Some(crate::PivotCacheItem::Empty),
            Self::StringIndex(_) | Self::SharedItemIndex(_) | Self::Number(_) => None,
        }
    }
}

pub(super) fn validate_pivot_table_config(config: &PivotTableConfig) -> Result<()> {
    if config.source_first_row > config.source_last_row
        || config.source_first_col > config.source_last_col
        || config.first_row > config.last_row
        || config.first_col > config.last_col
    {
        return Err(Error::InvalidCellReference(
            "PivotTable source or output range is reversed".to_string(),
        ));
    }
    if [
        config.source_first_col,
        config.source_last_col,
        config.first_col,
        config.last_col,
        config.first_data_col,
    ]
    .into_iter()
    .any(|col| col > u16::from(u8::MAX))
        || !(config.first_row..=config.last_row).contains(&config.first_header_row)
        || !(config.first_row..=config.last_row).contains(&config.first_data_row)
        || !(config.first_col..=config.last_col).contains(&config.first_data_col)
    {
        return Err(Error::InvalidCellReference(
            "PivotTable location is outside its BIFF8 output grid".to_string(),
        ));
    }
    u16::try_from(config.fields.len()).map_err(|_| {
        Error::InvalidData("PivotTable field count exceeds BIFF8 capacity".to_string())
    })?;
    let expected_rows = usize::from(config.source_last_row - config.source_first_row);
    if !config.source_data.is_empty() && config.source_data.len() != expected_rows {
        return Err(Error::InvalidData(format!(
            "PivotCache source row count {} does not match source range row count {expected_rows}",
            config.source_data.len()
        )));
    }
    let mut group_children = vec![None; config.fields.len()];
    for (field_index, field) in config.fields.iter().enumerate() {
        if let Some(crate::PivotCacheGrouping::Discrete(grouping)) = &field.grouping {
            let base = usize::from(grouping.base_field_index);
            if base >= config.fields.len() || base == field_index {
                return Err(Error::InvalidData(format!(
                    "PivotCache grouping field {field_index} has invalid base field {base}"
                )));
            }
            if group_children[base].replace(field_index).is_some() {
                return Err(Error::InvalidData(format!(
                    "PivotCache base field {base} has multiple grouping children"
                )));
            }
            let mut cursor = base;
            let mut seen = vec![false; config.fields.len()];
            while let Some(crate::PivotCacheGrouping::Discrete(parent)) =
                &config.fields[cursor].grouping
            {
                if seen[cursor] {
                    return Err(Error::InvalidData(
                        "PivotCache grouping chain contains a cycle".to_string(),
                    ));
                }
                seen[cursor] = true;
                cursor = usize::from(parent.base_field_index);
                if cursor >= config.fields.len() {
                    break;
                }
            }
        }
    }
    for (field_index, field) in config.fields.iter().enumerate() {
        u16::try_from(field.cache_items.len()).map_err(|_| {
            Error::InvalidData(format!(
                "PivotCache field {field_index} has too many shared items"
            ))
        })?;
        if field.is_numeric && !field.cache_items.is_empty() && field.grouping.is_none() {
            return Err(Error::InvalidData(format!(
                "numeric PivotCache field {field_index} must use inline Number rows"
            )));
        }
        for (item_index, item) in field.cache_items.iter().enumerate() {
            item.validate()?;
            if matches!(item, crate::PivotCacheItem::Number(_))
                && !matches!(field.grouping, Some(crate::PivotCacheGrouping::Numeric(_)))
            {
                return Err(Error::InvalidData(format!(
                    "non-numeric PivotCache field {field_index} cannot contain numeric shared item {item_index}"
                )));
            }
            if field.cache_items[..item_index].contains(item) {
                return Err(Error::InvalidData(format!(
                    "PivotCache field {field_index} contains duplicate shared item {item_index}"
                )));
            }
        }
        if let Some(grouping) = &field.grouping {
            let group_items = grouping.group_items();
            if group_items.is_empty() || group_items.len() > usize::from(u16::MAX) {
                return Err(Error::InvalidData(format!(
                    "PivotCache grouping field {field_index} has invalid group-item count"
                )));
            }
            for (index, item) in group_items.iter().enumerate() {
                item.validate()?;
                if group_items[..index].contains(item) {
                    return Err(Error::InvalidData(format!(
                        "PivotCache grouping field {field_index} has duplicate group item {index}"
                    )));
                }
            }
            match grouping {
                crate::PivotCacheGrouping::Numeric(value) => {
                    if !value.start.is_finite()
                        || !value.end.is_finite()
                        || !value.step.is_finite()
                        || value.start >= value.end
                        || value.step <= 0.0
                    {
                        return Err(Error::InvalidData(format!(
                            "PivotCache numeric grouping field {field_index} has invalid bounds or step"
                        )));
                    }
                    if field
                        .cache_items
                        .iter()
                        .any(|item| !matches!(item, crate::PivotCacheItem::Number(_)))
                    {
                        return Err(Error::InvalidData(format!(
                            "PivotCache numeric grouping field {field_index} has nonnumeric original items"
                        )));
                    }
                },
                crate::PivotCacheGrouping::Date(value) => {
                    if value.start >= value.end || value.step == 0 || value.step > i16::MAX as u16 {
                        return Err(Error::InvalidData(format!(
                            "PivotCache date grouping field {field_index} has invalid bounds or step"
                        )));
                    }
                    if field.cache_items.iter().any(|item| {
                        !matches!(
                            item,
                            crate::PivotCacheItem::DateTime(_) | crate::PivotCacheItem::Empty
                        )
                    }) {
                        return Err(Error::InvalidData(format!(
                            "PivotCache date grouping field {field_index} has invalid original items"
                        )));
                    }
                },
                crate::PivotCacheGrouping::Discrete(value) => {
                    if !field.cache_items.is_empty() || field.is_numeric {
                        return Err(Error::InvalidData(format!(
                            "discrete PivotCache grouping field {field_index} must be derived"
                        )));
                    }
                    let base_items = config.fields[usize::from(value.base_field_index)]
                        .cache_items
                        .len();
                    if value.item_to_group.len() != base_items {
                        return Err(Error::InvalidData(format!(
                            "PivotCache grouping field {field_index} mapping is not exhaustive"
                        )));
                    }
                    let mut used = vec![false; value.group_items.len()];
                    for mapped in &value.item_to_group {
                        let mapped = usize::from(*mapped);
                        if mapped >= used.len() {
                            return Err(Error::InvalidData(format!(
                                "PivotCache grouping field {field_index} mapping index is out of range"
                            )));
                        }
                        used[mapped] = true;
                    }
                    if used.iter().any(|used| !used) {
                        return Err(Error::InvalidData(format!(
                            "PivotCache grouping field {field_index} contains an unused group item"
                        )));
                    }
                },
            }
        }
        for item in &field.items {
            let visible_count = field
                .grouping
                .as_ref()
                .map_or(field.cache_items.len(), |grouping| {
                    grouping.group_items().len()
                });
            if item.item_type == 0 && usize::from(item.cache_index) >= visible_count {
                return Err(Error::InvalidData(format!(
                    "PivotTable field {field_index} SXVI cache index {} is out of range",
                    item.cache_index
                )));
            }
        }
    }
    for (row_index, row) in config.source_data.iter().enumerate() {
        let source_field_count = config
            .fields
            .iter()
            .filter(|field| !matches!(field.grouping, Some(crate::PivotCacheGrouping::Discrete(_))))
            .count();
        if row.len() != source_field_count {
            return Err(Error::InvalidData(format!(
                "PivotCache row {row_index} has {} values for {} fields",
                row.len(),
                source_field_count
            )));
        }
        let mut row_values = row.iter();
        for (field_index, field) in config.fields.iter().enumerate() {
            if matches!(field.grouping, Some(crate::PivotCacheGrouping::Discrete(_))) {
                continue;
            }
            let value = row_values.next().unwrap();
            if field.is_numeric && field.grouping.is_none() {
                match value {
                    PivotCacheValue::Number(number) if number.is_finite() => {},
                    PivotCacheValue::Number(_) => {
                        return Err(Error::InvalidData(format!(
                            "PivotCache row {row_index} field {field_index} is non-finite"
                        )));
                    },
                    _ => {
                        return Err(Error::InvalidData(format!(
                            "PivotCache row {row_index} field {field_index} does not match numeric field type"
                        )));
                    },
                }
                continue;
            }
            let index = match value {
                PivotCacheValue::StringIndex(index) => usize::from(*index),
                PivotCacheValue::SharedItemIndex(index) => usize::from(*index),
                PivotCacheValue::Number(number) if matches!(field.grouping, Some(crate::PivotCacheGrouping::Numeric(_))) => {
                    field.cache_items.iter().position(|candidate| candidate == &crate::PivotCacheItem::Number(*number)).ok_or_else(|| {
                        Error::InvalidData(format!("PivotCache row {row_index} field {field_index} numeric value is absent from original items"))
                    })?
                },
                PivotCacheValue::Number(_) => return Err(Error::InvalidData(format!("PivotCache row {row_index} field {field_index} does not match shared field type"))),
                typed => {
                    let item = typed.shared_item().expect("typed shared value");
                    field.cache_items.iter().position(|candidate| candidate == &item).ok_or_else(|| {
                        Error::InvalidData(format!("PivotCache row {row_index} field {field_index} value is absent from shared items"))
                    })?
                },
            };
            if index >= field.cache_items.len() {
                return Err(Error::InvalidData(format!(
                    "PivotCache row {row_index} field {field_index} shared index {index} is out of range"
                )));
            }
        }
    }
    Ok(())
}

/// A single pivot item.
#[derive(Debug, Clone)]
pub struct PivotItemConfig {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    /// Option flags.
    pub flags: u16,
    /// Cache index.
    pub cache_index: u16,
    /// Optional item name override.
    pub name: Option<String>,
}

/// A pivot data item (value field).
#[derive(Debug, Clone)]
pub struct PivotDataItemConfig {
    /// Index of the source field in the pivot cache.
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum, 1=Count, 2=Average, 3=Max, 4=Min, ...
    pub function: u16,
    /// Display format flags.
    pub display_format: u16,
    /// Base field index (for "show values as").
    pub base_field_index: u16,
    /// Base item index.
    pub base_item_index: u16,
    /// Number format index.
    pub num_format_index: u16,
    /// Optional name override.
    pub name: String,
}

fn column_to_letters(col: u16) -> String {
    let mut col_index = col as u32;
    let mut buf = Vec::new();

    loop {
        let rem = (col_index % 26) as u8;
        buf.push((b'A' + rem) as char);
        col_index /= 26;
        if col_index == 0 {
            break;
        }
        col_index -= 1;
    }

    buf.iter().rev().collect()
}

pub(super) fn a1_cell(row: u32, col: u16) -> String {
    let col_str = column_to_letters(col);
    let row_idx = row + 1;
    format!("{col_str}{row_idx}")
}

/// Cell value type for writing
#[derive(Debug, Clone)]
pub enum CellValue {
    /// String value
    String(String),
    /// Number value (f64)
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Formula (stored as string)
    Formula(String),
    /// Blank/empty cell
    Blank,
}

/// BIFF8 worksheet print/page setup written by `Writer::set_page_setup`.
#[derive(Debug, Clone, PartialEq)]
pub struct PageSetupOptions {
    pub print_headers: bool,
    pub print_gridlines: bool,
    pub header: String,
    pub footer: String,
    pub horizontally_centered: bool,
    pub vertically_centered: bool,
    pub left_margin_inches: f64,
    pub right_margin_inches: f64,
    pub top_margin_inches: f64,
    pub bottom_margin_inches: f64,
    pub paper_size: u16,
    pub scale_percent: u16,
    pub starting_page_number: Option<i16>,
    pub fit_width_pages: u16,
    pub fit_height_pages: u16,
    pub print_order: PrintOrder,
    pub orientation: Option<PrintOrientation>,
    pub black_and_white: bool,
    pub draft_quality: bool,
    pub comments: PrintComments,
    pub errors: PrintErrors,
    pub horizontal_resolution_dpi: u16,
    pub vertical_resolution_dpi: u16,
    pub header_margin_inches: f64,
    pub footer_margin_inches: f64,
    pub copies: u16,
    /// Opaque DEVMODE bytes. They are serialized but never interpreted or executed.
    pub printer_driver_data: Option<Vec<u8>>,
    /// Even/first-page header/footer text and display flags; `None` emits no
    /// `HeaderFooter` record.
    pub header_footer: Option<crate::HeaderFooter>,
}

impl Default for PageSetupOptions {
    fn default() -> Self {
        Self {
            print_headers: false,
            print_gridlines: false,
            header: String::new(),
            footer: String::new(),
            horizontally_centered: false,
            vertically_centered: false,
            left_margin_inches: 0.75,
            right_margin_inches: 0.75,
            top_margin_inches: 1.0,
            bottom_margin_inches: 1.0,
            paper_size: 1,
            scale_percent: 100,
            starting_page_number: None,
            fit_width_pages: 1,
            fit_height_pages: 1,
            print_order: PrintOrder::DownThenOver,
            orientation: Some(PrintOrientation::Portrait),
            black_and_white: false,
            draft_quality: false,
            comments: PrintComments::None,
            errors: PrintErrors::Displayed,
            horizontal_resolution_dpi: 600,
            vertical_resolution_dpi: 600,
            header_margin_inches: 0.5,
            footer_margin_inches: 0.5,
            copies: 1,
            printer_driver_data: None,
            header_footer: None,
        }
    }
}

/// BIFF8 worksheet default dimensions and outline workspace settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorksheetLayoutOptions {
    pub default_row_height_twips: u16,
    pub empty_rows_hidden: bool,
    pub default_row_height_unsynced: bool,
    pub thick_top_border: bool,
    pub thick_bottom_border: bool,
    pub default_column_width_chars: u16,
    pub max_row_outline_level: u8,
    pub max_column_outline_level: u8,
    pub row_gutter_width: u16,
    pub column_gutter_height: u16,
    pub show_automatic_page_breaks: bool,
    pub apply_styles_to_outlines: bool,
    pub summary_rows_below: bool,
    pub summary_columns_right: bool,
    pub fit_to_page: bool,
    pub synchronize_horizontal_scrolling: bool,
    pub synchronize_vertical_scrolling: bool,
    pub alternate_expression_evaluation: bool,
    pub alternate_formula_entry: bool,
}

impl Default for WorksheetLayoutOptions {
    fn default() -> Self {
        Self {
            default_row_height_twips: 255,
            empty_rows_hidden: false,
            default_row_height_unsynced: false,
            thick_top_border: false,
            thick_bottom_border: false,
            default_column_width_chars: 8,
            max_row_outline_level: 0,
            max_column_outline_level: 0,
            row_gutter_width: 0,
            column_gutter_height: 0,
            show_automatic_page_breaks: true,
            apply_styles_to_outlines: false,
            summary_rows_below: true,
            summary_columns_right: true,
            fit_to_page: false,
            synchronize_horizontal_scrolling: false,
            synchronize_vertical_scrolling: false,
            alternate_expression_evaluation: false,
            alternate_formula_entry: false,
        }
    }
}

impl WorksheetLayoutOptions {
    pub(crate) fn validate(self) -> Result<()> {
        if self.default_row_height_twips > 8179
            || (!self.empty_rows_hidden && self.default_row_height_twips == 0)
        {
            return Err(Error::InvalidData(
                "default row height must be 1..=8179, or 0..=8179 for hidden empty rows"
                    .to_string(),
            ));
        }
        if self.default_column_width_chars > 255 {
            return Err(Error::InvalidData(
                "default column width must be at most 255 characters".to_string(),
            ));
        }
        if self.max_row_outline_level > 7 || self.max_column_outline_level > 7 {
            return Err(Error::InvalidData(
                "worksheet outline levels must be 0..=7".to_string(),
            ));
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct WorkbookProtection {
    pub(super) protect_structure: bool,
    pub(super) protect_windows: bool,
    pub(super) password_hash: Option<u16>,
    pub(super) protect_revisions: bool,
    pub(super) revision_password_hash: Option<u16>,
}

#[derive(Debug, Clone)]
pub(super) struct FileSharing {
    pub(super) read_only_recommended: bool,
    pub(super) password_hash: Option<u16>,
    pub(super) user_name: String,
}
#[derive(Debug)]
pub(super) struct VbaWriteMetadata {
    pub workbook_code_name: String,
    pub project: litchi_vba::Payload,
}
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CalculationSettings {
    pub mode: crate::CalculationMode,
    pub maximum_iterations: u16,
    pub iteration_enabled: bool,
    pub iteration_delta: f64,
    pub full_precision: bool,
    pub reference_mode: crate::ReferenceMode,
    pub recalculate_before_save: bool,
    pub recalculation_engine_id: u32,
    /// Optional BIFF8 multithreaded-calculation metadata. This controls only
    /// serialized workbook settings; the writer never evaluates formulas.
    pub multithreaded_calculation: Option<crate::MultithreadedCalculation>,
    pub force_full_calculation: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorkbookEnvironmentOptions {
    pub template: bool,
    pub has_biff5_stream: bool,
    pub create_backup_copy: bool,
    pub object_display_mode: crate::ObjectDisplayMode,
    pub refresh_external_data_on_load: bool,
    pub save_external_link_values: bool,
    pub has_envelope: bool,
    pub envelope_visible: bool,
    pub envelope_initialized: bool,
    pub link_update_mode: crate::LinkUpdateMode,
    pub hide_unselected_table_borders: bool,
    pub supports_natural_language_formulas: bool,
    pub default_country_code: u16,
    pub current_country_code: u16,
}

/// Primary BIFF8 workbook window and sheet-tab navigation settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorkbookWindowOptions {
    pub horizontal_position_twips: i16,
    pub vertical_position_twips: i16,
    pub width_twips: i16,
    pub height_twips: i16,
    pub hidden: bool,
    pub minimized: bool,
    pub very_hidden: bool,
    pub show_horizontal_scrollbar: bool,
    pub show_vertical_scrollbar: bool,
    pub show_sheet_tabs: bool,
    pub group_dates_in_autofilter: bool,
    pub active_sheet_index: u16,
    pub first_visible_sheet_index: u16,
    pub selected_sheet_count: u16,
    pub sheet_tab_ratio_per_mille: u16,
}

/// BIFF8 built-in and custom function-category settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FunctionGroupOptions {
    pub built_in: crate::BuiltInFunctionCategories,
    pub custom_categories: Vec<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ExternalCacheRowOptions {
    pub row: u16,
    pub first_column: u8,
    pub values: Vec<crate::CachedValue>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ExternalSheetOptions {
    pub name: String,
    pub cache_rows: Vec<ExternalCacheRowOptions>,
}

/// An inert external-workbook directory/cache. The encoded path is serialized but never opened.
#[derive(Debug, Clone, PartialEq)]
pub struct ExternalWorkbookOptions {
    pub encoded_virtual_path: String,
    pub sheets: Vec<ExternalSheetOptions>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDefinedNameOptions {
    pub name: String,
    pub sheet_index: Option<u16>,
    pub built_in: bool,
    pub formula_bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddInFunctionOptions {
    pub name: String,
    pub unused_data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DdeOrOleItemOptions {
    pub name: String,
    pub automatic: bool,
    pub picture: bool,
    pub standard_document_name: bool,
    pub ole_link: bool,
    pub clipboard_format: crate::ClipboardFormat,
    pub displayed_as_icon: bool,
    pub storage_id: u32,
    pub matrix: Option<crate::ValueMatrix>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DdeOrOleLinkOptions {
    pub encoded_virtual_path: String,
    pub items: Vec<DdeOrOleItemOptions>,
}

fn validate_short_external_name(name: &str) -> Result<()> {
    if name.encode_utf16().count() > u8::MAX as usize {
        return Err(Error::InvalidData(
            "external name exceeds 255 UTF-16 code units".to_string(),
        ));
    }
    Ok(())
}

impl ExternalDefinedNameOptions {
    pub(crate) fn validate(&self, sheet_count: usize) -> Result<()> {
        validate_short_external_name(&self.name)?;
        if self
            .sheet_index
            .is_some_and(|index| usize::from(index) >= sheet_count)
        {
            return Err(Error::InvalidData(
                "external name sheet scope is out of range".to_string(),
            ));
        }
        if self.formula_bytes.len() > u16::MAX as usize
            || self
                .formula_bytes
                .first()
                .is_some_and(|token| !matches!(token, 0x1c | 0x3a | 0x3b | 0x3c | 0x3d))
        {
            return Err(Error::InvalidData(
                "invalid opaque external-name formula bytes".to_string(),
            ));
        }
        Ok(())
    }
}

impl AddInFunctionOptions {
    pub(crate) fn validate(&self) -> Result<()> {
        validate_short_external_name(&self.name)?;
        if self.unused_data.len() > u16::MAX as usize {
            return Err(Error::InvalidData(
                "add-in unused data exceeds u16".to_string(),
            ));
        }
        Ok(())
    }
}

impl DdeOrOleLinkOptions {
    pub(crate) fn validate(&self) -> Result<()> {
        let path_len = self.encoded_virtual_path.encode_utf16().count();
        let path_characters = self.encoded_virtual_path.chars().collect::<Vec<_>>();
        let has_ole_separator = path_characters
            .get(1..path_characters.len().saturating_sub(1))
            .is_some_and(|middle| middle.contains(&'\u{3}'));
        if !(3..=255).contains(&path_len) || !has_ole_separator {
            return Err(Error::InvalidData(
                "DDE/OLE virtual path must be a bounded encoded OLE-link".to_string(),
            ));
        }
        if self.items.is_empty() || self.items.len() > 4096 {
            return Err(Error::InvalidData(
                "DDE/OLE source must contain 1..=4096 items".to_string(),
            ));
        }
        for item in &self.items {
            validate_short_external_name(&item.name)?;
            if item.standard_document_name {
                if item.ole_link
                    || item.displayed_as_icon
                    || item.storage_id != 0
                    || item.name != "StdDocumentName"
                    || item.matrix.is_some()
                {
                    return Err(Error::InvalidData(
                        "invalid standard-document DDE item".to_string(),
                    ));
                }
            } else {
                if !item.ole_link && item.storage_id != 0 {
                    return Err(Error::InvalidData(
                        "DDE item cannot identify OLE link storage".to_string(),
                    ));
                }
                if item.displayed_as_icon && !item.ole_link {
                    return Err(Error::InvalidData(
                        "only an OLE item can be displayed as an icon".to_string(),
                    ));
                }
            }
            if let Some(matrix) = &item.matrix {
                matrix.validate()?;
            }
        }
        Ok(())
    }
}

impl ExternalWorkbookOptions {
    pub(crate) fn validate(&self) -> Result<()> {
        let path_len = self.encoded_virtual_path.encode_utf16().count();
        if !(1..=255).contains(&path_len) {
            return Err(Error::InvalidData(
                "external virtual path must be 1..=255 UTF-16 code units".to_string(),
            ));
        }
        if self.sheets.is_empty() || self.sheets.len() > 256 {
            return Err(Error::InvalidData(
                "external workbook must contain 1..=256 sheets".to_string(),
            ));
        }
        let invalid_sheet_char =
            |value: char| matches!(value, '\\' | '/' | '?' | '*' | '[' | ']' | ':');
        for sheet in &self.sheets {
            let name_len = sheet.name.encode_utf16().count();
            if !(1..=31).contains(&name_len)
                || sheet.name.chars().any(invalid_sheet_char)
                || sheet.name.starts_with('\'')
                || sheet.name.ends_with('\'')
            {
                return Err(Error::InvalidData(
                    "invalid external sheet name".to_string(),
                ));
            }
            if sheet.cache_rows.len() > i16::MAX as usize {
                return Err(Error::InvalidData(
                    "external sheet cache has too many CRN rows".to_string(),
                ));
            }
            let mut previous_end = None;
            for row in &sheet.cache_rows {
                if row.values.is_empty() || usize::from(row.first_column) + row.values.len() > 256 {
                    return Err(Error::InvalidData(
                        "external CRN column range is invalid".to_string(),
                    ));
                }
                if previous_end.is_some_and(|(previous_row, previous_column)| {
                    row.row < previous_row
                        || (row.row == previous_row
                            && usize::from(row.first_column) <= previous_column)
                }) {
                    return Err(Error::InvalidData(
                        "external CRN rows overlap or are out of order".to_string(),
                    ));
                }
                previous_end = Some((
                    row.row,
                    usize::from(row.first_column) + row.values.len() - 1,
                ));
                for value in &row.values {
                    match value {
                        crate::CachedValue::Number(number) if !number.is_finite() => {
                            return Err(Error::InvalidData(
                                "external cached number must be finite".to_string(),
                            ));
                        },
                        crate::CachedValue::Text(text) if text.encode_utf16().count() > 255 => {
                            return Err(Error::InvalidData(
                                "external cached string exceeds 255 UTF-16 code units".to_string(),
                            ));
                        },
                        _ => {},
                    }
                }
            }
        }
        Ok(())
    }
}

impl Default for FunctionGroupOptions {
    fn default() -> Self {
        Self {
            built_in: crate::BuiltInFunctionCategories::Fourteen,
            custom_categories: Vec::new(),
        }
    }
}

impl FunctionGroupOptions {
    pub(crate) fn validate(&self) -> Result<()> {
        if usize::from(self.built_in.count()) + self.custom_categories.len() > 256 {
            return Err(Error::InvalidData(
                "function category count exceeds 256".to_string(),
            ));
        }
        let mut unique = HashSet::with_capacity(self.custom_categories.len());
        for name in &self.custom_categories {
            if name.encode_utf16().count() > 32 {
                return Err(Error::InvalidData(
                    "function category name exceeds 32 UTF-16 code units".to_string(),
                ));
            }
            if !unique.insert(name) {
                return Err(Error::InvalidData(
                    "function category names must be unique".to_string(),
                ));
            }
        }
        Ok(())
    }
}

impl Default for WorkbookWindowOptions {
    fn default() -> Self {
        Self {
            horizontal_position_twips: i16::MAX,
            vertical_position_twips: i16::MAX,
            width_twips: 0x4b2d,
            height_twips: 0x1e62,
            hidden: false,
            minimized: false,
            very_hidden: false,
            show_horizontal_scrollbar: true,
            show_vertical_scrollbar: true,
            show_sheet_tabs: true,
            group_dates_in_autofilter: true,
            active_sheet_index: 0,
            first_visible_sheet_index: 0,
            selected_sheet_count: 1,
            sheet_tab_ratio_per_mille: 600,
        }
    }
}

impl WorkbookWindowOptions {
    pub(crate) fn validate_intrinsic(self) -> Result<()> {
        if self.width_twips < 1 || self.height_twips < 1 {
            return Err(Error::InvalidData(
                "workbook window dimensions must be positive".to_string(),
            ));
        }
        if self.sheet_tab_ratio_per_mille > 1000 {
            return Err(Error::InvalidData(
                "sheet tab ratio must be at most 1000".to_string(),
            ));
        }
        if self.very_hidden && !self.hidden {
            return Err(Error::InvalidData(
                "very hidden workbook windows must also be hidden".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn validate_for_sheet_count(self, sheet_count: usize) -> Result<()> {
        self.validate_intrinsic()?;
        if sheet_count == 0 || sheet_count > 4112 {
            return Err(Error::InvalidData(
                "RRTabId writer supports 1..=4112 sheets".to_string(),
            ));
        }
        if usize::from(self.active_sheet_index) >= sheet_count
            || usize::from(self.first_visible_sheet_index) >= sheet_count
            || usize::from(self.selected_sheet_count) > sheet_count
        {
            return Err(Error::InvalidData(
                "workbook window tab reference is outside the sheet collection".to_string(),
            ));
        }
        Ok(())
    }
}

impl Default for WorkbookEnvironmentOptions {
    fn default() -> Self {
        Self {
            template: false,
            has_biff5_stream: false,
            create_backup_copy: false,
            object_display_mode: crate::ObjectDisplayMode::ShowAll,
            refresh_external_data_on_load: false,
            save_external_link_values: true,
            has_envelope: false,
            envelope_visible: false,
            envelope_initialized: false,
            link_update_mode: crate::LinkUpdateMode::Prompt,
            hide_unselected_table_borders: false,
            supports_natural_language_formulas: false,
            default_country_code: 1,
            current_country_code: 1,
        }
    }
}

impl WorkbookEnvironmentOptions {
    pub(crate) fn book_bool_bits(self) -> u16 {
        u16::from(!self.save_external_link_values)
            | (u16::from(self.has_envelope) << 2)
            | (u16::from(self.envelope_visible) << 3)
            | (u16::from(self.envelope_initialized) << 4)
            | ((match self.link_update_mode {
                crate::LinkUpdateMode::Prompt => 0,
                crate::LinkUpdateMode::Never => 1,
                crate::LinkUpdateMode::Silent => 2,
            }) << 5)
            | (u16::from(self.hide_unselected_table_borders) << 8)
    }
}

impl Default for CalculationSettings {
    fn default() -> Self {
        Self {
            mode: crate::CalculationMode::Automatic,
            maximum_iterations: 100,
            iteration_enabled: false,
            iteration_delta: 0.001,
            full_precision: true,
            reference_mode: crate::ReferenceMode::A1,
            recalculate_before_save: true,
            recalculation_engine_id: 0x000E_EA35,
            multithreaded_calculation: None,
            force_full_calculation: false,
        }
    }
}
/// A complete caller-defined BIFF8 table-style family.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomTableStyles {
    differential_formats: Vec<DifferentialFormat>,
    catalog: TableStyles,
}

impl CustomTableStyles {
    pub fn try_new(
        differential_formats: Vec<DifferentialFormat>,
        catalog: TableStyles,
    ) -> Result<Self> {
        let value = Self {
            differential_formats,
            catalog,
        };
        value.validate_structure()?;
        Ok(value)
    }

    pub fn try_from_styles(
        differential_formats: Vec<DifferentialFormat>,
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
        custom_styles: Vec<TableStyle>,
    ) -> Result<Self> {
        Self::try_new(
            differential_formats,
            TableStyles::try_with_custom_styles(
                default_table_style,
                default_pivot_style,
                custom_styles,
            )?,
        )
    }

    pub fn differential_formats(&self) -> &[DifferentialFormat] {
        &self.differential_formats
    }

    pub const fn catalog(&self) -> &TableStyles {
        &self.catalog
    }

    fn validate_structure(&self) -> Result<()> {
        if self.differential_formats.len() > usize::from(u16::MAX) + 1 {
            return Err(Error::InvalidData(
                "custom table styles contain more than 65,536 DXFs".to_string(),
            ));
        }
        for differential_format in &self.differential_formats {
            differential_format.to_record_bytes()?;
        }
        self.catalog
            .to_family_record_bytes(self.differential_formats.len())?;
        Ok(())
    }

    pub(crate) fn validate(&self, formatting: &FormattingManager) -> Result<()> {
        self.validate_structure()?;
        for (dxf_index, differential_format) in self.differential_formats.iter().enumerate() {
            for property in differential_format.properties().properties() {
                if let XfProperty::NumberFormatId(number_format_id) = property
                    && !formatting.contains_number_format_id(*number_format_id)
                {
                    return Err(Error::InvalidData(format!(
                        "custom table-style DXF {dxf_index} references undefined number format {number_format_id}"
                    )));
                }
            }
        }
        Ok(())
    }
}

fn is_builtin_table_style(name: &str) -> bool {
    [
        ("TableStyleLight", 21u16),
        ("TableStyleMedium", 28),
        ("TableStyleDark", 11),
    ]
    .iter()
    .any(|(prefix, maximum)| {
        name.strip_prefix(prefix)
            .and_then(|suffix| suffix.parse::<u16>().ok())
            .is_some_and(|value| (1..=*maximum).contains(&value))
    })
}

pub(super) fn validate_list_object_style(
    name: &str,
    custom: Option<&CustomTableStyles>,
) -> Result<()> {
    if is_builtin_table_style(name)
        || custom.is_some_and(|styles| {
            styles.catalog().custom_styles().iter().any(|style| {
                style.name().eq_ignore_ascii_case(name) && style.is_available_for_tables()
            })
        })
    {
        return Ok(());
    }
    Err(Error::InvalidData(format!(
        "table style {name:?} is not a table-capable built-in or configured custom style"
    )))
}

pub(super) fn validate_list_object_relationships(
    worksheets: &[WritableWorksheet],
    custom: Option<&CustomTableStyles>,
    defined_names: &[super::named_range::DefinedName],
    defined_name_records: &[(DefinedNameRecordOptions, crate::DefinedNameFutureRecords)],
) -> Result<()> {
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    let defined_names = defined_names
        .iter()
        .map(|name| name.name.to_lowercase())
        .chain(
            defined_name_records
                .iter()
                .map(|(name, _)| name.name.to_lowercase()),
        )
        .collect::<HashSet<_>>();
    for worksheet in worksheets {
        for (index, table) in worksheet.list_objects.iter().enumerate() {
            table.validate()?;
            validate_list_object_style(table.style().unwrap().name(), custom)?;
            if !ids.insert(table.id()) || !names.insert(table.name().to_lowercase()) {
                return Err(Error::InvalidData(
                    "duplicate workbook table identifier or name".to_string(),
                ));
            }
            if defined_names.contains(&table.name().to_lowercase()) {
                return Err(Error::InvalidData(
                    "table name collides with a workbook defined name".to_string(),
                ));
            }
            if worksheet.list_objects[..index]
                .iter()
                .any(|existing| existing.range().overlaps(table.range()))
            {
                return Err(Error::InvalidData(
                    "table ranges overlap within the worksheet".to_string(),
                ));
            }
            if worksheet.auto_filter.is_some_and(|filter| {
                u32::from(table.range().first_row()) <= filter.last_row
                    && filter.first_row <= u32::from(table.range().last_row())
                    && table.range().first_column() <= filter.last_col
                    && filter.first_col <= table.range().last_column()
            }) {
                return Err(Error::InvalidData(
                    "table range overlaps the worksheet AutoFilter".to_string(),
                ));
            }
        }
        if let Some(sort) = &worksheet.sort_data
            && let crate::writer::sort::Parent::Table { id } = sort.parent()
            && !worksheet
                .list_objects
                .iter()
                .any(|table| table.id().value() == id)
        {
            return Err(Error::InvalidData(
                "table SortData references an unknown ListObject identifier".to_string(),
            ));
        }
    }
    Ok(())
}

pub(super) fn prepare_data_validation(
    validation: &DataValidation,
) -> Result<DataValidationBiffPayload> {
    if let Some(title) = validation.input_title.as_ref()
        && title.encode_utf16().count() > 32
    {
        return Err(Error::InvalidData(
            "Input message title must be at most 32 characters".to_string(),
        ));
    }
    if let Some(text) = validation.input_message.as_ref()
        && text.encode_utf16().count() > 255
    {
        return Err(Error::InvalidData(
            "Input message text must be at most 255 characters".to_string(),
        ));
    }
    if let Some(title) = validation.error_title.as_ref()
        && title.encode_utf16().count() > 32
    {
        return Err(Error::InvalidData(
            "Error message title must be at most 32 characters".to_string(),
        ));
    }
    if let Some(text) = validation.error_message.as_ref()
        && text.encode_utf16().count() > 225
    {
        return Err(Error::InvalidData(
            "Error message text must be at most 225 characters".to_string(),
        ));
    }
    validation.validation_type.to_biff_payload()
}

/// XLS file writer
///
/// Provides methods to create and modify XLS (BIFF8) files.
pub struct Writer {
    /// Worksheets to write
    pub(super) worksheets: Vec<WritableWorksheet>,
    /// Shared string table
    pub(super) shared_strings: Vec<String>,
    /// String to index mapping for deduplication
    pub(super) string_map: HashMap<String, u32>,
    /// Workbook-level defined names (named ranges).
    pub(super) defined_names: Vec<super::named_range::DefinedName>,
    pub(super) defined_name_records:
        Vec<(DefinedNameRecordOptions, crate::DefinedNameFutureRecords)>,
    pub(super) fmt: FormattingManager,
    /// Total number of string occurrences (including duplicates) for SST.cstTotal
    pub(super) sst_total: u32,
    pub(super) workbook_protection: Option<WorkbookProtection>,
    pub(super) file_sharing: Option<FileSharing>,
    /// Use 1904 date system (Mac) instead of 1900 (Windows)
    pub(super) use_1904_dates: bool,
    pub(super) calculation_settings: CalculationSettings,
    pub(super) vba_metadata: Option<VbaWriteMetadata>,
    pub(super) environment_options: WorkbookEnvironmentOptions,
    pub(super) workbook_window_options: WorkbookWindowOptions,
    pub(super) function_group_options: FunctionGroupOptions,
    pub(super) external_workbooks: Vec<ExternalWorkbookOptions>,
    pub(super) external_names: Vec<Vec<ExternalDefinedNameOptions>>,
    pub(super) add_in_functions: Vec<AddInFunctionOptions>,
    pub(super) dde_or_ole_links: Vec<DdeOrOleLinkOptions>,
    pub(super) custom_table_styles: Option<CustomTableStyles>,
    /// Optional inert root-level `[MS-XLS]` XML-map catalog.
    pub(super) xml_map: Option<crate::xml_map::MapInfo>,
    pub(super) book_ext: Option<crate::BookExt>,
    pub(super) theme: Option<crate::Theme>,
    /// MDX (OLAP cube) metadata emitted as the globals `METADATA` production.
    pub(super) mdx_metadata: Option<crate::MdxMetadata>,
    /// Real-time data (RTD) topics emitted as `RealTimeData` records.
    pub(super) real_time_data: Vec<crate::real_time_data::Record>,
    /// Web pages published from the workbook globals (`WebPub` records).
    pub(super) web_publications: Vec<crate::WebPub>,
    pub(super) xf_extensions: Vec<crate::XfExt>,
    pub(super) style_extensions: Vec<crate::StyleExt>,
    /// Optional inert Office Toolbars (`XCB`) stream.
    pub(super) toolbar: Option<crate::Wrapper<'static>>,
    pub(super) encryption: Option<WriterEncryption>,
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}
