//! XLS file writer implementation
//!
//! This module provides functionality to create and modify Microsoft Excel files
//! in the legacy binary format (.xls files) using the BIFF (Binary Interchange File Format).
//!
//! # Architecture
//!
//! The writer generates BIFF8 records and uses the OLE writer to create the
//! compound document structure. It supports:
//! - Creating workbooks with multiple worksheets
//! - Writing cell values (numbers, strings, formulas, booleans)
//! - Shared string table management
//! - Basic cell formatting
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_xls::Writer;
//!
//! let mut writer = Writer::new();
//! let sheet = writer.add_worksheet("Sheet1")?;
//!
//! // Write some data
//! writer.write_string(sheet, 0, 0, "Hello")?;
//! writer.write_number(sheet, 0, 1, 42.0)?;
//!
//! // Save the file
//! writer.save("output.xls")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::super::error::{Error, Result};
use super::biff::AutoFilterConditionWrite;
use super::formatting::{CellStyle, ExtendedFormat, FormattingManager};
use crate::EncryptionProfile;
use crate::encryption::{WriterEncryption, encrypt_workbook_for_write, validate_writer_encryption};
use crate::page_setup::{PrintComments, PrintErrors, PrintOrder, PrintOrientation};
use crate::{DifferentialFormat, ListObject, TableStyle, TableStyles, XfProperty};
use litchi_cfb::writer::OleWriter;
use std::collections::{HashMap, HashSet};
use zeroize::Zeroizing;

mod comment;
mod conditional_format;
mod data_validation;
mod named_range;
pub mod shape;
mod shape_group;
mod stream;
mod worksheet;

pub use self::comment::{CommentTextRunWrite, CommentWriteOptions};
pub use self::conditional_format::{
    ConditionalFormat, ConditionalFormat12Group, ConditionalFormat12Rule, ConditionalFormat12Type,
    ConditionalFormatGroup, ConditionalFormatOperator, ConditionalFormatRange,
    ConditionalFormatRule, ConditionalFormatType, ConditionalPattern,
};
use self::data_validation::DataValidationBiffPayload;
pub use self::data_validation::{
    DataValidation, DataValidationErrorStyle, DataValidationFormulaKind, DataValidationImeMode,
    DataValidationOperator, DataValidationOptions, DataValidationRange, DataValidationTableOptions,
    DataValidationType,
};
use self::named_range::DefinedName as InternalDefinedName;
pub use self::named_range::{DefinedName, DefinedNameRecordOptions};
pub use self::shape::{
    ShapeColor, ShapeFill, ShapeKind, ShapeLine, ShapeText, ShapeTextRun, ShapeWrite,
};
pub use self::shape_group::{ShapeGroupChild, ShapeGroupWrite};
use self::worksheet::{
    AutoFilterColumnDef, AutoFilterRange, CellPos, HorizontalPageBreak, Hyperlink, MergedRange,
    PivotCellXfRole, SheetProtection, SortConfig, VerticalPageBreak, WritableCell,
    WritablePivotDataItem, WritablePivotField, WritablePivotItem, WritablePivotTable,
    WritableWorksheet,
};

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
    fn shared_item(&self) -> Option<crate::PivotCacheItem> {
        match self {
            Self::Boolean(value) => Some(crate::PivotCacheItem::Boolean(*value)),
            Self::Error(value) => Some(crate::PivotCacheItem::Error(*value)),
            Self::DateTime(value) => Some(crate::PivotCacheItem::DateTime(*value)),
            Self::Empty => Some(crate::PivotCacheItem::Empty),
            Self::StringIndex(_) | Self::SharedItemIndex(_) | Self::Number(_) => None,
        }
    }
}

fn validate_pivot_table_config(config: &PivotTableConfig) -> Result<()> {
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

fn a1_cell(row: u32, col: u16) -> String {
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
    pub(super) fn validate(self) -> Result<()> {
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
struct WorkbookProtection {
    protect_structure: bool,
    protect_windows: bool,
    password_hash: Option<u16>,
    protect_revisions: bool,
    revision_password_hash: Option<u16>,
}

#[derive(Debug, Clone)]
struct FileSharing {
    read_only_recommended: bool,
    password_hash: Option<u16>,
    user_name: String,
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
    pub(super) fn validate(&self, sheet_count: usize) -> Result<()> {
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
    pub(super) fn validate(&self) -> Result<()> {
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
    pub(super) fn validate(&self) -> Result<()> {
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
    pub(super) fn validate(&self) -> Result<()> {
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
    pub(super) fn validate(&self) -> Result<()> {
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
    pub(super) fn validate_intrinsic(self) -> Result<()> {
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

    pub(super) fn validate_for_sheet_count(self, sheet_count: usize) -> Result<()> {
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
    pub(super) fn book_bool_bits(self) -> u16 {
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

    pub(super) fn validate(&self, formatting: &FormattingManager) -> Result<()> {
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

fn validate_list_object_relationships(
    worksheets: &[WritableWorksheet],
    custom: Option<&CustomTableStyles>,
    defined_names: &[InternalDefinedName],
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

fn prepare_data_validation(validation: &DataValidation) -> Result<DataValidationBiffPayload> {
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
    worksheets: Vec<WritableWorksheet>,
    /// Shared string table
    shared_strings: Vec<String>,
    /// String to index mapping for deduplication
    string_map: HashMap<String, u32>,
    /// Workbook-level defined names (named ranges).
    defined_names: Vec<InternalDefinedName>,
    defined_name_records: Vec<(DefinedNameRecordOptions, crate::DefinedNameFutureRecords)>,
    fmt: FormattingManager,
    /// Total number of string occurrences (including duplicates) for SST.cstTotal
    sst_total: u32,
    workbook_protection: Option<WorkbookProtection>,
    file_sharing: Option<FileSharing>,
    /// Use 1904 date system (Mac) instead of 1900 (Windows)
    use_1904_dates: bool,
    calculation_settings: CalculationSettings,
    vba_metadata: Option<VbaWriteMetadata>,
    environment_options: WorkbookEnvironmentOptions,
    workbook_window_options: WorkbookWindowOptions,
    function_group_options: FunctionGroupOptions,
    external_workbooks: Vec<ExternalWorkbookOptions>,
    external_names: Vec<Vec<ExternalDefinedNameOptions>>,
    add_in_functions: Vec<AddInFunctionOptions>,
    dde_or_ole_links: Vec<DdeOrOleLinkOptions>,
    custom_table_styles: Option<CustomTableStyles>,
    book_ext: Option<crate::BookExt>,
    theme: Option<crate::Theme>,
    /// MDX (OLAP cube) metadata emitted as the globals `METADATA` production.
    mdx_metadata: Option<crate::MdxMetadata>,
    /// Real-time data (RTD) topics emitted as `RealTimeData` records.
    real_time_data: Vec<crate::RealTimeData>,
    /// Web pages published from the workbook globals (`WebPub` records).
    web_publications: Vec<crate::WebPub>,
    xf_extensions: Vec<crate::XfExt>,
    style_extensions: Vec<crate::StyleExt>,
    encryption: Option<WriterEncryption>,
}

impl Writer {
    /// Create a new XLS writer
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            shared_strings: Vec::new(),
            string_map: HashMap::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            sst_total: 0,
            fmt: FormattingManager::new(),
            workbook_protection: None,
            file_sharing: None,
            use_1904_dates: false,
            calculation_settings: CalculationSettings::default(),
            vba_metadata: None,
            environment_options: WorkbookEnvironmentOptions::default(),
            workbook_window_options: WorkbookWindowOptions::default(),
            function_group_options: FunctionGroupOptions::default(),
            external_workbooks: Vec::new(),
            external_names: Vec::new(),
            add_in_functions: Vec::new(),
            dde_or_ole_links: Vec::new(),
            custom_table_styles: None,
            book_ext: None,
            theme: None,
            mdx_metadata: None,
            real_time_data: Vec::new(),
            web_publications: Vec::new(),
            xf_extensions: Vec::new(),
            style_extensions: Vec::new(),
            encryption: None,
        }
    }

    /// Configure BIFF8 password-to-open encryption for subsequent writes.
    ///
    /// Validation is atomic: an invalid password or profile leaves the current
    /// encryption configuration unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<()> {
        let password = password.into();
        validate_writer_encryption(&password, profile)?;
        self.encryption = Some(WriterEncryption {
            password: Zeroizing::new(password),
            profile,
        });
        Ok(())
    }

    /// Remove password-to-open encryption from subsequent writes.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured password-to-open encryption profile.
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name (max 31 characters)
    ///
    /// # Returns
    ///
    /// * `Result<usize, Error>` - Worksheet index or error
    pub fn add_worksheet(&mut self, name: &str) -> Result<usize> {
        // Validate worksheet name
        if name.is_empty() || name.len() > 31 {
            return Err(Error::InvalidData(
                "Worksheet name must be 1-31 characters".to_string(),
            ));
        }

        // Check for duplicate names
        if self.worksheets.iter().any(|ws| ws.name == name) {
            return Err(Error::InvalidData(format!(
                "Worksheet '{}' already exists",
                name
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        self.synchronize_workbook_window_selection();
        Ok(index)
    }

    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> Result<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::String(value.to_string()), format_id)
    }

    /// Write a number value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Numeric value
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> Result<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    pub fn write_number_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: f64,
        format_id: u16,
    ) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::InvalidData(
                "cell number must be finite for BIFF8 serialization".to_string(),
            ));
        }
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Number(value), format_id)
    }

    /// Write a boolean value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Boolean value
    pub fn write_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> Result<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    pub fn write_boolean_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Boolean(value), format_id)
    }

    /// Write a formula to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - Formula string (without leading '=')
    ///
    /// The supported BIFF8 formula subset includes constants, cell/range
    /// references, arithmetic/comparison operators, and built-in functions
    /// recognized by [`FormulaTokenizer`](crate::writer::FormulaTokenizer).
    pub fn write_formula(&mut self, sheet: usize, row: u32, col: u16, formula: &str) -> Result<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(
            sheet,
            pos,
            CellValue::Formula(formula.to_string()),
            format_id,
        )
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This is a thin wrapper around the internal `FormattingManager`
    /// and mirrors Apache POI's `HSSFDataFormat.getFormat` API. The
    /// returned index can be stored in `ExtendedFormat.format_index`
    /// to apply number formats to cells.
    pub fn register_number_format(&mut self, pattern: &str) -> u16 {
        self.fmt.register_number_format(pattern)
    }

    /// Register a reusable cell style defined by `CellStyle`.
    ///
    /// The returned identifier can be passed to the `write_*_with_format`
    /// methods to apply this style to individual cells.
    pub fn add_cell_style(&mut self, style: CellStyle) -> u16 {
        self.fmt.register_cell_style(style)
    }

    pub fn add_cell_format(&mut self, format: ExtendedFormat) -> u16 {
        self.fmt.add_format(format)
    }

    /// Installs a complete custom table-style family.
    ///
    /// Validation happens before assignment, so an error leaves the current
    /// writer configuration unchanged.
    pub fn set_custom_table_styles(&mut self, styles: CustomTableStyles) -> Result<()> {
        styles.validate(&self.fmt)?;
        self.custom_table_styles = Some(styles);
        Ok(())
    }

    /// Removes caller-defined table styles and restores the default write path.
    pub fn clear_custom_table_styles(&mut self) {
        self.custom_table_styles = None;
    }

    /// Adds a legacy BIFF8 worksheet table and writes its header captions.
    pub fn add_list_object(&mut self, sheet: usize, table: ListObject) -> Result<()> {
        table.validate()?;
        let style = table.style().ok_or_else(|| {
            Error::InvalidData("validated table is missing its style".to_string())
        })?;
        validate_list_object_style(style.name(), self.custom_table_styles.as_ref())?;
        if self
            .worksheets
            .iter()
            .flat_map(|worksheet| &worksheet.list_objects)
            .any(|existing| {
                existing.id() == table.id() || existing.name().eq_ignore_ascii_case(table.name())
            })
        {
            return Err(Error::InvalidData(
                "table identifier or name collides within the workbook".to_string(),
            ));
        }
        if self
            .defined_names
            .iter()
            .any(|name| name.name.eq_ignore_ascii_case(table.name()))
            || self
                .defined_name_records
                .iter()
                .any(|(name, _)| name.name.eq_ignore_ascii_case(table.name()))
        {
            return Err(Error::InvalidData(
                "table name collides with a workbook defined name".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet
            .list_objects
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
        let mut header_cells = Vec::new();
        for (offset, column) in table
            .columns()
            .iter()
            .enumerate()
            .filter(|_| table.has_header_row())
        {
            let key = (
                u32::from(table.range().first_row()),
                table.range().first_column() + offset as u16,
            );
            if let Some(cell) = worksheet.cells.get(&key)
                && !matches!(&cell.value, CellValue::String(value) if value == column.name())
            {
                return Err(Error::InvalidData(
                    "table header collides with a different cell value".to_string(),
                ));
            } else if !worksheet.cells.contains_key(&key) {
                header_cells.push(WritableCell::new(
                    CellPos::try_new(key.0, key.1)?,
                    CellValue::String(column.name().to_string()),
                    0,
                    None,
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.include_list_object_range(table.range());
        for cell in header_cells {
            worksheet.add_cell(cell);
        }
        worksheet.list_objects.push(table);
        Ok(())
    }

    pub fn clear_list_objects(&mut self, sheet: usize) -> Result<()> {
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .list_objects
            .clear();
        Ok(())
    }

    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 characters (Lbl.cch is a byte).
    fn validate_defined_name(name: &str) -> Result<()> {
        if name.is_empty() {
            return Err(Error::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        let char_count = name.chars().count();
        if char_count > u8::MAX as usize {
            return Err(Error::InvalidData(
                "Defined name must be at most 255 characters".to_string(),
            ));
        }

        Ok(())
    }

    fn hash_password(password: &str) -> u16 {
        let bytes = password.as_bytes();
        if bytes.is_empty() {
            return 0;
        }

        let mut hash: u16 = 0;
        for &b in bytes.iter().rev() {
            let high_bit = (hash >> 14) & 0x0001;
            hash = ((hash << 1) & 0x7FFF) | high_bit;
            hash ^= b as u16;
        }

        let high_bit = (hash >> 14) & 0x0001;
        hash = ((hash << 1) & 0x7FFF) | high_bit;
        hash ^= bytes.len() as u16;
        hash ^= 0xCE4B;
        hash
    }

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(Error::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        worksheet.add_hyperlink(Hyperlink {
            first_row: row,
            last_row: row,
            first_col: col,
            last_col: col,
            url: url.to_string(),
        });

        Ok(())
    }

    /// Add a canonical, macro-inert BIFF8 comment to a cell.
    pub fn add_comment(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
    ) -> Result<()> {
        self.add_comment_with_options(
            sheet,
            row,
            col,
            author,
            text,
            CommentWriteOptions::default(),
        )
    }

    /// Add a canonical BIFF8 comment with explicit visibility, anchor, rich runs, and GUID options.
    pub fn add_comment_with_options(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
        options: CommentWriteOptions,
    ) -> Result<()> {
        let (row, column) = comment::validate_comment(row, col, author, text, &options)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.comments.len() >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 canonical comment shapes".to_string(),
            ));
        }
        if worksheet
            .comments
            .iter()
            .any(|comment| comment.row == row && comment.column == column)
        {
            return Err(Error::InvalidData(
                "a cell cannot contain more than one comment".to_string(),
            ));
        }
        if let Some(guid) = options.guid
            && worksheet
                .comments
                .iter()
                .any(|comment| comment.options.guid == Some(guid))
        {
            return Err(Error::InvalidData(
                "comment GUID override is duplicated on the worksheet".to_string(),
            ));
        }
        let comment = comment::WritableComment::try_new(row, column, author, text, options)?;
        worksheet.add_comment(comment)
    }

    /// Add a validated, macro-inert primitive shape and return its worksheet OBJ identifier.
    pub fn add_shape(&mut self, sheet: usize, mut shape: ShapeWrite) -> Result<u16> {
        shape.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let reserved = collect_reserved_object_ids(worksheet, 0)?;
        let object_count =
            reserved
                .len()
                .checked_add(worksheet.comments.len())
                .ok_or(Error::Allocation(
                    "computing the worksheet drawing-object count",
                ))?;
        if object_count >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        let object_id = if let Some(requested) = shape.object_id {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
            requested
        } else {
            (1..u16::MAX)
                .find(|candidate| !reserved.contains(candidate))
                .ok_or_else(|| {
                    Error::InvalidData("worksheet object IDs are exhausted".to_string())
                })?
        };
        worksheet
            .shapes
            .try_reserve(1)
            .map_err(|_| Error::Allocation("reserving worksheet shape storage"))?;
        shape.object_id = Some(object_id);
        worksheet.shapes.push(shape);
        Ok(object_id)
    }

    /// Remove a primitive by its assigned OBJ identifier.
    pub fn remove_shape(&mut self, sheet: usize, object_id: u16) -> Result<ShapeWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shapes
            .iter()
            .position(|shape| shape.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shapes.remove(index))
    }

    /// Remove all writable primitive shapes from a worksheet.
    pub fn clear_shapes(&mut self, sheet: usize) -> Result<usize> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let count = worksheet.shapes.len();
        worksheet.shapes.clear();
        Ok(count)
    }

    /// Add a validated shape group and return the group's worksheet OBJ identifier.
    ///
    /// The group consumes one object ID for itself plus one per child; assigned
    /// child identifiers are stored back into the group before it is retained.
    pub fn add_shape_group(&mut self, sheet: usize, mut group: ShapeGroupWrite) -> Result<u16> {
        group.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let group_count = group.object_count()?;
        let mut reserved = collect_reserved_object_ids(worksheet, group_count)?;
        let object_count = reserved
            .len()
            .checked_add(worksheet.comments.len())
            .and_then(|count| count.checked_add(group_count))
            .ok_or(Error::Allocation(
                "computing the worksheet drawing-object count",
            ))?;
        if object_count > 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        for requested in group_object_ids(&group) {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
        }
        for requested in group_object_ids(&group) {
            reserved.insert(requested);
        }
        worksheet
            .shape_groups
            .try_reserve(1)
            .map_err(|_| Error::Allocation("reserving worksheet shape-group storage"))?;
        let group_id = assign_object_id(&mut reserved, group.object_id)?;
        group.object_id = Some(group_id);
        for child in &mut group.children {
            child.object_id = Some(assign_object_id(&mut reserved, child.object_id)?);
        }
        worksheet.shape_groups.push(group);
        Ok(group_id)
    }

    /// Remove a shape group by the group's assigned OBJ identifier.
    pub fn remove_shape_group(&mut self, sheet: usize, object_id: u16) -> Result<ShapeGroupWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shape_groups
            .iter()
            .position(|group| group.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shape_groups.remove(index))
    }

    pub fn set_auto_filter(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        MergedRange::try_new(first_row, last_row, first_col, last_col)?;
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.list_objects.iter().any(|table| {
            first_row <= u32::from(table.range().last_row())
                && u32::from(table.range().first_row()) <= last_row
                && first_col <= table.range().last_column()
                && table.range().first_column() <= last_col
        }) {
            return Err(Error::InvalidData(
                "set_auto_filter: range overlaps a worksheet table".to_string(),
            ));
        }
        let itab = sheet
            .checked_add(1)
            .and_then(|index| u16::try_from(index).ok())
            .ok_or_else(|| {
                Error::InvalidData(
                    "set_auto_filter: sheet index exceeds BIFF8 itab limit".to_string(),
                )
            })?;
        let target_sheet = u16::try_from(sheet).map_err(|_| {
            Error::InvalidData("set_auto_filter: sheet index exceeds BIFF8 limit".to_string())
        })?;

        let start_ref = a1_cell(first_row, first_col);
        let end_ref = a1_cell(last_row, last_col);
        let reference = format!("{start_ref}:{end_ref}");
        let defined_name = InternalDefinedName {
            name: "_FilterDatabase".to_string(),
            reference,
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(target_sheet),
            hidden: true,
            is_function: false,
            is_built_in: true,
            built_in_code: Some(0x0D),
        };

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.auto_filter = Some(AutoFilterRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });
        self.defined_names.retain(|n| {
            !(n.is_built_in && n.built_in_code == Some(0x0D) && n.local_sheet == Some(itab))
        });
        self.defined_names.push(defined_name);

        Ok(())
    }

    /// Add a filter condition to a specific column within the AutoFilter range.
    ///
    /// The AutoFilter range must first be set via [`Self::set_auto_filter`]. The
    /// `column_index` is 0-based relative to the filter range start column.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `column_index` — column within the filter range (0-based relative)
    /// * `join_or` — `true` to join conditions with OR, `false` for AND
    /// * `cond1` — first filter condition
    /// * `cond2` — second filter condition (use `AutoFilterConditionWrite::None` if unused)
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xls::writer::biff::AutoFilterConditionWrite;
    ///
    /// // Filter column 2: value > 100
    /// writer.add_filter_condition(
    ///     sheet_idx, 2, false,
    ///     AutoFilterConditionWrite::Number { operator: 0x04, value: 100.0 },
    ///     AutoFilterConditionWrite::None,
    /// )?;
    /// ```
    pub fn add_filter_condition(
        &mut self,
        sheet: usize,
        column_index: u16,
        join_or: bool,
        cond1: AutoFilterConditionWrite,
        cond2: AutoFilterConditionWrite,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let Some(filter) = worksheet.auto_filter else {
            return Err(Error::InvalidData(
                "add_filter_condition: call set_auto_filter first".to_string(),
            ));
        };
        let width = filter.last_col - filter.first_col + 1;
        if column_index >= width {
            return Err(Error::InvalidCellReference(format!(
                "AutoFilter relative column {column_index} exceeds its {width}-column range"
            )));
        }
        if worksheet
            .auto_filter_columns
            .iter()
            .any(|entry| entry.column_index == column_index)
        {
            return Err(Error::InvalidData(
                "AutoFilter column already has a condition".to_string(),
            ));
        }

        worksheet.add_auto_filter_column(AutoFilterColumnDef {
            column_index,
            join_or,
            condition1: cond1,
            condition2: cond2,
        });

        Ok(())
    }

    /// Set the sort configuration for a worksheet.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `case_sensitive` — whether sorting is case-sensitive
    /// * `sort_by_columns` — `true` for left-to-right sort, `false` for top-to-bottom
    /// * `keys` — up to 3 sort keys as `(column_index, descending)` tuples
    pub fn set_sort(
        &mut self,
        sheet: usize,
        case_sensitive: bool,
        sort_by_columns: bool,
        keys: &[(u16, bool)],
    ) -> Result<()> {
        if keys.is_empty() || keys.len() > 3 {
            return Err(Error::InvalidData(
                "set_sort: must provide 1..3 sort keys".to_string(),
            ));
        }
        if keys.iter().any(|(col, _)| *col > u16::from(u8::MAX)) {
            return Err(Error::InvalidCellReference(
                "sort key column is outside the BIFF8 grid".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.set_sort_config(SortConfig {
            case_sensitive,
            sort_by_columns,
            keys: keys.to_vec(),
        });

        Ok(())
    }

    /// Replace the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Unlike [`set_sort`](Self::set_sort), this preserves the complete
    /// `SortData` model, including an explicit range, more than three keys,
    /// custom lists, differential-format colors, and icon sets. The previous
    /// owned configuration is returned.
    pub fn put_sort(
        &mut self,
        sheet: usize,
        sort: crate::writer::sort::Config,
    ) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if let crate::writer::sort::Parent::Table { id } = sort.parent()
            && !worksheet
                .list_objects
                .iter()
                .any(|table| table.id().value() == id)
        {
            return Err(Error::InvalidData(
                "table SortData references an unknown ListObject identifier".to_string(),
            ));
        }
        Ok(worksheet.put_sort(sort))
    }

    /// Remove and return the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Removing an absent configuration succeeds and returns `None`.
    pub fn remove_sort(&mut self, sheet: usize) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        Ok(worksheet.remove_sort())
    }

    /// Add a pivot table definition to a worksheet.
    ///
    /// This writes the SX* record family (SXVS, SXVIEW, SXVD, SXVI, SXDI,
    /// SXPI) to the worksheet stream. The pivot table must be fully
    /// configured before calling this method.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `config` — pivot table configuration (see [`PivotTableConfig`])
    pub fn add_pivot_table(&mut self, sheet: usize, config: PivotTableConfig) -> Result<()> {
        validate_pivot_table_config(&config)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Generate pivot output cells BEFORE consuming config.fields / config.data_items.
        // Excel validates that DIMENSIONS and cell content are consistent with the
        // pivot table definition; missing cells cause a "corrupt file" repair dialog.
        Self::generate_pivot_output_cells(worksheet, &config)?;
        self.fmt.enable_pivot_xfs();

        let fields: Vec<WritablePivotField> = config
            .fields
            .into_iter()
            .map(|f| {
                let mut items: Vec<WritablePivotItem> = f
                    .items
                    .into_iter()
                    .map(|i| WritablePivotItem {
                        item_type: i.item_type,
                        flags: i.flags,
                        cache_index: i.cache_index,
                        name: i.name,
                    })
                    .collect();

                // Sort data items (item_type=0x0000) alphabetically by their
                // cache label to match Excel's default SXVI ordering.  Non-data
                // items (subtotals etc.) stay at the end.
                let data_end = items
                    .iter()
                    .position(|i| i.item_type != 0x0000)
                    .unwrap_or(items.len());
                items[..data_end].sort_unstable_by(|a, b| {
                    let al = f
                        .cache_items
                        .get(a.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    let bl = f
                        .cache_items
                        .get(b.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    al.cmp(&bl)
                });

                WritablePivotField {
                    axis: f.axis,
                    subtotal_count: f.subtotal_count,
                    subtotal_flags: f.subtotal_flags,
                    items,
                    name: f.name,
                    cache_name: f.cache_name,
                    cache_items: f.cache_items,
                    is_numeric: f.is_numeric,
                    grouping: f.grouping,
                }
            })
            .collect();

        let data_items: Vec<WritablePivotDataItem> = config
            .data_items
            .into_iter()
            .map(|d| WritablePivotDataItem {
                source_field_index: d.source_field_index,
                function: d.function,
                display_format: d.display_format,
                base_field_index: d.base_field_index,
                base_item_index: d.base_item_index,
                num_format_index: d.num_format_index,
                name: d.name,
            })
            .collect();

        worksheet.add_pivot_table(WritablePivotTable {
            name: config.name,
            source_type: config.source_type,
            source_sheet_name: config.source_sheet_name,
            source_first_row: config.source_first_row,
            source_last_row: config.source_last_row,
            source_first_col: config.source_first_col,
            source_last_col: config.source_last_col,
            first_row: config.first_row,
            last_row: config.last_row,
            first_col: config.first_col,
            last_col: config.last_col,
            first_header_row: config.first_header_row,
            first_data_row: config.first_data_row,
            first_data_col: config.first_data_col,
            data_field_name: config.data_field_name,
            data_axis: config.data_axis,
            data_position: config.data_position,
            fields,
            data_items,
            page_entries: config.page_entries,
            source_data: config.source_data,
        });

        Ok(())
    }

    /// Generate the cell data that Excel expects in the SXVIEW output area.
    ///
    /// The layout (for a single row-field, single col-field, single page-field,
    /// single data-field configuration) is:
    ///
    /// ```text
    /// (first_row-2, 0)       : page field name    (first_row-2, 1)       : "(All)"
    /// (first_row,   0)       : data item name      (first_row, first_data_col): "Column Labels"
    /// (first_header_row, 0)  : "Row Labels"        (fhr, fdc+j)           : col item names …
    /// (first_data_row+i, 0)  : row item name       (fdr+i, fdc+j)         : aggregated value
    /// (last_row, 0)          : "Grand Total"        (lr, fdc+j)            : column totals
    /// ```
    fn generate_pivot_output_cells(
        ws: &mut WritableWorksheet,
        cfg: &PivotTableConfig,
    ) -> Result<()> {
        // Identify fields per axis.
        let row_field = cfg.fields.iter().find(|f| f.axis == 0x0001);
        let col_field = cfg.fields.iter().find(|f| f.axis == 0x0002);
        let page_field = cfg.fields.iter().find(|f| f.axis == 0x0004);

        let data_item = cfg.data_items.first();

        // Helper: find the field index for a given field by cache_name.
        let field_idx_of =
            |name: &str| -> Option<usize> { cfg.fields.iter().position(|f| f.cache_name == name) };

        // Collect row/col item labels from cache_items, sorted alphabetically
        // to match Excel's default SXVI ordering.  Also build a mapping from
        // cache_index → sorted position so the aggregation grid uses the same
        // order as the output rows/columns.
        let (row_items, row_cache_to_sorted) = Self::sorted_cache_items(row_field);
        let (col_items, col_cache_to_sorted) = Self::sorted_cache_items(col_field);

        let fr = cfg.first_row;
        let fhr = cfg.first_header_row;
        let fdr = cfg.first_data_row;
        let fdc = cfg.first_data_col;
        let lr = cfg.last_row;
        let lc = cfg.last_col;
        let fc = cfg.first_col;

        let offset = |base: u16, amount: usize| -> Result<u16> {
            let amount = u16::try_from(amount).map_err(|_| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })?;
            base.checked_add(amount).ok_or_else(|| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })
        };
        let mut staged = Vec::new();
        let mut add = |row: u16,
                       col: u16,
                       value: CellValue,
                       pivot_xf_role: Option<PivotCellXfRole>|
         -> Result<()> {
            staged.push(WritableCell::new(
                CellPos::try_new(u32::from(row), col)?,
                value,
                0,
                pivot_xf_role,
            ));
            Ok(())
        };

        // --- Page field area (above SXVIEW range) ---
        if let Some(pf) = page_field {
            let page_row = fr.saturating_sub(2);
            add(
                page_row,
                0,
                CellValue::String(pf.cache_name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
            add(
                page_row,
                1,
                CellValue::String("(All)".to_string()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }

        // --- Row at first_row: data item name + "Column Labels" ---
        if let Some(di) = data_item {
            add(
                fr,
                fc,
                CellValue::String(di.name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }
        if col_field.is_some() {
            add(
                fr,
                fdc,
                CellValue::String("Column Labels".to_string()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }

        // --- Row at first_header_row: "Row Labels" + column item names + "Grand Total" ---
        add(
            fhr,
            fc,
            CellValue::String("Row Labels".to_string()),
            Some(PivotCellXfRole::HeaderAccent),
        )?;
        for (j, ci) in col_items.iter().enumerate() {
            add(
                fhr,
                offset(fdc, j)?,
                CellValue::String(ci.clone()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }
        add(
            fhr,
            lc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::HeaderPlain),
        )?;

        // --- Compute aggregated values from source_data ---
        let row_fi = row_field.and_then(|f| field_idx_of(&f.cache_name));
        let col_fi = col_field.and_then(|f| field_idx_of(&f.cache_name));
        let data_fi = data_item.map(|di| di.source_field_index as usize);

        let nr = row_items.len();
        let nc = col_items.len();
        let mut grid = vec![vec![0.0f64; nc]; nr];
        let mut row_totals = vec![0.0f64; nr];
        let mut col_totals = vec![0.0f64; nc];
        let mut grand_total = 0.0f64;

        for row_data in &cfg.source_data {
            // Map cache indices through the sorted permutation so that
            // grid positions match the alphabetically-sorted output.
            let ri = row_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    row_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let ci = col_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    col_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let val = data_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::Number(v)) => Some(*v),
                _ => None,
            });

            if let (Some(ri), Some(ci), Some(val)) = (ri, ci, val)
                && ri < nr
                && ci < nc
            {
                grid[ri][ci] += val;
                row_totals[ri] += val;
                col_totals[ci] += val;
                grand_total += val;
            }
        }

        // --- Data rows ---
        for (i, (ri_name, row_total)) in row_items.iter().zip(row_totals.iter()).enumerate() {
            let r = offset(fdr, i)?;
            add(
                r,
                fc,
                CellValue::String(ri_name.clone()),
                Some(PivotCellXfRole::RowLabel),
            )?;
            for (j, cell_val) in grid[i].iter().enumerate() {
                add(
                    r,
                    offset(fdc, j)?,
                    CellValue::Number(*cell_val),
                    Some(PivotCellXfRole::Value),
                )?;
            }
            add(
                r,
                lc,
                CellValue::Number(*row_total),
                Some(PivotCellXfRole::Value),
            )?;
        }

        // --- Grand total row ---
        add(
            lr,
            fc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::RowLabel),
        )?;
        for (j, col_total) in col_totals.iter().enumerate() {
            add(
                lr,
                offset(fdc, j)?,
                CellValue::Number(*col_total),
                Some(PivotCellXfRole::Value),
            )?;
        }
        add(
            lr,
            lc,
            CellValue::Number(grand_total),
            Some(PivotCellXfRole::Value),
        )?;
        for cell in staged {
            ws.add_cell(cell);
        }
        Ok(())
    }

    /// Sort a field's cache items alphabetically and return the sorted labels
    /// plus a mapping from original cache index to sorted position.
    ///
    /// Returns `(sorted_labels, cache_to_sorted)` where `cache_to_sorted[i]`
    /// gives the position of original cache item `i` in the sorted output.
    fn sorted_cache_items(field: Option<&PivotFieldConfig>) -> (Vec<String>, Vec<usize>) {
        let Some(f) = field else {
            return (Vec::new(), Vec::new());
        };

        // Build (original_index, label) pairs and sort by label.
        let mut indexed: Vec<(usize, String)> = f
            .cache_items
            .iter()
            .enumerate()
            .map(|(i, item)| (i, item.display_text()))
            .collect();
        indexed.sort_unstable_by(|a, b| a.1.cmp(&b.1));

        let sorted_labels: Vec<String> = indexed.iter().map(|(_, value)| value.clone()).collect();

        // cache_to_sorted[original_cache_idx] = position in sorted output
        let mut cache_to_sorted = vec![0usize; f.cache_items.len()];
        for (sorted_pos, (orig_idx, _)) in indexed.iter().enumerate() {
            cache_to_sorted[*orig_idx] = sorted_pos;
        }

        (sorted_labels, cache_to_sorted)
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers. More complex formulas will be rejected
    /// at serialization time to avoid emitting invalid BIFF payloads.
    pub fn define_name(&mut self, name: &str, reference: &str) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name: workbook must have at least one worksheet".to_string(),
            ));
        }

        // For now, workbook-scoped names that refer to cell ranges are
        // anchored to the first worksheet. Users who need explicit
        // sheet scoping can use `define_name_local`.
        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet: usize) -> Result<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            Error::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a workbook-scoped named range with a user-visible comment.
    pub fn define_name_with_comment(
        &mut self,
        name: &str,
        reference: &str,
        comment: &str,
    ) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name_with_comment: workbook must have at least one worksheet".to_string(),
            ));
        }

        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Remove all defined names with the given name.
    ///
    /// Returns `true` if at least one name was removed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.defined_names.len();
        self.defined_names.retain(|n| n.name != name);
        self.defined_names.len() < initial_len
    }

    /// Get all defined names in this workbook.
    pub fn named_ranges(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Add complete inert BIFF8 defined-name metadata.
    pub fn add_defined_name_record(&mut self, options: DefinedNameRecordOptions) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records
            .push((options, Default::default()));
        Ok(index)
    }

    /// Add complete inert `Lbl` metadata and its ordered BIFF8 future records.
    pub fn add_defined_name_record_with_future_records(
        &mut self,
        options: DefinedNameRecordOptions,
        future: crate::DefinedNameFutureRecords,
    ) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        named_range::validate_future_records(&future, options.serialized_name())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records.push((options, future));
        Ok(index)
    }

    /// Set the width of a column in character units.
    ///
    /// The column index is 0-based (0 = column A), matching the rest of the
    /// XLS writer API. The width is specified in the same units as Excel's
    /// UI, i.e. the number of characters of the "0" glyph in the default
    /// font. Internally this is converted to BIFF8 units of 1/256 characters
    /// for the COLINFO record.
    pub fn set_column_width(&mut self, sheet: usize, col: u16, width_chars: f64) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "set_column_width: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        if !(width_chars.is_finite()) || width_chars <= 0.0 {
            return Err(Error::InvalidData(
                "set_column_width: width must be a positive finite value".to_string(),
            ));
        }

        let max_units = 255u32 * 256u32; // Excel maximum column width
        let width_units_f = (width_chars * 256.0).round();
        if width_units_f <= 0.0 || width_units_f > max_units as f64 {
            return Err(Error::InvalidData(
                "set_column_width: width exceeds Excel's maximum (255 characters)".to_string(),
            ));
        }

        let width_units = width_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_column_width(col, width_units);
        Ok(())
    }

    /// Hide a column.
    pub fn hide_column(&mut self, sheet: usize, col: u16) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "hide_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_column(col);
        Ok(())
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, sheet: usize, col: u16) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "show_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_column(col);
        Ok(())
    }

    pub fn merge_cells(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        let range = MergedRange::try_new(first_row, last_row, first_col, last_col)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if worksheet
            .merged_ranges
            .iter()
            .any(|existing| range.overlaps(*existing))
        {
            return Err(Error::InvalidData("merged-cell ranges overlap".to_string()));
        }

        worksheet.add_merged_range(range);

        Ok(())
    }

    /// Configure freeze panes for the specified worksheet.
    ///
    /// Row and column indices are 0-based and represent the number of
    /// rows/columns at the top/left that remain frozen.
    pub fn freeze_panes(&mut self, sheet: usize, freeze_rows: u32, freeze_cols: u16) -> Result<()> {
        if freeze_rows == 0 && freeze_cols == 0 {
            let worksheet = self
                .worksheets
                .get_mut(sheet)
                .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
            worksheet.clear_freeze_panes();
            return Ok(());
        }
        let freeze_rows = u16::try_from(freeze_rows).map_err(|_| {
            Error::InvalidCellReference("freeze-panes row is outside the BIFF8 grid".to_string())
        })?;
        let freeze_cols = u8::try_from(freeze_cols).map_err(|_| {
            Error::InvalidCellReference("freeze-panes column is outside the BIFF8 grid".to_string())
        })?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.set_freeze_panes(freeze_rows, freeze_cols)
    }

    /// Remove any freeze panes from the specified worksheet.
    pub fn unfreeze_panes(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.clear_freeze_panes();
        Ok(())
    }

    /// Replace the worksheet's checked BIFF8 zoom scale.
    pub fn put_scale(
        &mut self,
        sheet: usize,
        scale: Option<crate::writer::view::Scale>,
    ) -> Result<Option<crate::writer::view::Scale>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(worksheet.put_scale(scale))
    }

    /// Replace a worksheet view after validating the complete prospective state.
    pub fn put_view(
        &mut self,
        sheet: usize,
        view: crate::writer::view::View,
    ) -> Result<crate::writer::view::View> {
        view.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(std::mem::replace(&mut worksheet.view, view))
    }

    /// Replace a worksheet pane and its selections as one failure-atomic edit.
    pub fn put_pane(
        &mut self,
        sheet: usize,
        pane: crate::writer::view::Pane,
        selections: Vec<crate::writer::view::Selection>,
    ) -> Result<(
        Option<crate::writer::view::Pane>,
        Vec<crate::writer::view::Selection>,
    )> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.view.put_pane(pane, selections)
    }

    /// Set the height of a row in points.
    ///
    /// The row index is 0-based (0 = first row), and the height is specified
    /// in typographic points. Internally this is converted to twips
    /// (1/20th of a point) for the BIFF8 ROW record.
    pub fn set_row_height(&mut self, sheet: usize, row: u32, height_points: f64) -> Result<()> {
        if !(height_points.is_finite()) || height_points <= 0.0 {
            return Err(Error::InvalidData(
                "set_row_height: height must be a positive finite value".to_string(),
            ));
        }

        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "set_row_height: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let height_units_f = (height_points * 20.0).round();
        if height_units_f <= 0.0 || height_units_f > u16::MAX as f64 {
            return Err(Error::InvalidData(
                "set_row_height: height exceeds BIFF8 row height limit".to_string(),
            ));
        }

        let height_units = height_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_row_height(row, height_units);
        Ok(())
    }

    /// Hide a row.
    pub fn hide_row(&mut self, sheet: usize, row: u32) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "hide_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_row(row);
        Ok(())
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, sheet: usize, row: u32) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "show_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_row(row);
        Ok(())
    }

    /// Add a data validation rule to the specified worksheet.
    pub fn add_data_validation(&mut self, sheet: usize, validation: DataValidation) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let range = validation.range;
        worksheet.add_data_validation(
            validation,
            payload,
            vec![range],
            DataValidationOptions::default(),
        );

        Ok(())
    }

    /// Add a validation with typed flags and additional target ranges.
    pub fn add_data_validation_with_options(
        &mut self,
        sheet: usize,
        validation: DataValidation,
        additional_ranges: &[DataValidationRange],
        options: DataValidationOptions,
    ) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;
        let range_count = 1usize
            .checked_add(additional_ranges.len())
            .ok_or_else(|| Error::InvalidData("DV range count overflows".to_string()))?;
        if range_count > 432 {
            return Err(Error::InvalidData("DV range count exceeds 432".to_string()));
        }
        let mut ranges = Vec::with_capacity(range_count);
        ranges.push(validation.range);
        ranges.extend_from_slice(additional_ranges);
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.add_data_validation(validation, payload, ranges, options);
        Ok(())
    }

    /// Configure worksheet-level DVAL window/dropdown metadata.
    pub fn set_data_validation_table_options(
        &mut self,
        sheet: usize,
        options: DataValidationTableOptions,
    ) -> Result<()> {
        if options.x_left > 65_535
            || options.y_top > 65_535
            || matches!(options.dropdown_object_id, Some(0))
            || options.dropdown_object_id.is_some_and(|id| id > 32_767)
        {
            return Err(Error::InvalidData(
                "DVAL metadata is out of range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.data_validation_table_options = Some(options);
        Ok(())
    }

    pub fn add_conditional_format(&mut self, sheet: usize, cf: ConditionalFormat) -> Result<()> {
        if cf.first_row > cf.last_row
            || cf.first_col > cf.last_col
            || cf.last_row > 65_535
            || cf.last_col > 255
        {
            return Err(Error::InvalidData(
                "add_conditional_format: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_conditional_format(cf);

        Ok(())
    }

    /// Add one legacy `CONDFMT` collection with ordered ranges and one to three ordered rules.
    pub fn add_conditional_format_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormatGroup,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > 3 {
            return Err(Error::InvalidData(
                "legacy conditional-format rule count must be 1..=3".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        for rule in &group.rules {
            rule.format_type.to_biff_payload()?;
        }
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .add_conditional_format_group(group);
        Ok(())
    }

    /// Add one future `CondFmt12` collection. Formula tokens and visual
    /// payloads are serialized exactly and are never evaluated.
    pub fn add_conditional_format12_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormat12Group,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "future conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "future conditional-format rule count must be 1..=65535".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "future conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.conditional_formats.len() + worksheet.conditional_formats12.len() >= 32_768 {
            return Err(Error::InvalidData(
                "conditional-format group count exceeds the 15-bit BIFF identifier space"
                    .to_string(),
            ));
        }
        let mut priorities = worksheet
            .conditional_formats12
            .iter()
            .flat_map(|existing| existing.rules.iter().map(|rule| rule.priority))
            .collect::<HashSet<_>>();
        for rule in &group.rules {
            if rule.priority == 0 || !priorities.insert(rule.priority) {
                return Err(Error::InvalidData(
                    "future conditional-format priorities must be nonzero and unique per sheet"
                        .to_string(),
                ));
            }
            if !matches!(rule.template, 0..=5 | 7..=12 | 15..=27 | 29 | 30) {
                return Err(Error::InvalidData(
                    "future conditional-format template is invalid".to_string(),
                ));
            }
            let between = matches!(
                rule.format_type,
                ConditionalFormat12Type::CellValue {
                    operator: ConditionalFormatOperator::Between
                        | ConditionalFormatOperator::NotBetween,
                    ..
                }
            );
            if let ConditionalFormat12Type::CellValue { formula2, .. } = &rule.format_type
                && between != formula2.is_some()
            {
                return Err(Error::InvalidData(
                        "between/not-between CF12 rules require two formulas; other comparisons require one".to_string(),
                    ));
            }
            let visual = matches!(
                rule.format_type,
                ConditionalFormat12Type::ColorScale { .. }
                    | ConditionalFormat12Type::DataBar { .. }
                    | ConditionalFormat12Type::IconSet { .. }
            );
            if visual && (rule.stop_if_true || rule.differential_format != [0, 0, 0, 0, 0, 0]) {
                return Err(Error::InvalidData(
                    "visual CF12 rules require an empty DXFN12 and cannot stop-if-true".to_string(),
                ));
            }
            let (condition_type, comparison, formula1, formula2, active_formula, payload) =
                rule.format_type.biff_parts();
            let config = crate::writer::biff::Cf12Config {
                condition_type,
                comparison,
                differential_format: &rule.differential_format,
                formula1,
                formula2,
                active_formula,
                stop_if_true: rule.stop_if_true,
                priority: rule.priority,
                template: rule.template,
                template_parameters: rule.template_parameters,
                rule_payload: payload,
            };
            crate::writer::biff::write_cf12(&mut Vec::new(), &config)?;
        }
        worksheet.add_conditional_format12_group(group);
        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        pos: CellPos,
        value: CellValue,
        format_id: u16,
    ) -> Result<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(Error::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_cell(WritableCell::new(pos, value, format_id, None));

        Ok(())
    }

    /// Set the date system (1900 vs 1904)
    ///
    /// # Arguments
    ///
    /// * `use_1904` - True to use 1904 date system (Mac), false for 1900 (Windows, default)
    pub fn set_1904_dates(&mut self, use_1904: bool) {
        self.use_1904_dates = use_1904;
    }

    pub fn set_workbook_environment(&mut self, options: WorkbookEnvironmentOptions) -> Result<()> {
        if options.refresh_external_data_on_load && !options.template {
            return Err(Error::InvalidData(
                "RefreshAll requires a template workbook".to_string(),
            ));
        }
        if (options.envelope_visible || options.envelope_initialized) && !options.has_envelope {
            return Err(Error::InvalidData(
                "envelope state flags require has_envelope".to_string(),
            ));
        }
        if !(1..=981).contains(&options.default_country_code)
            || !(1..=981).contains(&options.current_country_code)
        {
            return Err(Error::InvalidData(
                "country codes must be 1..=981".to_string(),
            ));
        }
        self.environment_options = options;
        Ok(())
    }

    /// Set the workbook extension flags emitted as a `BookExt` record
    /// (MS-XLS 2.4.23); `None` emits no record.
    pub fn set_book_ext(&mut self, book_ext: Option<crate::BookExt>) {
        self.book_ext = book_ext;
    }

    /// Append a real-time data (RTD) topic emitted as a `RealTimeData`
    /// record (MS-XLS 2.4.214) in the workbook globals.
    ///
    /// When the topic shares a prefix with the previously added topic, set
    /// [`crate::RealTimeData::common_prefix_len`] and store only the
    /// trailing sub-strings in `topic_segments`, matching the on-disk prefix
    /// compression.
    pub fn add_real_time_data(&mut self, topic: crate::RealTimeData) -> Result<()> {
        if let Some(cell) = topic
            .cells
            .iter()
            .find(|cell| usize::from(cell.sheet_index) >= self.worksheets.len())
        {
            return Err(Error::WorksheetNotFound(format!(
                "Sheet {}",
                cell.sheet_index
            )));
        }
        self.real_time_data.push(topic);
        Ok(())
    }

    /// Append a Web page published from the workbook globals, emitted as a
    /// `WebPub` record (MS-XLS 2.4.344).
    pub fn add_web_publication(&mut self, publication: crate::WebPub) -> Result<()> {
        publication.validate_for_write()?;
        self.web_publications.push(publication);
        Ok(())
    }

    /// Append a Web page published from a worksheet, emitted as a `WebPub`
    /// record (MS-XLS 2.4.344) in that sheet's substream.
    pub fn add_sheet_web_publication(
        &mut self,
        sheet: usize,
        publication: crate::WebPub,
    ) -> Result<()> {
        publication.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.web_publications.push(publication);
        Ok(())
    }

    /// Set a worksheet's default phonetic format and visible phonetic ranges
    /// (PHONETICINFO, MS-XLS 2.4.192); `None` emits no record.
    pub fn set_phonetic_info(
        &mut self,
        sheet: usize,
        phonetic_info: Option<crate::PhoneticInfo>,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.phonetic_info = phonetic_info;
        Ok(())
    }

    /// Set the document theme emitted as a `Theme` record (MS-XLS 2.4.326);
    /// `None` emits no record. Large custom theme contents are chunked into
    /// ContinueFrt12 records automatically.
    pub fn set_theme(&mut self, theme: Option<crate::Theme>) {
        self.theme = theme;
    }

    /// Set the MDX (OLAP cube) metadata emitted as the workbook globals
    /// `METADATA` production (MS-XLS 2.1); `None` emits no records. Oversized
    /// record payloads are chunked into ContinueFrt12 records automatically.
    pub fn set_mdx_metadata(&mut self, metadata: Option<crate::MdxMetadata>) {
        self.mdx_metadata = metadata;
    }

    /// Set the `XFExt` formatting property extensions (MS-XLS 2.4.355)
    /// emitted after the XF table. Each extension's `xf_index` is validated
    /// against the written XF record count when the workbook is saved.
    pub fn set_xf_extensions(&mut self, xf_extensions: Vec<crate::XfExt>) {
        self.xf_extensions = xf_extensions;
    }

    /// Set the `StyleExt` cell-style extensions (MS-XLS 2.4.270) emitted
    /// after the built-in STYLE records.
    pub fn set_style_extensions(&mut self, style_extensions: Vec<crate::StyleExt>) {
        self.style_extensions = style_extensions;
    }

    pub fn set_workbook_window(&mut self, options: WorkbookWindowOptions) -> Result<()> {
        options.validate_intrinsic()?;
        self.workbook_window_options = options;
        self.synchronize_workbook_window_selection();
        Ok(())
    }

    fn synchronize_workbook_window_selection(&mut self) {
        let sheet_count = self.worksheets.len();
        let selected_count = usize::from(self.workbook_window_options.selected_sheet_count);
        let active = usize::from(self.workbook_window_options.active_sheet_index);
        if selected_count == 0 || selected_count > sheet_count || active >= sheet_count {
            return;
        }
        let first_selected = active.min(sheet_count - selected_count);
        let selected_range = first_selected..first_selected + selected_count;
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            worksheet.view.select(selected_range.contains(&index));
        }
    }

    pub fn set_function_groups(&mut self, options: FunctionGroupOptions) -> Result<()> {
        options.validate()?;
        self.function_group_options = options;
        Ok(())
    }

    pub fn add_external_workbook_link(
        &mut self,
        options: ExternalWorkbookOptions,
    ) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "external supporting-book count exceeds resource bound".to_string(),
            ));
        }
        let index = self.external_workbooks.len();
        self.external_workbooks.push(options);
        self.external_names.push(Vec::new());
        Ok(index)
    }

    fn external_name_count(&self) -> usize {
        self.external_names.iter().map(Vec::len).sum::<usize>()
            + self.add_in_functions.len()
            + self
                .dde_or_ole_links
                .iter()
                .map(|link| link.items.len())
                .sum::<usize>()
    }

    pub fn add_external_defined_name(
        &mut self,
        external_workbook: usize,
        options: ExternalDefinedNameOptions,
    ) -> Result<usize> {
        let book = self
            .external_workbooks
            .get(external_workbook)
            .ok_or_else(|| {
                Error::InvalidData("external workbook index is out of range".to_string())
            })?;
        options.validate(book.sheets.len())?;
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let names = &mut self.external_names[external_workbook];
        let index = names.len();
        names.push(options);
        Ok(index)
    }

    pub fn add_add_in_function(&mut self, options: AddInFunctionOptions) -> Result<usize> {
        options.validate()?;
        if self.add_in_functions.is_empty()
            && self.external_workbooks.len() + self.dde_or_ole_links.len() >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "add-in function count exceeds resource bound".to_string(),
            ));
        }
        let index = self.add_in_functions.len();
        self.add_in_functions.push(options);
        Ok(index)
    }

    pub fn add_dde_or_ole_link(&mut self, options: DdeOrOleLinkOptions) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self
            .external_name_count()
            .checked_add(options.items.len())
            .is_none_or(|count| count > 4096)
        {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let index = self.dde_or_ole_links.len();
        self.dde_or_ole_links.push(options);
        Ok(index)
    }

    pub fn set_calculation_settings(&mut self, settings: CalculationSettings) -> Result<()> {
        if !(1..=32_767).contains(&settings.maximum_iterations) {
            return Err(Error::InvalidData(
                "maximum calculation iterations must be 1..=32767".to_string(),
            ));
        }
        if !settings.iteration_delta.is_finite() || settings.iteration_delta < 0.0 {
            return Err(Error::InvalidData(
                "calculation iteration delta must be finite and non-negative".to_string(),
            ));
        }
        self.calculation_settings = settings;
        Ok(())
    }

    pub fn set_recalculation_pending(&mut self, sheet: usize, pending: bool) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.formulas_pending_recalculation = pending;
        Ok(())
    }

    pub fn set_scenario_manager(
        &mut self,
        sheet: usize,
        manager: crate::ScenarioManager,
    ) -> Result<()> {
        manager.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = Some(manager);
        Ok(())
    }

    pub fn clear_scenario_manager(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = None;
        Ok(())
    }

    /// Configure an inert BIFF8 data-consolidation directory for a worksheet.
    pub fn set_consolidation(
        &mut self,
        sheet: usize,
        consolidation: crate::Consolidation,
    ) -> Result<()> {
        consolidation.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = Some(consolidation);
        Ok(())
    }

    pub fn clear_consolidation(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = None;
        Ok(())
    }

    /// Configure a complete inert VBA project with safe default limits.
    pub fn set_vba(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
    ) -> Result<()> {
        self.set_vba_with(workbook_code_name, project, &litchi_vba::Limits::default())
    }

    /// Configure a complete inert VBA project using explicit resource limits.
    ///
    /// Module source is serialized but never compiled, interpreted, or run.
    /// Validation and serialization finish before the writer state is changed.
    pub fn set_vba_with(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        let payload = project.finish(limits)?;
        self.put_vba(workbook_code_name, payload)
    }

    /// Configure an already validated and serialized inert VBA project.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    pub fn put_vba(
        &mut self,
        workbook_code_name: &str,
        payload: litchi_vba::Payload,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        self.vba_metadata = Some(VbaWriteMetadata {
            workbook_code_name: workbook_code_name.to_string(),
            project: payload,
        });
        Ok(())
    }

    /// Remove the configured project and all worksheet VBA code names.
    pub fn clear_vba(&mut self) {
        self.vba_metadata = None;
        for worksheet in &mut self.worksheets {
            worksheet.vba_code_name = None;
        }
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_metadata.is_some()
    }

    /// Set a worksheet's tab color as a BIFF8 palette index (SHEETEXT
    /// `icvPlain`, MS-XLS 2.4.259). Valid indices are 0x08 through 0x3F;
    /// `None` clears an explicitly set color.
    pub fn set_worksheet_tab_color(&mut self, sheet: usize, tab_color: Option<u8>) -> Result<()> {
        if let Some(index) = tab_color
            && !(0x08..=0x3F).contains(&index)
        {
            return Err(Error::InvalidData(format!(
                "sheet tab color index {index:#04X} is outside the Icv palette"
            )));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.tab_color = tab_color;
        Ok(())
    }

    /// Author a what-if data table (MS-XLS 2.4.319) anchored at a formula
    /// cell. The anchor cell is written as a `PtgTbl` formula immediately
    /// followed by the `Table` record; it must lie outside the table range
    /// and must not already carry a value.
    pub fn add_data_table(
        &mut self,
        sheet: usize,
        anchor_row: u32,
        anchor_col: u16,
        table: crate::DataTable,
    ) -> Result<()> {
        let anchor_pos = CellPos::try_new(anchor_row, anchor_col)?;
        let range = table.range();
        let inside = (u32::from(range.first_row())..=u32::from(range.last_row()))
            .contains(&anchor_row)
            && (range.first_col()..=range.last_col()).contains(&anchor_pos.col());
        if inside {
            return Err(Error::InvalidData(
                "data-table anchor formula cell must lie outside the table range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet
            .data_tables
            .iter()
            .any(|(row, col, _)| (*row, *col) == (anchor_row, anchor_col))
        {
            return Err(Error::InvalidData(
                "duplicate data-table anchor cell".to_string(),
            ));
        }
        if let Some(cell) = worksheet.cells.get(&(anchor_row, anchor_col)) {
            if !matches!(cell.value, CellValue::Blank) {
                return Err(Error::InvalidData(
                    "data-table anchor cell already carries a value".to_string(),
                ));
            }
        } else {
            worksheet.add_cell(WritableCell::new(anchor_pos, CellValue::Blank, 0, None));
        }
        worksheet.data_tables.push((anchor_row, anchor_col, table));
        Ok(())
    }

    pub fn set_worksheet_vba_code_name(
        &mut self,
        sheet: usize,
        code_name: Option<&str>,
    ) -> Result<()> {
        if self.vba_metadata.is_none() && code_name.is_some() {
            return Err(Error::InvalidData(
                "worksheet VBA code names require an enabled VBA project".to_string(),
            ));
        }
        if let Some(value) = code_name {
            crate::vba::validate_code_name(value)?;
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.vba_code_name = code_name.map(str::to_string);
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_worksheet_layout(
        &mut self,
        sheet: usize,
        options: WorksheetLayoutOptions,
    ) -> Result<()> {
        options.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_layout = options;
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_page_setup(&mut self, sheet: usize, options: PageSetupOptions) -> Result<()> {
        let valid_margin = |value: f64| value.is_finite() && (0.0..49.0).contains(&value);
        if !valid_margin(options.left_margin_inches)
            || !valid_margin(options.right_margin_inches)
            || !valid_margin(options.top_margin_inches)
            || !valid_margin(options.bottom_margin_inches)
            || !valid_margin(options.header_margin_inches)
            || !valid_margin(options.footer_margin_inches)
        {
            return Err(Error::InvalidData(
                "page margins must be finite and between 0 and 49 inches".to_string(),
            ));
        }
        if options.header.encode_utf16().count() > 255
            || options.footer.encode_utf16().count() > 255
        {
            return Err(Error::InvalidData(
                "header and footer must not exceed 255 UTF-16 code units".to_string(),
            ));
        }
        if (118..=255).contains(&options.paper_size)
            || !(10..=400).contains(&options.scale_percent)
            || options.fit_width_pages > 32767
            || options.fit_height_pages > 32767
            || options.horizontal_resolution_dpi == 0
            || options.vertical_resolution_dpi == 0
            || options.copies == 0
            || options.copies > 32767
        {
            return Err(Error::InvalidData(
                "page setup contains an out-of-range dimension".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = Some(options);
        Ok(())
    }

    pub fn clear_page_setup(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = None;
        worksheet.horizontal_page_breaks.clear();
        worksheet.vertical_page_breaks.clear();
        Ok(())
    }

    /// Add a horizontal break at the first row below the break.
    pub fn add_horizontal_page_break(
        &mut self,
        sheet: usize,
        row: u32,
        col_start: u16,
        col_end: u16,
    ) -> Result<()> {
        let page_break = HorizontalPageBreak::try_new(row, col_start, col_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.horizontal_page_breaks.len() >= 1026 {
            return Err(Error::InvalidData(
                "horizontal page-break count exceeds 1026".to_string(),
            ));
        }
        if worksheet
            .horizontal_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "horizontal page-break ranges overlap".to_string(),
            ));
        }
        worksheet.horizontal_page_breaks.push(page_break);
        Ok(())
    }

    /// Add a vertical break at the first column right of the break.
    pub fn add_vertical_page_break(
        &mut self,
        sheet: usize,
        column: u16,
        row_start: u32,
        row_end: u32,
    ) -> Result<()> {
        let page_break = VerticalPageBreak::try_new(column, row_start, row_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.vertical_page_breaks.len() >= 255 {
            return Err(Error::InvalidData(
                "vertical page-break count exceeds 255".to_string(),
            ));
        }
        if worksheet
            .vertical_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "vertical page-break ranges overlap".to_string(),
            ));
        }
        worksheet.vertical_page_breaks.push(page_break);
        Ok(())
    }

    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        protect_structure: bool,
        protect_windows: bool,
    ) {
        if !protect_structure && !protect_windows && password.is_none() {
            self.workbook_protection = None;
            return;
        }

        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_structure = protect_structure;
        protection.protect_windows = protect_windows;
        protection.password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    pub fn unprotect_workbook(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_structure = false;
            protection.protect_windows = false;
            protection.password_hash = None;
            self.workbook_protection = protection.protect_revisions.then_some(protection);
        }
    }

    /// Configure legacy shared-workbook revision protection.
    pub fn protect_revisions(&mut self, password: Option<&str>) {
        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_revisions = true;
        protection.revision_password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    /// Remove shared-workbook revision protection.
    pub fn unprotect_revisions(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_revisions = false;
            protection.revision_password_hash = None;
            self.workbook_protection = (protection.protect_structure
                || protection.protect_windows
                || protection.password_hash.is_some())
            .then_some(protection);
        }
    }

    /// Configure read-only recommendation and an optional write-reservation password.
    pub fn set_file_sharing(
        &mut self,
        read_only_recommended: bool,
        password: Option<&str>,
        user_name: &str,
    ) -> Result<()> {
        if user_name.encode_utf16().count() > 54 {
            return Err(Error::InvalidData(
                "FILESHARING username exceeds 54 UTF-16 code units".to_string(),
            ));
        }
        self.file_sharing = Some(FileSharing {
            read_only_recommended,
            password_hash: password.map(Self::hash_password),
            user_name: user_name.to_string(),
        });
        Ok(())
    }

    pub fn clear_file_sharing(&mut self) {
        self.file_sharing = None;
    }

    pub fn protect_sheet(
        &mut self,
        sheet: usize,
        password: Option<&str>,
        protect_objects: bool,
        protect_scenarios: bool,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let password_hash = password.map(Self::hash_password);
        worksheet.sheet_protection = Some(SheetProtection {
            protect_objects,
            protect_scenarios,
            password_hash,
        });

        Ok(())
    }

    pub fn unprotect_sheet(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_protection = None;
        Ok(())
    }

    /// Save the XLS file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), Error>` - Success or error
    ///
    /// # Implementation Status
    ///
    /// ✅ Basic structure generation (BOF, EOF, workbook globals)
    /// ✅ Cell record generation (Number, LabelSST, BoolErr)
    /// ✅ Shared string table (SST)
    /// ✅ Formula tokenization for the supported BIFF8 writer subset
    /// ❌ Cell formatting (XF records)
    /// ❌ Column widths / row heights
    /// ❌ Merged cells
    /// ❌ Named ranges
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(&mut ole_writer, &streams)?;

        // Save to file
        ole_writer.save(path)?;

        Ok(())
    }

    /// Write to a writer (useful for testing and in-memory generation)
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer
    ///
    /// # Returns
    ///
    /// * `Result<(), Error>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(&mut self, writer: &mut W) -> Result<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(&mut ole_writer, &streams)?;

        // Write to the provided writer
        ole_writer.write_to(writer)?;

        Ok(())
    }

    fn populate_compound_document(
        &self,
        ole_writer: &mut OleWriter,
        streams: &stream::WorkbookStreams,
    ) -> Result<()> {
        ole_writer.create_stream(&["Workbook"], &streams.workbook)?;
        if let Some(metadata) = &self.vba_metadata {
            ole_writer.create_storage(&["_VBA_PROJECT_CUR"])?;
            metadata
                .project
                .write_into(ole_writer, &["_VBA_PROJECT_CUR"])?;
        }

        // Pivot cache storage: _SX_DB_CUR/XXXX. Stream names use four-digit
        // uppercase hexadecimal identifiers, matching the legacy convention.
        if !streams.pivot_caches.is_empty() {
            ole_writer.create_storage(&["_SX_DB_CUR"])?;
            for (id, data) in &streams.pivot_caches {
                let name = format!("{:04X}", id);
                ole_writer.create_stream(&["_SX_DB_CUR", &name], data)?;
            }
        }
        Ok(())
    }

    /// Build the shared string table from all string cells
    fn build_shared_strings(&mut self) {
        self.shared_strings.clear();
        self.string_map.clear();
        self.sst_total = 0;

        // Collect all unique strings from all worksheets
        for worksheet in &self.worksheets {
            for cell in worksheet.cells.values() {
                if let CellValue::String(ref s) = cell.value {
                    // Count total occurrences
                    self.sst_total = self.sst_total.saturating_add(1);
                    // Insert unique strings
                    if !self.string_map.contains_key(s) {
                        let index = self.shared_strings.len() as u32;
                        self.string_map.insert(s.clone(), index);
                        self.shared_strings.push(s.clone());
                    }
                }
            }
        }
    }

    /// Generate the complete Workbook stream (plus pivot cache streams) with
    /// all BIFF records.
    fn generate_workbook_streams(&self) -> Result<stream::WorkbookStreams> {
        let mut streams = stream::generate_workbook_stream(
            self.use_1904_dates,
            self.calculation_settings,
            self.vba_metadata.as_ref(),
            self.environment_options,
            self.workbook_window_options,
            &self.function_group_options,
            &self.external_workbooks,
            &self.external_names,
            &self.add_in_functions,
            &self.dde_or_ole_links,
            &self.fmt,
            self.custom_table_styles.as_ref(),
            &self.defined_names,
            &self.defined_name_records,
            &self.shared_strings,
            self.sst_total,
            self.workbook_protection,
            self.file_sharing.as_ref(),
            self.book_ext.as_ref(),
            self.theme.as_ref(),
            self.mdx_metadata.as_ref(),
            &self.real_time_data,
            &self.web_publications,
            &self.xf_extensions,
            &self.style_extensions,
            &self.worksheets,
            &self.string_map,
        )?;
        if let Some(encryption) = &self.encryption {
            streams.workbook = encrypt_workbook_for_write(streams.workbook, encryption)?;
        }
        Ok(streams)
    }

    /// Get the number of worksheets in this workbook
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get worksheet name by index
    pub fn get_worksheet_name(&self, index: usize) -> Option<&str> {
        self.worksheets.get(index).map(|w| w.name.as_str())
    }

    // Implementation status notes:
    // ✅ Building shared string table (SST) with deduplication - IMPLEMENTED
    // ✅ Generating BIFF8 records for supported cell types - Number, LabelSST, BoolErr, Formula
    // ❌ Worksheet management (rename, delete, reorder) - Future enhancement
    // ❌ Cell formatting (fonts, colors, borders, number formats) - Future enhancement
    // ❌ Column widths and row heights - Future enhancement
    // ❌ Merged cells - Future enhancement
    // ✅ Named ranges (simple A1-style, workbook and sheet scoped) - IMPLEMENTED
    // ✅ Formula parsing and tokenization for the supported writer subset
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}

/// Iterate every OBJ identifier requested or assigned inside a shape group.
fn group_object_ids(group: &ShapeGroupWrite) -> impl Iterator<Item = u16> + '_ {
    group
        .object_id
        .into_iter()
        .chain(group.children.iter().filter_map(|child| child.object_id))
}

/// Collect every existing worksheet drawing-object identifier with fallible
/// capacity reservation for any IDs the caller will add before retaining data.
fn collect_reserved_object_ids(
    worksheet: &WritableWorksheet,
    additional: usize,
) -> Result<HashSet<u16>> {
    let pivot_capacity = worksheet
        .pivot_tables
        .iter()
        .try_fold(0usize, |count, table| {
            count
                .checked_add(table.page_entries.len())
                .ok_or(Error::Allocation(
                    "computing worksheet pivot-object ID capacity",
                ))
        })?;
    let group_capacity = worksheet
        .shape_groups
        .iter()
        .try_fold(0usize, |count, group| {
            count
                .checked_add(group.object_count()?)
                .ok_or(Error::Allocation(
                    "computing worksheet shape-group ID capacity",
                ))
        })?;
    let capacity = pivot_capacity
        .checked_add(worksheet.shapes.len())
        .and_then(|count| count.checked_add(group_capacity))
        .and_then(|count| count.checked_add(additional))
        .ok_or(Error::Allocation(
            "computing worksheet drawing-object ID capacity",
        ))?;
    let mut reserved = HashSet::new();
    reserved
        .try_reserve(capacity)
        .map_err(|_| Error::Allocation("reserving worksheet drawing-object ID storage"))?;
    reserved.extend(
        worksheet
            .pivot_tables
            .iter()
            .flat_map(|table| table.page_entries.iter().map(|entry| entry.2))
            .filter(|id| *id != 0 && *id != u16::MAX),
    );
    reserved.extend(worksheet.shapes.iter().filter_map(|shape| shape.object_id));
    reserved.extend(worksheet.shape_groups.iter().flat_map(group_object_ids));
    Ok(reserved)
}

/// Reserve the requested OBJ identifier or the first free canonical one.
fn assign_object_id(reserved: &mut HashSet<u16>, requested: Option<u16>) -> Result<u16> {
    let object_id = match requested {
        Some(object_id) => object_id,
        None => (1..u16::MAX)
            .find(|candidate| !reserved.contains(candidate))
            .ok_or_else(|| Error::InvalidData("worksheet object IDs are exhausted".to_string()))?,
    };
    if requested.is_none() {
        reserved.insert(object_id);
    }
    Ok(object_id)
}

/// Implementation notes for BIFF record generation:
///
/// All core BIFF8 records have been implemented in the `biff` module:
/// - ✅ write_bof() - Beginning of File (0x0809)
/// - ✅ write_eof() - End of File (0x000A)
/// - ✅ write_codepage() - Code page (0x0042)
/// - ✅ write_date1904() - Date system (0x0022)
/// - ✅ write_window1() - Workbook window properties (0x003D)
/// - ✅ write_boundsheet() - Sheet metadata (0x0085)
/// - ✅ write_dimensions() - Worksheet dimensions (0x0200)
/// - ✅ write_sst() - Shared string table with CONTINUE support (0x00FC)
/// - ✅ write_number() - Floating point cell (0x0203)
/// - ✅ write_labelsst() - String cell (0x00FD)
/// - ✅ write_boolerr() - Boolean/error cell (0x0205)
/// - ✅ write_formula() - Formula cell with RPN tokens (0x0006)
/// - ✅ write_continue() - Continuation record (0x003C)
///
/// Future enhancements:
/// - XF records (0x00E0) - For cell formatting
/// - FONT records (0x0031) - For font definitions
/// - FORMAT records (0x041E) - For number formats
/// - COLINFO records (0x007D) - For column widths
/// - ROW records (0x0208) - For row heights
/// - MERGEDCELLS records (0x00E5) - For merged cell ranges
/// - NAME records (0x0018) - For named ranges
#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::sheet::Cell;
    use std::io::Cursor;

    #[test]
    fn test_create_writer() {
        let writer = Writer::new();
        assert_eq!(writer.worksheets.len(), 0);
        assert_eq!(writer.shared_strings.len(), 0);
    }

    #[test]
    fn test_add_worksheet() {
        let mut writer = Writer::new();
        let idx = writer.add_worksheet("Sheet1").unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.worksheets.len(), 1);
        assert_eq!(writer.worksheets[0].name, "Sheet1");
    }

    #[test]
    fn test_add_multiple_worksheets() {
        let mut writer = Writer::new();
        let idx1 = writer.add_worksheet("Sheet1").unwrap();
        let idx2 = writer.add_worksheet("Sheet2").unwrap();
        let idx3 = writer.add_worksheet("Sheet3").unwrap();

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 2);
        assert_eq!(writer.worksheets.len(), 3);
    }

    #[test]
    fn test_add_worksheet_empty_name() {
        let mut writer = Writer::new();
        let result = writer.add_worksheet("");
        assert!(result.is_err());
    }

    #[test]
    fn test_add_worksheet_long_name() {
        let mut writer = Writer::new();
        let long_name = "A".repeat(50);
        let result = writer.add_worksheet(&long_name);
        assert!(result.is_err()); // Name too long
    }

    #[test]
    fn test_add_worksheet_duplicate_name() {
        let mut writer = Writer::new();
        writer.add_worksheet("Sheet1").unwrap();
        let result = writer.add_worksheet("Sheet1");
        assert!(result.is_err());
    }

    #[test]
    fn test_write_string() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert_eq!(cell.row(), 0);
        assert_eq!(cell.col(), 0);
        assert!(matches!(&cell.value, CellValue::String(s) if s == "Hello"));
    }

    #[test]
    fn test_write_number() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 42.5).unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert!(matches!(&cell.value, CellValue::Number(n) if *n == 42.5));
    }

    #[test]
    fn non_finite_cell_numbers_are_rejected_before_mutation() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.write_number(sheet, 0, 0, value)
            }));
            assert!(matches!(
                result,
                Ok(Err(Error::InvalidData(message)))
                    if message == "cell number must be finite for BIFF8 serialization"
            ));
            assert!(writer.worksheets[sheet].cells.is_empty());
        }
    }

    #[test]
    fn cell_grid_bounds_are_atomic_and_never_unwind() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        let max = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.write_number(sheet, u32::from(u16::MAX), u16::from(u8::MAX), 42.5)
        }));
        assert!(matches!(max, Ok(Ok(()))));
        let state = |writer: &Writer| {
            let worksheet = &writer.worksheets[sheet];
            (
                worksheet.cells.len(),
                worksheet.first_row,
                worksheet.last_row,
                worksheet.first_col,
                worksheet.last_col,
            )
        };
        let max_state = (1, 65_535, 65_536, 255, 256);
        assert_eq!(state(&writer), max_state);

        let oversized_row = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.write_string(sheet, 65_536, 0, "outside")
        }));
        assert!(matches!(
            oversized_row,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(state(&writer), max_state);

        let adversarial_row = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.write_number(sheet, u32::MAX, 0, 1.0)
        }));
        assert!(matches!(
            adversarial_row,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(state(&writer), max_state);

        let oversized_col = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.write_formula(sheet, 0, 256, "1")
        }));
        assert!(matches!(
            oversized_col,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(state(&writer), max_state);

        let table = crate::DataTable::one_variable(
            crate::DataTableRange::new(2, 8, 3, 5).unwrap(),
            false,
            crate::DataTableInputCell::Deleted,
        );
        let adversarial_col = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_data_table(sheet, 0, u16::MAX, table)
        }));
        assert!(matches!(
            adversarial_col,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(state(&writer), max_state);
        assert!(writer.worksheets[sheet].data_tables.is_empty());

        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.set_position(0);
        let workbook = crate::Workbook::new(output).unwrap();
        let cell = workbook
            .xls_worksheet(0)
            .unwrap()
            .get_cell(u32::from(u16::MAX), u32::from(u8::MAX))
            .unwrap();
        assert!(matches!(
            cell.value(),
            litchi_core::sheet::CellValue::Float(value) if *value == 42.5
        ));
    }

    #[test]
    fn worksheet_location_bounds_are_atomic_and_never_unwind() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Grid").unwrap();

        writer.merge_cells(sheet, 65_534, 65_535, 254, 255).unwrap();
        writer
            .add_horizontal_page_break(sheet, 65_535, 0, 16_383)
            .unwrap();
        writer
            .add_vertical_page_break(sheet, 255, 0, 65_535)
            .unwrap();
        let max_range = DataValidationRange::new(65_535, 65_535, 255, 255).unwrap();
        writer
            .add_data_validation(
                sheet,
                DataValidation::new(max_range, DataValidationType::Any),
            )
            .unwrap();

        let state = |writer: &Writer| {
            let worksheet = &writer.worksheets[sheet];
            (
                worksheet.merged_ranges.len(),
                worksheet.horizontal_page_breaks.len(),
                worksheet.vertical_page_breaks.len(),
                worksheet.data_validations.len(),
                writer.real_time_data.len(),
                writer.web_publications.len(),
            )
        };
        let max_state = (1, 1, 1, 1, 0, 0);
        assert_eq!(state(&writer), max_state);

        for outcome in [
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.merge_cells(sheet, 65_536, 65_536, 0, 0)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.merge_cells(sheet, 0, 0, 256, 256)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.merge_cells(sheet, 2, 1, 0, 0)
            })),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }
        let merge_overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.merge_cells(sheet, 65_535, 65_535, 255, 255)
        }));
        assert!(matches!(merge_overlap, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(state(&writer), max_state);

        for outcome in [
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.add_horizontal_page_break(sheet, 65_536, 0, 1)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.add_horizontal_page_break(sheet, 0, 0, 16_384)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.add_vertical_page_break(sheet, 256, 0, 1)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.add_vertical_page_break(sheet, 0, 0, 65_536)
            })),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }

        let overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_horizontal_page_break(sheet, 65_535, 1, 2)
        }));
        assert!(matches!(overlap, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(state(&writer), max_state);
        let overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_vertical_page_break(sheet, 255, 1, 2)
        }));
        assert!(matches!(overlap, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(state(&writer), max_state);

        for outcome in [
            std::panic::catch_unwind(|| DataValidationRange::new(65_536, 65_536, 0, 0)),
            std::panic::catch_unwind(|| DataValidationRange::new(0, 0, 256, 256)),
            std::panic::catch_unwind(|| DataValidationRange::new(1, 0, 0, 0)),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }

        let too_many_ranges = vec![max_range; 432];
        let count_error = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_data_validation_with_options(
                sheet,
                DataValidation::new(max_range, DataValidationType::Any),
                &too_many_ranges,
                DataValidationOptions::default(),
            )
        }));
        assert!(matches!(count_error, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(state(&writer), max_state);

        let invalid_rule = DataValidation::new(
            max_range,
            DataValidationType::Whole {
                operator: DataValidationOperator::Between,
                value1: 1,
                value2: None,
            },
        );
        let rule_error = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_data_validation(sheet, invalid_rule)
        }));
        assert!(matches!(rule_error, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(state(&writer), max_state);

        assert!(matches!(
            crate::DataTableInputCell::present(65_535, 255),
            Ok(crate::DataTableInputCell::Present {
                row: 65_535,
                col: 255
            })
        ));
        for outcome in [
            std::panic::catch_unwind(|| crate::DataTableInputCell::present(65_536, 0)),
            std::panic::catch_unwind(|| crate::DataTableInputCell::present(u32::MAX, 0)),
            std::panic::catch_unwind(|| crate::DataTableInputCell::present(0, 256)),
            std::panic::catch_unwind(|| crate::DataTableInputCell::present(0, u16::MAX)),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }
        let invalid_publication = crate::WebPub {
            source: crate::WebSourceType::Workbook,
            page_type: crate::WebPageType::ViewOnly,
            range: Some(crate::WebPubRange::new(0, 0, 0, 0).unwrap()),
            auto_republish: false,
            single_file: false,
            style_id: 1,
            source_name: None,
            file_destination: "x.htm".to_string(),
            div_id: String::new(),
            title: "x".to_string(),
            chart_shape_id: None,
            reserved: Vec::new(),
        };
        let invalid_publication = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_web_publication(invalid_publication)
        }));
        assert!(matches!(
            invalid_publication,
            Ok(Err(Error::InvalidData(_)))
        ));
        assert_eq!(state(&writer), max_state);

        for outcome in [
            std::panic::catch_unwind(|| crate::WebPubRange::new(65_536, 65_536, 0, 0)),
            std::panic::catch_unwind(|| crate::WebPubRange::new(0, 0, 256, 256)),
            std::panic::catch_unwind(|| crate::WebPubRange::new(1, 0, 0, 0)),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }

        for outcome in [
            std::panic::catch_unwind(|| crate::RtdCell::new(65_536, 0, 0)),
            std::panic::catch_unwind(|| crate::RtdCell::new(0, 256, 0)),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(state(&writer), max_state);
        }
        let invalid_topic = crate::RealTimeData {
            common_prefix_len: 0,
            topic_segments: vec!["server".to_string(), String::new()],
            topic: "server".to_string(),
            value: crate::RtdValue::Integer(1),
            cells: vec![crate::RtdCell::new(0, 0, 1).unwrap()],
        };
        let missing_sheet = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_real_time_data(invalid_topic)
        }));
        assert!(matches!(
            missing_sheet,
            Ok(Err(Error::WorksheetNotFound(_)))
        ));
        assert_eq!(state(&writer), max_state);

        writer
            .set_page_setup(sheet, PageSetupOptions::default())
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.set_position(0);
        let workbook = crate::Workbook::new(output).unwrap();
        let worksheet = workbook.xls_worksheet(sheet).unwrap();
        assert!(
            worksheet
                .merged_cells()
                .iter()
                .any(|range| range.first_row == 65_534
                    && range.last_row == 65_535
                    && range.first_col == 254
                    && range.last_col == 255)
        );
        let page_setup = worksheet.page_setup().unwrap();
        assert_eq!(page_setup.horizontal_page_breaks()[0].position(), 65_535);
        assert_eq!(page_setup.horizontal_page_breaks()[0].range_end(), 16_383);
        assert_eq!(page_setup.vertical_page_breaks()[0].position(), 255);
        assert_eq!(page_setup.vertical_page_breaks()[0].range_end(), 65_535);
    }

    #[test]
    fn page_break_entry_limits_are_failure_atomic() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Break limits").unwrap();

        for row in 0..1_026 {
            writer.add_horizontal_page_break(sheet, row, 0, 1).unwrap();
        }
        for column in 0..255 {
            writer.add_vertical_page_break(sheet, column, 0, 1).unwrap();
        }
        assert_eq!(writer.worksheets[sheet].horizontal_page_breaks.len(), 1_026);
        assert_eq!(writer.worksheets[sheet].vertical_page_breaks.len(), 255);

        let horizontal = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_horizontal_page_break(sheet, 1_026, 0, 1)
        }));
        assert!(matches!(horizontal, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(writer.worksheets[sheet].horizontal_page_breaks.len(), 1_026);

        let vertical = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_vertical_page_break(sheet, 255, 0, 1)
        }));
        assert!(matches!(vertical, Ok(Err(Error::InvalidData(_)))));
        assert_eq!(writer.worksheets[sheet].vertical_page_breaks.len(), 255);
    }

    #[test]
    fn filter_sort_and_pivot_locations_fail_before_mutation() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Locations").unwrap();
        writer.set_auto_filter(sheet, 0, 65_535, 0, 255).unwrap();
        let filter_state = |writer: &Writer| {
            let worksheet = &writer.worksheets[sheet];
            (
                worksheet.auto_filter.map(|range| {
                    (
                        range.first_row,
                        range.last_row,
                        range.first_col,
                        range.last_col,
                    )
                }),
                worksheet.auto_filter_columns.len(),
                worksheet.sort_config.is_some(),
                worksheet.pivot_tables.len(),
                worksheet.cells.len(),
                writer.defined_names.len(),
            )
        };
        let initial = (Some((0, 65_535, 0, 255)), 0, false, 0, 0, 1);
        assert_eq!(filter_state(&writer), initial);

        for outcome in [
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.set_auto_filter(sheet, 0, 65_536, 0, 0)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.set_auto_filter(sheet, 0, 0, 0, 256)
            })),
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                writer.set_auto_filter(sheet, 1, 0, 0, 0)
            })),
        ] {
            assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
            assert_eq!(filter_state(&writer), initial);
        }

        let invalid_filter = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_filter_condition(
                sheet,
                256,
                false,
                AutoFilterConditionWrite::None,
                AutoFilterConditionWrite::None,
            )
        }));
        assert!(matches!(
            invalid_filter,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(filter_state(&writer), initial);

        let invalid_sort = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.set_sort(sheet, false, false, &[(256, false)])
        }));
        assert!(matches!(
            invalid_sort,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(filter_state(&writer), initial);

        let invalid_pivot = PivotTableConfig {
            name: "InvalidPivot".to_string(),
            source_type: 1,
            source_sheet_name: "Locations".to_string(),
            source_first_row: 0,
            source_last_row: 0,
            source_first_col: 0,
            source_last_col: 256,
            first_row: 0,
            last_row: 1,
            first_col: 0,
            last_col: 1,
            first_header_row: 0,
            first_data_row: 1,
            first_data_col: 1,
            data_field_name: "Values".to_string(),
            data_axis: 0,
            data_position: 0,
            fields: Vec::new(),
            data_items: Vec::new(),
            page_entries: Vec::new(),
            source_data: Vec::new(),
        };
        let invalid_pivot = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_pivot_table(sheet, invalid_pivot)
        }));
        assert!(matches!(
            invalid_pivot,
            Ok(Err(Error::InvalidCellReference(_)))
        ));
        assert_eq!(filter_state(&writer), initial);
    }

    #[test]
    fn test_write_boolean() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_boolean(sheet, 0, 0, true).unwrap();
        writer.write_boolean(sheet, 1, 0, false).unwrap();

        assert_eq!(writer.worksheets[0].cells.len(), 2);
        assert!(matches!(
            writer.worksheets[0].cells.get(&(0, 0)).unwrap().value,
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            writer.worksheets[0].cells.get(&(1, 0)).unwrap().value,
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_write_formula() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_formula(sheet, 0, 0, "SUM(A1:B1)").unwrap();

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert!(matches!(&cell.value, CellValue::Formula(f) if f == "SUM(A1:B1)"));
    }

    #[test]
    fn invalid_formula_reference_returns_error_without_unwinding() {
        let outcome = std::panic::catch_unwind(|| -> Result<()> {
            let mut writer = Writer::new();
            let sheet = writer.add_worksheet("Sheet1")?;
            writer.write_formula(sheet, 0, 0, "ZZZZ1")?;
            writer.write_to(&mut Cursor::new(Vec::new()))
        });

        assert!(matches!(
            outcome,
            Ok(Err(Error::InvalidCellReference(reference))) if reference == "ZZZZ1"
        ));
    }

    #[test]
    fn test_formula_round_trips_through_xls_reader() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 2.0).unwrap();
        writer.write_number(sheet, 0, 1, 3.0).unwrap();
        writer.write_formula(sheet, 0, 2, "SUM(A1:B1)").unwrap();
        writer
            .write_formula(sheet, 0, 3, "IF(TRUE,\"a\"\"b\",FALSE)")
            .unwrap();

        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.set_position(0);
        let workbook = crate::Workbook::new(output).unwrap();
        let formula_cell = workbook.xls_worksheet(0).unwrap().get_cell(0, 2).unwrap();

        assert!(formula_cell.is_formula());
        assert_eq!(formula_cell.formula(), Some("=SUM((A1:B1))"));
        assert!(!formula_cell.formula_bytes().unwrap().is_empty());
        assert_eq!(
            workbook
                .xls_worksheet(0)
                .unwrap()
                .get_cell(0, 3)
                .unwrap()
                .formula(),
            Some("=IF(TRUE,\"a\"\"b\",FALSE)")
        );
    }

    #[test]
    fn test_write_multiple_cells() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        writer.write_string(sheet, 0, 0, "A1").unwrap();
        writer.write_string(sheet, 0, 1, "B1").unwrap();
        writer.write_string(sheet, 1, 0, "A2").unwrap();
        writer.write_string(sheet, 1, 1, "B2").unwrap();

        assert_eq!(writer.worksheets[0].cells.len(), 4);
    }

    #[test]
    fn test_shared_strings_build() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        writer.write_string(sheet, 0, 1, "Hello").unwrap();
        writer.write_string(sheet, 1, 0, "World").unwrap();

        // Build shared strings table (normally done during write)
        writer.build_shared_strings();

        // Should only have 2 unique strings
        assert_eq!(writer.shared_strings.len(), 2);
    }

    #[test]
    fn test_write_to_memory() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Test").unwrap();
        writer.write_number(sheet, 0, 1, 123.45).unwrap();

        let mut cursor = Cursor::new(Vec::new());
        let result = writer.write_to(&mut cursor);
        assert!(result.is_ok());

        let data = cursor.into_inner();
        assert!(!data.is_empty());
        // Should start with OLE compound document signature
        assert_eq!(
            &data[0..8],
            [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
        );
    }

    #[test]
    fn test_save_to_file() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();

        let temp_path = std::env::temp_dir().join("test_xls_writer.xls");
        let result = writer.save(&temp_path);
        assert!(result.is_ok());

        // Verify file was created
        assert!(temp_path.exists());

        // Clean up
        let _ = std::fs::remove_file(temp_path);
    }

    #[test]
    fn test_xls_writer_default() {
        let writer: Writer = Default::default();
        assert_eq!(writer.worksheets.len(), 0);
        assert_eq!(writer.shared_strings.len(), 0);
    }

    #[test]
    fn test_xlscellvalue_variants() {
        let string_val = CellValue::String("test".to_string());
        let number_val = CellValue::Number(42.0);
        let bool_val = CellValue::Boolean(true);
        let formula_val = CellValue::Formula("A1+B1".to_string());
        let blank_val = CellValue::Blank;

        assert!(matches!(string_val, CellValue::String(_)));
        assert!(matches!(number_val, CellValue::Number(_)));
        assert!(matches!(bool_val, CellValue::Boolean(_)));
        assert!(matches!(formula_val, CellValue::Formula(_)));
        assert!(matches!(blank_val, CellValue::Blank));
    }

    #[test]
    fn test_xlscellvalue_debug() {
        let val = CellValue::String("test".to_string());
        let debug = format!("{:?}", val);
        assert!(debug.contains("String"));
    }

    #[test]
    fn test_xlscellvalue_clone() {
        let val = CellValue::Number(42.0);
        let cloned = val.clone();
        assert!(matches!(cloned, CellValue::Number(42.0)));
    }

    #[test]
    fn test_writablecell_creation() {
        let cell = WritableCell::new(
            CellPos::try_new(5, 3).unwrap(),
            CellValue::String("Test".to_string()),
            15,
            None,
        );

        assert_eq!(cell.row(), 5);
        assert_eq!(cell.col(), 3);
        assert_eq!(cell.format_idx, 15);
        assert!(matches!(
            CellPos::try_new(65_536, 0),
            Err(Error::InvalidCellReference(_))
        ));
        assert!(matches!(
            CellPos::try_new(0, 256),
            Err(Error::InvalidCellReference(_))
        ));
    }

    #[test]
    fn test_writableworksheet_creation() {
        let ws = WritableWorksheet::new("TestSheet".to_string());
        assert_eq!(ws.name, "TestSheet");
        assert!(ws.cells.is_empty());
        assert!(ws.merged_ranges.is_empty());
        assert!(ws.column_widths.is_empty());
    }

    #[test]
    fn test_writableworksheet_add_cell() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let cell = WritableCell::new(
            CellPos::try_new(0, 0).unwrap(),
            CellValue::Number(100.0),
            0,
            None,
        );
        ws.add_cell(cell);
        assert_eq!(ws.cells.len(), 1);
    }

    #[test]
    fn test_writableworksheet_set_column_width() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        ws.set_column_width(0, 2560); // ~10 characters
        assert_eq!(ws.column_widths.get(&0), Some(&2560));
    }

    #[test]
    fn test_writableworksheet_merge_cells() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        ws.add_merged_range(MergedRange::try_new(0, 1, 0, 2).unwrap()); // Merge A1:C2
        assert_eq!(ws.merged_ranges.len(), 1);
        assert_eq!(ws.merged_ranges[0].fields(), (0, 1, 0, 2));
    }

    #[test]
    fn test_writableworksheet_freeze_panes() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        assert!(ws.view.pane().is_none());
        ws.set_freeze_panes(1, 2).unwrap();
        let pane = ws.view.pane().unwrap();
        assert_eq!(pane.vertical(), 1);
        assert_eq!(pane.horizontal(), 2);
    }

    #[test]
    fn test_writableworksheet_add_conditional_format() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let cf = ConditionalFormat {
            first_row: 0,
            last_row: 10,
            first_col: 0,
            last_col: 0,
            format_type: ConditionalFormatType::Formula {
                formula: "A1>100".to_string(),
            },
            pattern: None,
        };
        ws.add_conditional_format(cf);
        assert_eq!(ws.conditional_formats.len(), 1);
    }

    #[test]
    fn test_writableworksheet_add_data_validation() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let dv = DataValidation {
            range: DataValidationRange::new(0, 10, 0, 0).unwrap(),
            validation_type: DataValidationType::List {
                values: vec!["Option1".to_string(), "Option2".to_string()],
            },
            show_input_message: true,
            input_title: None,
            input_message: None,
            show_error_alert: true,
            error_title: None,
            error_message: None,
        };
        let payload = dv.validation_type.to_biff_payload().unwrap();
        ws.add_data_validation(
            dv,
            payload,
            vec![DataValidationRange::new(0, 9, 0, 0).unwrap()],
            DataValidationOptions::default(),
        );
        assert_eq!(ws.data_validations.len(), 1);
    }

    #[test]
    fn test_writableworksheet_add_hyperlink() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let link = Hyperlink {
            first_row: 0,
            last_row: 0,
            first_col: 0,
            last_col: 0,
            url: "https://example.com".to_string(),
        };
        ws.add_hyperlink(link);
        assert_eq!(ws.hyperlinks.len(), 1);
        assert_eq!(ws.hyperlinks[0].url, "https://example.com");
    }

    #[test]
    fn test_xls_defined_name_basic() {
        let name = DefinedName {
            name: "TestRange".to_string(),
            reference: "A1:B10".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(0),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        assert_eq!(name.name, "TestRange");
        assert_eq!(name.reference, "A1:B10");
        assert_eq!(name.target_sheet, Some(0));
    }

    #[test]
    fn test_xls_defined_name_to_biff_formula_area() {
        let name = DefinedName {
            name: "TestRange".to_string(),
            reference: "A1:B10".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(0),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        let formula = name.to_biff_formula().unwrap();
        assert!(!formula.is_empty());
    }

    #[test]
    fn test_xls_defined_name_normalizes_reversed_area_corners() {
        let forward = DefinedName {
            name: "Forward".to_string(),
            reference: "A1:B10".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(0),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        let reversed = DefinedName {
            name: "Reversed".to_string(),
            reference: "B10:A1".to_string(),
            ..forward.clone()
        };

        assert_eq!(
            reversed.to_biff_formula().unwrap(),
            forward.to_biff_formula().unwrap()
        );
    }

    #[test]
    fn test_xls_defined_name_to_biff_formula_single() {
        let name = DefinedName {
            name: "SingleCell".to_string(),
            reference: "C5".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: None,
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        let formula = name.to_biff_formula().unwrap();
        assert!(!formula.is_empty());
    }
}
