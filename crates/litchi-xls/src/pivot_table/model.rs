//! Typed PivotTable and PivotCache models.
//!
//! This module owns semantic values, validation, and the public object model;
//! BIFF record framing and parsing live in the sibling codec module.

use super::codec::{SXVD_TYPE, cache_invalid};
use crate::error::{Error, Result};

/// A BIFF8 cell error stored in an `SXERROR` PivotCache item.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum PivotCacheError {
    Null = 0x00,
    DivisionByZero = 0x07,
    Value = 0x0F,
    Reference = 0x17,
    Name = 0x1D,
    Number = 0x24,
    NotAvailable = 0x2A,
}

impl PivotCacheError {
    pub const fn code(self) -> u16 {
        self as u16
    }
}

impl TryFrom<u16> for PivotCacheError {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self, Self::Error> {
        match value {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2A => Ok(Self::NotAvailable),
            _ => Err(cache_invalid(
                0x00CB,
                format!("invalid BIFF error code 0x{value:04X}"),
            )),
        }
    }
}

/// Lossless calendar value stored by `SXDATETIME`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PivotCacheDateTime {
    year: u16,
    month: u16,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PivotCacheDateGroupUnit {
    Seconds = 1,
    Minutes = 2,
    Hours = 3,
    Days = 4,
    Months = 5,
    Quarters = 6,
    Years = 7,
}

impl TryFrom<u16> for PivotCacheDateGroupUnit {
    type Error = Error;
    fn try_from(value: u16) -> Result<Self, Self::Error> {
        match value {
            1 => Ok(Self::Seconds),
            2 => Ok(Self::Minutes),
            3 => Ok(Self::Hours),
            4 => Ok(Self::Days),
            5 => Ok(Self::Months),
            6 => Ok(Self::Quarters),
            7 => Ok(Self::Years),
            _ => Err(cache_invalid(
                0x00D8,
                format!("unsupported date grouping unit {value}"),
            )),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheNumericGrouping {
    pub start: f64,
    pub end: f64,
    pub step: f64,
    pub auto_start: bool,
    pub auto_end: bool,
    pub group_items: Vec<PivotCacheItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheDateGrouping {
    pub unit: PivotCacheDateGroupUnit,
    pub start: PivotCacheDateTime,
    pub end: PivotCacheDateTime,
    pub step: u16,
    pub auto_start: bool,
    pub auto_end: bool,
    pub group_items: Vec<PivotCacheItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheDiscreteGrouping {
    pub base_field_index: u16,
    pub group_items: Vec<PivotCacheItem>,
    /// One group-item ordinal for every original item in the base field.
    pub item_to_group: Vec<u16>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheGrouping {
    Numeric(PivotCacheNumericGrouping),
    Date(PivotCacheDateGrouping),
    Discrete(PivotCacheDiscreteGrouping),
}

impl PivotCacheGrouping {
    pub(crate) fn group_items(&self) -> &[PivotCacheItem] {
        match self {
            Self::Numeric(value) => &value.group_items,
            Self::Date(value) => &value.group_items,
            Self::Discrete(value) => &value.group_items,
        }
    }
}

impl PivotCacheDateTime {
    pub fn try_new(
        year: u16,
        month: u16,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
    ) -> Result<Self> {
        let value = Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
        };
        value.validate()?;
        Ok(value)
    }

    fn validate(self) -> Result<()> {
        let leap = self.year.is_multiple_of(4)
            && (!self.year.is_multiple_of(100) || self.year.is_multiple_of(400));
        let max_day = match self.month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 if leap || self.year == 1900 => 29,
            2 => 28,
            _ => 0,
        };
        let sentinel = self.year == 1900
            && self.month == 1
            && self.day == 0
            && self.hour == 0
            && self.minute == 0
            && self.second == 0;
        if !(1900..=9999).contains(&self.year)
            || max_day == 0
            || (!sentinel && !(1..=max_day).contains(&self.day))
            || self.hour > 23
            || self.minute > 59
            || self.second > 59
        {
            return Err(cache_invalid(
                0x00CE,
                format!(
                    "invalid PivotCache date/time {:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                    self.year, self.month, self.day, self.hour, self.minute, self.second
                ),
            ));
        }
        Ok(())
    }

    pub const fn year(self) -> u16 {
        self.year
    }
    pub const fn month(self) -> u16 {
        self.month
    }
    pub const fn day(self) -> u8 {
        self.day
    }
    pub const fn hour(self) -> u8 {
        self.hour
    }
    pub const fn minute(self) -> u8 {
        self.minute
    }
    pub const fn second(self) -> u8 {
        self.second
    }
}

/// A typed PivotCache shared or resolved row item.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheItem {
    String(String),
    Number(f64),
    Boolean(bool),
    Error(PivotCacheError),
    DateTime(PivotCacheDateTime),
    Empty,
}

impl From<String> for PivotCacheItem {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<&str> for PivotCacheItem {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl PivotCacheItem {
    pub(crate) fn validate(&self) -> Result<()> {
        match self {
            Self::String(value) if value.is_empty() => Err(cache_invalid(
                0x00CD,
                "empty strings must use PivotCacheItem::Empty",
            )),
            Self::Number(value) if !value.is_finite() => Err(cache_invalid(
                0x00C9,
                "PivotCache numeric items must be finite",
            )),
            Self::DateTime(value) => value.validate(),
            _ => Ok(()),
        }
    }

    pub(crate) fn display_text(&self) -> String {
        match self {
            Self::String(value) => value.clone(),
            Self::Number(value) => value.to_string(),
            Self::Boolean(value) => if *value { "TRUE" } else { "FALSE" }.to_string(),
            Self::Error(value) => format!("#{value:?}"),
            Self::DateTime(value) => format!(
                "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                value.year, value.month, value.day, value.hour, value.minute, value.second
            ),
            Self::Empty => String::new(),
        }
    }
}

pub(crate) fn pivot_cache_data_flags(items: &[PivotCacheItem]) -> u16 {
    let (mut text, mut integer, mut double, mut date, mut blank) =
        (false, false, false, false, false);
    for item in items {
        match item {
            PivotCacheItem::String(_) | PivotCacheItem::Boolean(_) | PivotCacheItem::Error(_) => {
                text = true
            },
            PivotCacheItem::Number(value) if value.fract() == 0.0 => integer = true,
            PivotCacheItem::Number(_) => double = true,
            PivotCacheItem::DateTime(_) => date = true,
            PivotCacheItem::Empty => blank = true,
        }
    }
    if date && blank && !text && !integer && !double {
        return 0x0980;
    }
    if blank {
        text = true;
    }
    const FLAGS: [u16; 16] = [
        0x0000, 0x0480, 0x0520, 0x05A0, 0x0560, 0x05E0, 0x0520, 0x05A0, 0x0900, 0x0D80, 0x0D00,
        0x0D80, 0x0D00, 0x0D80, 0x0D00, 0x0D80,
    ];
    FLAGS[usize::from(text)
        | (usize::from(integer) << 1)
        | (usize::from(double) << 2)
        | (usize::from(date) << 3)]
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCacheField {
    pub(crate) name: String,
    pub(crate) flags: u16,
    pub(crate) group_parent: Option<u16>,
    pub(crate) group_base: Option<u16>,
    pub(crate) items: Vec<PivotCacheItem>,
    pub(crate) grouping: Option<PivotCacheGrouping>,
}
impl PivotCacheField {
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn flags(&self) -> u16 {
        self.flags
    }
    pub fn items(&self) -> &[PivotCacheItem] {
        &self.items
    }
    pub const fn group_parent(&self) -> Option<u16> {
        self.group_parent
    }
    pub const fn group_base(&self) -> Option<u16> {
        self.group_base
    }
    pub fn grouping(&self) -> Option<&PivotCacheGrouping> {
        self.grouping.as_ref()
    }
    pub(crate) fn replace_grouping(&mut self, grouping: Option<PivotCacheGrouping>) {
        self.group_base = match &grouping {
            Some(PivotCacheGrouping::Discrete(value)) => Some(value.base_field_index),
            _ => None,
        };
        self.grouping = grouping;
    }
    pub(crate) fn set_group_parent(&mut self, parent: Option<u16>) {
        self.group_parent = parent;
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PivotCache {
    pub(crate) stream_id: u16,
    pub(crate) flags: u16,
    pub(crate) record_count: u32,
    pub(crate) fields: Vec<PivotCacheField>,
    pub(crate) rows: Vec<Vec<PivotCacheItem>>,
}
impl PivotCache {
    pub const fn stream_id(&self) -> u16 {
        self.stream_id
    }
    pub const fn flags(&self) -> u16 {
        self.flags
    }
    pub const fn record_count(&self) -> u32 {
        self.record_count
    }
    pub fn fields(&self) -> &[PivotCacheField] {
        &self.fields
    }
    pub fn rows(&self) -> &[Vec<PivotCacheItem>] {
        &self.rows
    }
    pub(crate) fn fields_mut(&mut self) -> &mut [PivotCacheField] {
        &mut self.fields
    }
}

/// Pivot field axis placement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAxis {
    /// No axis (hidden / unused).
    None,
    /// Row axis.
    Row,
    /// Column axis.
    Column,
    /// Page (filter) axis.
    Page,
    /// Data (values) axis.
    Data,
}

impl PivotAxis {
    pub(crate) fn from_u16(val: u16) -> Result<Self> {
        match val {
            0x0000 => Ok(Self::None),
            0x0001 => Ok(Self::Row),
            0x0002 => Ok(Self::Column),
            0x0004 => Ok(Self::Page),
            0x0008 => Ok(Self::Data),
            _ => Err(cache_invalid(
                SXVD_TYPE,
                format!("invalid PivotTable axis 0x{val:04X}"),
            )),
        }
    }
    pub const fn code(self) -> u16 {
        match self {
            Self::None => 0,
            Self::Row => 1,
            Self::Column => 2,
            Self::Page => 4,
            Self::Data => 8,
        }
    }
}

/// Aggregation function for data items.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotFunction {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevP,
    Var,
    VarP,
    Unknown(u16),
}

impl PivotFunction {
    pub(crate) fn from_u16(val: u16) -> Self {
        match val {
            0x0000 => Self::Sum,
            0x0001 => Self::Count,
            0x0002 => Self::Average,
            0x0003 => Self::Max,
            0x0004 => Self::Min,
            0x0005 => Self::Product,
            0x0006 => Self::CountNums,
            0x0007 => Self::StdDev,
            0x0008 => Self::StdDevP,
            0x0009 => Self::Var,
            0x000A => Self::VarP,
            other => Self::Unknown(other),
        }
    }
    pub const fn code(self) -> u16 {
        match self {
            Self::Sum => 0,
            Self::Count => 1,
            Self::Average => 2,
            Self::Max => 3,
            Self::Min => 4,
            Self::Product => 5,
            Self::CountNums => 6,
            Self::StdDev => 7,
            Self::StdDevP => 8,
            Self::Var => 9,
            Self::VarP => 10,
            Self::Unknown(value) => value,
        }
    }
}

// ---------------------------------------------------------------------------
// SXVIEW — View Definition
// ---------------------------------------------------------------------------

/// Parsed SXVIEW record (pivot table header / definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewDef {
    /// First row of the pivot table output range.
    pub first_row: u16,
    /// Last row of the pivot table output range.
    pub last_row: u16,
    /// First column of the pivot table output range.
    pub first_col: u16,
    /// Last column of the pivot table output range.
    pub last_col: u16,
    /// First header row.
    pub first_header_row: u16,
    /// First data row (body).
    pub first_data_row: u16,
    /// First data column (body).
    pub first_data_col: u16,
    /// Zero-based index into the workbook's global PivotCache list.
    pub cache_index: u16,
    /// Number of row fields.
    pub row_field_count: u16,
    /// Number of column fields.
    pub col_field_count: u16,
    /// Number of page (filter) fields.
    pub page_field_count: u16,
    /// Number of data (value) fields.
    pub data_field_count: u16,
    /// Total number of data rows in the source.
    pub data_row_count: u16,
    /// Total number of visible data columns.
    pub data_col_count: u16,
    /// Total number of fields (dimensions).
    pub field_count: u16,
    /// Axis used for the data field header (when >1 data field).
    pub data_axis: PivotAxis,
    /// Position of data field label within the axis.
    pub data_position: u16,
    /// SXVIEW option flags.
    pub flags: u16,
    /// Built-in PivotTable auto-format index.
    pub auto_format_index: u16,
    /// Name of the pivot table.
    pub name: String,
    /// Name of the data field header (e.g. "Values").
    pub data_field_name: String,
}

/// Parsed SXVD record (single pivot field definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewField {
    /// Axis this field is assigned to.
    pub axis: PivotAxis,
    /// Number of subtotals.
    pub subtotal_count: u16,
    /// Subtotal function bitmask.
    pub subtotal_flags: u16,
    /// Number of items in this field.
    pub item_count: u16,
    /// Optional field name override (empty string = use source name).
    pub name: Option<String>,
    /// SXVI records owned by this field.
    pub items: Vec<PivotViewItem>,
    /// Mandatory legacy extended-field properties.
    pub extension: Option<PivotViewFieldExtension>,
    /// Losslessly preserved future records scoped to this field.
    pub additional_extensions: Vec<PivotAdditionalExtension>,
}

/// Item type within a pivot field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotItemType {
    Data,
    Default,
    Sum,
    CountA,
    Average,
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
    Unknown(u16),
}

impl PivotItemType {
    pub(crate) fn from_u16(val: u16) -> Self {
        match val {
            0xFE => Self::Data,
            0xFF => Self::Default,
            0x00 => Self::Sum,
            0x01 => Self::CountA,
            0x02 => Self::Average,
            0x03 => Self::Max,
            0x04 => Self::Min,
            0x05 => Self::Product,
            0x06 => Self::Count,
            0x07 => Self::StdDev,
            0x08 => Self::StdDevP,
            0x09 => Self::Var,
            0x0A => Self::VarP,
            0x0B => Self::Grand,
            0x0C => Self::Blank,
            other => Self::Unknown(other),
        }
    }

    pub const fn code(self) -> u16 {
        match self {
            Self::Data => 0xFE,
            Self::Default => 0xFF,
            Self::Sum => 0x00,
            Self::CountA => 0x01,
            Self::Average => 0x02,
            Self::Max => 0x03,
            Self::Min => 0x04,
            Self::Product => 0x05,
            Self::Count => 0x06,
            Self::StdDev => 0x07,
            Self::StdDevP => 0x08,
            Self::Var => 0x09,
            Self::VarP => 0x0A,
            Self::Grand => 0x0B,
            Self::Blank => 0x0C,
            Self::Unknown(value) => value,
        }
    }
}

/// Parsed SXVI record (pivot field item).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewItem {
    /// Item type.
    pub item_type: PivotItemType,
    /// Option flags.
    pub flags: u16,
    /// Cache index.
    pub cache_index: u16,
    /// Optional item name override.
    pub name: Option<String>,
}

/// Parsed SXDI record (data/value field definition).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotDataItem {
    /// Index of the source field in the pivot cache.
    pub source_field_index: u16,
    /// Aggregation function.
    pub function: PivotFunction,
    /// Display format flags.
    pub display_format: u16,
    /// Index into SXVD for base field (used for "show values as").
    pub base_field_index: u16,
    /// Index into SXVI for base item.
    pub base_item_index: u16,
    /// Number format index.
    pub num_format_index: u16,
    /// Optional name override.
    pub name: String,
}

/// Pivot cache source type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotSourceType {
    /// Data from a worksheet range.
    Worksheet,
    /// Data from an external source.
    External,
    /// Consolidation ranges.
    Consolidation,
    /// Data from a named range / scenario.
    Scenario,
    /// Unknown source type.
    Unknown(u16),
}

impl PivotSourceType {
    pub(crate) fn from_u16(val: u16) -> Self {
        match val {
            0x0001 => Self::Worksheet,
            0x0002 => Self::External,
            0x0004 => Self::Consolidation,
            0x0010 => Self::Scenario,
            other => Self::Unknown(other),
        }
    }
}

/// A single page field entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageFieldEntry {
    /// Index into SXVI for the selected item.
    pub item_index: u16,
    /// Index into SXVD for the field.
    pub field_index: u16,
    /// Object ID (unused in most cases).
    pub object_id: u16,
}

/// Typed page-field selection derived from an SXPI selected-item ordinal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotPageSelection {
    Item(u16),
    All,
}

impl PageFieldEntry {
    pub const fn selection(self) -> PivotPageSelection {
        if self.item_index == 0x7FFD {
            PivotPageSelection::All
        } else {
            PivotPageSelection::Item(self.item_index)
        }
    }
}

/// One row or column axis entry from SXIVD.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAxisField {
    Field(u16),
    DataLayout,
}

/// Legacy extended properties attached to one SXVD field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewFieldExtension {
    pub flags: u32,
    pub auto_sort_data_index: Option<u16>,
    pub auto_show_data_index: Option<u16>,
    pub number_format_index: u16,
    pub subtotal_name: Option<String>,
    /// Reserved bytes retained exactly for roundtrip-aware readers.
    pub reserved: [u8; 8],
}

/// One visible layout line from an SXLI row/column line array.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotLayoutLine {
    pub repeated_item_count: u16,
    pub item_type: u16,
    pub custom_name_flags: u16,
    pub item_indices: Vec<u16>,
}

/// Extended PivotTable view properties from SXEX.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewExtension {
    pub format_count: u16,
    pub select_count: u16,
    pub page_rows: u16,
    pub page_cols: u16,
    pub flags: u32,
    pub error_string: Option<String>,
    pub null_string: Option<String>,
    pub tag: Option<String>,
    pub page_field_style: Option<String>,
    pub table_style: Option<String>,
    pub vacate_style: Option<String>,
}

/// Query/Pivot tag metadata from QsiSxTag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotQueryTag {
    pub table_type: u16,
    pub flags: u16,
    pub options: u32,
    pub last_refresh_version: u8,
    pub minimum_refresh_version: u8,
    pub first_created_version: u8,
    pub table_name: String,
    /// Bytes following the table name, retained exactly.
    pub trailing_payload: Vec<u8>,
}

/// Excel 9+ PivotTable layout metadata from SXVIEWEX9.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewEx9 {
    pub frt_flags: u16,
    pub report_flags: u32,
    pub view_flags: u32,
    pub auto_format_index: u16,
    pub grand_total_name: String,
}

/// Losslessly preserved SXADDL view- or field-extension record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotAdditionalExtension {
    pub class: u8,
    pub kind: u8,
    pub reserved: u16,
    pub payload: Vec<u8>,
}
