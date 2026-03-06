//! Pivot table record parsing for XLS BIFF8 files.
//!
//! Parses the family of SX* records that define pivot table structures:
//!
//! - **SXVIEW** (0x00B0): View definition — the main pivot table header.
//! - **SXVD** (0x00B1): View field — describes a single field (dimension).
//! - **SXVI** (0x00B2): View item — a single item within a field.
//! - **SXDI** (0x00C5): Data item — describes a data field (value area).
//! - **SXVS** (0x00E3): View source — source type of the pivot cache.
//! - **SXPI** (0x00B6): Page item — page field entries.
//!
//! # References
//!
//! - MS-XLS sections 2.4.271–2.4.283
//! - Apache POI `org.apache.poi.hssf.record.pivottable.*`

use crate::common::binary;
use crate::ole::xls::error::{XlsError, XlsResult};

// ---------------------------------------------------------------------------
// Record type constants
// ---------------------------------------------------------------------------

/// SXVIEW record type.
pub const SXVIEW_TYPE: u16 = 0x00B0;
/// SXVD (View Fields) record type.
pub const SXVD_TYPE: u16 = 0x00B1;
/// SXVI (View Item) record type.
pub const SXVI_TYPE: u16 = 0x00B2;
/// SXPI (Page Item) record type.
pub const SXPI_TYPE: u16 = 0x00B6;
/// SXDI (Data Item) record type.
pub const SXDI_TYPE: u16 = 0x00C5;
/// SXVS (View Source) record type.
pub const SXVS_TYPE: u16 = 0x00E3;

// ---------------------------------------------------------------------------
// Axis constants (used by SXVD and SXDI)
// ---------------------------------------------------------------------------

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
    fn from_u16(val: u16) -> Self {
        match val {
            0x0001 => Self::Row,
            0x0002 => Self::Column,
            0x0004 => Self::Page,
            0x0008 => Self::Data,
            _ => Self::None,
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
    fn from_u16(val: u16) -> Self {
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
}

// ---------------------------------------------------------------------------
// SXVIEW — View Definition
// ---------------------------------------------------------------------------

/// Parsed SXVIEW record (pivot table header / definition).
#[derive(Debug, Clone)]
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
    /// Total number of fields (dimensions).
    pub field_count: u16,
    /// Axis used for the data field header (when >1 data field).
    pub data_axis: PivotAxis,
    /// Position of data field label within the axis.
    pub data_position: u16,
    /// Name of the pivot table.
    pub name: String,
    /// Name of the data field header (e.g. "Values").
    pub data_field_name: String,
}

/// Parse an SXVIEW record.
///
/// Layout (Apache POI `ViewDefinitionRecord`):
/// ```text
///  0  u16  rwFirst
///  2  u16  rwLast
///  4  u16  colFirst
///  6  u16  colLast
///  8  u16  rwFirstHead
/// 10  u16  rwFirstData
/// 12  u16  colFirstData
/// 14  u16  cDimRw       (row field count)
/// 16  u16  cDimCol
/// 18  u16  cDimPg
/// 20  u16  cDimData
/// 22  u16  cRw          (data row count)
/// 24  u16  cDim         (total field count)
/// 26  u16  cItm         (unused)
/// 28  u16  cITMData     (unused)
/// 30  u16  sxaxis4Data
/// 32  u16  ipos4Data
/// 34  u16  cchName      (length of name)
/// 36  u16  cchData      (length of data field name)
/// 38  var  name (XLUnicodeStringNoCch)
///     var  dataField (XLUnicodeStringNoCch)
/// ```
pub fn parse_sxview(data: &[u8]) -> XlsResult<PivotViewDef> {
    if data.len() < 38 {
        return Err(XlsError::InvalidLength {
            expected: 38,
            found: data.len(),
        });
    }

    let first_row = binary::read_u16_le_at(data, 0)?;
    let last_row = binary::read_u16_le_at(data, 2)?;
    let first_col = binary::read_u16_le_at(data, 4)?;
    let last_col = binary::read_u16_le_at(data, 6)?;
    let first_header_row = binary::read_u16_le_at(data, 8)?;
    let first_data_row = binary::read_u16_le_at(data, 10)?;
    let first_data_col = binary::read_u16_le_at(data, 12)?;
    let row_field_count = binary::read_u16_le_at(data, 14)?;
    let col_field_count = binary::read_u16_le_at(data, 16)?;
    let page_field_count = binary::read_u16_le_at(data, 18)?;
    let data_field_count = binary::read_u16_le_at(data, 20)?;
    let data_row_count = binary::read_u16_le_at(data, 22)?;
    let field_count = binary::read_u16_le_at(data, 24)?;
    // skip cItm (26), cITMData (28)
    let data_axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 30)?);
    let data_position = binary::read_u16_le_at(data, 32)?;
    let cch_name = binary::read_u16_le_at(data, 34)? as usize;
    let cch_data = binary::read_u16_le_at(data, 36)? as usize;

    let mut offset = 38;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;
    let data_field_name = read_xl_string_no_cch(data, &mut offset, cch_data)?;

    Ok(PivotViewDef {
        first_row,
        last_row,
        first_col,
        last_col,
        first_header_row,
        first_data_row,
        first_data_col,
        row_field_count,
        col_field_count,
        page_field_count,
        data_field_count,
        data_row_count,
        field_count,
        data_axis,
        data_position,
        name,
        data_field_name,
    })
}

// ---------------------------------------------------------------------------
// SXVD — View Field
// ---------------------------------------------------------------------------

/// Parsed SXVD record (single pivot field definition).
#[derive(Debug, Clone)]
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
}

/// Parse an SXVD record.
///
/// Layout:
/// ```text
///  0  u16  sxaxis   (axis)
///  2  u16  cSub     (subtotal count)
///  4  u16  grbitSub (subtotal flags)
///  6  u16  cItm     (item count)
///  8  u16  cchName  (0xFFFF = not present)
/// 10  var  name (XLUnicodeStringNoCch)  — only if cchName != 0xFFFF
/// ```
pub fn parse_sxvd(data: &[u8]) -> XlsResult<PivotViewField> {
    if data.len() < 10 {
        return Err(XlsError::InvalidLength {
            expected: 10,
            found: data.len(),
        });
    }

    let axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 0)?);
    let subtotal_count = binary::read_u16_le_at(data, 2)?;
    let subtotal_flags = binary::read_u16_le_at(data, 4)?;
    let item_count = binary::read_u16_le_at(data, 6)?;
    let cch_name = binary::read_u16_le_at(data, 8)?;

    let name = if cch_name != 0xFFFF {
        let mut offset = 10;
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    } else {
        None
    };

    Ok(PivotViewField {
        axis,
        subtotal_count,
        subtotal_flags,
        item_count,
        name,
    })
}

// ---------------------------------------------------------------------------
// SXVI — View Item
// ---------------------------------------------------------------------------

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
    fn from_u16(val: u16) -> Self {
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
}

/// Parsed SXVI record (pivot field item).
#[derive(Debug, Clone)]
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

/// Parse an SXVI record.
///
/// Layout:
/// ```text
///  0  u16  itmType
///  2  u16  grbitItem
///  4  u16  iCache
///  6  u16  cchName  (0xFFFF = not present)
///  8  var  name
/// ```
pub fn parse_sxvi(data: &[u8]) -> XlsResult<PivotViewItem> {
    if data.len() < 8 {
        return Err(XlsError::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }

    let item_type = PivotItemType::from_u16(binary::read_u16_le_at(data, 0)?);
    let flags = binary::read_u16_le_at(data, 2)?;
    let cache_index = binary::read_u16_le_at(data, 4)?;
    let cch_name = binary::read_u16_le_at(data, 6)?;

    let name = if cch_name != 0xFFFF {
        let mut offset = 8;
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    } else {
        None
    };

    Ok(PivotViewItem {
        item_type,
        flags,
        cache_index,
        name,
    })
}

// ---------------------------------------------------------------------------
// SXDI — Data Item
// ---------------------------------------------------------------------------

/// Parsed SXDI record (data/value field definition).
#[derive(Debug, Clone)]
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

/// Parse an SXDI record.
///
/// Layout (POI `DataItemRecord`):
/// ```text
///  0  u16  isxvdData   (source field index)
///  2  u16  iiftab      (aggregation function)
///  4  u16  df          (display format)
///  6  u16  isxvd       (base field index)
///  8  u16  isxvi       (base item index)
/// 10  u16  ifmt        (number format)
/// 12  u16  cchName
/// 14  var  name
/// ```
pub fn parse_sxdi(data: &[u8]) -> XlsResult<PivotDataItem> {
    if data.len() < 14 {
        return Err(XlsError::InvalidLength {
            expected: 14,
            found: data.len(),
        });
    }

    let source_field_index = binary::read_u16_le_at(data, 0)?;
    let function = PivotFunction::from_u16(binary::read_u16_le_at(data, 2)?);
    let display_format = binary::read_u16_le_at(data, 4)?;
    let base_field_index = binary::read_u16_le_at(data, 6)?;
    let base_item_index = binary::read_u16_le_at(data, 8)?;
    let num_format_index = binary::read_u16_le_at(data, 10)?;
    let cch_name = binary::read_u16_le_at(data, 12)? as usize;

    let mut offset = 14;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;

    Ok(PivotDataItem {
        source_field_index,
        function,
        display_format,
        base_field_index,
        base_item_index,
        num_format_index,
        name,
    })
}

// ---------------------------------------------------------------------------
// SXVS — View Source
// ---------------------------------------------------------------------------

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
    fn from_u16(val: u16) -> Self {
        match val {
            0x0001 => Self::Worksheet,
            0x0002 => Self::External,
            0x0004 => Self::Consolidation,
            0x0010 => Self::Scenario,
            other => Self::Unknown(other),
        }
    }
}

/// Parse an SXVS record (2 bytes: source type).
pub fn parse_sxvs(data: &[u8]) -> XlsResult<PivotSourceType> {
    if data.len() < 2 {
        return Err(XlsError::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(PivotSourceType::from_u16(binary::read_u16_le_at(data, 0)?))
}

// ---------------------------------------------------------------------------
// SXPI — Page Item
// ---------------------------------------------------------------------------

/// A single page field entry.
#[derive(Debug, Clone, Copy)]
pub struct PageFieldEntry {
    /// Index into SXVI for the selected item.
    pub item_index: u16,
    /// Index into SXVD for the field.
    pub field_index: u16,
    /// Object ID (unused in most cases).
    pub object_id: u16,
}

/// Parse an SXPI record.
///
/// Each entry is 6 bytes: `(isxvi: u16, isxvd: u16, idObj: u16)`.
/// The number of entries is `data.len() / 6`.
pub fn parse_sxpi(data: &[u8]) -> XlsResult<Vec<PageFieldEntry>> {
    let entry_count = data.len() / 6;
    let mut entries = Vec::with_capacity(entry_count);

    for i in 0..entry_count {
        let offset = i * 6;
        if offset + 6 > data.len() {
            break;
        }
        entries.push(PageFieldEntry {
            item_index: binary::read_u16_le_at(data, offset)?,
            field_index: binary::read_u16_le_at(data, offset + 2)?,
            object_id: binary::read_u16_le_at(data, offset + 4)?,
        });
    }

    Ok(entries)
}

// ---------------------------------------------------------------------------
// Aggregate: PivotTable
// ---------------------------------------------------------------------------

/// Complete pivot table definition aggregated from multiple SX* records.
#[derive(Debug, Clone)]
pub struct PivotTable {
    /// View definition (SXVIEW).
    pub view: PivotViewDef,
    /// Source type (SXVS).
    pub source_type: PivotSourceType,
    /// Field definitions (SXVD records, in order).
    pub fields: Vec<PivotViewField>,
    /// All items across all fields (SXVI records, in order).
    pub items: Vec<PivotViewItem>,
    /// Data field definitions (SXDI records).
    pub data_items: Vec<PivotDataItem>,
    /// Page field entries (SXPI records).
    pub page_entries: Vec<PageFieldEntry>,
}

impl PivotTable {
    /// Create a new pivot table from its view definition.
    pub fn new(view: PivotViewDef) -> Self {
        Self {
            source_type: PivotSourceType::Worksheet,
            fields: Vec::with_capacity(view.field_count as usize),
            items: Vec::new(),
            data_items: Vec::with_capacity(view.data_field_count as usize),
            page_entries: Vec::with_capacity(view.page_field_count as usize),
            view,
        }
    }
}

// ---------------------------------------------------------------------------
// String helper
// ---------------------------------------------------------------------------

/// Read an XLUnicodeStringNoCch: 1-byte flags then `cch` chars.
fn read_xl_string_no_cch(data: &[u8], offset: &mut usize, cch: usize) -> XlsResult<String> {
    if cch == 0 {
        return Ok(String::new());
    }

    if *offset >= data.len() {
        return Err(XlsError::InvalidLength {
            expected: *offset + 1,
            found: data.len(),
        });
    }

    let flags = data[*offset];
    *offset += 1;
    let is_utf16 = flags & 0x01 != 0;

    if is_utf16 {
        let byte_len = cch * 2;
        if *offset + byte_len > data.len() {
            return Err(XlsError::InvalidLength {
                expected: *offset + byte_len,
                found: data.len(),
            });
        }
        let words: Vec<u16> = data[*offset..*offset + byte_len]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        *offset += byte_len;
        String::from_utf16(&words)
            .map_err(|e| XlsError::InvalidData(format!("Invalid UTF-16 in pivot string: {}", e)))
    } else {
        if *offset + cch > data.len() {
            return Err(XlsError::InvalidLength {
                expected: *offset + cch,
                found: data.len(),
            });
        }
        let s: String = data[*offset..*offset + cch]
            .iter()
            .map(|&b| b as char)
            .collect();
        *offset += cch;
        Ok(s)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_sxvs() {
        let data = 0x0001u16.to_le_bytes();
        assert_eq!(parse_sxvs(&data).unwrap(), PivotSourceType::Worksheet);
    }

    #[test]
    fn test_parse_sxpi_two_entries() {
        let mut data = Vec::new();
        // Entry 1
        data.extend_from_slice(&1u16.to_le_bytes()); // isxvi
        data.extend_from_slice(&0u16.to_le_bytes()); // isxvd
        data.extend_from_slice(&0u16.to_le_bytes()); // idObj
        // Entry 2
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());

        let entries = parse_sxpi(&data).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].item_index, 1);
        assert_eq!(entries[1].field_index, 1);
    }

    #[test]
    fn test_parse_sxvd_no_name() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0001u16.to_le_bytes()); // axis = Row
        data.extend_from_slice(&0u16.to_le_bytes()); // cSub
        data.extend_from_slice(&0u16.to_le_bytes()); // grbitSub
        data.extend_from_slice(&5u16.to_le_bytes()); // cItm
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cchName = not present

        let field = parse_sxvd(&data).unwrap();
        assert_eq!(field.axis, PivotAxis::Row);
        assert_eq!(field.item_count, 5);
        assert!(field.name.is_none());
    }

    #[test]
    fn test_parse_sxvi_data_item() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x00FEu16.to_le_bytes()); // itmType = Data
        data.extend_from_slice(&0u16.to_le_bytes()); // flags
        data.extend_from_slice(&3u16.to_le_bytes()); // iCache
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // no name

        let item = parse_sxvi(&data).unwrap();
        assert_eq!(item.item_type, PivotItemType::Data);
        assert_eq!(item.cache_index, 3);
        assert!(item.name.is_none());
    }
}
