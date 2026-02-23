//! Pivot table support for XLSB

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};

/// Pivot field item
#[derive(Debug, Clone)]
pub struct PivotFieldItem {
    /// Item name
    pub name: String,
    /// Item type (0=data, 1=default, 2=sum, etc.)
    pub item_type: u8,
}

impl PivotFieldItem {
    pub fn new(name: String, item_type: u8) -> Self {
        PivotFieldItem { name, item_type }
    }
}

/// Pivot field
#[derive(Debug, Clone)]
pub struct PivotField {
    /// Field name
    pub name: String,
    /// Field items
    pub items: Vec<PivotFieldItem>,
    /// Show all items
    pub show_all: bool,
    /// Default subtotal
    pub default_subtotal: bool,
}

impl PivotField {
    pub fn new(name: String) -> Self {
        PivotField {
            name,
            items: Vec::new(),
            show_all: false,
            default_subtotal: true,
        }
    }

    pub fn add_item(&mut self, item: PivotFieldItem) {
        self.items.push(item);
    }
}

/// Pivot table location
#[derive(Debug, Clone)]
pub struct PivotLocation {
    /// First row (0-based)
    pub row_first: u32,
    /// Last row (0-based)
    pub row_last: u32,
    /// First column (0-based)
    pub col_first: u32,
    /// Last column (0-based)
    pub col_last: u32,
}

impl PivotLocation {
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        PivotLocation {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        Ok(PivotLocation {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        })
    }

    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(16);
        data.extend_from_slice(&self.row_first.to_le_bytes());
        data.extend_from_slice(&self.row_last.to_le_bytes());
        data.extend_from_slice(&self.col_first.to_le_bytes());
        data.extend_from_slice(&self.col_last.to_le_bytes());
        data
    }
}

/// Pivot cache source
#[derive(Debug, Clone)]
pub struct PivotCacheSource {
    /// Source type (0=worksheet, 1=external, 2=consolidation, etc.)
    pub source_type: u8,
    /// Worksheet source range (for worksheet type)
    pub worksheet_source: Option<String>,
    /// Connection ID (for external type)
    pub connection_id: Option<u32>,
}

impl PivotCacheSource {
    pub fn worksheet(range: String) -> Self {
        PivotCacheSource {
            source_type: 0,
            worksheet_source: Some(range),
            connection_id: None,
        }
    }

    pub fn external(connection_id: u32) -> Self {
        PivotCacheSource {
            source_type: 1,
            worksheet_source: None,
            connection_id: Some(connection_id),
        }
    }
}

/// Pivot cache definition
#[derive(Debug, Clone)]
pub struct PivotCacheDefinition {
    /// Cache ID
    pub cache_id: u32,
    /// Source
    pub source: PivotCacheSource,
    /// Fields
    pub fields: Vec<PivotField>,
    /// Record count
    pub record_count: u32,
}

impl PivotCacheDefinition {
    pub fn new(cache_id: u32, source: PivotCacheSource) -> Self {
        PivotCacheDefinition {
            cache_id,
            source,
            fields: Vec::new(),
            record_count: 0,
        }
    }

    pub fn add_field(&mut self, field: PivotField) {
        self.fields.push(field);
    }
}

/// Row/column field reference
#[derive(Debug, Clone)]
pub struct PivotFieldReference {
    /// Field index
    pub field_index: u32,
}

impl PivotFieldReference {
    pub fn new(field_index: u32) -> Self {
        PivotFieldReference { field_index }
    }

    pub fn serialize(&self) -> Vec<u8> {
        self.field_index.to_le_bytes().to_vec()
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        Ok(PivotFieldReference {
            field_index: binary::read_u32_le_at(data, 0)?,
        })
    }
}

/// Data field (value field)
#[derive(Debug, Clone)]
pub struct PivotDataField {
    /// Field index
    pub field_index: u32,
    /// Subtotal function (0=average, 1=count, 2=count nums, 3=max, 4=min, 5=product, 6=stdev, 7=stdevp, 8=sum, 9=var, 10=varp)
    pub subtotal: u8,
    /// Custom name
    pub name: Option<String>,
}

impl PivotDataField {
    pub fn new(field_index: u32, subtotal: u8) -> Self {
        PivotDataField {
            field_index,
            subtotal,
            name: None,
        }
    }

    pub fn with_name(mut self, name: String) -> Self {
        self.name = Some(name);
        self
    }
}

/// Pivot table definition
#[derive(Debug, Clone)]
pub struct PivotTable {
    /// Table name
    pub name: String,
    /// Cache ID
    pub cache_id: u32,
    /// Location on worksheet
    pub location: PivotLocation,
    /// Row fields
    pub row_fields: Vec<PivotFieldReference>,
    /// Column fields
    pub column_fields: Vec<PivotFieldReference>,
    /// Data fields
    pub data_fields: Vec<PivotDataField>,
    /// Page fields (filters)
    pub page_fields: Vec<PivotFieldReference>,
    /// Show grand totals for rows
    pub row_grand_totals: bool,
    /// Show grand totals for columns
    pub col_grand_totals: bool,
}

impl PivotTable {
    pub fn new(name: String, cache_id: u32, location: PivotLocation) -> Self {
        PivotTable {
            name,
            cache_id,
            location,
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            data_fields: Vec::new(),
            page_fields: Vec::new(),
            row_grand_totals: true,
            col_grand_totals: true,
        }
    }

    pub fn add_row_field(&mut self, field: PivotFieldReference) {
        self.row_fields.push(field);
    }

    pub fn add_column_field(&mut self, field: PivotFieldReference) {
        self.column_fields.push(field);
    }

    pub fn add_data_field(&mut self, field: PivotDataField) {
        self.data_fields.push(field);
    }

    pub fn add_page_field(&mut self, field: PivotFieldReference) {
        self.page_fields.push(field);
    }

    pub fn parse_header(data: &[u8]) -> XlsbResult<(String, u32)> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let cache_id = binary::read_u32_le_at(data, 0)?;
        let offset = 4;

        // Read name string
        let (name, _) = read_optional_string(&data[offset..])?;
        let name = name.unwrap_or_default();

        Ok((name, cache_id))
    }

    pub fn serialize_header(&self) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&self.cache_id.to_le_bytes());
        write_optional_string(&mut data, Some(&self.name));
        data
    }
}

/// Read optional string from XLSB data
fn read_optional_string(data: &[u8]) -> XlsbResult<(Option<String>, usize)> {
    if data.len() < 4 {
        return Ok((None, 0));
    }

    let len = binary::read_u32_le_at(data, 0)? as usize;
    if len == 0 {
        return Ok((None, 4));
    }

    if data.len() < 4 + len * 2 {
        return Ok((None, 4));
    }

    let mut chars = Vec::with_capacity(len);
    for i in 0..len {
        let ch = binary::read_u16_le_at(data, 4 + i * 2)?;
        chars.push(ch);
    }

    let string = String::from_utf16_lossy(&chars);
    Ok((Some(string), 4 + len * 2))
}

/// Write optional string to XLSB data
fn write_optional_string(data: &mut Vec<u8>, s: Option<&str>) {
    if let Some(s) = s {
        let chars: Vec<u16> = s.encode_utf16().collect();
        data.extend_from_slice(&(chars.len() as u32).to_le_bytes());
        for &ch in &chars {
            data.extend_from_slice(&ch.to_le_bytes());
        }
    } else {
        data.extend_from_slice(&0u32.to_le_bytes());
    }
}
