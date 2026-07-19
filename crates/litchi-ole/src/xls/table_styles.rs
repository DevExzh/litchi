//! BIFF8 custom table and PivotTable style catalogs.

use std::collections::HashSet;

use super::differential_format::{DXF_RECORD_TYPE, validate_frt_header, write_frt_header};
use super::{XlsError, XlsResult};

pub(crate) const TABLE_STYLES_RECORD_TYPE: u16 = 0x088E;
pub(crate) const TABLE_STYLE_RECORD_TYPE: u16 = 0x088F;
pub(crate) const TABLE_STYLE_ELEMENT_RECORD_TYPE: u16 = 0x0890;
const STYLE_RECORD_TYPE: u16 = 0x0293;
const BUILT_IN_STYLE_COUNT: u32 = 144;
const TABLE_STYLES_FIXED_LEN: usize = 20;
const TABLE_STYLE_FIXED_LEN: usize = 20;
const TABLE_STYLE_ELEMENT_LEN: usize = 24;
const MAX_STYLE_NAME_UNITS: usize = 255;
const MAX_TABLE_STYLE_ELEMENTS: usize = 28;
const MAX_CUSTOM_TABLE_STYLES: usize = 65_536;
const MAX_DIFFERENTIAL_FORMATS: usize = 65_536;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn record_bytes(record_type: u16, payload: Vec<u8>) -> XlsResult<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .map_err(|_| invalid(record_type, "payload length exceeds BIFF u16"))?;
    let mut data = Vec::with_capacity(4 + payload.len());
    data.extend_from_slice(&record_type.to_le_bytes());
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(&payload);
    Ok(data)
}

/// A typed, zero-based reference into the workbook's global DXF table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct XlsDifferentialFormatId(u32);

impl XlsDifferentialFormatId {
    pub const fn new(index: u32) -> Self {
        Self(index)
    }

    pub const fn index(self) -> u32 {
        self.0
    }
}

/// One of the 28 regions that can be formatted by a BIFF8 table style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsTableStyleRegion {
    WholeTable,
    HeaderRow,
    TotalRow,
    FirstColumn,
    LastColumn,
    FirstRowStripe,
    SecondRowStripe,
    FirstColumnStripe,
    SecondColumnStripe,
    FirstHeaderCell,
    LastHeaderCell,
    FirstTotalCell,
    LastTotalCell,
    FirstSubtotalColumn,
    SecondSubtotalColumn,
    ThirdSubtotalColumn,
    FirstSubtotalRow,
    SecondSubtotalRow,
    ThirdSubtotalRow,
    BlankRow,
    FirstColumnSubheading,
    SecondColumnSubheading,
    ThirdColumnSubheading,
    FirstRowSubheading,
    SecondRowSubheading,
    ThirdRowSubheading,
    PageFieldLabels,
    PageFieldValues,
}

impl XlsTableStyleRegion {
    fn from_u32(value: u32) -> XlsResult<Self> {
        Ok(match value {
            0 => Self::WholeTable,
            1 => Self::HeaderRow,
            2 => Self::TotalRow,
            3 => Self::FirstColumn,
            4 => Self::LastColumn,
            5 => Self::FirstRowStripe,
            6 => Self::SecondRowStripe,
            7 => Self::FirstColumnStripe,
            8 => Self::SecondColumnStripe,
            9 => Self::FirstHeaderCell,
            10 => Self::LastHeaderCell,
            11 => Self::FirstTotalCell,
            12 => Self::LastTotalCell,
            13 => Self::FirstSubtotalColumn,
            14 => Self::SecondSubtotalColumn,
            15 => Self::ThirdSubtotalColumn,
            16 => Self::FirstSubtotalRow,
            17 => Self::SecondSubtotalRow,
            18 => Self::ThirdSubtotalRow,
            19 => Self::BlankRow,
            20 => Self::FirstColumnSubheading,
            21 => Self::SecondColumnSubheading,
            22 => Self::ThirdColumnSubheading,
            23 => Self::FirstRowSubheading,
            24 => Self::SecondRowSubheading,
            25 => Self::ThirdRowSubheading,
            26 => Self::PageFieldLabels,
            27 => Self::PageFieldValues,
            _ => {
                return Err(invalid(
                    TABLE_STYLE_ELEMENT_RECORD_TYPE,
                    format!("reserved table-style region {value}"),
                ));
            },
        })
    }

    const fn to_u32(self) -> u32 {
        match self {
            Self::WholeTable => 0,
            Self::HeaderRow => 1,
            Self::TotalRow => 2,
            Self::FirstColumn => 3,
            Self::LastColumn => 4,
            Self::FirstRowStripe => 5,
            Self::SecondRowStripe => 6,
            Self::FirstColumnStripe => 7,
            Self::SecondColumnStripe => 8,
            Self::FirstHeaderCell => 9,
            Self::LastHeaderCell => 10,
            Self::FirstTotalCell => 11,
            Self::LastTotalCell => 12,
            Self::FirstSubtotalColumn => 13,
            Self::SecondSubtotalColumn => 14,
            Self::ThirdSubtotalColumn => 15,
            Self::FirstSubtotalRow => 16,
            Self::SecondSubtotalRow => 17,
            Self::ThirdSubtotalRow => 18,
            Self::BlankRow => 19,
            Self::FirstColumnSubheading => 20,
            Self::SecondColumnSubheading => 21,
            Self::ThirdColumnSubheading => 22,
            Self::FirstRowSubheading => 23,
            Self::SecondRowSubheading => 24,
            Self::ThirdRowSubheading => 25,
            Self::PageFieldLabels => 26,
            Self::PageFieldValues => 27,
        }
    }

    pub const fn is_stripe(self) -> bool {
        matches!(
            self,
            Self::FirstRowStripe
                | Self::SecondRowStripe
                | Self::FirstColumnStripe
                | Self::SecondColumnStripe
        )
    }
}

/// One `TableStyleElement` and its reference to a global differential format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsTableStyleElement {
    region: XlsTableStyleRegion,
    size: u32,
    differential_format: XlsDifferentialFormatId,
}

impl XlsTableStyleElement {
    pub fn try_new(
        region: XlsTableStyleRegion,
        differential_format: XlsDifferentialFormatId,
    ) -> XlsResult<Self> {
        Self::try_with_size(region, 1, differential_format)
    }

    pub fn try_with_stripe_size(
        region: XlsTableStyleRegion,
        stripe_size: u8,
        differential_format: XlsDifferentialFormatId,
    ) -> XlsResult<Self> {
        if !region.is_stripe() {
            return Err(invalid(
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "stripe size can only be set for stripe regions",
            ));
        }
        Self::try_with_size(region, u32::from(stripe_size), differential_format)
    }

    fn try_with_size(
        region: XlsTableStyleRegion,
        size: u32,
        differential_format: XlsDifferentialFormatId,
    ) -> XlsResult<Self> {
        if !(1..=9).contains(&size) {
            return Err(invalid(
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                format!("table-style element size {size} is outside 1..=9"),
            ));
        }
        Ok(Self {
            region,
            size,
            differential_format,
        })
    }

    pub const fn region(&self) -> XlsTableStyleRegion {
        self.region
    }

    pub fn stripe_size(&self) -> Option<u8> {
        self.region.is_stripe().then_some(self.size as u8)
    }

    pub const fn differential_format(&self) -> XlsDifferentialFormatId {
        self.differential_format
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
        if data.len() != TABLE_STYLE_ELEMENT_LEN {
            return Err(invalid(
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                format!(
                    "TableStyleElement payload has {} bytes; expected {TABLE_STYLE_ELEMENT_LEN}",
                    data.len()
                ),
            ));
        }
        validate_frt_header(data, TABLE_STYLE_ELEMENT_RECORD_TYPE)?;
        Self::try_with_size(
            XlsTableStyleRegion::from_u32(read_u32(
                data,
                12,
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "TableStyleElement.tseType",
            )?)?,
            read_u32(
                data,
                16,
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "TableStyleElement.size",
            )?,
            XlsDifferentialFormatId::new(read_u32(
                data,
                20,
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "TableStyleElement.index",
            )?),
        )
    }

    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        Self::try_with_size(self.region, self.size, self.differential_format)?;
        let mut data = Vec::with_capacity(TABLE_STYLE_ELEMENT_LEN);
        write_frt_header(&mut data, TABLE_STYLE_ELEMENT_RECORD_TYPE);
        data.extend_from_slice(&self.region.to_u32().to_le_bytes());
        data.extend_from_slice(&self.size.to_le_bytes());
        data.extend_from_slice(&self.differential_format.index().to_le_bytes());
        Ok(data)
    }

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        record_bytes(TABLE_STYLE_ELEMENT_RECORD_TYPE, self.to_payload()?)
    }
}

/// A user-defined table style and its ordered style elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsTableStyle {
    name: String,
    available_for_tables: bool,
    available_for_pivot_tables: bool,
    elements: Vec<XlsTableStyleElement>,
    declared_element_count: u32,
}

impl XlsTableStyle {
    pub fn try_new(
        name: impl Into<String>,
        available_for_tables: bool,
        available_for_pivot_tables: bool,
        elements: Vec<XlsTableStyleElement>,
    ) -> XlsResult<Self> {
        let declared_element_count = u32::try_from(elements.len()).map_err(|_| {
            invalid(
                TABLE_STYLE_RECORD_TYPE,
                "table-style element count overflows",
            )
        })?;
        let value = Self {
            name: name.into(),
            available_for_tables,
            available_for_pivot_tables,
            elements,
            declared_element_count,
        };
        value.validate_complete()?;
        Ok(value)
    }

    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn is_available_for_tables(&self) -> bool {
        self.available_for_tables
    }
    pub const fn is_available_for_pivot_tables(&self) -> bool {
        self.available_for_pivot_tables
    }
    pub fn elements(&self) -> &[XlsTableStyleElement] {
        &self.elements
    }
    pub const fn declared_element_count(&self) -> u32 {
        self.declared_element_count
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
        if data.len() < TABLE_STYLE_FIXED_LEN
            || data.len() > TABLE_STYLE_FIXED_LEN + MAX_STYLE_NAME_UNITS * 2
        {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                format!("TableStyle payload has invalid length {}", data.len()),
            ));
        }
        validate_frt_header(data, TABLE_STYLE_RECORD_TYPE)?;
        let flags = read_u16(data, 12, TABLE_STYLE_RECORD_TYPE, "TableStyle flags")?;
        if flags & !0x0006 != 0 {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle reserved flag bits must be zero",
            ));
        }
        let declared_element_count =
            read_u32(data, 14, TABLE_STYLE_RECORD_TYPE, "TableStyle.ctse")?;
        if declared_element_count > MAX_TABLE_STYLE_ELEMENTS as u32 {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                format!("TableStyle declares {declared_element_count} elements; maximum is 28"),
            ));
        }
        let name_units = usize::from(read_u16(
            data,
            18,
            TABLE_STYLE_RECORD_TYPE,
            "TableStyle.cchName",
        )?);
        if !(1..=MAX_STYLE_NAME_UNITS).contains(&name_units)
            || data.len() != TABLE_STYLE_FIXED_LEN + name_units * 2
        {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle name length is invalid",
            ));
        }
        Ok(Self {
            name: decode_utf16(
                &data[TABLE_STYLE_FIXED_LEN..],
                TABLE_STYLE_RECORD_TYPE,
                "name",
            )?,
            available_for_tables: flags & 0x0004 != 0,
            available_for_pivot_tables: flags & 0x0002 != 0,
            elements: Vec::with_capacity(declared_element_count as usize),
            declared_element_count,
        })
    }

    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        self.validate_complete()?;
        let units = self.name.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::with_capacity(TABLE_STYLE_FIXED_LEN + units.len() * 2);
        write_frt_header(&mut data, TABLE_STYLE_RECORD_TYPE);
        let flags = (u16::from(self.available_for_pivot_tables) << 1)
            | (u16::from(self.available_for_tables) << 2);
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&(self.elements.len() as u32).to_le_bytes());
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        Ok(data)
    }

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        record_bytes(TABLE_STYLE_RECORD_TYPE, self.to_payload()?)
    }

    fn validate_header(&self) -> XlsResult<()> {
        let units = self.name.encode_utf16().count();
        if !(1..=MAX_STYLE_NAME_UNITS).contains(&units) {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle name must contain 1..=255 UTF-16 code units",
            ));
        }
        if self.declared_element_count > MAX_TABLE_STYLE_ELEMENTS as u32 {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle element count exceeds 28",
            ));
        }
        Ok(())
    }

    fn validate_complete(&self) -> XlsResult<()> {
        self.validate_header()?;
        if self.elements.len() != self.declared_element_count as usize {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                format!(
                    "TableStyle declares {} elements but contains {}",
                    self.declared_element_count,
                    self.elements.len()
                ),
            ));
        }
        let mut regions = HashSet::with_capacity(self.elements.len());
        for element in &self.elements {
            element.to_payload()?;
            if !regions.insert(element.region) {
                return Err(invalid(
                    TABLE_STYLE_ELEMENT_RECORD_TYPE,
                    "duplicate region in one TableStyle collection",
                ));
            }
        }
        Ok(())
    }
}

/// The `TableStyles` catalog header and all following custom style records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsTableStyles {
    total_style_count: u32,
    default_table_style: String,
    default_pivot_style: String,
    custom_styles: Vec<XlsTableStyle>,
}

impl XlsTableStyles {
    /// Constructs a catalog header. Use `try_with_custom_styles` for a complete custom catalog.
    pub fn try_new(
        total_style_count: u32,
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
    ) -> XlsResult<Self> {
        let value = Self {
            total_style_count,
            default_table_style: default_table_style.into(),
            default_pivot_style: default_pivot_style.into(),
            custom_styles: Vec::new(),
        };
        value.validate_header()?;
        Ok(value)
    }

    pub fn try_with_custom_styles(
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
        custom_styles: Vec<XlsTableStyle>,
    ) -> XlsResult<Self> {
        if custom_styles.len() > MAX_CUSTOM_TABLE_STYLES {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                format!("custom table-style count exceeds resource cap {MAX_CUSTOM_TABLE_STYLES}"),
            ));
        }
        let total_style_count = BUILT_IN_STYLE_COUNT
            .checked_add(custom_styles.len() as u32)
            .ok_or_else(|| invalid(TABLE_STYLES_RECORD_TYPE, "table-style count overflows"))?;
        let value = Self {
            total_style_count,
            default_table_style: default_table_style.into(),
            default_pivot_style: default_pivot_style.into(),
            custom_styles,
        };
        value.validate_complete(usize::MAX)?;
        Ok(value)
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
        let maximum = TABLE_STYLES_FIXED_LEN + MAX_STYLE_NAME_UNITS * 4;
        if !(TABLE_STYLES_FIXED_LEN..=maximum).contains(&data.len()) {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                format!("TableStyles payload has invalid length {}", data.len()),
            ));
        }
        validate_frt_header(data, TABLE_STYLES_RECORD_TYPE)?;
        let total_style_count = read_u32(data, 12, TABLE_STYLES_RECORD_TYPE, "TableStyles.cts")?;
        let table_units = usize::from(read_u16(
            data,
            16,
            TABLE_STYLES_RECORD_TYPE,
            "TableStyles.cchDefTableStyle",
        )?);
        let pivot_units = usize::from(read_u16(
            data,
            18,
            TABLE_STYLES_RECORD_TYPE,
            "TableStyles.cchDefPivotStyle",
        )?);
        if table_units > MAX_STYLE_NAME_UNITS || pivot_units > MAX_STYLE_NAME_UNITS {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                "TableStyles default name exceeds 255 UTF-16 code units",
            ));
        }
        let table_end = TABLE_STYLES_FIXED_LEN + table_units * 2;
        let pivot_end = table_end + pivot_units * 2;
        if pivot_end != data.len() {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                "TableStyles name lengths do not consume the payload exactly",
            ));
        }
        Self::try_new(
            total_style_count,
            decode_utf16(
                &data[TABLE_STYLES_FIXED_LEN..table_end],
                TABLE_STYLES_RECORD_TYPE,
                "default table style",
            )?,
            decode_utf16(
                &data[table_end..pivot_end],
                TABLE_STYLES_RECORD_TYPE,
                "default PivotTable style",
            )?,
        )
    }

    pub const fn total_style_count(&self) -> u32 {
        self.total_style_count
    }
    pub const fn built_in_style_count(&self) -> u32 {
        BUILT_IN_STYLE_COUNT
    }
    pub const fn custom_style_count(&self) -> u32 {
        self.total_style_count - BUILT_IN_STYLE_COUNT
    }
    pub const fn has_custom_styles(&self) -> bool {
        self.custom_style_count() != 0
    }
    pub fn default_table_style(&self) -> &str {
        &self.default_table_style
    }
    pub fn default_pivot_style(&self) -> &str {
        &self.default_pivot_style
    }
    pub fn custom_styles(&self) -> &[XlsTableStyle] {
        &self.custom_styles
    }

    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        self.validate_header()?;
        let table_units = self.default_table_style.encode_utf16().collect::<Vec<_>>();
        let pivot_units = self.default_pivot_style.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::with_capacity(
            TABLE_STYLES_FIXED_LEN + (table_units.len() + pivot_units.len()) * 2,
        );
        write_frt_header(&mut data, TABLE_STYLES_RECORD_TYPE);
        data.extend_from_slice(&self.total_style_count.to_le_bytes());
        data.extend_from_slice(&(table_units.len() as u16).to_le_bytes());
        data.extend_from_slice(&(pivot_units.len() as u16).to_le_bytes());
        data.extend(table_units.into_iter().flat_map(u16::to_le_bytes));
        data.extend(pivot_units.into_iter().flat_map(u16::to_le_bytes));
        Ok(data)
    }

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        record_bytes(TABLE_STYLES_RECORD_TYPE, self.to_payload()?)
    }

    /// Serializes the complete `TABLESTYLES` ABNF family in record order.
    pub fn to_family_record_bytes(&self, differential_format_count: usize) -> XlsResult<Vec<u8>> {
        self.validate_complete(differential_format_count)?;
        let mut data = self.to_record_bytes()?;
        for style in &self.custom_styles {
            data.extend_from_slice(&style.to_record_bytes()?);
            for element in &style.elements {
                data.extend_from_slice(&element.to_record_bytes()?);
            }
        }
        Ok(data)
    }

    fn validate_header(&self) -> XlsResult<()> {
        if self.total_style_count < BUILT_IN_STYLE_COUNT {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                "TableStyles cts must be at least 144",
            ));
        }
        for (name, field) in [
            (&self.default_table_style, "default table style"),
            (&self.default_pivot_style, "default PivotTable style"),
        ] {
            if name.encode_utf16().count() > MAX_STYLE_NAME_UNITS {
                return Err(invalid(
                    TABLE_STYLES_RECORD_TYPE,
                    format!("TableStyles {field} exceeds 255 UTF-16 code units"),
                ));
            }
        }
        if self.custom_style_count() as usize > MAX_CUSTOM_TABLE_STYLES {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                format!("custom table-style count exceeds resource cap {MAX_CUSTOM_TABLE_STYLES}"),
            ));
        }
        Ok(())
    }

    fn validate_complete(&self, differential_format_count: usize) -> XlsResult<()> {
        self.validate_header()?;
        if self.custom_styles.len() != self.custom_style_count() as usize {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                format!(
                    "TableStyles declares {} custom styles but contains {}",
                    self.custom_style_count(),
                    self.custom_styles.len()
                ),
            ));
        }
        let mut names = HashSet::with_capacity(self.custom_styles.len());
        for style in &self.custom_styles {
            style.validate_complete()?;
            if !names.insert(style.name.to_lowercase()) {
                return Err(invalid(
                    TABLE_STYLE_RECORD_TYPE,
                    format!("duplicate custom table-style name {:?}", style.name),
                ));
            }
            for element in &style.elements {
                if element.differential_format.index() as usize >= differential_format_count {
                    return Err(invalid(
                        TABLE_STYLE_ELEMENT_RECORD_TYPE,
                        format!(
                            "table-style element references DXF {} but only {differential_format_count} exist",
                            element.differential_format.index()
                        ),
                    ));
                }
            }
        }
        Ok(())
    }
}

fn decode_utf16(data: &[u8], record_type: u16, field: &str) -> XlsResult<String> {
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(record_type, format!("{field} contains invalid UTF-16")))
}

pub(crate) struct TableStylesCollector {
    header: Option<XlsTableStyles>,
    styles: Vec<XlsTableStyle>,
    current_style: Option<XlsTableStyle>,
    family_closed: bool,
    saw_style_record: bool,
    differential_format_count: usize,
}

impl TableStylesCollector {
    pub(crate) fn new() -> Self {
        Self {
            header: None,
            styles: Vec::new(),
            current_style: None,
            family_closed: false,
            saw_style_record: false,
            differential_format_count: 0,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if self
            .current_style
            .as_ref()
            .is_some_and(|style| style.elements.len() == style.declared_element_count as usize)
            && record_type != TABLE_STYLE_ELEMENT_RECORD_TYPE
        {
            self.finish_current_style()?;
        }

        match record_type {
            DXF_RECORD_TYPE => {
                if self.saw_style_record || self.header.is_some() {
                    return Err(invalid(
                        DXF_RECORD_TYPE,
                        "DXF record appears after the STYLES or TABLESTYLES collection",
                    ));
                }
                if self.differential_format_count == MAX_DIFFERENTIAL_FORMATS {
                    return Err(invalid(
                        DXF_RECORD_TYPE,
                        format!("DXF count exceeds resource cap {MAX_DIFFERENTIAL_FORMATS}"),
                    ));
                }
                self.differential_format_count += 1;
            },
            STYLE_RECORD_TYPE => {
                if self.header.is_some() {
                    self.close_family()?;
                }
                self.saw_style_record = true;
            },
            TABLE_STYLES_RECORD_TYPE => {
                if !self.saw_style_record {
                    return Err(invalid(
                        TABLE_STYLES_RECORD_TYPE,
                        "TableStyles appears before the required STYLES collection",
                    ));
                }
                if self.header.is_some() || self.family_closed {
                    return Err(invalid(
                        TABLE_STYLES_RECORD_TYPE,
                        "duplicate or noncontiguous TableStyles",
                    ));
                }
                self.header = Some(XlsTableStyles::parse_payload(data)?);
            },
            TABLE_STYLE_RECORD_TYPE => {
                if self.header.is_none() || self.family_closed {
                    return Err(invalid(
                        TABLE_STYLE_RECORD_TYPE,
                        "TableStyle appears outside TABLESTYLES",
                    ));
                }
                if self.current_style.is_some() {
                    return Err(invalid(
                        TABLE_STYLE_RECORD_TYPE,
                        "TableStyle appears before the prior element count is satisfied",
                    ));
                }
                if self.styles.len() == MAX_CUSTOM_TABLE_STYLES {
                    return Err(invalid(
                        TABLE_STYLE_RECORD_TYPE,
                        format!(
                            "custom style count exceeds resource cap {MAX_CUSTOM_TABLE_STYLES}"
                        ),
                    ));
                }
                let style = XlsTableStyle::parse_payload(data)?;
                if style.declared_element_count == 0 {
                    self.styles.push(style);
                } else {
                    self.current_style = Some(style);
                }
            },
            TABLE_STYLE_ELEMENT_RECORD_TYPE => {
                if self.family_closed {
                    return Err(invalid(
                        TABLE_STYLE_ELEMENT_RECORD_TYPE,
                        "TableStyleElement appears after TABLESTYLES ended",
                    ));
                }
                let style = self.current_style.as_mut().ok_or_else(|| {
                    invalid(
                        TABLE_STYLE_ELEMENT_RECORD_TYPE,
                        "TableStyleElement has no preceding TableStyle",
                    )
                })?;
                if style.elements.len() == style.declared_element_count as usize {
                    return Err(invalid(
                        TABLE_STYLE_ELEMENT_RECORD_TYPE,
                        "TableStyleElement exceeds the preceding ctse count",
                    ));
                }
                style
                    .elements
                    .push(XlsTableStyleElement::parse_payload(data)?);
            },
            _ if self.header.is_some() => self.close_family()?,
            _ => {},
        }
        Ok(())
    }

    fn finish_current_style(&mut self) -> XlsResult<()> {
        if let Some(style) = self.current_style.take() {
            style.validate_complete()?;
            self.styles.push(style);
        }
        Ok(())
    }

    fn close_family(&mut self) -> XlsResult<()> {
        if let Some(style) = &self.current_style {
            if style.elements.len() != style.declared_element_count as usize {
                return Err(invalid(
                    TABLE_STYLE_ELEMENT_RECORD_TYPE,
                    "TABLESTYLES ended before the declared element count was satisfied",
                ));
            }
        }
        self.finish_current_style()?;
        self.family_closed = true;
        Ok(())
    }

    pub(crate) fn finish(
        mut self,
        differential_format_count: usize,
    ) -> XlsResult<Option<XlsTableStyles>> {
        if self.header.is_none() {
            return Ok(None);
        }
        self.close_family()?;
        if self.differential_format_count != differential_format_count {
            return Err(invalid(
                DXF_RECORD_TYPE,
                "DXF collector count disagrees with formatting ownership",
            ));
        }
        let mut header = self.header.take().unwrap();
        header.custom_styles = self.styles;
        header.validate_complete(differential_format_count)?;
        Ok(Some(header))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_REFERENCE_HEX: &str = concat!(
        "8e08000000000000000000009000000011001100",
        "5400610062006c0065005300740079006c0065004d0065006400690075006d003900",
        "5000690076006f0074005300740079006c0065004c00690067006800740031003600",
    );

    fn decode_hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| {
                let digit = |byte: u8| match byte {
                    b'0'..=b'9' => byte - b'0',
                    b'a'..=b'f' => byte - b'a' + 10,
                    _ => panic!("invalid test hex"),
                };
                digit(pair[0]) << 4 | digit(pair[1])
            })
            .collect()
    }

    fn empty_style_record(name: &str, count: u32) -> Vec<u8> {
        let mut style = XlsTableStyle::try_new(name, true, false, vec![]).unwrap();
        style.declared_element_count = count;
        let units = name.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::new();
        write_frt_header(&mut data, TABLE_STYLE_RECORD_TYPE);
        data.extend_from_slice(&4u16.to_le_bytes());
        data.extend_from_slice(&count.to_le_bytes());
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        data
    }

    #[test]
    fn parses_and_round_trips_poi_reference_header() {
        let bytes = decode_hex(POI_REFERENCE_HEX);
        let styles = XlsTableStyles::parse_payload(&bytes).unwrap();
        assert_eq!(styles.total_style_count(), 144);
        assert_eq!(styles.custom_style_count(), 0);
        assert_eq!(styles.default_table_style(), "TableStyleMedium9");
        assert_eq!(styles.default_pivot_style(), "PivotStyleLight16");
        assert_eq!(styles.to_payload().unwrap(), bytes);
    }

    #[test]
    fn complete_custom_family_round_trips_all_record_types() {
        let element = XlsTableStyleElement::try_with_stripe_size(
            XlsTableStyleRegion::FirstRowStripe,
            2,
            XlsDifferentialFormatId::new(0),
        )
        .unwrap();
        let style =
            XlsTableStyle::try_new("Custom One", true, true, vec![element.clone()]).unwrap();
        let catalog = XlsTableStyles::try_with_custom_styles(
            "TableStyleMedium2",
            "PivotStyleLight16",
            vec![style.clone()],
        )
        .unwrap();
        let bytes = catalog.to_family_record_bytes(1).unwrap();
        assert!(bytes.windows(2).any(|value| value == [0x8F, 0x08]));
        assert!(bytes.windows(2).any(|value| value == [0x90, 0x08]));
        assert_eq!(
            XlsTableStyle::parse_payload(&style.to_payload().unwrap())
                .unwrap()
                .name(),
            "Custom One"
        );
        assert_eq!(
            XlsTableStyleElement::parse_payload(&element.to_payload().unwrap()).unwrap(),
            element
        );
    }

    #[test]
    fn collector_enforces_abnf_counts_and_references() {
        let header = XlsTableStyles::try_new(145, "TableStyleMedium2", "PivotStyleLight16")
            .unwrap()
            .to_payload()
            .unwrap();
        let dxf = crate::xls::differential_format::XlsDifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        let style = empty_style_record("Custom", 1);
        let element = XlsTableStyleElement::try_new(
            XlsTableStyleRegion::WholeTable,
            XlsDifferentialFormatId::new(0),
        )
        .unwrap()
        .to_payload()
        .unwrap();

        let mut collector = TableStylesCollector::new();
        collector.feed_record(DXF_RECORD_TYPE, &dxf).unwrap();
        collector.feed_record(STYLE_RECORD_TYPE, &[]).unwrap();
        collector
            .feed_record(TABLE_STYLES_RECORD_TYPE, &header)
            .unwrap();
        collector
            .feed_record(TABLE_STYLE_RECORD_TYPE, &style)
            .unwrap();
        collector
            .feed_record(TABLE_STYLE_ELEMENT_RECORD_TYPE, &element)
            .unwrap();
        let catalog = collector.finish(1).unwrap().unwrap();
        assert_eq!(catalog.custom_styles().len(), 1);

        let mut missing = TableStylesCollector::new();
        missing.feed_record(DXF_RECORD_TYPE, &dxf).unwrap();
        missing.feed_record(STYLE_RECORD_TYPE, &[]).unwrap();
        missing
            .feed_record(TABLE_STYLES_RECORD_TYPE, &header)
            .unwrap();
        missing
            .feed_record(TABLE_STYLE_RECORD_TYPE, &style)
            .unwrap();
        assert!(missing.feed_record(0x000A, &[]).is_err());

        let bad_reference = XlsTableStyleElement::try_new(
            XlsTableStyleRegion::WholeTable,
            XlsDifferentialFormatId::new(1),
        )
        .unwrap()
        .to_payload()
        .unwrap();
        let mut hostile = TableStylesCollector::new();
        hostile.feed_record(DXF_RECORD_TYPE, &dxf).unwrap();
        hostile.feed_record(STYLE_RECORD_TYPE, &[]).unwrap();
        hostile
            .feed_record(TABLE_STYLES_RECORD_TYPE, &header)
            .unwrap();
        hostile
            .feed_record(TABLE_STYLE_RECORD_TYPE, &style)
            .unwrap();
        hostile
            .feed_record(TABLE_STYLE_ELEMENT_RECORD_TYPE, &bad_reference)
            .unwrap();
        assert!(hostile.finish(1).is_err());
    }

    #[test]
    fn rejects_hostile_headers_lengths_regions_duplicates_and_caps() {
        let reference = decode_hex(POI_REFERENCE_HEX);
        assert!(XlsTableStyles::parse_payload(&reference[..19]).is_err());
        let mut bad = reference.clone();
        bad[0] = 0;
        assert!(XlsTableStyles::parse_payload(&bad).is_err());
        let mut bad = reference;
        bad[12..16].copy_from_slice(&143u32.to_le_bytes());
        assert!(XlsTableStyles::parse_payload(&bad).is_err());
        assert!(XlsTableStyles::try_new(144, "x".repeat(256), "").is_err());

        let element = XlsTableStyleElement::try_new(
            XlsTableStyleRegion::WholeTable,
            XlsDifferentialFormatId::new(0),
        )
        .unwrap();
        assert!(
            XlsTableStyle::try_new("dup", true, false, vec![element.clone(), element]).is_err()
        );
        assert!(
            XlsTableStyleElement::try_with_stripe_size(
                XlsTableStyleRegion::HeaderRow,
                2,
                XlsDifferentialFormatId::new(0),
            )
            .is_err()
        );
        let mut bad_element = XlsTableStyleElement::try_new(
            XlsTableStyleRegion::WholeTable,
            XlsDifferentialFormatId::new(0),
        )
        .unwrap()
        .to_payload()
        .unwrap();
        bad_element[12..16].copy_from_slice(&28u32.to_le_bytes());
        assert!(XlsTableStyleElement::parse_payload(&bad_element).is_err());
    }
}
