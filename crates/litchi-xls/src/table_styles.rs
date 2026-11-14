//! BIFF8 custom table and PivotTable style catalogs.

use std::collections::HashSet;

use super::differential_format::{DXF_RECORD_TYPE, validate_frt_header, write_frt_header};
use super::{Error, Result};

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

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn record_bytes(record_type: u16, payload: Vec<u8>) -> Result<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .map_err(|_error| invalid(record_type, "payload length exceeds BIFF u16"))?;
    let mut data = Vec::with_capacity(4 + payload.len());
    data.extend_from_slice(&record_type.to_le_bytes());
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(&payload);
    Ok(data)
}

/// A typed, zero-based reference into the workbook's global DXF table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct DifferentialFormatId(u32);

impl DifferentialFormatId {
    #[must_use]
    pub const fn new(index: u32) -> Self {
        Self(index)
    }

    #[must_use]
    pub const fn index(self) -> u32 {
        self.0
    }
}

/// One of the 28 regions that can be formatted by a BIFF8 table style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableStyleRegion {
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

impl TableStyleRegion {
    fn from_u32(value: u32) -> Result<Self> {
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

    #[must_use]
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
pub struct TableStyleElement {
    region: TableStyleRegion,
    size: u32,
    differential_format: DifferentialFormatId,
}

impl TableStyleElement {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(
        region: TableStyleRegion,
        differential_format: DifferentialFormatId,
    ) -> Result<Self> {
        Self::try_with_size(region, 1, differential_format)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_with_stripe_size(
        region: TableStyleRegion,
        stripe_size: u8,
        differential_format: DifferentialFormatId,
    ) -> Result<Self> {
        if !region.is_stripe() {
            return Err(invalid(
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "stripe size can only be set for stripe regions",
            ));
        }
        Self::try_with_size(region, u32::from(stripe_size), differential_format)
    }

    fn try_with_size(
        region: TableStyleRegion,
        size: u32,
        differential_format: DifferentialFormatId,
    ) -> Result<Self> {
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

    #[must_use]
    pub const fn region(&self) -> TableStyleRegion {
        self.region
    }

    #[must_use]
    pub fn stripe_size(&self) -> Option<u8> {
        self.region
            .is_stripe()
            .then_some(crate::utils::truncate_u32_to_u8(self.size))
    }

    #[must_use]
    pub const fn differential_format(&self) -> DifferentialFormatId {
        self.differential_format
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
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
            TableStyleRegion::from_u32(read_u32(
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
            DifferentialFormatId::new(read_u32(
                data,
                20,
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "TableStyleElement.index",
            )?),
        )
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        Self::try_with_size(self.region, self.size, self.differential_format)?;
        let mut data = Vec::with_capacity(TABLE_STYLE_ELEMENT_LEN);
        write_frt_header(&mut data, TABLE_STYLE_ELEMENT_RECORD_TYPE);
        data.extend_from_slice(&self.region.to_u32().to_le_bytes());
        data.extend_from_slice(&self.size.to_le_bytes());
        data.extend_from_slice(&self.differential_format.index().to_le_bytes());
        Ok(data)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes(TABLE_STYLE_ELEMENT_RECORD_TYPE, self.to_payload()?)
    }
}

/// A user-defined table style and its ordered style elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyle {
    name: String,
    available_for_tables: bool,
    available_for_pivot_tables: bool,
    elements: Vec<TableStyleElement>,
    declared_element_count: u32,
}

impl TableStyle {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(
        name: impl Into<String>,
        available_for_tables: bool,
        available_for_pivot_tables: bool,
        elements: Vec<TableStyleElement>,
    ) -> Result<Self> {
        let declared_element_count = u32::try_from(elements.len()).map_err(|_error| {
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

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub const fn is_available_for_tables(&self) -> bool {
        self.available_for_tables
    }
    #[must_use]
    pub const fn is_available_for_pivot_tables(&self) -> bool {
        self.available_for_pivot_tables
    }
    #[must_use]
    pub fn elements(&self) -> &[TableStyleElement] {
        &self.elements
    }
    #[must_use]
    pub const fn declared_element_count(&self) -> u32 {
        self.declared_element_count
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
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
        if declared_element_count > crate::utils::truncate_usize_to_u32(MAX_TABLE_STYLE_ELEMENTS) {
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

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        self.validate_complete()?;
        let units = self.name.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::with_capacity(TABLE_STYLE_FIXED_LEN + units.len() * 2);
        write_frt_header(&mut data, TABLE_STYLE_RECORD_TYPE);
        let flags = (u16::from(self.available_for_pivot_tables) << 1)
            | (u16::from(self.available_for_tables) << 2);
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(
            &crate::utils::truncate_usize_to_u32(self.elements.len()).to_le_bytes(),
        );
        data.extend_from_slice(&crate::utils::truncate_usize_to_u16(units.len()).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        Ok(data)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes(TABLE_STYLE_RECORD_TYPE, self.to_payload()?)
    }

    fn validate_header(&self) -> Result<()> {
        let units = self.name.encode_utf16().count();
        if !(1..=MAX_STYLE_NAME_UNITS).contains(&units) {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle name must contain 1..=255 UTF-16 code units",
            ));
        }
        if self.declared_element_count
            > crate::utils::truncate_usize_to_u32(MAX_TABLE_STYLE_ELEMENTS)
        {
            return Err(invalid(
                TABLE_STYLE_RECORD_TYPE,
                "TableStyle element count exceeds 28",
            ));
        }
        Ok(())
    }

    fn validate_complete(&self) -> Result<()> {
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
pub struct TableStyles {
    total_style_count: u32,
    default_table_style: String,
    default_pivot_style: String,
    custom_styles: Vec<TableStyle>,
}

impl TableStyles {
    /// Constructs a catalog header. Use `try_with_custom_styles` for a complete custom catalog.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(
        total_style_count: u32,
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            total_style_count,
            default_table_style: default_table_style.into(),
            default_pivot_style: default_pivot_style.into(),
            custom_styles: Vec::new(),
        };
        value.validate_header()?;
        Ok(value)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_with_custom_styles(
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
        custom_styles: Vec<TableStyle>,
    ) -> Result<Self> {
        if custom_styles.len() > MAX_CUSTOM_TABLE_STYLES {
            return Err(invalid(
                TABLE_STYLES_RECORD_TYPE,
                format!("custom table-style count exceeds resource cap {MAX_CUSTOM_TABLE_STYLES}"),
            ));
        }
        let total_style_count = BUILT_IN_STYLE_COUNT
            .checked_add(crate::utils::truncate_usize_to_u32(custom_styles.len()))
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

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
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

    #[must_use]
    pub const fn total_style_count(&self) -> u32 {
        self.total_style_count
    }
    #[must_use]
    pub const fn built_in_style_count(&self) -> u32 {
        BUILT_IN_STYLE_COUNT
    }
    #[must_use]
    pub const fn custom_style_count(&self) -> u32 {
        self.total_style_count - BUILT_IN_STYLE_COUNT
    }
    #[must_use]
    pub const fn has_custom_styles(&self) -> bool {
        self.custom_style_count() != 0
    }
    #[must_use]
    pub fn default_table_style(&self) -> &str {
        &self.default_table_style
    }
    #[must_use]
    pub fn default_pivot_style(&self) -> &str {
        &self.default_pivot_style
    }
    #[must_use]
    pub fn custom_styles(&self) -> &[TableStyle] {
        &self.custom_styles
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        self.validate_header()?;
        let table_units = self.default_table_style.encode_utf16().collect::<Vec<_>>();
        let pivot_units = self.default_pivot_style.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::with_capacity(
            TABLE_STYLES_FIXED_LEN + (table_units.len() + pivot_units.len()) * 2,
        );
        write_frt_header(&mut data, TABLE_STYLES_RECORD_TYPE);
        data.extend_from_slice(&self.total_style_count.to_le_bytes());
        data.extend_from_slice(
            &crate::utils::truncate_usize_to_u16(table_units.len()).to_le_bytes(),
        );
        data.extend_from_slice(
            &crate::utils::truncate_usize_to_u16(pivot_units.len()).to_le_bytes(),
        );
        data.extend(table_units.into_iter().flat_map(u16::to_le_bytes));
        data.extend(pivot_units.into_iter().flat_map(u16::to_le_bytes));
        Ok(data)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes(TABLE_STYLES_RECORD_TYPE, self.to_payload()?)
    }

    /// Serializes the complete `TABLESTYLES` ABNF family in record order.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_family_record_bytes(&self, differential_format_count: usize) -> Result<Vec<u8>> {
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

    fn validate_header(&self) -> Result<()> {
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

    fn validate_complete(&self, differential_format_count: usize) -> Result<()> {
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

fn decode_utf16(data: &[u8], record_type: u16, field: &str) -> Result<String> {
    let units = data
        .as_chunks::<2>()
        .0
        .iter()
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_error| invalid(record_type, format!("{field} contains invalid UTF-16")))
}

/// Where the workbook globals stand relative to the DXF records, the STYLES
/// collection and the TABLESTYLES collection that follows it.
///
/// ```text
/// BeforeStyles --STYLE--> AfterStyles --STYLE--> AfterStyles
/// AfterStyles  --TableStyles--> Catalog { ended: false }
/// Catalog { ended: false } --TableStyle | TableStyleElement--> Catalog { ended: false }
/// Catalog { .. } --STYLE or any record outside the family--> Catalog { ended: true }
/// ```
///
/// DXF records are accepted only in `BeforeStyles`. TableStyles is refused
/// in `BeforeStyles` (no STYLES yet) and in `Catalog` (duplicate). TableStyle
/// is refused outside `Catalog { ended: false }`, TableStyleElement in
/// `Catalog { ended: true }`. Ending the catalog first completes the open
/// TableStyle, which fails when its declared element count is not met.
enum TableStylesBlockState {
    BeforeStyles,
    AfterStyles,
    Catalog { header: TableStyles, ended: bool },
}

pub(crate) struct TableStylesCollector {
    state: TableStylesBlockState,
    styles: Vec<TableStyle>,
    current_style: Option<TableStyle>,
    differential_format_count: usize,
}

impl TableStylesCollector {
    pub(crate) fn new() -> Self {
        Self {
            state: TableStylesBlockState::BeforeStyles,
            styles: Vec::new(),
            current_style: None,
            differential_format_count: 0,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
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
                if !matches!(self.state, TableStylesBlockState::BeforeStyles) {
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
            STYLE_RECORD_TYPE => match self.state {
                TableStylesBlockState::BeforeStyles | TableStylesBlockState::AfterStyles => {
                    self.state = TableStylesBlockState::AfterStyles;
                },
                TableStylesBlockState::Catalog { .. } => self.close_family()?,
            },
            TABLE_STYLES_RECORD_TYPE => match self.state {
                TableStylesBlockState::BeforeStyles => {
                    return Err(invalid(
                        TABLE_STYLES_RECORD_TYPE,
                        "TableStyles appears before the required STYLES collection",
                    ));
                },
                TableStylesBlockState::Catalog { .. } => {
                    return Err(invalid(
                        TABLE_STYLES_RECORD_TYPE,
                        "duplicate or noncontiguous TableStyles",
                    ));
                },
                TableStylesBlockState::AfterStyles => {
                    self.state = TableStylesBlockState::Catalog {
                        header: TableStyles::parse_payload(data)?,
                        ended: false,
                    };
                },
            },
            TABLE_STYLE_RECORD_TYPE => {
                if !matches!(
                    self.state,
                    TableStylesBlockState::Catalog { ended: false, .. }
                ) {
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
                let style = TableStyle::parse_payload(data)?;
                if style.declared_element_count == 0 {
                    self.styles.push(style);
                } else {
                    self.current_style = Some(style);
                }
            },
            TABLE_STYLE_ELEMENT_RECORD_TYPE => {
                if matches!(
                    self.state,
                    TableStylesBlockState::Catalog { ended: true, .. }
                ) {
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
                style.elements.push(TableStyleElement::parse_payload(data)?);
            },
            _ if matches!(self.state, TableStylesBlockState::Catalog { .. }) => {
                self.close_family()?;
            },
            _ => {},
        }
        Ok(())
    }

    fn finish_current_style(&mut self) -> Result<()> {
        if let Some(style) = self.current_style.take() {
            style.validate_complete()?;
            self.styles.push(style);
        }
        Ok(())
    }

    /// Completes the open TableStyle, if any, at the end of the catalog.
    fn end_current_style(&mut self) -> Result<()> {
        if let Some(style) = &self.current_style
            && style.elements.len() != style.declared_element_count as usize
        {
            return Err(invalid(
                TABLE_STYLE_ELEMENT_RECORD_TYPE,
                "TABLESTYLES ended before the declared element count was satisfied",
            ));
        }
        self.finish_current_style()
    }

    fn close_family(&mut self) -> Result<()> {
        self.end_current_style()?;
        match &mut self.state {
            TableStylesBlockState::BeforeStyles | TableStylesBlockState::AfterStyles => {},
            TableStylesBlockState::Catalog { ended, .. } => *ended = true,
        }
        Ok(())
    }

    pub(crate) fn finish(
        mut self,
        differential_format_count: usize,
    ) -> Result<Option<TableStyles>> {
        let mut header =
            match std::mem::replace(&mut self.state, TableStylesBlockState::BeforeStyles) {
                TableStylesBlockState::BeforeStyles | TableStylesBlockState::AfterStyles => {
                    return Ok(None);
                },
                TableStylesBlockState::Catalog { header, .. } => header,
            };
        self.end_current_style()?;
        if self.differential_format_count != differential_format_count {
            return Err(invalid(
                DXF_RECORD_TYPE,
                "DXF collector count disagrees with formatting ownership",
            ));
        }
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
            .as_chunks::<2>()
            .0
            .iter()
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
        let mut style = TableStyle::try_new(name, true, false, vec![]).unwrap();
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
        let styles = TableStyles::parse_payload(&bytes).unwrap();
        assert_eq!(styles.total_style_count(), 144);
        assert_eq!(styles.custom_style_count(), 0);
        assert_eq!(styles.default_table_style(), "TableStyleMedium9");
        assert_eq!(styles.default_pivot_style(), "PivotStyleLight16");
        assert_eq!(styles.to_payload().unwrap(), bytes);
    }

    #[test]
    fn complete_custom_family_round_trips_all_record_types() {
        let element = TableStyleElement::try_with_stripe_size(
            TableStyleRegion::FirstRowStripe,
            2,
            DifferentialFormatId::new(0),
        )
        .unwrap();
        let style = TableStyle::try_new("Custom One", true, true, vec![element.clone()]).unwrap();
        let catalog = TableStyles::try_with_custom_styles(
            "TableStyleMedium2",
            "PivotStyleLight16",
            vec![style.clone()],
        )
        .unwrap();
        let bytes = catalog.to_family_record_bytes(1).unwrap();
        assert!(bytes.windows(2).any(|value| value == [0x8F, 0x08]));
        assert!(bytes.windows(2).any(|value| value == [0x90, 0x08]));
        assert_eq!(
            TableStyle::parse_payload(&style.to_payload().unwrap())
                .unwrap()
                .name(),
            "Custom One"
        );
        assert_eq!(
            TableStyleElement::parse_payload(&element.to_payload().unwrap()).unwrap(),
            element
        );
    }

    #[test]
    fn collector_enforces_abnf_counts_and_references() {
        let header = TableStyles::try_new(145, "TableStyleMedium2", "PivotStyleLight16")
            .unwrap()
            .to_payload()
            .unwrap();
        let dxf = crate::differential_format::DifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        let style = empty_style_record("Custom", 1);
        let element =
            TableStyleElement::try_new(TableStyleRegion::WholeTable, DifferentialFormatId::new(0))
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

        let bad_reference =
            TableStyleElement::try_new(TableStyleRegion::WholeTable, DifferentialFormatId::new(1))
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
        assert!(TableStyles::parse_payload(&reference[..19]).is_err());
        let mut bad = reference.clone();
        bad[0] = 0;
        assert!(TableStyles::parse_payload(&bad).is_err());
        let mut bad = reference;
        bad[12..16].copy_from_slice(&143u32.to_le_bytes());
        assert!(TableStyles::parse_payload(&bad).is_err());
        assert!(TableStyles::try_new(144, "x".repeat(256), "").is_err());

        let element =
            TableStyleElement::try_new(TableStyleRegion::WholeTable, DifferentialFormatId::new(0))
                .unwrap();
        assert!(TableStyle::try_new("dup", true, false, vec![element.clone(), element]).is_err());
        assert!(
            TableStyleElement::try_with_stripe_size(
                TableStyleRegion::HeaderRow,
                2,
                DifferentialFormatId::new(0),
            )
            .is_err()
        );
        let mut bad_element =
            TableStyleElement::try_new(TableStyleRegion::WholeTable, DifferentialFormatId::new(0))
                .unwrap()
                .to_payload()
                .unwrap();
        bad_element[12..16].copy_from_slice(&28u32.to_le_bytes());
        assert!(TableStyleElement::parse_payload(&bad_element).is_err());
    }
}

#[cfg(test)]
mod collection_order_tests {
    use super::*;

    fn style_record(name: &str, count: u32) -> Vec<u8> {
        let units = name.encode_utf16().collect::<Vec<_>>();
        let mut data = Vec::new();
        write_frt_header(&mut data, TABLE_STYLE_RECORD_TYPE);
        data.extend_from_slice(&4u16.to_le_bytes());
        data.extend_from_slice(&count.to_le_bytes());
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        data
    }

    /// Feeds one record per letter until the first error: `X` a DXF, `S` a
    /// STYLE, `H` the TableStyles header, `T` a TableStyle declaring one
    /// element, `U` a TableStyle declaring none, `E` a TableStyleElement and
    /// `O` an unrelated record. The answer lists `.` per accepted record, then
    /// the error text, or what `finish` makes of the records (with the DXF
    /// count the stream declared).
    fn outcome(letters: &str) -> String {
        let header = TableStyles::try_new(145, "TableStyleMedium2", "PivotStyleLight16")
            .unwrap()
            .to_payload()
            .unwrap();
        let dxf = crate::differential_format::DifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        let element =
            TableStyleElement::try_new(TableStyleRegion::WholeTable, DifferentialFormatId::new(0))
                .unwrap()
                .to_payload()
                .unwrap();
        let mut collector = TableStylesCollector::new();
        let mut differential_formats = 0;
        let mut styles = 0;
        let mut answers = String::new();
        for letter in letters.chars() {
            let (record_type, data) = match letter {
                'X' => {
                    differential_formats += 1;
                    (DXF_RECORD_TYPE, dxf.clone())
                },
                'S' => (STYLE_RECORD_TYPE, Vec::new()),
                'H' => (TABLE_STYLES_RECORD_TYPE, header.clone()),
                'T' | 'U' => {
                    styles += 1;
                    let count = u32::from(letter == 'T');
                    (
                        TABLE_STYLE_RECORD_TYPE,
                        style_record(&format!("Custom{styles}"), count),
                    )
                },
                'E' => (TABLE_STYLE_ELEMENT_RECORD_TYPE, element.clone()),
                'O' => (0x000A, Vec::new()),
                other => panic!("unknown test record letter {other}"),
            };
            if let Err(error) = collector.feed_record(record_type, &data) {
                return format!("{letters} => {answers} {error}");
            }
            answers.push('.');
        }
        let finished = match collector.finish(differential_formats) {
            Ok(None) => "none".to_string(),
            Ok(Some(catalog)) => format!("catalog custom={}", catalog.custom_styles().len()),
            Err(error) => error.to_string(),
        };
        format!("{letters} => {answers} finish: {finished}")
    }

    /// The DXF, STYLES and TABLESTYLES collections are fed every order of up
    /// to three of their records, and longer orders through a complete custom
    /// catalog and past its end.
    #[test]
    fn table_style_collections_accept_and_reject_record_orders() {
        let mut cases = vec![String::new()];
        let mut previous = vec![String::new()];
        for _ in 0..3 {
            let mut next = Vec::new();
            for prefix in &previous {
                for letter in ['X', 'S', 'H', 'T', 'E', 'O'] {
                    next.push(format!("{prefix}{letter}"));
                }
            }
            cases.extend(next.iter().cloned());
            previous = next;
        }
        cases.extend(
            [
                "XSH", "XSHO", "XSHTE", "XSHTEO", "XSHTEOS", "XSHTEOE", "XSHTEOT", "XSHTEOH",
                "XSHTES", "XSHTESH", "XSHTET", "XSHTETE", "XSHTEE", "XSHTEU", "XSHTEUO", "XSHUTE",
                "XSHUO", "XSHUE", "XSHTO", "XSHTS", "XSHTH", "XSHTU", "XSHOS", "XSHOSX", "XSHSX",
                "XSHSH", "XSHST", "XSHSE", "XSOH", "XSSH", "SXH", "XSHOO",
            ]
            .map(str::to_string),
        );
        let actual: Vec<String> = cases.iter().map(|case| outcome(case)).collect();
        assert_eq!(actual, TABLE_STYLE_ORDER_EXPECTED);
    }

    const TABLE_STYLE_ORDER_EXPECTED: &[&str] = &[
        " =>  finish: none",
        "X => . finish: none",
        "S => . finish: none",
        "H =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "T =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "E =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "O => . finish: none",
        "XX => .. finish: none",
        "XS => .. finish: none",
        "XH => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XT => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XO => .. finish: none",
        "SX => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SS => .. finish: none",
        "SH => .. finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "ST => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "SE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SO => .. finish: none",
        "HX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HT =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "TX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TT =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "EX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ES =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ET =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OX => .. finish: none",
        "OS => .. finish: none",
        "OH => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OT => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OO => .. finish: none",
        "XXX => ... finish: none",
        "XXS => ... finish: none",
        "XXH => .. Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XXT => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XXE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XXO => ... finish: none",
        "XSX => .. Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "XSS => ... finish: none",
        "XSH => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "XST => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XSE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XSO => ... finish: none",
        "XHX => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XHS => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XHH => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XHT => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XHE => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XHO => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XTX => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XTS => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XTH => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XTT => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XTE => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XTO => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XEX => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XES => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XEH => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XET => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XEE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XEO => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XOX => ... finish: none",
        "XOS => ... finish: none",
        "XOH => .. Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "XOT => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XOE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XOO => ... finish: none",
        "SXX => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SXS => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SXH => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SXT => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SXE => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SXO => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SSX => .. Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SSS => ... finish: none",
        "SSH => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "SST => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "SSE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SSO => ... finish: none",
        "SHX => .. Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SHS => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "SHH => .. Invalid record 0x088E: duplicate or noncontiguous TableStyles",
        "SHT => ... finish: Invalid record 0x0890: TABLESTYLES ended before the declared element count was satisfied",
        "SHE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SHO => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "STX => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "STS => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "STH => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "STT => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "STE => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "STO => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "SEX => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SES => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SEH => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SET => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SEE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SEO => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SOX => .. Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "SOS => ... finish: none",
        "SOH => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "SOT => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "SOE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "SOO => ... finish: none",
        "HXX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HXS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HXH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HXT =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HXE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HXO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HSX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HSS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HSH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HST =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HSE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HSO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHT =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HHO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTT =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HTO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HEX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HES =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HEH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HET =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HEE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HEO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOX =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOS =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOH =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOT =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOE =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "HOO =>  Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "TXX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TXS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TXH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TXT =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TXE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TXO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TSX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TSS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TSH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TST =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TSE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TSO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THT =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "THO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTT =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TTO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TEX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TES =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TEH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TET =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TEE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TEO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOX =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOS =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOH =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOT =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOE =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "TOO =>  Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "EXX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EXS =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EXH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EXT =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EXE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EXO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ESX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ESS =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ESH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EST =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ESE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ESO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHS =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHT =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EHO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETS =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETT =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "ETO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EEX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EES =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EEH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EET =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EEE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EEO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOX =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOS =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOH =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOT =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOE =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "EOO =>  Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OXX => ... finish: none",
        "OXS => ... finish: none",
        "OXH => .. Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OXT => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OXE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OXO => ... finish: none",
        "OSX => .. Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "OSS => ... finish: none",
        "OSH => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "OST => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OSE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OSO => ... finish: none",
        "OHX => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OHS => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OHH => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OHT => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OHE => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OHO => . Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OTX => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OTS => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OTH => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OTT => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OTE => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OTO => . Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OEX => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OES => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OEH => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OET => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OEE => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OEO => . Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OOX => ... finish: none",
        "OOS => ... finish: none",
        "OOH => .. Invalid record 0x088E: TableStyles appears before the required STYLES collection",
        "OOT => .. Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "OOE => .. Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "OOO => ... finish: none",
        "XSH => ... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "XSHO => .... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "XSHTE => ..... finish: catalog custom=1",
        "XSHTEO => ...... finish: catalog custom=1",
        "XSHTEOS => ....... finish: catalog custom=1",
        "XSHTEOE => ...... Invalid record 0x0890: TableStyleElement appears after TABLESTYLES ended",
        "XSHTEOT => ...... Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XSHTEOH => ...... Invalid record 0x088E: duplicate or noncontiguous TableStyles",
        "XSHTES => ...... finish: catalog custom=1",
        "XSHTESH => ...... Invalid record 0x088E: duplicate or noncontiguous TableStyles",
        "XSHTET => ...... finish: Invalid record 0x0890: TABLESTYLES ended before the declared element count was satisfied",
        "XSHTETE => ....... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 2",
        "XSHTEE => ..... Invalid record 0x0890: TableStyleElement exceeds the preceding ctse count",
        "XSHTEU => ...... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 2",
        "XSHTEUO => ....... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 2",
        "XSHUTE => ...... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 2",
        "XSHUO => ..... finish: catalog custom=1",
        "XSHUE => .... Invalid record 0x0890: TableStyleElement has no preceding TableStyle",
        "XSHTO => .... Invalid record 0x0890: TABLESTYLES ended before the declared element count was satisfied",
        "XSHTS => .... Invalid record 0x0890: TABLESTYLES ended before the declared element count was satisfied",
        "XSHTH => .... Invalid record 0x088E: duplicate or noncontiguous TableStyles",
        "XSHTU => .... Invalid record 0x088F: TableStyle appears before the prior element count is satisfied",
        "XSHOS => ..... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "XSHOSX => ..... Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "XSHSX => .... Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "XSHSH => .... Invalid record 0x088E: duplicate or noncontiguous TableStyles",
        "XSHST => .... Invalid record 0x088F: TableStyle appears outside TABLESTYLES",
        "XSHSE => .... Invalid record 0x0890: TableStyleElement appears after TABLESTYLES ended",
        "XSOH => .... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "XSSH => .... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
        "SXH => . Invalid record 0x088D: DXF record appears after the STYLES or TABLESTYLES collection",
        "XSHOO => ..... finish: Invalid record 0x088E: TableStyles declares 1 custom styles but contains 0",
    ];
}
