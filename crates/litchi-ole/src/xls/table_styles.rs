//! BIFF8 default table and PivotTable style catalog metadata.

use super::{XlsError, XlsResult};

pub(crate) const TABLE_STYLES_RECORD_TYPE: u16 = 0x088E;
const BUILT_IN_STYLE_COUNT: u32 = 144;
const FIXED_PAYLOAD_LEN: usize = 20;
const MAX_STYLE_NAME_UNITS: usize = 255;
const MAX_PAYLOAD_LEN: usize = FIXED_PAYLOAD_LEN + MAX_STYLE_NAME_UNITS * 4;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: TABLE_STYLES_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> XlsResult<u16> {
    data.get(offset..offset + 2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| invalid(format!("truncated TableStyles {field}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> XlsResult<u32> {
    data.get(offset..offset + 4)
        .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()))
        .ok_or_else(|| invalid(format!("truncated TableStyles {field}")))
}

/// Default table-style catalog declared in the workbook globals stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsTableStyles {
    total_style_count: u32,
    default_table_style: String,
    default_pivot_style: String,
}

impl XlsTableStyles {
    pub fn try_new(
        total_style_count: u32,
        default_table_style: impl Into<String>,
        default_pivot_style: impl Into<String>,
    ) -> XlsResult<Self> {
        let value = Self {
            total_style_count,
            default_table_style: default_table_style.into(),
            default_pivot_style: default_pivot_style.into(),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
        if !(FIXED_PAYLOAD_LEN..=MAX_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(format!(
                "TableStyles payload has {} bytes; expected 20..={MAX_PAYLOAD_LEN}",
                data.len()
            )));
        }
        if read_u16(data, 0, "frtHeader.rt")? != TABLE_STYLES_RECORD_TYPE {
            return Err(invalid(
                "TableStyles future-record type does not match 0x088E",
            ));
        }
        if read_u16(data, 2, "frtHeader.grbitFrt")? != 0 {
            return Err(invalid("TableStyles future-record flags must be zero"));
        }
        if data[4..12].iter().any(|&byte| byte != 0) {
            return Err(invalid(
                "TableStyles future-record reserved bytes must be zero",
            ));
        }

        let total_style_count = read_u32(data, 12, "cts")?;
        let table_units = usize::from(read_u16(data, 16, "cchDefTableStyle")?);
        let pivot_units = usize::from(read_u16(data, 18, "cchDefPivotStyle")?);
        if table_units > MAX_STYLE_NAME_UNITS || pivot_units > MAX_STYLE_NAME_UNITS {
            return Err(invalid(
                "TableStyles default name exceeds 255 UTF-16 code units",
            ));
        }
        let table_bytes = table_units
            .checked_mul(2)
            .ok_or_else(|| invalid("TableStyles default table style length overflows"))?;
        let pivot_bytes = pivot_units
            .checked_mul(2)
            .ok_or_else(|| invalid("TableStyles default PivotTable style length overflows"))?;
        let table_end = FIXED_PAYLOAD_LEN
            .checked_add(table_bytes)
            .ok_or_else(|| invalid("TableStyles table style range overflows"))?;
        let pivot_end = table_end
            .checked_add(pivot_bytes)
            .ok_or_else(|| invalid("TableStyles PivotTable style range overflows"))?;
        if pivot_end != data.len() {
            return Err(invalid(
                "TableStyles name lengths do not consume the payload exactly",
            ));
        }
        let default_table_style =
            decode_utf16(&data[FIXED_PAYLOAD_LEN..table_end], "default table style")?;
        let default_pivot_style =
            decode_utf16(&data[table_end..pivot_end], "default PivotTable style")?;
        Self::try_new(total_style_count, default_table_style, default_pivot_style)
    }

    pub fn total_style_count(&self) -> u32 {
        self.total_style_count
    }
    pub const fn built_in_style_count(&self) -> u32 {
        BUILT_IN_STYLE_COUNT
    }
    pub fn custom_style_count(&self) -> u32 {
        self.total_style_count - BUILT_IN_STYLE_COUNT
    }
    pub fn has_custom_styles(&self) -> bool {
        self.custom_style_count() != 0
    }
    pub fn default_table_style(&self) -> &str {
        &self.default_table_style
    }
    pub fn default_pivot_style(&self) -> &str {
        &self.default_pivot_style
    }

    /// Serialize the complete TableStyles payload deterministically.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        let (table_units, pivot_units, size) = self.validate()?;
        let mut data = Vec::with_capacity(size);
        data.extend_from_slice(&TABLE_STYLES_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[0u8; 8]);
        data.extend_from_slice(&self.total_style_count.to_le_bytes());
        data.extend_from_slice(&(table_units.len() as u16).to_le_bytes());
        data.extend_from_slice(&(pivot_units.len() as u16).to_le_bytes());
        data.extend(table_units.into_iter().flat_map(u16::to_le_bytes));
        data.extend(pivot_units.into_iter().flat_map(u16::to_le_bytes));
        Ok(data)
    }

    /// Serialize the complete BIFF record including its four-byte record header.
    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        let payload = self.to_payload()?;
        let length = u16::try_from(payload.len())
            .map_err(|_| invalid("TableStyles payload length exceeds BIFF u16"))?;
        let mut data = Vec::with_capacity(4 + payload.len());
        data.extend_from_slice(&TABLE_STYLES_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&length.to_le_bytes());
        data.extend_from_slice(&payload);
        Ok(data)
    }

    fn validate(&self) -> XlsResult<(Vec<u16>, Vec<u16>, usize)> {
        if self.total_style_count < BUILT_IN_STYLE_COUNT {
            return Err(invalid("TableStyles cts must be at least 144"));
        }
        let table_units = self.default_table_style.encode_utf16().collect::<Vec<_>>();
        let pivot_units = self.default_pivot_style.encode_utf16().collect::<Vec<_>>();
        if table_units.len() > MAX_STYLE_NAME_UNITS || pivot_units.len() > MAX_STYLE_NAME_UNITS {
            return Err(invalid(
                "TableStyles default name exceeds 255 UTF-16 code units",
            ));
        }
        let size = table_units
            .len()
            .checked_add(pivot_units.len())
            .and_then(|units| units.checked_mul(2))
            .and_then(|bytes| bytes.checked_add(FIXED_PAYLOAD_LEN))
            .ok_or_else(|| invalid("TableStyles serialized size overflows"))?;
        if size > MAX_PAYLOAD_LEN {
            return Err(invalid(
                "TableStyles exceeds its specification-derived size cap",
            ));
        }
        Ok((table_units, pivot_units, size))
    }
}

fn decode_utf16(data: &[u8], field: &str) -> XlsResult<String> {
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units)
        .map_err(|_| invalid(format!("TableStyles {field} contains invalid UTF-16")))
}

pub(crate) struct TableStylesCollector {
    value: Option<XlsTableStyles>,
}

impl TableStylesCollector {
    pub(crate) fn new() -> Self {
        Self { value: None }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if record_type != TABLE_STYLES_RECORD_TYPE {
            return Ok(());
        }
        if self.value.is_some() {
            return Err(invalid("duplicate TableStyles record"));
        }
        self.value = Some(XlsTableStyles::parse_payload(data)?);
        Ok(())
    }

    pub(crate) fn finish(self) -> Option<XlsTableStyles> {
        self.value
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

    #[test]
    fn parses_and_round_trips_poi_reference() {
        let bytes = decode_hex(POI_REFERENCE_HEX);
        let styles = XlsTableStyles::parse_payload(&bytes).unwrap();
        assert_eq!(styles.total_style_count(), 144);
        assert_eq!(styles.custom_style_count(), 0);
        assert_eq!(styles.default_table_style(), "TableStyleMedium9");
        assert_eq!(styles.default_pivot_style(), "PivotStyleLight16");
        assert_eq!(styles.to_payload().unwrap(), bytes);
    }

    #[test]
    fn constructs_unicode_catalog_and_full_record() {
        let styles = XlsTableStyles::try_new(146, "\u{8868}\u{683c}", "\u{900f}\u{89c6}").unwrap();
        assert_eq!(styles.built_in_style_count(), 144);
        assert_eq!(styles.custom_style_count(), 2);
        let record = styles.to_record_bytes().unwrap();
        assert_eq!(&record[..2], &[0x8e, 0x08]);
        assert_eq!(XlsTableStyles::parse_payload(&record[4..]).unwrap(), styles);
    }

    #[test]
    fn rejects_malformed_headers_counts_lengths_and_encoding() {
        let reference = decode_hex(POI_REFERENCE_HEX);
        assert!(XlsTableStyles::parse_payload(&reference[..19]).is_err());
        let mut data = reference.clone();
        data[0] = 0;
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        let mut data = reference.clone();
        data[2] = 1;
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        let mut data = reference.clone();
        data[4] = 1;
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        let mut data = reference.clone();
        data[12..16].copy_from_slice(&143u32.to_le_bytes());
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        let mut data = reference.clone();
        data[16..18].copy_from_slice(&18u16.to_le_bytes());
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        let mut data = reference;
        data[20..22].copy_from_slice(&0xD800u16.to_le_bytes());
        assert!(XlsTableStyles::parse_payload(&data).is_err());
        assert!(XlsTableStyles::try_new(144, "x".repeat(256), "").is_err());

        let valid = XlsTableStyles::try_new(144, "TableStyleMedium2", "PivotStyleLight16")
            .unwrap()
            .to_payload()
            .unwrap();
        let mut collector = TableStylesCollector::new();
        collector
            .feed_record(TABLE_STYLES_RECORD_TYPE, &valid)
            .unwrap();
        assert!(
            collector
                .feed_record(TABLE_STYLES_RECORD_TYPE, &valid)
                .is_err()
        );
    }
}
