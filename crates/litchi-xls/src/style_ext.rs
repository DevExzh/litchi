//! BIFF8 `StyleExt` record (MS-XLS 2.4.270): cell-style extensions.
//!
//! A `StyleExt` follows its `Style` record and carries the style category,
//! visibility/customization flags, the style name as an `LPWideString`, and
//! an `XFProps` formatting property array shared with the DXF machinery.

use super::differential_format::XfProperties;
use super::{Error, Result};

/// Record type of the `StyleExt` record.
pub(crate) const STYLE_EXT_RECORD_TYPE: u16 = 0x0892;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Maximum length of the style name in UTF-16 code units.
const MAX_STYLE_NAME_CHARS: usize = 255;
/// `builtInData` value when the style is not built in.
const NOT_BUILT_IN_DATA: u16 = 0xFFFF;

// Flag bits of the byte following the FrtHeader.
const BUILT_IN: u8 = 0x01;
const HIDDEN: u8 = 0x02;
const CUSTOM: u8 = 0x04;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: STYLE_EXT_RECORD_TYPE,
        message: message.into(),
    }
}

/// The style category of a `StyleExt` (`iCategory`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleCategory {
    /// Custom style.
    Custom,
    /// Good, bad, neutral style.
    GoodBadNeutral,
    /// Data model style.
    DataModel,
    /// Title and heading style.
    TitleAndHeading,
    /// Themed cell style.
    ThemedCell,
    /// Number format style.
    NumberFormat,
}

impl StyleCategory {
    fn from_code(value: u8) -> Result<Self> {
        Ok(match value {
            0x00 => Self::Custom,
            0x01 => Self::GoodBadNeutral,
            0x02 => Self::DataModel,
            0x03 => Self::TitleAndHeading,
            0x04 => Self::ThemedCell,
            0x05 => Self::NumberFormat,
            value => return Err(invalid(format!("reserved style category {value:#04X}"))),
        })
    }

    fn code(self) -> u8 {
        match self {
            Self::Custom => 0x00,
            Self::GoodBadNeutral => 0x01,
            Self::DataModel => 0x02,
            Self::TitleAndHeading => 0x03,
            Self::ThemedCell => 0x04,
            Self::NumberFormat => 0x05,
        }
    }
}

/// Typed `StyleExt` record content (MS-XLS 2.4.270).
#[derive(Debug, Clone, PartialEq)]
pub struct StyleExt {
    built_in: bool,
    hidden: bool,
    custom: bool,
    category: StyleCategory,
    /// Raw `BuiltInStyle` data for built-in styles; `None` for custom styles.
    built_in_data: Option<u16>,
    name: String,
    properties: XfProperties,
}

impl StyleExt {
    /// A style extension; the name is limited to 255 UTF-16 code units, and
    /// `custom` requires a built-in style.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(
        built_in: bool,
        category: StyleCategory,
        name: String,
        properties: XfProperties,
    ) -> Result<Self> {
        if name.encode_utf16().count() > MAX_STYLE_NAME_CHARS {
            return Err(invalid("style name exceeds 255 UTF-16 code units"));
        }
        Ok(Self {
            built_in,
            hidden: false,
            custom: false,
            category,
            built_in_data: built_in.then_some(0),
            name,
            properties,
        })
    }

    #[must_use]
    pub const fn built_in(&self) -> bool {
        self.built_in
    }
    #[must_use]
    pub const fn hidden(&self) -> bool {
        self.hidden
    }
    #[must_use]
    pub const fn custom(&self) -> bool {
        self.custom
    }
    #[must_use]
    pub const fn category(&self) -> StyleCategory {
        self.category
    }
    #[must_use]
    pub const fn built_in_data(&self) -> Option<u16> {
        self.built_in_data
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub fn properties(&self) -> &XfProperties {
        &self.properties
    }

    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
    }

    /// Mark the built-in style as customized; requires a built-in style.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn set_custom(&mut self, custom: bool) -> Result<()> {
        if custom && !self.built_in {
            return Err(invalid("fCustom requires a built-in style"));
        }
        self.custom = custom;
        Ok(())
    }

    /// Parse a `StyleExt` record payload.
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < FRT_HEADER_LEN + 4 {
            return Err(Error::InvalidLength {
                expected: FRT_HEADER_LEN + 4,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != STYLE_EXT_RECORD_TYPE {
            return Err(invalid("StyleExt FrtHeader.rt mismatch"));
        }
        let flags = data[12];
        let built_in = flags & BUILT_IN != 0;
        let custom = flags & CUSTOM != 0;
        if custom && !built_in {
            return Err(invalid("fCustom requires a built-in style"));
        }
        let category = StyleCategory::from_code(data[13])?;
        let built_in_data_raw = u16::from_le_bytes([data[14], data[15]]);
        let built_in_data = if built_in {
            Some(built_in_data_raw)
        } else {
            if built_in_data_raw != NOT_BUILT_IN_DATA {
                return Err(invalid("custom style must carry 0xFFFF builtInData"));
            }
            None
        };
        if data.len() < FRT_HEADER_LEN + 6 {
            return Err(Error::InvalidLength {
                expected: FRT_HEADER_LEN + 6,
                found: data.len(),
            });
        }
        let name_chars = usize::from(u16::from_le_bytes([data[16], data[17]]));
        if name_chars > MAX_STYLE_NAME_CHARS {
            return Err(invalid("style name exceeds 255 UTF-16 code units"));
        }
        let name_end = FRT_HEADER_LEN
            .checked_add(6)
            .and_then(|start| start.checked_add(name_chars * 2))
            .ok_or_else(|| invalid("style name length overflow"))?;
        let name_bytes = data
            .get(FRT_HEADER_LEN + 6..name_end)
            .ok_or(Error::InvalidLength {
                expected: name_end,
                found: data.len(),
            })?;
        let units = name_bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let name = String::from_utf16(&units)
            .map_err(|_error| invalid("style name contains invalid UTF-16"))?;
        let properties =
            XfProperties::parse(data.get(name_end..).ok_or(Error::InvalidLength {
                expected: name_end,
                found: data.len(),
            })?)?;
        Ok(Self {
            built_in,
            hidden: flags & HIDDEN != 0,
            custom,
            category,
            built_in_data,
            name,
            properties,
        })
    }

    /// Serialize back to a complete `StyleExt` record payload.
    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        if self.custom && !self.built_in {
            return Err(invalid("fCustom requires a built-in style"));
        }
        let name_units = self.name.encode_utf16().count();
        if name_units > MAX_STYLE_NAME_CHARS {
            return Err(invalid("style name exceeds 255 UTF-16 code units"));
        }
        let mut flags = 0u8;
        if self.built_in {
            flags |= BUILT_IN;
        }
        if self.hidden {
            flags |= HIDDEN;
        }
        if self.custom {
            flags |= CUSTOM;
        }
        let mut payload = Vec::new();
        payload.extend_from_slice(&STYLE_EXT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        payload.push(flags);
        payload.push(self.category.code());
        payload.extend_from_slice(
            &self
                .built_in_data
                .unwrap_or(NOT_BUILT_IN_DATA)
                .to_le_bytes(),
        );
        payload.extend_from_slice(&crate::utils::truncate_usize_to_u16(name_units).to_le_bytes());
        for unit in self.name.encode_utf16() {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        payload.extend_from_slice(&self.properties.to_bytes()?);
        Ok(payload)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::differential_format::XfProperty;

    fn record(flags: u8, category: u8, built_in_data: u16, name: &str, xf_props: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STYLE_EXT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 10]);
        data.push(flags);
        data.push(category);
        data.extend_from_slice(&built_in_data.to_le_bytes());
        let units = name.encode_utf16().collect::<Vec<_>>();
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data.extend_from_slice(xf_props);
        data
    }

    fn xf_props_bytes() -> Vec<u8> {
        XfProperties::try_new(vec![XfProperty::Locked(true)])
            .unwrap()
            .to_bytes()
            .unwrap()
    }

    #[test]
    fn parses_built_in_style_extension() {
        let data = record(0x03, 0x04, 0x0000, "Heading 1", &xf_props_bytes());
        let parsed = StyleExt::parse(&data).unwrap();
        assert!(parsed.built_in());
        assert!(parsed.hidden());
        assert!(!parsed.custom());
        assert_eq!(parsed.category(), StyleCategory::ThemedCell);
        assert_eq!(parsed.built_in_data(), Some(0));
        assert_eq!(parsed.name(), "Heading 1");
        assert_eq!(
            parsed.properties().properties(),
            &[XfProperty::Locked(true)]
        );
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn parses_custom_style_extension() {
        let data = record(0x00, 0x00, NOT_BUILT_IN_DATA, "My Style", &xf_props_bytes());
        let parsed = StyleExt::parse(&data).unwrap();
        assert!(!parsed.built_in());
        assert_eq!(parsed.built_in_data(), None);
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated.
        assert!(StyleExt::parse(&[0; 8]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = record(0x00, 0x00, NOT_BUILT_IN_DATA, "x", &xf_props_bytes());
        wrong_rt[0..2].copy_from_slice(&0x087Du16.to_le_bytes());
        assert!(StyleExt::parse(&wrong_rt).is_err());
        // fCustom without fBuiltIn.
        assert!(
            StyleExt::parse(&record(
                0x04,
                0x00,
                NOT_BUILT_IN_DATA,
                "x",
                &xf_props_bytes()
            ))
            .is_err()
        );
        // Reserved category.
        assert!(
            StyleExt::parse(&record(
                0x00,
                0x06,
                NOT_BUILT_IN_DATA,
                "x",
                &xf_props_bytes()
            ))
            .is_err()
        );
        // Custom style with built-in data.
        assert!(StyleExt::parse(&record(0x00, 0x00, 0x0001, "x", &xf_props_bytes())).is_err());
        // Overlong name.
        let long = "x".repeat(256);
        assert!(
            StyleExt::parse(&record(
                0x00,
                0x00,
                NOT_BUILT_IN_DATA,
                &long,
                &xf_props_bytes()
            ))
            .is_err()
        );
        // Builder validation.
        let props = XfProperties::default();
        let mut value =
            StyleExt::try_new(false, StyleCategory::Custom, "s".to_string(), props).unwrap();
        assert!(value.set_custom(true).is_err());
    }
}
