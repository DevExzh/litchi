//! Typed hyperlink records for XLSB.
//!
//! This module owns the semantic representation and `BrtHLink` codec from
//! `[MS-XLSB]` section 2.4.693. Worksheet and relationship orchestration
//! remains in the host crate.

use thiserror::Error;

use crate::raw::{Cursor, Writer};

/// The fixed-width `rfx` prefix of a `BrtHLink` payload.
pub const PREFIX_LEN: usize = 16;

/// Result type for hyperlink parsing and serialization.
pub type Result<T> = std::result::Result<T, Error>;

/// A typed `BrtHLink` failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The payload does not contain the fixed-width `rfx` prefix.
    #[error("invalid BrtHLink payload length: expected at least {expected} bytes, found {found}")]
    InvalidLength {
        /// Minimum payload length.
        expected: usize,
        /// Actual payload length.
        found: usize,
    },
    /// A scalar or string failed validated BIFF12 decoding or encoding.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
}

/// Hyperlink information for a cell or range of cells.
///
/// When reading existing XLSB files, the `r_id` identifies the destination in
/// the worksheet relationship part. The relationship target is intentionally
/// kept as an optional writer-side value; resolving package relationships is a
/// host concern.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    /// First row (zero-based).
    pub row_first: u32,
    /// Last row (zero-based, inclusive).
    pub row_last: u32,
    /// First column (zero-based).
    pub col_first: u32,
    /// Last column (zero-based, inclusive).
    pub col_last: u32,
    /// Relationship ID (points to an external link).
    pub r_id: String,
    /// Location within the destination document or workbook.
    pub location: Option<String>,
    /// Tooltip text.
    pub tooltip: Option<String>,
    /// Display text.
    pub display: Option<String>,
    /// External hyperlink target URL (writer-side only).
    pub target: Option<String>,
}

impl Hyperlink {
    /// Create a hyperlink with an explicit relationship ID.
    #[must_use]
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32, r_id: String) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location: None,
            tooltip: None,
            display: None,
            target: None,
        }
    }

    /// Create an internal hyperlink pointing to a workbook location.
    #[must_use]
    pub fn new_internal(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        location: String,
    ) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id: String::new(),
            location: Some(location),
            tooltip: None,
            display: None,
            target: None,
        }
    }

    /// Create an external hyperlink pointing to a URL.
    #[must_use]
    pub fn new_external(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        target: String,
    ) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id: String::new(),
            location: None,
            tooltip: None,
            display: None,
            target: Some(target),
        }
    }

    /// Set a fragment or workbook location.
    #[must_use]
    pub fn with_location(mut self, location: String) -> Self {
        self.location = Some(location);
        self
    }

    /// Set tooltip text.
    #[must_use]
    pub fn with_tooltip(mut self, tooltip: String) -> Self {
        self.tooltip = Some(tooltip);
        self
    }

    /// Set display text.
    #[must_use]
    pub fn with_display(mut self, display: String) -> Self {
        self.display = Some(display);
        self
    }

    /// Parse a `BrtHLink` payload.
    ///
    /// The specification requires all four wide-string fields. For
    /// compatibility with older host readers, a payload ending after any
    /// complete field is accepted for the remaining optional textual fields;
    /// a present but truncated field is still rejected by the raw cursor.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < PREFIX_LEN {
            return Err(Error::InvalidLength {
                expected: PREFIX_LEN,
                found: data.len(),
            });
        }

        let mut cursor = Cursor::new(data, "BrtHLink");
        let row_first = cursor.read_u32()?;
        let row_last = cursor.read_u32()?;
        let col_first = cursor.read_u32()?;
        let col_last = cursor.read_u32()?;
        let r_id = cursor.read_wide_string()?;
        let location = read_optional_string(&mut cursor)?;
        let tooltip = read_optional_string(&mut cursor)?;
        let display = read_optional_string(&mut cursor)?;
        cursor.finish()?;

        Ok(Self {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location,
            tooltip,
            display,
            target: None,
        })
    }

    /// Serialize a `BrtHLink` payload with validated resource limits.
    pub fn try_serialize(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer.write_u32(self.row_first)?;
        writer.write_u32(self.row_last)?;
        writer.write_u32(self.col_first)?;
        writer.write_u32(self.col_last)?;
        writer.write_wide_string(&self.r_id)?;
        write_optional_string(&mut writer, self.location.as_deref())?;
        write_optional_string(&mut writer, self.tooltip.as_deref())?;
        write_optional_string(&mut writer, self.display.as_deref())?;
        Ok(writer.finish())
    }

    /// Serialize to a `BrtHLink` payload.
    ///
    /// This infallible method preserves the historical host API. Callers
    /// processing untrusted or unusually large values should use
    /// [`Self::try_serialize`] so the raw-kernel resource limits are surfaced.
    #[must_use]
    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&self.row_first.to_le_bytes());
        data.extend_from_slice(&self.row_last.to_le_bytes());
        data.extend_from_slice(&self.col_first.to_le_bytes());
        data.extend_from_slice(&self.col_last.to_le_bytes());
        write_wide_string_legacy(&mut data, &self.r_id);
        write_wide_string_legacy(&mut data, self.location.as_deref().unwrap_or_default());
        write_wide_string_legacy(&mut data, self.tooltip.as_deref().unwrap_or_default());
        write_wide_string_legacy(&mut data, self.display.as_deref().unwrap_or_default());
        data
    }
}

fn read_optional_string(cursor: &mut Cursor<'_>) -> Result<Option<String>> {
    if cursor.remaining() == 0 {
        return Ok(None);
    }
    let value = cursor.read_wide_string()?;
    Ok((!value.is_empty()).then_some(value))
}

fn write_optional_string(writer: &mut Writer<Vec<u8>>, value: Option<&str>) -> Result<()> {
    writer.write_wide_string(value.unwrap_or_default())?;
    Ok(())
}

fn write_wide_string_legacy(data: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().count();
    data.extend_from_slice(&(units as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builders_preserve_hyperlink_semantics() {
        let link = Hyperlink::new_external(0, 1, 2, 3, "https://example.test".to_string())
            .with_location("#section".to_string())
            .with_tooltip("Open".to_string())
            .with_display("Example".to_string());

        assert_eq!(link.row_last, 1);
        assert_eq!(link.col_first, 2);
        assert_eq!(link.target.as_deref(), Some("https://example.test"));
        assert_eq!(link.location.as_deref(), Some("#section"));
    }

    #[test]
    fn serialize_parse_round_trip() {
        let link = Hyperlink::new(0, 4, 1, 3, "rId7".to_string())
            .with_location("Sheet2!A1".to_string())
            .with_tooltip("Go".to_string())
            .with_display("Display".to_string());

        let encoded = link.try_serialize().expect("valid hyperlink strings");
        let parsed = Hyperlink::parse(&encoded).expect("valid hyperlink payload");
        assert_eq!(
            parsed,
            Hyperlink {
                target: None,
                ..link
            }
        );
    }

    #[test]
    fn missing_optional_strings_are_accepted_for_compatibility() {
        let mut encoded = Vec::new();
        encoded.extend_from_slice(&0_u32.to_le_bytes());
        encoded.extend_from_slice(&1_u32.to_le_bytes());
        encoded.extend_from_slice(&2_u32.to_le_bytes());
        encoded.extend_from_slice(&3_u32.to_le_bytes());
        encoded.extend_from_slice(&0_u32.to_le_bytes());

        let parsed = Hyperlink::parse(&encoded).expect("relationship ID is required");
        assert_eq!(parsed.r_id, "");
        assert!(parsed.location.is_none());
    }

    #[test]
    fn rejects_truncated_prefix_and_string() {
        assert!(matches!(
            Hyperlink::parse(&[0; PREFIX_LEN - 1]),
            Err(Error::InvalidLength { .. })
        ));

        let mut encoded = vec![0; PREFIX_LEN];
        encoded.extend_from_slice(&2_u32.to_le_bytes());
        encoded.extend_from_slice(&[0, 0]);
        assert!(matches!(Hyperlink::parse(&encoded), Err(Error::Wire(_))));
    }
}
