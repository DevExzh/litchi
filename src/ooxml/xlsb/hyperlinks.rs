//! Hyperlink support for XLSB

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};
use crate::ooxml::xlsb::records::wide_str_with_len;

/// Hyperlink information
///
/// Represents a hyperlink in a cell or range of cells.
///
/// When reading existing XLSB files, we only see the relationship ID (`r_id`)
/// stored in the `BrtHLink` record. The actual external URL is stored in the
/// OPC relationships part and is currently not resolved here. For writer-side
/// usage we allow an optional `target` URL which is used to create the
/// appropriate external relationship and generate a concrete `r_id` during
/// XLSB writing.
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// First row (0-based)
    pub row_first: u32,
    /// Last row (0-based, inclusive)
    pub row_last: u32,
    /// First column (0-based)
    pub col_first: u32,
    /// Last column (0-based, inclusive)
    pub col_last: u32,
    /// Relationship ID (points to external link)
    pub r_id: String,
    /// Location within document (e.g., sheet reference)
    pub location: Option<String>,
    /// Tooltip text
    pub tooltip: Option<String>,
    /// Display text
    pub display: Option<String>,
    /// External hyperlink target URL (writer-side only, `None` for parsed links)
    pub target: Option<String>,
}

impl Hyperlink {
    /// Create a new hyperlink with an explicit relationship ID.
    ///
    /// This constructor is primarily intended for low-level scenarios where
    /// the caller already knows the `rId` that will be used in the
    /// `sheetX.bin.rels` part. For typical writer usage prefer
    /// [`Hyperlink::new_external`], which takes a URL and lets the writer
    /// create the relationship.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
    ///
    /// let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string());
    /// ```
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32, r_id: String) -> Self {
        Hyperlink {
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

    /// Create a new internal hyperlink that points to a location inside the
    /// workbook.
    ///
    /// The `location` uses standard A1-style references such as
    /// `"Sheet2!A1"`. No external OPC relationship is created for these
    /// links; instead, the location is stored directly in the `BrtHLink`
    /// record and the `r_id` field is left empty, as specified in
    /// [MS-XLSB] 2.4.355.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
    ///
    /// // Hyperlink from A1 to A3 on the same sheet (0-based coordinates)
    /// let link = Hyperlink::new_internal(0, 0, 0, 0, "Sheet1!A3".to_string());
    /// ```
    pub fn new_internal(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        location: String,
    ) -> Self {
        Hyperlink {
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

    /// Create a new external hyperlink that points to a URL.
    ///
    /// The URL will be used by the writer to create an external OPC
    /// relationship of type `relationships::HYPERLINK`. A concrete `rId` will
    /// be generated automatically and injected into the `BrtHLink` record at
    /// write time.
    pub fn new_external(
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        target: String,
    ) -> Self {
        Hyperlink {
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

    /// Set location (e.g., "Sheet1!A1")
    pub fn with_location(mut self, location: String) -> Self {
        self.location = Some(location);
        self
    }

    /// Set tooltip
    pub fn with_tooltip(mut self, tooltip: String) -> Self {
        self.tooltip = Some(tooltip);
        self
    }

    /// Set display text
    pub fn with_display(mut self, display: String) -> Self {
        self.display = Some(display);
        self
    }

    /// Parse from XLSB BrtHLink record
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        let row_first = binary::read_u32_le_at(data, 0)?;
        let row_last = binary::read_u32_le_at(data, 4)?;
        let col_first = binary::read_u32_le_at(data, 8)?;
        let col_last = binary::read_u32_le_at(data, 12)?;

        let mut offset = 16;

        // Read relationship ID
        let (r_id, consumed) = wide_str_with_len(&data[offset..])?;
        offset += consumed;

        // Read location (optional)
        let (location, consumed) = if offset < data.len() {
            let (loc, c) = wide_str_with_len(&data[offset..])?;
            (if loc.is_empty() { None } else { Some(loc) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read tooltip (optional)
        let (tooltip, consumed) = if offset < data.len() {
            let (tt, c) = wide_str_with_len(&data[offset..])?;
            (if tt.is_empty() { None } else { Some(tt) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read display text (optional)
        let display = if offset < data.len() {
            let (disp, _) = wide_str_with_len(&data[offset..])?;
            if disp.is_empty() { None } else { Some(disp) }
        } else {
            None
        };

        Ok(Hyperlink {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location,
            tooltip,
            display,
            // For parsed hyperlinks we currently keep only the r_id and
            // textual properties. The external URL lives in the
            // sheetX.bin.rels part and is not resolved here.
            target: None,
        })
    }

    /// Serialize to XLSB BrtHLink record
    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::new();

        // Write range
        data.extend_from_slice(&self.row_first.to_le_bytes());
        data.extend_from_slice(&self.row_last.to_le_bytes());
        data.extend_from_slice(&self.col_first.to_le_bytes());
        data.extend_from_slice(&self.col_last.to_le_bytes());

        // Write r_id (RelID)
        Self::write_wide_string(&mut data, &self.r_id);

        // Write location (optional)
        if let Some(ref loc) = self.location {
            Self::write_wide_string(&mut data, loc);
        } else {
            Self::write_wide_string(&mut data, "");
        }

        // Write tooltip (optional)
        if let Some(ref tt) = self.tooltip {
            Self::write_wide_string(&mut data, tt);
        } else {
            Self::write_wide_string(&mut data, "");
        }

        // Write display (optional). Excel accepts either an empty string or a
        // friendly display name. For round-trip compatibility we preserve any
        // parsed display text and only fall back to the empty string when no
        // display is set.
        if let Some(ref disp) = self.display {
            Self::write_wide_string(&mut data, disp);
        } else {
            Self::write_wide_string(&mut data, "");
        }

        data
    }

    /// Helper to write a wide string
    fn write_wide_string(data: &mut Vec<u8>, s: &str) {
        let utf16: Vec<u16> = s.encode_utf16().collect();
        data.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
        for code_unit in utf16 {
            data.extend_from_slice(&code_unit.to_le_bytes());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_builder() {
        let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string())
            .with_location("Sheet1!A1".to_string())
            .with_tooltip("Click here".to_string());

        assert_eq!(link.location, Some("Sheet1!A1".to_string()));
        assert_eq!(link.tooltip, Some("Click here".to_string()));
    }

    #[test]
    fn test_internal_hyperlink_builder() {
        let link = Hyperlink::new_internal(0, 1, 2, 3, "Sheet2!B5".to_string())
            .with_tooltip("Go to Sheet2".to_string());

        assert_eq!(link.row_first, 0);
        assert_eq!(link.row_last, 1);
        assert_eq!(link.col_first, 2);
        assert_eq!(link.col_last, 3);
        assert!(link.r_id.is_empty());
        assert_eq!(link.location, Some("Sheet2!B5".to_string()));
        assert_eq!(link.tooltip, Some("Go to Sheet2".to_string()));
        assert!(link.target.is_none());
    }
}
