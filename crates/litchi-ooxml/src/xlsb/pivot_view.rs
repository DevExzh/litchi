//! Compatibility adapter for the owner PivotTable-view framing codec.
//!
//! Workbook, worksheet, relationship, and package orchestration remain in
//! this crate. The bounded BIFF12 framing and lossless stream retention live
//! in [`litchi_xlsb::pivot_view`].

use crate::xlsb::error::{XlsbError, XlsbResult};

/// A PivotTable definition stream with validated enclosing records, preserved verbatim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbPivotTableViewPart {
    inner: litchi_xlsb::pivot_view::PivotTableViewPart,
}

impl XlsbPivotTableViewPart {
    /// Parse a complete PivotTable part while retaining every original byte.
    pub fn from_bytes(bytes: Vec<u8>) -> XlsbResult<Self> {
        litchi_xlsb::pivot_view::PivotTableViewPart::from_bytes(bytes)
            .map(|inner| Self { inner })
            .map_err(map_owner_error)
    }

    /// Unique PivotTable view name (`irstName`).
    pub fn name(&self) -> &str {
        self.inner.name()
    }

    /// Workbook PivotCache identifier (`idCache`).
    pub fn cache_id(&self) -> u32 {
        self.inner.cache_id()
    }

    /// Data functionality level that created the view (`bVerSxMacro`).
    pub fn version_created(&self) -> u8 {
        self.inner.version_created()
    }

    /// Complete original PivotTable definition stream.
    pub fn as_bytes(&self) -> &[u8] {
        self.inner.as_bytes()
    }
}

fn map_owner_error(error: litchi_xlsb::pivot_view::Error) -> XlsbError {
    match error {
        litchi_xlsb::pivot_view::Error::Wire(error) => XlsbError::Wire(error),
        litchi_xlsb::pivot_view::Error::InvalidLength { expected, found } => {
            XlsbError::InvalidLength { expected, found }
        },
        litchi_xlsb::pivot_view::Error::Invalid(message) => XlsbError::InvalidFormula(message),
        other => XlsbError::InvalidFormula(other.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_xlsb::raw::{Writer, kind};

    fn view_stream(name: &str, cache_id: u32) -> Vec<u8> {
        let mut begin = vec![0u8; 32];
        begin[28..32].copy_from_slice(&cache_id.to_le_bytes());
        begin.extend_from_slice(&(name.encode_utf16().count() as u32).to_le_bytes());
        for unit in name.encode_utf16() {
            begin.extend_from_slice(&unit.to_le_bytes());
        }
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer.write_record(kind::BEGIN_SX_VIEW, &begin).unwrap();
        writer
            .write_record(kind::BEGIN_SX_LOCATION, &[0; 36])
            .unwrap();
        writer.write_record(kind::END_SX_LOCATION, &[]).unwrap();
        writer.write_record(kind::END_SX_VIEW, &[]).unwrap();
        bytes
    }

    #[test]
    fn preserves_complete_view_stream_and_extracts_binding() {
        let bytes = view_stream("Revenue Pivot", 17);
        let view = XlsbPivotTableViewPart::from_bytes(bytes.clone()).unwrap();
        assert_eq!(view.name(), "Revenue Pivot");
        assert_eq!(view.cache_id(), 17);
        assert_eq!(view.version_created(), 0);
        assert_eq!(view.as_bytes(), bytes);
    }

    #[test]
    fn refuses_truncation_and_records_outside_view() {
        let mut truncated = view_stream("P", 1);
        truncated.pop();
        assert!(XlsbPivotTableViewPart::from_bytes(truncated).is_err());

        let mut trailing = view_stream("P", 1);
        Writer::new(&mut trailing)
            .write_record(kind::END_SX_LOCATION, &[])
            .unwrap();
        assert!(XlsbPivotTableViewPart::from_bytes(trailing).is_err());
    }
}
