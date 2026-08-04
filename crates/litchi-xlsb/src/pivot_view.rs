//! Bounded, lossless PivotTable-view framing for XLSB.
//!
//! The PivotTable part is an extensible BIFF12 record collection described by
//! [MS-XLSB] sections 2.1.7.40, 2.4.278, and 2.4.631. The owner validates the
//! enclosing `BrtBeginSXView`/`BrtEndSXView` records and extracts the fields
//! needed for package binding, while retaining every original byte so newer
//! PivotTable records survive a read/write cycle.

use crate::raw::{Cursor, Header, Kind, Limits};
use std::fmt;
use thiserror::Error as ThisError;

const MAX_PIVOT_TABLE_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_PIVOT_TABLE_RECORDS: usize = 1_000_000;
const MAX_PIVOT_VIEW_NAME_UNITS: usize = 255;
const PIVOT_VIEW_LIMITS: Limits =
    Limits::new(MAX_PIVOT_TABLE_PART_BYTES, MAX_PIVOT_VIEW_NAME_UNITS);

/// Result type for the standalone PivotTable-view codec.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the bounded PivotTable-view codec.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A BIFF12 header or scalar failed raw validation.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    /// A fixed-width field or enclosing stream boundary is malformed.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A PivotTable framing or identity invariant is invalid.
    #[error("invalid PivotTable view: {0}")]
    InvalidFormula(String),
}

/// A complete PivotTable definition stream with validated framing.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableViewPart {
    name: String,
    cache_id: u32,
    version_created: u8,
    bytes: Vec<u8>,
}

impl PivotTableViewPart {
    /// Parse a complete PivotTable part while retaining every original byte.
    ///
    /// Only the enclosing view records and the binding fields in
    /// `BrtBeginSXView` are interpreted. Inner records remain opaque so
    /// extension records and features outside this crate's typed model are
    /// preserved losslessly.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > MAX_PIVOT_TABLE_PART_BYTES {
            return Err(Error::InvalidLength {
                expected: MAX_PIVOT_TABLE_PART_BYTES,
                found: bytes.len(),
            });
        }

        let mut offset = 0usize;
        let mut count = 0usize;
        let mut begin = None;
        let mut ended = false;

        while offset < bytes.len() {
            count = count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormula("PivotTable record count overflow".to_string())
            })?;
            if count > MAX_PIVOT_TABLE_RECORDS {
                return Err(Error::InvalidFormula(
                    "PivotTable record count exceeds the safety limit".to_string(),
                ));
            }

            let tail = bytes.get(offset..).ok_or_else(|| {
                Error::InvalidFormula(format!("PivotTable record offset {offset} is invalid"))
            })?;
            let (header, header_len) = Header::parse(tail, PIVOT_VIEW_LIMITS)?;
            let payload_start = offset.checked_add(header_len).ok_or_else(|| {
                Error::InvalidFormula("PivotTable record header offset overflow".to_string())
            })?;
            let payload_end = payload_start.checked_add(header.len()).ok_or_else(|| {
                Error::InvalidFormula("PivotTable record payload offset overflow".to_string())
            })?;
            if payload_end > bytes.len() {
                let remaining = bytes.len().saturating_sub(payload_start);
                return Err(Error::InvalidFormula(format!(
                    "PivotTable record {} declares {} bytes with only {remaining} remaining",
                    header.kind(),
                    header.len()
                )));
            }
            let payload = &bytes[payload_start..payload_end];

            let record_kind: Kind = header.kind();
            match record_kind {
                crate::raw::kind::BEGIN_SX_VIEW => {
                    if begin.is_some() || count != 1 || ended {
                        return Err(Error::InvalidFormula(
                            "PivotTable has duplicate or misplaced BrtBeginSXView".to_string(),
                        ));
                    }
                    begin = Some(parse_begin_view(payload)?);
                },
                crate::raw::kind::END_SX_VIEW => {
                    if begin.is_none() || ended || !payload.is_empty() {
                        return Err(Error::InvalidFormula(
                            "PivotTable has malformed BrtEndSXView".to_string(),
                        ));
                    }
                    ended = true;
                    if payload_end != bytes.len() {
                        return Err(Error::InvalidFormula(
                            "PivotTable has records after BrtEndSXView".to_string(),
                        ));
                    }
                },
                _ if begin.is_none() || ended => {
                    return Err(Error::InvalidFormula(
                        "PivotTable record lies outside BrtBeginSXView collection".to_string(),
                    ));
                },
                _ => {},
            }

            offset = payload_end;
        }

        let (name, cache_id, version_created) = begin
            .ok_or_else(|| Error::InvalidFormula("PivotTable omits BrtBeginSXView".to_string()))?;
        if !ended {
            return Err(Error::InvalidFormula(
                "PivotTable omits BrtEndSXView".to_string(),
            ));
        }

        Ok(Self {
            name,
            cache_id,
            version_created,
            bytes,
        })
    }

    /// Unique PivotTable view name (`irstName`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Workbook PivotCache identifier (`idCache`).
    #[must_use]
    pub const fn cache_id(&self) -> u32 {
        self.cache_id
    }

    /// Data functionality level that created the view (`bVerSxMacro`).
    #[must_use]
    pub const fn version_created(&self) -> u8 {
        self.version_created
    }

    /// Complete original PivotTable definition stream.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }
}

fn parse_begin_view(data: &[u8]) -> Result<(String, u32, u8)> {
    if data.len() < 36 {
        return Err(Error::InvalidLength {
            expected: 36,
            found: data.len(),
        });
    }

    let cache_id_bytes = data.get(28..32).ok_or(Error::InvalidLength {
        expected: 32,
        found: data.len(),
    })?;
    let cache_id =
        u32::from_le_bytes(
            cache_id_bytes
                .try_into()
                .map_err(|_| Error::InvalidLength {
                    expected: 4,
                    found: cache_id_bytes.len(),
                })?,
        );
    let mut cursor = Cursor::with_limits(&data[32..], "BrtBeginSXView.irstName", PIVOT_VIEW_LIMITS);
    let name = cursor.read_wide_string()?;
    let units = name.encode_utf16().count();
    if units == 0 || units > MAX_PIVOT_VIEW_NAME_UNITS || name.contains('\0') {
        return Err(Error::InvalidFormula(
            "PivotTable view name is empty, too long, or contains NUL".to_string(),
        ));
    }
    Ok((name, cache_id, data[0]))
}

impl fmt::Display for PivotTableViewPart {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PivotTableViewPart")
            .field("name", &self.name)
            .field("cache_id", &self.cache_id)
            .field("version_created", &self.version_created)
            .field("bytes", &self.bytes.len())
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::{Writer, kind};

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
        let view = PivotTableViewPart::from_bytes(bytes.clone()).unwrap();
        assert_eq!(view.name(), "Revenue Pivot");
        assert_eq!(view.cache_id(), 17);
        assert_eq!(view.version_created(), 0);
        assert_eq!(view.as_bytes(), bytes);
    }

    #[test]
    fn refuses_truncation_and_records_outside_view() {
        let mut truncated = view_stream("P", 1);
        truncated.pop();
        assert!(PivotTableViewPart::from_bytes(truncated).is_err());

        let mut trailing = view_stream("P", 1);
        Writer::new(&mut trailing)
            .write_record(kind::END_SX_LOCATION, &[])
            .unwrap();
        assert!(PivotTableViewPart::from_bytes(trailing).is_err());
    }
}
