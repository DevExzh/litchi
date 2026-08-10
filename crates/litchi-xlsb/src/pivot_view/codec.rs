#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! Bounded BIFF12 PivotTable-view framing codec.

use super::model::Part;
use super::{Error, Result};
use crate::raw::{Cursor, Header, Kind, Limits};

const MAX_PIVOT_TABLE_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_PIVOT_TABLE_RECORDS: usize = 1_000_000;
const MAX_PIVOT_VIEW_NAME_UNITS: usize = 255;
const PIVOT_VIEW_LIMITS: Limits =
    Limits::new(MAX_PIVOT_TABLE_PART_BYTES, MAX_PIVOT_VIEW_NAME_UNITS);

impl Part {
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
            count = count
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("PivotTable record count overflow".to_string()))?;
            if count > MAX_PIVOT_TABLE_RECORDS {
                return Err(Error::Invalid(
                    "PivotTable record count exceeds the safety limit".to_string(),
                ));
            }

            let tail = bytes.get(offset..).ok_or_else(|| {
                Error::Invalid(format!("PivotTable record offset {offset} is invalid"))
            })?;
            let (header, header_len) = Header::parse(tail, PIVOT_VIEW_LIMITS)?;
            let payload_start = offset.checked_add(header_len).ok_or_else(|| {
                Error::Invalid("PivotTable record header offset overflow".to_string())
            })?;
            let payload_end = payload_start.checked_add(header.len()).ok_or_else(|| {
                Error::Invalid("PivotTable record payload offset overflow".to_string())
            })?;
            if payload_end > bytes.len() {
                let remaining = bytes.len().saturating_sub(payload_start);
                return Err(Error::Invalid(format!(
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
                        return Err(Error::Invalid(
                            "PivotTable has duplicate or misplaced BrtBeginSXView".to_string(),
                        ));
                    }
                    begin = Some(parse_begin_view(payload)?);
                },
                crate::raw::kind::END_SX_VIEW => {
                    if begin.is_none() || ended || !payload.is_empty() {
                        return Err(Error::Invalid(
                            "PivotTable has malformed BrtEndSXView".to_string(),
                        ));
                    }
                    ended = true;
                    if payload_end != bytes.len() {
                        return Err(Error::Invalid(
                            "PivotTable has records after BrtEndSXView".to_string(),
                        ));
                    }
                },
                _ if begin.is_none() || ended => {
                    return Err(Error::Invalid(
                        "PivotTable record lies outside BrtBeginSXView collection".to_string(),
                    ));
                },
                _ => {},
            }

            offset = payload_end;
        }

        let (name, cache_id, version_created) =
            begin.ok_or_else(|| Error::Invalid("PivotTable omits BrtBeginSXView".to_string()))?;
        if !ended {
            return Err(Error::Invalid("PivotTable omits BrtEndSXView".to_string()));
        }

        Ok(Part::new(name, cache_id, version_created, bytes))
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
        return Err(Error::Invalid(
            "PivotTable view name is empty, too long, or contains NUL".to_string(),
        ));
    }
    Ok((name, cache_id, data[0]))
}
