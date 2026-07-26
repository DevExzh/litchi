//! Bounded, lossless PivotTable-view parts for XLSB authoring.
//!
//! The PivotTable definition stream (MS-XLSB 2.1.7.40) has a large,
//! extensible record grammar. This type deliberately preserves the complete
//! stream instead of projecting it through the older partial model in
//! `pivot_tables.rs`. The package writer uses the parsed view name and cache
//! identifier to validate PivotChart and PivotCache relationships.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::{XlsbRecordHeader, record_types};
use litchi_core::binary;
use std::io::{Cursor, ErrorKind, Read};

const MAX_PIVOT_TABLE_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_PIVOT_TABLE_RECORDS: usize = 1_000_000;

/// A PivotTable definition stream with validated enclosing records, preserved verbatim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbPivotTableViewPart {
    name: String,
    cache_id: u32,
    version_created: u8,
    bytes: Vec<u8>,
}

impl XlsbPivotTableViewPart {
    /// Parse a complete PivotTable part while retaining every original byte.
    ///
    /// The enclosing view collection, its name, cache identifier, record
    /// count, and stream boundaries are validated. Inner records remain
    /// opaque so extension records and features outside Litchi's typed
    /// PivotTable model are not discarded.
    pub fn from_bytes(bytes: Vec<u8>) -> XlsbResult<Self> {
        if bytes.len() > MAX_PIVOT_TABLE_PART_BYTES {
            return Err(XlsbError::InvalidLength {
                expected: MAX_PIVOT_TABLE_PART_BYTES,
                found: bytes.len(),
            });
        }

        let mut cursor = Cursor::new(bytes.as_slice());
        let mut count = 0usize;
        let mut begin = None;
        let mut ended = false;
        while usize::try_from(cursor.position()).unwrap_or(usize::MAX) < bytes.len() {
            count = count.checked_add(1).ok_or_else(|| {
                XlsbError::InvalidFormula("PivotTable record count overflow".to_string())
            })?;
            if count > MAX_PIVOT_TABLE_RECORDS {
                return Err(XlsbError::InvalidFormula(
                    "PivotTable record count exceeds the safety limit".to_string(),
                ));
            }
            let record = read_complete_record(&mut cursor)?;
            match record.record_type {
                record_types::BEGIN_SX_VIEW => {
                    if begin.is_some() || count != 1 || ended {
                        return Err(XlsbError::InvalidFormula(
                            "PivotTable has duplicate or misplaced BrtBeginSXView".to_string(),
                        ));
                    }
                    begin = Some(parse_begin_view(&record.data)?);
                },
                record_types::END_SX_VIEW => {
                    if begin.is_none() || ended || !record.data.is_empty() {
                        return Err(XlsbError::InvalidFormula(
                            "PivotTable has malformed BrtEndSXView".to_string(),
                        ));
                    }
                    ended = true;
                    if usize::try_from(cursor.position()).unwrap_or(usize::MAX) != bytes.len() {
                        return Err(XlsbError::InvalidFormula(
                            "PivotTable has records after BrtEndSXView".to_string(),
                        ));
                    }
                },
                _ if begin.is_none() || ended => {
                    return Err(XlsbError::InvalidFormula(
                        "PivotTable record lies outside BrtBeginSXView collection".to_string(),
                    ));
                },
                _ => {},
            }
        }
        let (name, cache_id, version_created) = begin.ok_or_else(|| {
            XlsbError::InvalidFormula("PivotTable omits BrtBeginSXView".to_string())
        })?;
        if !ended {
            return Err(XlsbError::InvalidFormula(
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
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Workbook PivotCache identifier (`idCache`).
    pub fn cache_id(&self) -> u32 {
        self.cache_id
    }

    /// Data functionality level that created the view (`bVerSxMacro`).
    pub fn version_created(&self) -> u8 {
        self.version_created
    }

    /// Complete original PivotTable definition stream.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }
}

struct CompleteRecord {
    record_type: u16,
    data: Vec<u8>,
}

fn read_complete_record(cursor: &mut Cursor<&[u8]>) -> XlsbResult<CompleteRecord> {
    let start = cursor.position();
    let header = XlsbRecordHeader::read(cursor).map_err(|error| match error {
        XlsbError::Io(io_error) if io_error.kind() == ErrorKind::UnexpectedEof => {
            XlsbError::InvalidFormula(format!(
                "PivotTable has a truncated record header at byte {start}"
            ))
        },
        other => other,
    })?;
    let remaining = cursor
        .get_ref()
        .len()
        .saturating_sub(usize::try_from(cursor.position()).unwrap_or(usize::MAX));
    if header.data_len > remaining {
        return Err(XlsbError::InvalidFormula(format!(
            "PivotTable record {} declares {} bytes with only {remaining} remaining",
            header.record_type, header.data_len
        )));
    }
    let mut data = vec![0u8; header.data_len];
    cursor.read_exact(&mut data)?;
    Ok(CompleteRecord {
        record_type: header.record_type,
        data,
    })
}

fn parse_begin_view(data: &[u8]) -> XlsbResult<(String, u32, u8)> {
    if data.len() < 36 {
        return Err(XlsbError::InvalidLength {
            expected: 36,
            found: data.len(),
        });
    }
    let cache_id = binary::read_u32_le_at(data, 28)?;
    let (name, consumed) = crate::xlsb::records::wide_str_with_len(&data[32..])?;
    if consumed > data.len() - 32 {
        return Err(XlsbError::InvalidFormula(
            "PivotTable view name overruns BrtBeginSXView".to_string(),
        ));
    }
    let units = name.encode_utf16().count();
    if units == 0 || units > 255 || name.contains('\0') {
        return Err(XlsbError::InvalidFormula(
            "PivotTable view name is empty, too long, or contains NUL".to_string(),
        ));
    }
    Ok((name, cache_id, data[0]))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::writer::RecordWriter;

    fn view_stream(name: &str, cache_id: u32) -> Vec<u8> {
        let mut begin = vec![0u8; 32];
        begin[28..32].copy_from_slice(&cache_id.to_le_bytes());
        begin.extend_from_slice(&(name.encode_utf16().count() as u32).to_le_bytes());
        for unit in name.encode_utf16() {
            begin.extend_from_slice(&unit.to_le_bytes());
        }
        let mut bytes = Vec::new();
        let mut writer = RecordWriter::new(&mut bytes);
        writer
            .write_record(record_types::BEGIN_SX_VIEW, &begin)
            .unwrap();
        writer
            .write_record(record_types::BEGIN_SX_LOCATION, &[0; 36])
            .unwrap();
        writer
            .write_record(record_types::END_SX_LOCATION, &[])
            .unwrap();
        writer.write_record(record_types::END_SX_VIEW, &[]).unwrap();
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
        RecordWriter::new(&mut trailing)
            .write_record(record_types::END_SX_LOCATION, &[])
            .unwrap();
        assert!(XlsbPivotTableViewPart::from_bytes(trailing).is_err());
    }
}
