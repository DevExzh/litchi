//! Semantic XLSB records built on the shared BIFF12 wire kernel.

use crate::named_ranges::Definition;
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Header, Kind, Limits, kind};
use litchi_core::binary;
use std::io::Read;

/// Decode one strict BIFF12 wide string and report consumed bytes.
pub(crate) fn decode_string(buf: &[u8]) -> Result<(String, usize)> {
    let mut cursor = Cursor::new(buf, "XLWideString");
    let value = cursor.read_wide_string()?;
    Ok((value, cursor.position()))
}

/// Private streaming adapter for seek-based worksheet readers.
///
/// Header validation remains owned by `crate::raw::Header`; this adapter
/// only retains the declared length until the caller lends a reusable payload
/// buffer.
pub(crate) struct Stream<R> {
    reader: R,
    pending_len: Option<usize>,
    limits: Limits,
}

impl<R: Read> Stream<R> {
    pub(crate) fn new(reader: R) -> Self {
        Self {
            reader,
            pending_len: None,
            limits: Limits::DEFAULT,
        }
    }

    pub(crate) fn read_type(&mut self) -> Result<Kind> {
        if self.pending_len.is_some() {
            return Err(Error::Unrecognized {
                typ: "BIFF12 stream".to_string(),
                val: "payload must be read before the next header".to_string(),
            });
        }
        let header = Header::read(&mut self.reader, self.limits)?
            .ok_or_else(|| Error::UnexpectedEndOfStream("BIFF12 record header".to_string()))?;
        self.pending_len = Some(header.len());
        Ok(header.kind())
    }

    pub(crate) fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize> {
        let len = self.pending_len.take().ok_or_else(|| Error::Unrecognized {
            typ: "BIFF12 stream".to_string(),
            val: "record header must be read before its payload".to_string(),
        })?;
        buf.resize(len, 0);
        self.reader.read_exact(buf)?;
        Ok(len)
    }
}

/// Workbook properties record.
#[derive(Debug, Clone)]
pub struct WorkbookPropRecord {
    pub is_date1904: bool,
}

impl WorkbookPropRecord {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.is_empty() {
            return Ok(WorkbookPropRecord { is_date1904: false });
        }

        let flags = data[0];
        let is_date1904 = (flags & 0x01) != 0;

        Ok(WorkbookPropRecord { is_date1904 })
    }
}

/// Bundle sheet record (worksheet metadata)
#[derive(Debug, Clone)]
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct BundleSheetRecord {
    pub id: u32,
    pub name: String,
    pub state: u32,
    pub rel_id: Option<String>,
}

impl BundleSheetRecord {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let state = binary::read_u32_le_at(data, 0)?;
        if state > 2 {
            return Err(Error::Unrecognized {
                typ: "BrtBundleSh hsState".to_string(),
                val: state.to_string(),
            });
        }
        let current_id = binary::read_u32_le_at(data, 4)?;
        let (id, strings_offset) = if (1..=0xFFFF).contains(&current_id) {
            (current_id, 8)
        } else {
            // Excel beta XLSB files have an undocumented extra four bytes
            // before iTabID. This layout is also recognized by Apache POI.
            if data.len() < 16 {
                return Err(Error::Unrecognized {
                    typ: "BrtBundleSh iTabID".to_string(),
                    val: current_id.to_string(),
                });
            }
            let beta_id = binary::read_u32_le_at(data, 8)?;
            if !(1..=0xFFFF).contains(&beta_id) {
                return Err(Error::Unrecognized {
                    typ: "BrtBundleSh iTabID".to_string(),
                    val: format!("current {current_id}, beta {beta_id}"),
                });
            }
            (beta_id, 12)
        };

        let (rel_id, rel_consumed) = if binary::read_u32_le_at(data, strings_offset)? == u32::MAX {
            (None, 4)
        } else {
            let (value, consumed) = decode_string(&data[strings_offset..])?;
            if value.is_empty() {
                return Err(Error::Unrecognized {
                    typ: "BrtBundleSh strRelID".to_string(),
                    val: "empty relationship ID".to_string(),
                });
            }
            (Some(value), consumed)
        };
        if rel_id.is_none() && state != 2 {
            return Err(Error::Unrecognized {
                typ: "BrtBundleSh strRelID".to_string(),
                val: "NULL relationship on a sheet that is not very hidden".to_string(),
            });
        }
        let name_offset = strings_offset
            .checked_add(rel_consumed)
            .ok_or_else(|| Error::Encoding("BrtBundleSh relationship size overflow".to_string()))?;
        let (name, name_consumed) = decode_string(&data[name_offset..])?;
        if name_offset + name_consumed != data.len() {
            return Err(Error::Unrecognized {
                typ: "BrtBundleSh".to_string(),
                val: format!(
                    "{} trailing bytes",
                    data.len() - name_offset - name_consumed
                ),
            });
        }
        let name_len = name.encode_utf16().count();
        if name_len == 0
            || name_len > 31
            || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
            || name.starts_with('\'')
            || name.ends_with('\'')
        {
            return Err(Error::Unrecognized {
                typ: "BrtBundleSh strName".to_string(),
                val: name,
            });
        }

        Ok(BundleSheetRecord {
            id,
            name,
            state,
            rel_id,
        })
    }
}

/// Row header record
#[derive(Debug, Clone)]
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
pub struct RowHeaderRecord {
    pub row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
impl RowHeaderRecord {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let first_col = binary::read_u16_le_at(data, 4)?;
        let last_col = binary::read_u16_le_at(data, 6)?;

        Ok(RowHeaderRecord {
            row,
            first_col,
            last_col,
        })
    }
}

/// Cell value types
#[derive(Debug, Clone)]
pub enum CellValue {
    Blank,
    Bool(bool),
    Error(u8),
    Real(f64),
    String(String),
    Isst(u32), // Index into shared string table
    Formula {
        value: Box<CellValue>,
        formula: Option<Vec<u8>>, // Raw formula bytes
    },
}

/// Cell record base
#[derive(Debug, Clone)]
pub struct CellRecord {
    pub row: u32,
    pub col: u16,
    pub value: CellValue,
}

impl CellRecord {
    pub fn parse(record_type: Kind, data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 4)?;

        let value = match record_type {
            kind::CELL_BLANK => CellValue::Blank,
            kind::CELL_BOOL => {
                if data.len() < 7 {
                    return Err(Error::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                let mut cursor = Cursor::new(&data[6..], "BrtCellBool");
                CellValue::Bool(cursor.read_bool8()?)
            },
            kind::CELL_ERROR => {
                if data.len() < 7 {
                    return Err(Error::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                CellValue::Error(data[6])
            },
            kind::CELL_REAL => {
                if data.len() < 14 {
                    return Err(Error::InvalidLength {
                        expected: 14,
                        found: data.len(),
                    });
                }
                CellValue::Real(binary::read_f64_le_at(data, 6)?)
            },
            kind::CELL_ST => {
                let (string, _) = decode_string(&data[6..])?;
                CellValue::String(string)
            },
            kind::CELL_ISST => {
                if data.len() < 10 {
                    return Err(Error::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                CellValue::Isst(binary::read_u32_le_at(data, 6)?)
            },
            kind::CELL_RK => {
                if data.len() < 10 {
                    return Err(Error::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                let mut cursor = Cursor::new(&data[6..], "BrtCellRk");
                CellValue::Real(cursor.read_rk()?)
            },
            // Formula records - parse formula bytes and cached value
            kind::FMLA_STRING => {
                if data.len() < 10 {
                    return Err(Error::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                // Skip style_id (4 bytes) and flags (1 byte) and formula length (4 bytes)
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len {
                    return Err(Error::InvalidLength {
                        expected: 10 + formula_len,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();

                // Read cached string value after formula
                let (string, _) = decode_string(&data[10 + formula_len..])?;
                CellValue::Formula {
                    value: Box::new(CellValue::String(string)),
                    formula: Some(formula_bytes),
                }
            },
            kind::FMLA_NUM => {
                if data.len() < 18 {
                    return Err(Error::InvalidLength {
                        expected: 18,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 8 {
                    return Err(Error::InvalidLength {
                        expected: 10 + formula_len + 8,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let num_value = binary::read_f64_le_at(data, 10 + formula_len)?;
                CellValue::Formula {
                    value: Box::new(CellValue::Real(num_value)),
                    formula: Some(formula_bytes),
                }
            },
            kind::FMLA_BOOL => {
                if data.len() < 11 {
                    return Err(Error::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(Error::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let mut cursor = Cursor::new(&data[10 + formula_len..], "BrtFmlaBool cached value");
                let bool_value = cursor.read_bool8()?;
                CellValue::Formula {
                    value: Box::new(CellValue::Bool(bool_value)),
                    formula: Some(formula_bytes),
                }
            },
            kind::FMLA_ERROR => {
                if data.len() < 11 {
                    return Err(Error::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(Error::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let error_code = data[10 + formula_len];
                CellValue::Formula {
                    value: Box::new(CellValue::Error(error_code)),
                    formula: Some(formula_bytes),
                }
            },
            _ => return Err(Error::InvalidRecordType(record_type.get())),
        };

        Ok(CellRecord { row, col, value })
    }
}

/// Column information record
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
#[derive(Debug, Clone)]
pub struct ColInfoRecord {
    pub first_col: u32,
    pub last_col: u32,
    pub width: f64,
    pub style_xf: u32,
    pub custom_width: bool,
    pub hidden: bool,
    pub best_fit: bool,
}

impl ColInfoRecord {
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let first_col = binary::read_u32_le_at(data, 0)?;
        let last_col = binary::read_u32_le_at(data, 4)?;
        // Width is stored as 256ths of a character
        let width_raw = binary::read_u32_le_at(data, 8)?;
        let width = width_raw as f64 / 256.0;

        let style_xf = if data.len() >= 16 {
            binary::read_u32_le_at(data, 12)?
        } else {
            0
        };

        let flags = if data.len() >= 18 {
            binary::read_u16_le_at(data, 16)?
        } else {
            0
        };

        let custom_width = (flags & 0x0002) != 0;
        let hidden = (flags & 0x0001) != 0;
        let best_fit = (flags & 0x0008) != 0;

        Ok(ColInfoRecord {
            first_col,
            last_col,
            width,
            style_xf,
            custom_width,
            hidden,
            best_fit,
        })
    }
}

/// Merged cell record
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
#[derive(Debug, Clone)]
pub struct MergeCellRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl MergeCellRecord {
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        Ok(MergeCellRecord {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        })
    }
}

/// Hyperlink record
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
#[derive(Debug, Clone)]
pub struct HyperlinkRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
    pub r_id: String,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

impl HyperlinkRecord {
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 16 {
            return Err(Error::InvalidLength {
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
        let (r_id, consumed) = decode_string(&data[offset..])?;
        offset += consumed;

        // Read location (optional)
        let (location, consumed) = if offset < data.len() {
            let (loc, c) = decode_string(&data[offset..])?;
            (if loc.is_empty() { None } else { Some(loc) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read tooltip (optional)
        let (tooltip, consumed) = if offset < data.len() {
            let (tt, c) = decode_string(&data[offset..])?;
            (if tt.is_empty() { None } else { Some(tt) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read display text (optional)
        let display = if offset < data.len() {
            let (disp, _) = decode_string(&data[offset..])?;
            if disp.is_empty() { None } else { Some(disp) }
        } else {
            None
        };

        Ok(HyperlinkRecord {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location,
            tooltip,
            display,
        })
    }
}

/// Named range record
#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
#[derive(Debug, Clone)]
pub struct NameRecord {
    pub name: String,
    pub formula: Option<Vec<u8>>,
    pub sheet_id: Option<u32>,
    pub hidden: bool,
    pub function: bool,
}

impl NameRecord {
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub fn parse(data: &[u8]) -> Result<Self> {
        let named_range = Definition::parse(data)?;

        Ok(NameRecord {
            name: named_range.name,
            formula: named_range.formula,
            sheet_id: named_range.sheet_id,
            hidden: named_range.hidden,
            function: named_range.function,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn wide_string(value: &str) -> Vec<u8> {
        let mut data = (value.encode_utf16().count() as u32).to_le_bytes().to_vec();
        for unit in value.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_current_and_excel_beta_bundle_sheets() {
        let mut current = 0u32.to_le_bytes().to_vec();
        current.extend_from_slice(&7u32.to_le_bytes());
        current.extend_from_slice(&wide_string("rId7"));
        current.extend_from_slice(&wide_string("Data"));
        let sheet = BundleSheetRecord::parse(&current).unwrap();
        assert_eq!(sheet.id, 7);
        assert_eq!(sheet.state, 0);
        assert_eq!(sheet.rel_id.as_deref(), Some("rId7"));
        assert_eq!(sheet.name, "Data");

        let mut beta = 0u64.to_le_bytes().to_vec();
        beta.extend_from_slice(&8u32.to_le_bytes());
        beta.extend_from_slice(&wide_string("rId8"));
        beta.extend_from_slice(&wide_string("Legacy"));
        let sheet = BundleSheetRecord::parse(&beta).unwrap();
        assert_eq!(sheet.id, 8);
        assert_eq!(sheet.rel_id.as_deref(), Some("rId8"));
        assert_eq!(sheet.name, "Legacy");
    }

    #[test]
    fn rejects_malformed_bundle_sheet_metadata() {
        let mut invalid = 0u32.to_le_bytes().to_vec();
        invalid.extend_from_slice(&0u32.to_le_bytes());
        invalid.extend_from_slice(&0u32.to_le_bytes());
        assert!(matches!(
            BundleSheetRecord::parse(&invalid),
            Err(Error::Unrecognized { .. })
        ));

        let mut null_visible = 0u32.to_le_bytes().to_vec();
        null_visible.extend_from_slice(&1u32.to_le_bytes());
        null_visible.extend_from_slice(&u32::MAX.to_le_bytes());
        null_visible.extend_from_slice(&wide_string("Module"));
        assert!(matches!(
            BundleSheetRecord::parse(&null_visible),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn cell_records_delegate_strict_boolean_and_rk_values_to_the_wire_cursor() {
        let mut rk = 4_u32.to_le_bytes().to_vec();
        rk.extend_from_slice(&2_u16.to_le_bytes());
        rk.extend_from_slice(&((((-42_i32) as u32) << 2) | 0x02).to_le_bytes());
        assert!(matches!(
            CellRecord::parse(kind::CELL_RK, &rk).unwrap().value,
            CellValue::Real(value) if value == -42.0
        ));

        let mut scaled = 5_u32.to_le_bytes().to_vec();
        scaled.extend_from_slice(&3_u16.to_le_bytes());
        scaled.extend_from_slice(&((1234_u32 << 2) | 0x03).to_le_bytes());
        assert!(matches!(
            CellRecord::parse(kind::CELL_RK, &scaled).unwrap().value,
            CellValue::Real(value) if value == 12.34
        ));

        let mut invalid_bool = 6_u32.to_le_bytes().to_vec();
        invalid_bool.extend_from_slice(&4_u16.to_le_bytes());
        invalid_bool.push(2);
        assert!(matches!(
            CellRecord::parse(kind::CELL_BOOL, &invalid_bool),
            Err(Error::Wire(crate::Error::InvalidBool { value: 2, .. }))
        ));

        let mut invalid_formula_bool = 7_u32.to_le_bytes().to_vec();
        invalid_formula_bool.extend_from_slice(&5_u16.to_le_bytes());
        invalid_formula_bool.extend_from_slice(&0_u32.to_le_bytes());
        invalid_formula_bool.push(2);
        assert!(matches!(
            CellRecord::parse(kind::FMLA_BOOL, &invalid_formula_bool),
            Err(Error::Wire(crate::Error::InvalidBool { value: 2, .. }))
        ));
    }
}
