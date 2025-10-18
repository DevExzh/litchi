//! XLSB cells reader implementation

use std::io::{Read, Seek};
use crate::sheet::CellValue;
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::RecordIter;
use crate::ooxml::xlsb::cell::XlsbCell;
use crate::ole::binary;

/// Dimensions of a worksheet
#[derive(Debug, Clone, Copy)]
#[allow(dead_code)]
pub struct Dimensions {
    pub start: (u32, u32),
    pub end: (u32, u32),
}

#[allow(dead_code)]
impl Dimensions {
    pub fn len(&self) -> usize {
        ((self.end.0 - self.start.0 + 1) * (self.end.1 - self.start.1 + 1)) as usize
    }
}

/// XLSB cells reader
#[allow(dead_code)]
pub struct XlsbCellsReader<RS>
where
    RS: Read + Seek,
{
    iter: RecordIter<RS>,
    shared_strings: Vec<String>,
    dimensions: Dimensions,
    current_row: u32,
    buf: Vec<u8>,
}

impl<RS> XlsbCellsReader<RS>
where
    RS: Read + Seek,
{
    pub fn new(
        mut iter: RecordIter<RS>,
        shared_strings: Vec<String>,
    ) -> XlsbResult<Self> {
        let mut buf = Vec::with_capacity(1024);

        // Skip to BrtWsDim (worksheet dimensions)
        let _ = iter.next_skip_blocks(
            0x0094, // BrtWsDim
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
        let dimensions = Self::parse_dimensions(&buf[..16]);

        // Skip to BrtBeginSheetData
        let _ = iter.next_skip_blocks(
            0x0091, // BrtBeginSheetData
            &[
                (0x0085, Some(0x0086)), // Views
                (0x0025, Some(0x0026)), // AC blocks
                (0x01E5, None),         // BrtWsFmtInfo
                (0x0186, Some(0x0187)), // Col Infos
            ],
            &mut buf,
        )?;

        Ok(XlsbCellsReader {
            iter,
            shared_strings,
            dimensions,
            current_row: 0,
            buf,
        })
    }

    #[allow(dead_code)]
    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> XlsbResult<Option<XlsbCell>> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;

            if typ == 0x0092 {
                // BrtEndSheetData
                return Ok(None);
            }

            let _ = self.iter.fill_buffer(&mut self.buf)?;

            match typ {
                0x0000 => {
                    // BrtRowHdr
                    self.current_row = binary::read_u32_le_at(&self.buf, 0)?;
                }
                0x0001 => {
                    // BrtCellBlank
                    if self.buf.len() >= 6 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::Empty)));
                    }
                }
                0x0002 => {
                    // BrtCellRk
                    if self.buf.len() >= 12 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let rk_val = binary::read_u32_le_at(&self.buf, 8)?;
                        let value = Self::parse_rk_value(rk_val);
                        return Ok(Some(XlsbCell::new(self.current_row, col, value)));
                    }
                }
                0x0003 => {
                    // BrtCellError
                    if self.buf.len() >= 9 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let error_code = self.buf[8];
                        let error_msg = match error_code {
                            0x00 => "#NULL!",
                            0x07 => "#DIV/0!",
                            0x0F => "#VALUE!",
                            0x17 => "#REF!",
                            0x1D => "#NAME?",
                            0x24 => "#NUM!",
                            0x2A => "#N/A",
                            0x2B => "#GETTING_DATA",
                            _ => "#ERR!",
                        };
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::Error(error_msg.to_string()))));
                    }
                }
                0x0004 => {
                    // BrtCellBool
                    if self.buf.len() >= 9 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let value = self.buf[8] != 0;
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::Bool(value))));
                    }
                }
                0x0005 => {
                    // BrtCellReal
                    if self.buf.len() >= 16 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let value = binary::read_f64_le_at(&self.buf, 8)?;
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::Float(value))));
                    }
                }
                0x0006 => {
                    // BrtCellSt
                    if self.buf.len() >= 8 {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let (string, _) = super::records::wide_str_with_len(&self.buf[8..])?;
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::String(string))));
                    }
                }
                       0x0007 => {
                           // BrtCellIsst
                           if self.buf.len() >= 12 {
                               let col = binary::read_u32_le_at(&self.buf, 0)?;
                               let idx = binary::read_u32_le_at(&self.buf, 8)? as usize;
                               let value = if idx < self.shared_strings.len() {
                                   CellValue::String(self.shared_strings[idx].clone())
                               } else {
                                   CellValue::Error("Invalid SST index".to_string())
                               };
                               return Ok(Some(XlsbCell::new(self.current_row, col, value)));
                           }
                       }
                _ => {
                    // Skip unknown records
                }
            }
        }
    }

    fn parse_dimensions(buf: &[u8]) -> Dimensions {
        Dimensions {
            start: (
                binary::read_u32_le_at(buf, 0).unwrap_or(0),
                binary::read_u32_le_at(buf, 8).unwrap_or(0),
            ),
            end: (
                binary::read_u32_le_at(buf, 4).unwrap_or(0),
                binary::read_u32_le_at(buf, 12).unwrap_or(0),
            ),
        }
    }

    fn parse_rk_value(rk: u32) -> CellValue {
        let d100 = (rk & 0x02) != 0;
        let is_int = (rk & 0x01) != 0;

        if is_int {
            let int_val = (rk >> 2) as i32;
            let value = if d100 {
                if int_val % 100 != 0 {
                    int_val as f64 / 100.0
                } else {
                    (int_val / 100) as f64
                }
            } else {
                int_val as f64
            };
            CellValue::Int(value as i64)
        } else {
            let mut float_bits = [0u8; 8];
            let masked_rk = rk & 0xFFFFFFFC;
            // RK floats use the lower 30 bits as the upper 32 bits of a double
            // In little-endian, this goes in the last 4 bytes
            float_bits[4..8].copy_from_slice(&masked_rk.to_le_bytes());
            let mut value = f64::from_le_bytes(float_bits);
            value = if d100 { value / 100.0 } else { value };

            // Check if it's a whole number
            if value == value.round() && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
                CellValue::Int(value as i64)
            } else {
                CellValue::Float(value)
            }
        }
    }
}
