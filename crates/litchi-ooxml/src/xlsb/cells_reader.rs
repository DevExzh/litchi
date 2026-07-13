//! XLSB cells reader implementation

use crate::xlsb::cell::XlsbCell;
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{CellParsedFormula, FormulaGroup, FormulaGroupKind};
use crate::xlsb::hyperlinks::Hyperlink;
use crate::xlsb::merged_cells::MergedCell;
use crate::xlsb::records::{RecordIter, record_types};
use litchi_core::binary;
use litchi_core::sheet::CellValue;
use std::io::{Read, Seek};
use std::sync::Arc;

struct ParsedFormulaCell {
    col: u32,
    cached_value: CellValue,
    formula: CellParsedFormula,
    flags: u16,
}

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
    pending_record: Option<(u16, Vec<u8>)>,
    formula_groups: Vec<Arc<FormulaGroup>>,
    /// Merged cells found in the worksheet
    pub merged_cells: Vec<MergedCell>,
    /// Hyperlinks found in the worksheet
    pub hyperlinks: Vec<Hyperlink>,
}

impl<RS> XlsbCellsReader<RS>
where
    RS: Read + Seek,
{
    pub fn new(mut iter: RecordIter<RS>, shared_strings: Vec<String>) -> XlsbResult<Self> {
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
            pending_record: None,
            formula_groups: Vec::new(),
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
        })
    }

    #[allow(dead_code)]
    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> XlsbResult<Option<XlsbCell>> {
        loop {
            self.buf.clear();
            let typ = if let Some((typ, data)) = self.pending_record.take() {
                self.buf = data;
                typ
            } else {
                let typ = self.iter.read_type()?;
                let _ = self.iter.fill_buffer(&mut self.buf)?;
                typ
            };

            if typ == 0x0092 {
                // BrtEndSheetData - continue to read advanced features
                self.read_advanced_features()?;
                return Ok(None);
            }

            match typ {
                0x0000 => {
                    // BrtRowHdr
                    self.current_row = binary::read_u32_le_at(&self.buf, 0)?;
                },
                0x0001
                    // BrtCellBlank
                    if self.buf.len() >= 4 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        return Ok(Some(XlsbCell::new(self.current_row, col, CellValue::Empty)));
                    },
                0x0002
                    // BrtCellRk
                    if self.buf.len() >= 12 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let rk_val = binary::read_u32_le_at(&self.buf, 8)?;
                        let value = Self::parse_rk_value(rk_val);
                        return Ok(Some(XlsbCell::new(self.current_row, col, value)));
                    },
                0x0003
                    // BrtCellError
                    if self.buf.len() >= 9 => {
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
                        return Ok(Some(XlsbCell::new(
                            self.current_row,
                            col,
                            CellValue::Error(error_msg.to_string()),
                        )));
                    },
                0x0004
                    // BrtCellBool
                    if self.buf.len() >= 9 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let value = self.buf[8] != 0;
                        return Ok(Some(XlsbCell::new(
                            self.current_row,
                            col,
                            CellValue::Bool(value),
                        )));
                    },
                0x0005
                    // BrtCellReal
                    if self.buf.len() >= 16 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let value = binary::read_f64_le_at(&self.buf, 8)?;
                        return Ok(Some(XlsbCell::new(
                            self.current_row,
                            col,
                            CellValue::Float(value),
                        )));
                    },
                0x0006
                    // BrtCellSt
                    if self.buf.len() >= 8 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let (string, _) = super::records::wide_str_with_len(&self.buf[8..])?;
                        return Ok(Some(XlsbCell::new(
                            self.current_row,
                            col,
                            CellValue::String(string),
                        )));
                    },
                0x0007
                    // BrtCellIsst
                    if self.buf.len() >= 12 => {
                        let col = binary::read_u32_le_at(&self.buf, 0)?;
                        let idx = binary::read_u32_le_at(&self.buf, 8)? as usize;
                        let value = if idx < self.shared_strings.len() {
                            CellValue::String(self.shared_strings[idx].clone())
                        } else {
                            CellValue::Error("Invalid SST index".to_string())
                        };
                        return Ok(Some(XlsbCell::new(self.current_row, col, value)));
                    },
                0x0008 => {
                    // BrtFmlaString - formula with string result
                    if self.buf.len() < 12 {
                        return Err(XlsbError::InvalidLength {
                            expected: 12,
                            found: self.buf.len(),
                        });
                    }
                    let col = binary::read_u32_le_at(&self.buf, 0)?;
                    let (string, consumed) =
                        super::records::wide_str_with_len(&self.buf[8..])?;
                    let parsed =
                        self.parse_formula_cell(col, CellValue::String(string), 8 + consumed)?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                0x0009 => {
                    // BrtFmlaNum - formula with numeric result
                    if self.buf.len() < 18 {
                        return Err(XlsbError::InvalidLength {
                            expected: 18,
                            found: self.buf.len(),
                        });
                    }
                    let col = binary::read_u32_le_at(&self.buf, 0)?;
                    let num_value = binary::read_f64_le_at(&self.buf, 8)?;
                    let parsed =
                        self.parse_formula_cell(col, CellValue::Float(num_value), 16)?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                0x000A => {
                    // BrtFmlaBool - formula with boolean result
                    if self.buf.len() < 11 {
                        return Err(XlsbError::InvalidLength {
                            expected: 11,
                            found: self.buf.len(),
                        });
                    }
                    let col = binary::read_u32_le_at(&self.buf, 0)?;
                    let bool_value = self.buf[8] != 0;
                    let parsed = self.parse_formula_cell(col, CellValue::Bool(bool_value), 9)?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                0x000B => {
                    // BrtFmlaError - formula with error result
                    if self.buf.len() < 11 {
                        return Err(XlsbError::InvalidLength {
                            expected: 11,
                            found: self.buf.len(),
                        });
                    }
                    let col = binary::read_u32_le_at(&self.buf, 0)?;
                    let error_msg = Self::error_text(self.buf[8]);
                    let parsed = self.parse_formula_cell(
                        col,
                        CellValue::Error(error_msg.to_string()),
                        9,
                    )?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                record_types::ARR_FMLA | record_types::SHR_FMLA => {
                    return Err(XlsbError::InvalidFormula(
                        "array/shared formula definition is not immediately preceded by a formula cell"
                            .to_string(),
                    ));
                },
                _ => {
                    // Skip unknown records
                },
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

    fn parse_formula_cell(
        &self,
        col: u32,
        cached_value: CellValue,
        flags_offset: usize,
    ) -> XlsbResult<ParsedFormulaCell> {
        let formula_offset = flags_offset.checked_add(2).ok_or_else(|| {
            XlsbError::InvalidFormula("formula record offset overflow".to_string())
        })?;
        if self.buf.len() < formula_offset {
            return Err(XlsbError::InvalidLength {
                expected: formula_offset,
                found: self.buf.len(),
            });
        }
        let flags = binary::read_u16_le_at(&self.buf, flags_offset)?;
        if flags & !0x0002 != 0 {
            return Err(XlsbError::InvalidFormula(format!(
                "invalid GrbitFmla flags 0x{flags:04X}"
            )));
        }
        let (formula, consumed) = CellParsedFormula::parse(&self.buf[formula_offset..])?;
        if formula_offset + consumed != self.buf.len() {
            return Err(XlsbError::InvalidFormula(format!(
                "formula record has {} trailing bytes",
                self.buf.len() - formula_offset - consumed
            )));
        }
        Ok(ParsedFormulaCell {
            col,
            cached_value,
            formula,
            flags,
        })
    }

    fn resolve_formula_record(&mut self, parsed: ParsedFormulaCell) -> XlsbResult<XlsbCell> {
        let position = (self.current_row, parsed.col);
        let exp_cell = parsed.formula.exp_cell()?;

        let next_type = self.iter.read_type()?;
        let mut next_data = Vec::new();
        let _ = self.iter.fill_buffer(&mut next_data)?;
        let new_group = match next_type {
            record_types::ARR_FMLA => Some(FormulaGroup::parse_array(&next_data)?),
            record_types::SHR_FMLA => Some(FormulaGroup::parse_shared(&next_data)?),
            _ => {
                self.pending_record = Some((next_type, next_data));
                None
            },
        };

        let group = if let Some(group) = new_group {
            if exp_cell.is_none() {
                return Err(XlsbError::InvalidFormula(format!(
                    "{:?} formula definition is not preceded by PtgExp",
                    group.kind
                )));
            }
            if exp_cell != Some(position) {
                return Err(XlsbError::InvalidFormula(format!(
                    "group definition at ({}, {}) is referenced by PtgExp ({}, {})",
                    position.0,
                    position.1,
                    exp_cell.map_or(u32::MAX, |target| target.0),
                    exp_cell.map_or(u32::MAX, |target| target.1)
                )));
            }
            if group.formula.rgce.first() == Some(&crate::xlsb::formula::ptg_types::PTG_EXP) {
                return Err(XlsbError::InvalidFormula(
                    "array/shared formula definition cannot contain PtgExp".to_string(),
                ));
            }
            match group.kind {
                FormulaGroupKind::Array if group.range.top_left() != position => {
                    return Err(XlsbError::InvalidFormula(format!(
                        "BrtArrFmla range {} is not anchored at {}",
                        group.range.to_a1(),
                        crate::xlsb::utils::cell_reference(position.0, position.1)
                    )));
                },
                FormulaGroupKind::Shared if group.range.top_left() != position => {
                    return Err(XlsbError::InvalidFormula(format!(
                        "BrtShrFmla range {} is not anchored at {}",
                        group.range.to_a1(),
                        crate::xlsb::utils::cell_reference(position.0, position.1)
                    )));
                },
                _ => {},
            }
            let group = Arc::new(group);
            self.formula_groups.push(Arc::clone(&group));
            Some(group)
        } else if let Some(target) = exp_cell {
            self.formula_groups
                .iter()
                .rev()
                .find(|group| {
                    group.range.top_left() == target && group.range.contains(position.0, position.1)
                })
                .cloned()
        } else {
            None
        };

        if let Some(group) = group {
            Ok(XlsbCell::new_grouped_formula(
                position.0,
                position.1,
                parsed.cached_value,
                parsed.formula,
                parsed.flags,
                group,
            ))
        } else if exp_cell.is_some() {
            Err(XlsbError::InvalidFormula(format!(
                "PtgExp cell {} has no array/shared formula definition",
                crate::xlsb::utils::cell_reference(position.0, position.1)
            )))
        } else {
            Ok(XlsbCell::new_formula_binary(
                position.0,
                position.1,
                parsed.cached_value,
                parsed.formula,
                parsed.flags,
            ))
        }
    }

    fn error_text(error_code: u8) -> &'static str {
        match error_code {
            0x00 => "#NULL!",
            0x07 => "#DIV/0!",
            0x0F => "#VALUE!",
            0x17 => "#REF!",
            0x1D => "#NAME?",
            0x24 => "#NUM!",
            0x2A => "#N/A",
            0x2B => "#GETTING_DATA",
            _ => "#ERR!",
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

    /// Read advanced features after sheet data
    ///
    /// This reads merged cells, hyperlinks, and other advanced features
    /// that appear after the sheet data section.
    fn read_advanced_features(&mut self) -> XlsbResult<()> {
        loop {
            self.buf.clear();

            // Try to read next record type
            let typ = match self.iter.read_type() {
                Ok(t) => t,
                Err(_) => break, // End of stream
            };

            let _ = self.iter.fill_buffer(&mut self.buf)?;

            match typ {
                0x00B1 => {
                    // BrtBeginMergeCells - start of merged cells section
                    self.read_merged_cells()?;
                },
                0x00B0 => {
                    // BrtMergeCell - single merged cell
                    if let Ok(merged) = MergedCell::parse(&self.buf) {
                        self.merged_cells.push(merged);
                    }
                },
                0x01EE => {
                    // BrtHLink - hyperlink
                    if let Ok(hyperlink) = Hyperlink::parse(&self.buf) {
                        self.hyperlinks.push(hyperlink);
                    }
                },
                0x0082 => {
                    // BrtEndSheet - end of worksheet
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }

        Ok(())
    }

    /// Read merged cells section
    fn read_merged_cells(&mut self) -> XlsbResult<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;

            if typ == 0x00B2 {
                // BrtEndMergeCells
                break;
            }

            if typ == 0x00B0 {
                // BrtMergeCell
                if let Ok(merged) = MergedCell::parse(&self.buf) {
                    self.merged_cells.push(merged);
                }
            }
        }

        Ok(())
    }
}
