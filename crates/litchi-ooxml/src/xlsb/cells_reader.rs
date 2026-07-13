//! XLSB cells reader implementation

use crate::xlsb::cell::{CellHeader, XlsbCell};
use crate::xlsb::conditional_formatting::{
    Cfvo, ColorScale, ConditionalFormatColor, ConditionalFormatting, ConditionalFormattingRule,
    DataBar, IconSet, parse_classic_header,
};
use crate::xlsb::data_validation::{
    DataValidation, DataValidationSettings, parse_collection_settings, parse_dval_list,
};
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaGroup, FormulaGroupKind, FormulaResolutionContext,
};
use crate::xlsb::hyperlinks::Hyperlink;
use crate::xlsb::merged_cells::MergedCell;
use crate::xlsb::records::{RecordIter, record_types};
use crate::xlsb::shared_strings::SharedString;
use crate::xlsb::worksheet::{
    XlsbAutoFilter, XlsbColumnInfo, XlsbRowInfo, XlsbSheetProtection, XlsbStrongProtection,
};
use litchi_core::binary;
use litchi_core::sheet::CellValue;
use std::io::{Read, Seek};
use std::sync::Arc;

struct ParsedFormulaCell {
    header: CellHeader,
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
pub struct XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    iter: RecordIter<RS>,
    shared_strings: &'a [SharedString],
    formula_context: &'a FormulaResolutionContext,
    cell_xf_count: usize,
    dimensions: Dimensions,
    current_row: u32,
    last_row: Option<u32>,
    buf: Vec<u8>,
    pending_record: Option<(u16, Vec<u8>)>,
    formula_groups: Vec<Arc<FormulaGroup>>,
    /// Merged cells found in the worksheet
    pub merged_cells: Vec<MergedCell>,
    /// Hyperlinks found in the worksheet
    pub hyperlinks: Vec<Hyperlink>,
    /// Column formatting records found before sheet data.
    pub column_infos: Vec<XlsbColumnInfo>,
    /// Row header metadata found within sheet data.
    pub row_infos: Vec<XlsbRowInfo>,
    /// Worksheet AutoFilter range.
    pub auto_filter: Option<XlsbAutoFilter>,
    /// Worksheet protection settings.
    pub sheet_protection: Option<XlsbSheetProtection>,
    /// ISO strong password-verifier metadata.
    pub strong_sheet_protection: Option<XlsbStrongProtection>,
    /// Classic worksheet data-validation rules.
    pub data_validations: Vec<DataValidation>,
    /// UI settings from the classic validation collection.
    pub data_validation_settings: Option<DataValidationSettings>,
    /// UI settings from the Office 2013 validation collection.
    pub data_validation14_settings: Option<DataValidationSettings>,
    /// Classic conditional-formatting blocks in stream order.
    pub conditional_formattings: Vec<ConditionalFormatting>,
}

impl<'a, RS> XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub fn new(
        mut iter: RecordIter<RS>,
        shared_strings: &'a [SharedString],
        formula_context: &'a FormulaResolutionContext,
        cell_xf_count: usize,
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
        if buf.len() != 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: buf.len(),
            });
        }
        let dimensions = Self::parse_dimensions(&buf);

        // Read worksheet preamble through BrtBeginSheetData, retaining column
        // formatting while safely ignoring unrelated and future records.
        let mut column_infos = Vec::new();
        loop {
            let typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;
            if typ == 0x0091 {
                break;
            }
            if typ == record_types::COL_INFO {
                column_infos.push(Self::parse_column_info(&buf, cell_xf_count)?);
            }
        }

        Ok(XlsbCellsReader {
            iter,
            shared_strings,
            formula_context,
            cell_xf_count,
            dimensions,
            current_row: 0,
            last_row: None,
            buf,
            pending_record: None,
            formula_groups: Vec::new(),
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            column_infos,
            row_infos: Vec::new(),
            auto_filter: None,
            sheet_protection: None,
            strong_sheet_protection: None,
            data_validations: Vec::new(),
            data_validation_settings: None,
            data_validation14_settings: None,
            conditional_formattings: Vec::new(),
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

            let cell_header = if matches!(typ, 0x0001..=0x000B | record_types::CELL_R_STRING) {
                Some(self.parse_cell_header()?)
            } else {
                None
            };

            match typ {
                0x0000 => {
                    // BrtRowHdr
                    let info = Self::parse_row_info(
                        &self.buf,
                        self.cell_xf_count,
                        self.last_row,
                    )?;
                    self.current_row = info.row;
                    self.last_row = Some(info.row);
                    self.row_infos.push(info);
                },
                0x0001
                    // BrtCellBlank
                    if self.buf.len() >= 8 => {
                        let header = cell_header.unwrap();
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Empty,
                        )));
                    },
                0x0002
                    // BrtCellRk
                    if self.buf.len() >= 12 => {
                        let header = cell_header.unwrap();
                        let rk_val = binary::read_u32_le_at(&self.buf, 8)?;
                        let value = Self::parse_rk_value(rk_val);
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            value,
                        )));
                    },
                0x0003
                    // BrtCellError
                    if self.buf.len() >= 9 => {
                        let header = cell_header.unwrap();
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
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Error(error_msg.to_string()),
                        )));
                    },
                0x0004
                    // BrtCellBool
                    if self.buf.len() >= 9 => {
                        let header = cell_header.unwrap();
                        let value = self.buf[8] != 0;
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Bool(value),
                        )));
                    },
                0x0005
                    // BrtCellReal
                    if self.buf.len() >= 16 => {
                        let header = cell_header.unwrap();
                        let value = binary::read_f64_le_at(&self.buf, 8)?;
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Float(value),
                        )));
                    },
                0x0006
                    // BrtCellSt
                    if self.buf.len() >= 8 => {
                        let header = cell_header.unwrap();
                        let (string, _) = super::records::wide_str_with_len(&self.buf[8..])?;
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            CellValue::String(string),
                        )));
                    },
                0x0007
                    // BrtCellIsst
                    if self.buf.len() >= 12 => {
                        let header = cell_header.unwrap();
                        let idx = binary::read_u32_le_at(&self.buf, 8)? as usize;
                        let value = if idx < self.shared_strings.len() {
                            CellValue::String(self.shared_strings[idx].text.clone())
                        } else {
                            CellValue::Error("Invalid SST index".to_string())
                        };
                        return Ok(Some(XlsbCell::new_styled(
                            self.current_row,
                            header,
                            value,
                        )));
                    },
                record_types::CELL_R_STRING => {
                    if self.buf.len() < 9 {
                        return Err(XlsbError::InvalidLength {
                            expected: 9,
                            found: self.buf.len(),
                        });
                    }
                    let header = cell_header.unwrap();
                    let rich_string = SharedString::parse(&self.buf[8..])?;
                    return Ok(Some(XlsbCell::new_rich_string(
                        self.current_row,
                        header,
                        rich_string,
                    )));
                },
                0x0008 => {
                    // BrtFmlaString - formula with string result
                    if self.buf.len() < 12 {
                        return Err(XlsbError::InvalidLength {
                            expected: 12,
                            found: self.buf.len(),
                        });
                    }
                    let header = cell_header.unwrap();
                    let (string, consumed) =
                        super::records::wide_str_with_len(&self.buf[8..])?;
                    let parsed =
                        self.parse_formula_cell(header, CellValue::String(string), 8 + consumed)?;
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
                    let header = cell_header.unwrap();
                    let num_value = binary::read_f64_le_at(&self.buf, 8)?;
                    let parsed =
                        self.parse_formula_cell(header, CellValue::Float(num_value), 16)?;
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
                    let header = cell_header.unwrap();
                    let bool_value = self.buf[8] != 0;
                    let parsed = self.parse_formula_cell(
                        header,
                        CellValue::Bool(bool_value),
                        9,
                    )?;
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
                    let header = cell_header.unwrap();
                    let error_msg = Self::error_text(self.buf[8]);
                    let parsed = self.parse_formula_cell(
                        header,
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

    fn parse_column_info(data: &[u8], cell_xf_count: usize) -> XlsbResult<XlsbColumnInfo> {
        if data.len() != 18 {
            return Err(XlsbError::InvalidLength {
                expected: 18,
                found: data.len(),
            });
        }
        let first_column = binary::read_u32_le_at(data, 0)?;
        let last_column = binary::read_u32_le_at(data, 4)?;
        let width_raw = binary::read_u32_le_at(data, 8)?;
        let style_id = binary::read_u32_le_at(data, 12)?;
        let flags = binary::read_u16_le_at(data, 16)?;
        if first_column > last_column || last_column >= 0x4000 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtColInfo range".to_string(),
                val: format!("{first_column}..={last_column}"),
            });
        }
        if width_raw > 0xFFFF {
            return Err(XlsbError::Unrecognized {
                typ: "BrtColInfo coldx".to_string(),
                val: width_raw.to_string(),
            });
        }
        if style_id as usize >= cell_xf_count {
            return Err(XlsbError::Unrecognized {
                typ: "BrtColInfo ixfe".to_string(),
                val: format!("{style_id} (cell XF count {cell_xf_count})"),
            });
        }
        if flags & !0x170F != 0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtColInfo flags".to_string(),
                val: format!("0x{flags:04X}"),
            });
        }
        Ok(XlsbColumnInfo {
            first_column,
            last_column,
            width: f64::from(width_raw) / 256.0,
            style_id,
            hidden: flags & 0x0001 != 0,
            user_set_width: flags & 0x0002 != 0,
            best_fit: flags & 0x0004 != 0,
            show_phonetic: flags & 0x0008 != 0,
            outline_level: ((flags >> 8) & 7) as u8,
            collapsed: flags & 0x1000 != 0,
        })
    }

    fn parse_row_info(
        data: &[u8],
        cell_xf_count: usize,
        previous_row: Option<u32>,
    ) -> XlsbResult<XlsbRowInfo> {
        if data.len() < 17 {
            return Err(XlsbError::InvalidLength {
                expected: 17,
                found: data.len(),
            });
        }
        let row = binary::read_u32_le_at(data, 0)?;
        let raw_style_id = binary::read_u32_le_at(data, 4)?;
        let height_twips = binary::read_u16_le_at(data, 8)?;
        let flags1 = data[10];
        let flags2 = data[11];
        let phonetic_flags = data[12];
        let span_count = binary::read_u32_le_at(data, 13)? as usize;
        if span_count > 16 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtRowHdr ccolspan".to_string(),
                val: span_count.to_string(),
            });
        }
        let expected = span_count
            .checked_mul(8)
            .and_then(|size| size.checked_add(17))
            .ok_or_else(|| XlsbError::Encoding("BrtRowHdr size overflow".to_string()))?;
        if data.len() != expected {
            return Err(XlsbError::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        if row >= 0x10_0000 || previous_row.is_some_and(|previous| row <= previous) {
            return Err(XlsbError::Unrecognized {
                typ: "BrtRowHdr rw".to_string(),
                val: row.to_string(),
            });
        }
        if height_twips > 0x2000 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtRowHdr miyRw".to_string(),
                val: height_twips.to_string(),
            });
        }
        if flags1 & 0xFC != 0 || flags2 & 0x80 != 0 || phonetic_flags & 0xFE != 0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtRowHdr flags".to_string(),
                val: format!("0x{flags1:02X}/0x{flags2:02X}/0x{phonetic_flags:02X}"),
            });
        }
        let style_applied = flags2 & 0x40 != 0;
        if style_applied && raw_style_id as usize >= cell_xf_count {
            return Err(XlsbError::Unrecognized {
                typ: "BrtRowHdr ixfe".to_string(),
                val: format!("{raw_style_id} (cell XF count {cell_xf_count})"),
            });
        }

        let mut column_spans = Vec::with_capacity(span_count);
        let mut previous_segment = None;
        for span in data[17..].chunks_exact(8) {
            let first = binary::read_u32_le_at(span, 0)?;
            let last = binary::read_u32_le_at(span, 4)?;
            let segment = first / 1024;
            if first > last
                || last >= 0x4000
                || last / 1024 != segment
                || previous_segment.is_some_and(|previous| segment <= previous)
            {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtRowHdr BrtColSpan".to_string(),
                    val: format!("{first}..={last}"),
                });
            }
            previous_segment = Some(segment);
            column_spans.push((first, last));
        }

        Ok(XlsbRowInfo {
            row,
            style_id: style_applied.then_some(raw_style_id),
            height: (flags2 & 0x20 != 0).then_some(f64::from(height_twips) / 20.0),
            extra_ascender: flags1 & 1 != 0,
            extra_descender: flags1 & 2 != 0,
            outline_level: flags2 & 7,
            collapsed: flags2 & 0x08 != 0,
            hidden: flags2 & 0x10 != 0,
            show_phonetic: phonetic_flags & 1 != 0,
            column_spans,
        })
    }

    fn parse_cell_header(&self) -> XlsbResult<CellHeader> {
        Self::decode_cell_header(&self.buf, self.cell_xf_count)
    }

    fn decode_cell_header(data: &[u8], cell_xf_count: usize) -> XlsbResult<CellHeader> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let flags = data[7];
        if flags & 0xFE != 0 {
            return Err(XlsbError::Unrecognized {
                typ: "Cell flags".to_string(),
                val: format!("0x{flags:02X}"),
            });
        }
        let col = binary::read_u32_le_at(data, 0)?;
        let style_id = u32::from(data[4]) | (u32::from(data[5]) << 8) | (u32::from(data[6]) << 16);
        if style_id as usize >= cell_xf_count {
            return Err(XlsbError::Unrecognized {
                typ: "Cell iStyleRef".to_string(),
                val: format!("{style_id} (cell XF count {cell_xf_count})"),
            });
        }
        Ok(CellHeader {
            col,
            style_id,
            show_phonetic: flags & 1 != 0,
        })
    }

    fn parse_formula_cell(
        &self,
        header: CellHeader,
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
            header,
            cached_value,
            formula,
            flags,
        })
    }

    fn resolve_formula_record(&mut self, parsed: ParsedFormulaCell) -> XlsbResult<XlsbCell> {
        let position = (self.current_row, parsed.header.col);
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
                parsed.header,
                parsed.cached_value,
                parsed.formula,
                parsed.flags,
                group,
                self.formula_context,
            ))
        } else if exp_cell.is_some() {
            Err(XlsbError::InvalidFormula(format!(
                "PtgExp cell {} has no array/shared formula definition",
                crate::xlsb::utils::cell_reference(position.0, position.1)
            )))
        } else {
            Ok(XlsbCell::new_formula_binary(
                position.0,
                parsed.header,
                parsed.cached_value,
                parsed.formula,
                parsed.flags,
                self.formula_context,
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
                record_types::BEGIN_A_FILTER => {
                    if self.auto_filter.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginAFilter".to_string(),
                            val: "duplicate worksheet AutoFilter".to_string(),
                        });
                    }
                    self.auto_filter = Some(Self::parse_auto_filter(&self.buf)?);
                    self.consume_auto_filter_records()?;
                },
                record_types::SHEET_PROTECTION => {
                    if self.sheet_protection.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtSheetProtection".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    self.sheet_protection = Some(Self::parse_sheet_protection(&self.buf)?);
                },
                record_types::SHEET_PROTECTION_ISO => {
                    if self.sheet_protection.is_some() || self.strong_sheet_protection.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "duplicate protection record".to_string(),
                        });
                    }
                    let (strong, iso_flags) = Self::parse_strong_sheet_protection(&self.buf)?;
                    self.buf.clear();
                    let next_type = self.iter.read_type()?;
                    let _ = self.iter.fill_buffer(&mut self.buf)?;
                    if next_type != record_types::SHEET_PROTECTION {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "not immediately followed by BrtSheetProtection".to_string(),
                        });
                    }
                    let base = Self::parse_sheet_protection(&self.buf)?;
                    if base.password_hash.is_some()
                        || Self::sheet_protection_flags(&base) != iso_flags
                    {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "following protection record does not match".to_string(),
                        });
                    }
                    self.sheet_protection = Some(base);
                    self.strong_sheet_protection = Some(strong);
                },
                record_types::BEGIN_D_VALS => {
                    if self.data_validation_settings.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginDVals".to_string(),
                            val: "duplicate collection".to_string(),
                        });
                    }
                    let (settings, count) = parse_collection_settings(&self.buf, false)?;
                    self.data_validation_settings = Some(settings);
                    self.consume_classic_data_validations(count)?;
                },
                record_types::BEGIN_D_VALS14 => {
                    if self.data_validation14_settings.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginDVals14".to_string(),
                            val: "duplicate collection".to_string(),
                        });
                    }
                    let (settings, count) = parse_collection_settings(&self.buf, true)?;
                    self.data_validation14_settings = Some(settings);
                    self.consume_extension_data_validations(count)?;
                },
                record_types::BEGIN_COND_FORMATTING => {
                    let (mut formatting, count, base) = parse_classic_header(&self.buf)?;
                    self.consume_conditional_formatting(&mut formatting, count, base)?;
                    self.conditional_formattings.push(formatting);
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

    fn consume_classic_data_validations(&mut self, expected_count: u32) -> XlsbResult<()> {
        let start = self.data_validations.len();
        let mut pending_list = None;
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::D_VAL_LIST => {
                    if pending_list.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtDValList".to_string(),
                            val: "consecutive list overrides".to_string(),
                        });
                    }
                    pending_list = Some(parse_dval_list(&self.buf)?);
                },
                record_types::D_VAL => {
                    let rule = DataValidation::parse_classic(
                        &self.buf,
                        pending_list.take(),
                        self.formula_context,
                    )?;
                    self.data_validations.push(rule);
                },
                record_types::END_D_VALS => {
                    if !self.buf.is_empty() {
                        return Err(XlsbError::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    if pending_list.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtDValList".to_string(),
                            val: "not followed by BrtDVal".to_string(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginDVals collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let found = self.data_validations.len() - start;
        if found != expected_count as usize {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginDVals count".to_string(),
                val: format!("declared {expected_count}, found {found}"),
            });
        }
        Ok(())
    }

    fn consume_extension_data_validations(&mut self, expected_count: u32) -> XlsbResult<()> {
        let start = self.data_validations.len();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::D_VAL14 => {
                    self.data_validations
                        .push(DataValidation::parse_extension14(
                            &self.buf,
                            self.formula_context,
                        )?)
                },
                record_types::END_D_VALS14 => {
                    if !self.buf.is_empty() {
                        return Err(XlsbError::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginDVals14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let found = self.data_validations.len() - start;
        if found != expected_count as usize {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginDVals14 count".to_string(),
                val: format!("declared {expected_count}, found {found}"),
            });
        }
        Ok(())
    }

    fn consume_conditional_formatting(
        &mut self,
        formatting: &mut ConditionalFormatting,
        expected_count: u32,
        base: (u32, u32),
    ) -> XlsbResult<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::BEGIN_CF_RULE => {
                    let mut rule = ConditionalFormattingRule::parse_with_context(
                        &self.buf,
                        base,
                        self.formula_context,
                    )?;
                    if formatting
                        .rules
                        .iter()
                        .chain(self.conditional_formattings.iter().flat_map(|cf| &cf.rules))
                        .any(|existing| existing.priority == rule.priority)
                    {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginCFRule priority".to_string(),
                            val: format!("duplicate {}", rule.priority),
                        });
                    }
                    self.consume_conditional_rule(&mut rule, base)?;
                    formatting.rules.push(rule);
                },
                record_types::END_COND_FORMATTING => {
                    if !self.buf.is_empty() {
                        return Err(XlsbError::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginConditionalFormatting collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if formatting.rules.len() != expected_count as usize {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginConditionalFormatting count".to_string(),
                val: format!(
                    "declared {expected_count}, found {}",
                    formatting.rules.len()
                ),
            });
        }
        Ok(())
    }

    fn consume_conditional_rule(
        &mut self,
        rule: &mut ConditionalFormattingRule,
        base: (u32, u32),
    ) -> XlsbResult<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::BEGIN_COLOR_SCALE => {
                    if rule.color_scale.is_some() || !self.buf.is_empty() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginColorScale".to_string(),
                            val: "duplicate record or nonempty payload".to_string(),
                        });
                    }
                    rule.color_scale = Some(self.consume_color_scale(base)?);
                },
                record_types::BEGIN_DATABAR => {
                    if rule.data_bar.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginDatabar".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    rule.data_bar = Some(self.consume_data_bar(&begin, base)?);
                },
                record_types::BEGIN_ICON_SET => {
                    if rule.icon_set.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "BrtBeginIconSet".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    rule.icon_set = Some(self.consume_icon_set(&begin, base)?);
                },
                record_types::END_CF_RULE => {
                    if !self.buf.is_empty() {
                        return Err(XlsbError::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginCFRule collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let valid_visual = match rule.rule_type {
            crate::xlsb::conditional_formatting::CfRuleType::ColorScale => {
                rule.color_scale.is_some() && rule.data_bar.is_none() && rule.icon_set.is_none()
            },
            crate::xlsb::conditional_formatting::CfRuleType::DataBar => {
                rule.color_scale.is_none() && rule.data_bar.is_some() && rule.icon_set.is_none()
            },
            crate::xlsb::conditional_formatting::CfRuleType::IconSet => {
                rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_some()
            },
            _ => rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_none(),
        };
        if !valid_visual {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginCFRule collection".to_string(),
                val: "visualization records do not match rule type".to_string(),
            });
        }
        Ok(())
    }

    fn consume_color_scale(&mut self, base: (u32, u32)) -> XlsbResult<ColorScale> {
        let mut cfvos = Vec::new();
        let mut colors = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::CFVO if colors.is_empty() => cfvos.push(Cfvo::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                record_types::COLOR => colors.push(ConditionalFormatColor::parse(&self.buf)?),
                record_types::END_COLOR_SCALE if self.buf.is_empty() => break,
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginColorScale collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if !(cfvos.len() == 2 || cfvos.len() == 3) || colors.len() != cfvos.len() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginColorScale collection".to_string(),
                val: format!("{} thresholds and {} colors", cfvos.len(), colors.len()),
            });
        }
        if cfvos.first().is_some_and(|cfvo| cfvo.cfvo_type == 3)
            || cfvos.last().is_some_and(|cfvo| cfvo.cfvo_type == 2)
            || (cfvos.len() == 3 && matches!(cfvos[1].cfvo_type, 2 | 3))
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginColorScale collection".to_string(),
                val: "invalid min/mid/max threshold type".to_string(),
            });
        }
        let has_middle = colors.len() == 3;
        let mut cfvos = cfvos.into_iter();
        let min_cfvo = cfvos.next().ok_or_else(|| XlsbError::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing minimum threshold".to_string(),
        })?;
        let middle_cfvo = if has_middle { cfvos.next() } else { None };
        let max_cfvo = cfvos.next().ok_or_else(|| XlsbError::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing maximum threshold".to_string(),
        })?;
        let mut colors = colors.into_iter();
        let min_color_record = colors.next().ok_or_else(|| XlsbError::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing minimum color".to_string(),
        })?;
        let mid_color_record = if has_middle { colors.next() } else { None };
        let max_color_record = colors.next().ok_or_else(|| XlsbError::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing maximum color".to_string(),
        })?;
        Ok(ColorScale {
            min_cfvo,
            mid_cfvo: middle_cfvo,
            max_cfvo,
            min_color: min_color_record.argb.unwrap_or(0),
            mid_color: mid_color_record.and_then(|color| color.argb),
            max_color: max_color_record.argb.unwrap_or(0),
            min_color_record,
            mid_color_record,
            max_color_record,
        })
    }

    fn consume_data_bar(&mut self, begin: &[u8], base: (u32, u32)) -> XlsbResult<DataBar> {
        if begin.len() != 3 || begin[0] > begin[1] || begin[1] > 100 || begin[2] > 1 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginDatabar".to_string(),
                val: "invalid width or show-value field".to_string(),
            });
        }
        let mut cfvos = Vec::new();
        let mut color = None;
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::CFVO if color.is_none() => cfvos.push(Cfvo::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                record_types::COLOR if color.is_none() => {
                    color = Some(ConditionalFormatColor::parse(&self.buf)?)
                },
                record_types::END_DATABAR if self.buf.is_empty() => break,
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginDatabar collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != 2 || color.is_none() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: format!("{} thresholds, color={}", cfvos.len(), color.is_some()),
            });
        }
        if cfvos[0].cfvo_type == 3 || cfvos[1].cfvo_type == 2 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: "invalid minimum/maximum threshold type".to_string(),
            });
        }
        let [min_cfvo, max_cfvo]: [Cfvo; 2] =
            cfvos.try_into().map_err(|_| XlsbError::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: "invalid threshold count".to_string(),
            })?;
        let color_record = color.ok_or_else(|| XlsbError::Unrecognized {
            typ: "BrtBeginDatabar collection".to_string(),
            val: "missing color".to_string(),
        })?;
        Ok(DataBar {
            min_cfvo,
            max_cfvo,
            color: color_record.argb.unwrap_or(0),
            show_value: begin[2] != 0,
            min_length: begin[0],
            max_length: begin[1],
            color_record,
        })
    }

    fn consume_icon_set(&mut self, begin: &[u8], base: (u32, u32)) -> XlsbResult<IconSet> {
        if begin.len() != 6 {
            return Err(XlsbError::InvalidLength {
                expected: 6,
                found: begin.len(),
            });
        }
        let icon_set = binary::read_u32_le_at(begin, 0)?;
        let flags = binary::read_u16_le_at(begin, 4)?;
        if icon_set > 16 || flags & !0x7e != 0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginIconSet".to_string(),
                val: format!("set {icon_set}, flags 0x{flags:04X}"),
            });
        }
        let expected = if icon_set <= 7 {
            3
        } else if icon_set <= 12 {
            4
        } else {
            5
        };
        let mut cfvos = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                record_types::CFVO => cfvos.push(Cfvo::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                record_types::END_ICON_SET if self.buf.is_empty() => break,
                _ => {
                    return Err(XlsbError::Unrecognized {
                        typ: "BrtBeginIconSet collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != expected {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginIconSet collection".to_string(),
                val: format!("expected {expected} thresholds, found {}", cfvos.len()),
            });
        }
        if cfvos
            .iter()
            .any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3) || !cfvo.save_greater_than_or_equal)
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginIconSet collection".to_string(),
                val: "invalid threshold type or fSaveGTE flag".to_string(),
            });
        }
        Ok(IconSet {
            icon_set_type: icon_set as u8,
            cfvos,
            show_value: flags & 0x02 == 0,
            reverse: flags & 0x04 == 0,
        })
    }

    fn parse_auto_filter(data: &[u8]) -> XlsbResult<XlsbAutoFilter> {
        if data.len() != 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        let first_row = binary::read_u32_le_at(data, 0)?;
        let last_row = binary::read_u32_le_at(data, 4)?;
        let first_column = binary::read_u32_le_at(data, 8)?;
        let last_column = binary::read_u32_le_at(data, 12)?;
        if first_row > last_row
            || last_row >= 0x10_0000
            || first_column > last_column
            || last_column >= 0x4000
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginAFilter rfx".to_string(),
                val: format!(
                    "rows {first_row}..={last_row}, columns {first_column}..={last_column}"
                ),
            });
        }
        Ok(XlsbAutoFilter {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }

    fn consume_auto_filter_records(&mut self) -> XlsbResult<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            if typ == record_types::END_A_FILTER {
                if !self.buf.is_empty() {
                    return Err(XlsbError::InvalidLength {
                        expected: 0,
                        found: self.buf.len(),
                    });
                }
                return Ok(());
            }
            if typ == record_types::BEGIN_A_FILTER {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtBeginAFilter".to_string(),
                    val: "nested AutoFilter".to_string(),
                });
            }
        }
    }

    fn parse_sheet_protection(data: &[u8]) -> XlsbResult<XlsbSheetProtection> {
        if data.len() != 66 {
            return Err(XlsbError::InvalidLength {
                expected: 66,
                found: data.len(),
            });
        }
        let password = binary::read_u16_le_at(data, 0)?;
        let mut flags = [false; 16];
        for (index, flag) in flags.iter_mut().enumerate() {
            let value = binary::read_u32_le_at(data, 2 + index * 4)?;
            if value > 1 {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtSheetProtection Boolean".to_string(),
                    val: format!("field {index}: {value}"),
                });
            }
            *flag = value != 0;
        }
        Ok(XlsbSheetProtection {
            password_hash: (password != 0).then_some(password),
            locked: flags[0],
            allow_edit_objects: flags[1],
            allow_edit_scenarios: flags[2],
            allow_format_cells: flags[3],
            allow_format_columns: flags[4],
            allow_format_rows: flags[5],
            allow_insert_columns: flags[6],
            allow_insert_rows: flags[7],
            allow_insert_hyperlinks: flags[8],
            allow_delete_columns: flags[9],
            allow_delete_rows: flags[10],
            allow_select_locked_cells: flags[11],
            allow_sort: flags[12],
            allow_auto_filter: flags[13],
            allow_pivot_tables: flags[14],
            allow_select_unlocked_cells: flags[15],
        })
    }

    fn sheet_protection_flags(protection: &XlsbSheetProtection) -> [bool; 16] {
        [
            protection.locked,
            protection.allow_edit_objects,
            protection.allow_edit_scenarios,
            protection.allow_format_cells,
            protection.allow_format_columns,
            protection.allow_format_rows,
            protection.allow_insert_columns,
            protection.allow_insert_rows,
            protection.allow_insert_hyperlinks,
            protection.allow_delete_columns,
            protection.allow_delete_rows,
            protection.allow_select_locked_cells,
            protection.allow_sort,
            protection.allow_auto_filter,
            protection.allow_pivot_tables,
            protection.allow_select_unlocked_cells,
        ]
    }

    fn parse_strong_sheet_protection(
        data: &[u8],
    ) -> XlsbResult<(XlsbStrongProtection, [bool; 16])> {
        if data.len() < 83 {
            return Err(XlsbError::InvalidLength {
                expected: 83,
                found: data.len(),
            });
        }
        let spin_count = binary::read_u32_le_at(data, 0)?;
        if spin_count > 10_000_000 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtSheetProtectionIso dwSpinCount".to_string(),
                val: spin_count.to_string(),
            });
        }
        let mut flags = [false; 16];
        for (index, flag) in flags.iter_mut().enumerate() {
            let value = binary::read_u32_le_at(data, 4 + index * 4)?;
            if value > 1 {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtSheetProtectionIso Boolean".to_string(),
                    val: format!("field {index}: {value}"),
                });
            }
            *flag = value != 0;
        }
        let mut offset = 68;
        let hash = Self::read_length_prefixed_bytes(data, &mut offset, "rgbHash")?;
        if hash.is_empty() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtSheetProtectionIso rgbHash".to_string(),
                val: "empty".to_string(),
            });
        }
        let salt = Self::read_length_prefixed_bytes(data, &mut offset, "rgbSalt")?;
        let algorithm = Self::read_nullable_wide_string(data, &mut offset)?
            .filter(|value| !value.is_empty())
            .ok_or_else(|| XlsbError::Unrecognized {
                typ: "BrtSheetProtectionIso szAlgName".to_string(),
                val: "null or empty".to_string(),
            })?;
        if offset != data.len() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtSheetProtectionIso".to_string(),
                val: format!("{} trailing bytes", data.len() - offset),
            });
        }
        Ok((
            XlsbStrongProtection {
                spin_count,
                hash,
                salt,
                algorithm,
            },
            flags,
        ))
    }

    fn read_length_prefixed_bytes(
        data: &[u8],
        offset: &mut usize,
        field: &str,
    ) -> XlsbResult<Vec<u8>> {
        let data_offset = offset.checked_add(4).ok_or_else(|| {
            XlsbError::Encoding(format!("BrtSheetProtectionIso {field} offset overflow"))
        })?;
        if data_offset > data.len() {
            return Err(XlsbError::InvalidLength {
                expected: data_offset,
                found: data.len(),
            });
        }
        let count = binary::read_u32_le_at(data, *offset)? as usize;
        let end = data_offset.checked_add(count).ok_or_else(|| {
            XlsbError::Encoding(format!("BrtSheetProtectionIso {field} size overflow"))
        })?;
        if end > data.len() {
            return Err(XlsbError::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        *offset = end;
        Ok(data[data_offset..end].to_vec())
    }

    fn read_nullable_wide_string(data: &[u8], offset: &mut usize) -> XlsbResult<Option<String>> {
        let text_offset = offset
            .checked_add(4)
            .ok_or_else(|| XlsbError::Encoding("ISO algorithm offset overflow".to_string()))?;
        if text_offset > data.len() {
            return Err(XlsbError::InvalidLength {
                expected: text_offset,
                found: data.len(),
            });
        }
        let count = binary::read_u32_le_at(data, *offset)?;
        if count == u32::MAX {
            *offset = text_offset;
            return Ok(None);
        }
        let byte_count = (count as usize)
            .checked_mul(2)
            .ok_or_else(|| XlsbError::Encoding("ISO algorithm size overflow".to_string()))?;
        let end = text_offset
            .checked_add(byte_count)
            .ok_or_else(|| XlsbError::Encoding("ISO algorithm end overflow".to_string()))?;
        if end > data.len() {
            return Err(XlsbError::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let units = data[text_offset..end]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let value = String::from_utf16(&units)
            .map_err(|error| XlsbError::Encoding(format!("invalid ISO algorithm: {error}")))?;
        *offset = end;
        Ok(Some(value))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::writer::RecordWriter;
    use litchi_core::sheet::Cell;

    fn rich_string_worksheet() -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut writer = RecordWriter::new(&mut bytes);
        writer.write_record(0x0094, &[0; 16]).unwrap();
        writer
            .write_record(record_types::BEGIN_COL_INFOS, &[])
            .unwrap();
        let mut column = 2u32.to_le_bytes().to_vec();
        column.extend_from_slice(&4u32.to_le_bytes());
        column.extend_from_slice(&4096u32.to_le_bytes());
        column.extend_from_slice(&0u32.to_le_bytes());
        column.extend_from_slice(&0x130Fu16.to_le_bytes());
        writer
            .write_record(record_types::COL_INFO, &column)
            .unwrap();
        writer
            .write_record(record_types::END_COL_INFOS, &[])
            .unwrap();
        writer.write_record(0x0091, &[]).unwrap();

        let mut row = 4u32.to_le_bytes().to_vec();
        row.extend_from_slice(&0u32.to_le_bytes());
        row.extend_from_slice(&500u16.to_le_bytes());
        row.extend_from_slice(&[3, 0x7A, 1]);
        row.extend_from_slice(&1u32.to_le_bytes());
        row.extend_from_slice(&2u32.to_le_bytes());
        row.extend_from_slice(&2u32.to_le_bytes());
        writer.write_record(record_types::ROW_HDR, &row).unwrap();

        let mut cell = 2u32.to_le_bytes().to_vec();
        cell.extend_from_slice(&[0, 0, 0, 1]);
        cell.push(1);
        cell.extend_from_slice(&2u32.to_le_bytes());
        cell.extend_from_slice(&[b'A', 0, b'B', 0]);
        cell.extend_from_slice(&2u32.to_le_bytes());
        cell.extend_from_slice(&[0, 0, 3, 0, 1, 0, 5, 0]);
        writer
            .write_record(record_types::CELL_R_STRING, &cell)
            .unwrap();
        writer.write_record(0x0092, &[]).unwrap();
        writer.write_record(0x0082, &[]).unwrap();
        bytes
    }

    #[test]
    fn decodes_and_validates_cell_style_header() {
        type Reader<'a> = XlsbCellsReader<'a, std::io::Cursor<&'a [u8]>>;

        let header = [7, 0, 0, 0, 0x34, 0x12, 0, 1];
        assert_eq!(
            Reader::decode_cell_header(&header, 0x1235).unwrap(),
            CellHeader {
                col: 7,
                style_id: 0x1234,
                show_phonetic: true,
            }
        );
        assert!(matches!(
            Reader::decode_cell_header(&header, 0x1234),
            Err(XlsbError::Unrecognized { .. })
        ));

        let reserved = [0, 0, 0, 0, 0, 0, 0, 2];
        assert!(matches!(
            Reader::decode_cell_header(&reserved, 1),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn parses_iso_sheet_protection_metadata_and_matching_flags() {
        type Reader<'a> = XlsbCellsReader<'a, std::io::Cursor<&'a [u8]>>;
        let flags = [
            true, false, true, false, true, false, true, false, true, false, true, false, true,
            false, true, false,
        ];
        let mut iso = 250_000u32.to_le_bytes().to_vec();
        for flag in flags {
            iso.extend_from_slice(&u32::from(flag).to_le_bytes());
        }
        iso.extend_from_slice(&3u32.to_le_bytes());
        iso.extend_from_slice(&[1, 2, 3]);
        iso.extend_from_slice(&2u32.to_le_bytes());
        iso.extend_from_slice(&[4, 5]);
        iso.extend_from_slice(&7u32.to_le_bytes());
        for unit in "SHA-512".encode_utf16() {
            iso.extend_from_slice(&unit.to_le_bytes());
        }

        let (strong, parsed_flags) = Reader::parse_strong_sheet_protection(&iso).unwrap();
        assert_eq!(strong.spin_count, 250_000);
        assert_eq!(strong.hash, vec![1, 2, 3]);
        assert_eq!(strong.salt, vec![4, 5]);
        assert_eq!(strong.algorithm, "SHA-512");
        assert_eq!(parsed_flags, flags);

        let mut base = 0u16.to_le_bytes().to_vec();
        for flag in flags {
            base.extend_from_slice(&u32::from(flag).to_le_bytes());
        }
        let protection = Reader::parse_sheet_protection(&base).unwrap();
        assert_eq!(Reader::sheet_protection_flags(&protection), flags);

        let mut worksheet = Vec::new();
        let mut writer = RecordWriter::new(&mut worksheet);
        writer.write_record(0x0094, &[0; 16]).unwrap();
        writer.write_record(0x0091, &[]).unwrap();
        writer.write_record(0x0092, &[]).unwrap();
        writer
            .write_record(record_types::SHEET_PROTECTION_ISO, &iso)
            .unwrap();
        writer
            .write_record(record_types::SHEET_PROTECTION, &base)
            .unwrap();
        writer.write_record(0x0082, &[]).unwrap();
        let formula_context = FormulaResolutionContext::default();
        let iter = RecordIter::new(std::io::Cursor::new(worksheet));
        let mut reader = XlsbCellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(reader.next_cell().unwrap().is_none());
        assert_eq!(reader.sheet_protection.unwrap(), protection);
        assert_eq!(reader.strong_sheet_protection.unwrap(), strong);
    }

    #[test]
    fn reads_inline_rich_string_cells() {
        let bytes = rich_string_worksheet();
        let formula_context = FormulaResolutionContext::default();
        let iter = RecordIter::new(std::io::Cursor::new(bytes));
        let mut reader = XlsbCellsReader::new(iter, &[], &formula_context, 1).unwrap();

        let cell = reader.next_cell().unwrap().unwrap();
        assert_eq!(cell.row(), 4);
        assert_eq!(cell.column(), 2);
        assert_eq!(cell.value(), &CellValue::String("AB".to_string()));
        assert!(cell.show_phonetic());
        let rich = cell.rich_string().unwrap();
        assert_eq!(rich.text, "AB");
        assert_eq!(rich.runs.len(), 2);
        assert_eq!(rich.runs[0].font_id, 3);
        assert_eq!(rich.runs[1].font_id, 5);
        assert!(reader.next_cell().unwrap().is_none());

        assert_eq!(reader.column_infos.len(), 1);
        let column = &reader.column_infos[0];
        assert_eq!((column.first_column, column.last_column), (2, 4));
        assert_eq!(column.width, 16.0);
        assert!(column.hidden);
        assert!(column.user_set_width);
        assert!(column.best_fit);
        assert!(column.show_phonetic);
        assert_eq!(column.outline_level, 3);
        assert!(column.collapsed);

        assert_eq!(reader.row_infos.len(), 1);
        let row = &reader.row_infos[0];
        assert_eq!(row.row, 4);
        assert_eq!(row.style_id, Some(0));
        assert_eq!(row.height, Some(25.0));
        assert!(row.extra_ascender);
        assert!(row.extra_descender);
        assert_eq!(row.outline_level, 2);
        assert!(row.collapsed);
        assert!(row.hidden);
        assert!(row.show_phonetic);
        assert_eq!(row.column_spans, vec![(2, 2)]);
    }

    #[test]
    fn rejects_a_validation_collection_with_a_mismatched_count() {
        let mut worksheet = Vec::new();
        let mut writer = RecordWriter::new(&mut worksheet);
        writer.write_record(0x0094, &[0; 16]).unwrap();
        writer.write_record(0x0091, &[]).unwrap();
        writer.write_record(0x0092, &[]).unwrap();
        let mut begin = vec![0; 14];
        begin.extend_from_slice(&1u32.to_le_bytes());
        writer
            .write_record(record_types::BEGIN_D_VALS, &begin)
            .unwrap();
        writer.write_record(record_types::END_D_VALS, &[]).unwrap();
        writer.write_record(0x0082, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = RecordIter::new(std::io::Cursor::new(worksheet));
        let mut reader = XlsbCellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(matches!(
            reader.next_cell(),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn rejects_conditional_formatting_with_a_mismatched_rule_count() {
        let mut worksheet = Vec::new();
        let mut writer = RecordWriter::new(&mut worksheet);
        writer.write_record(0x0094, &[0; 16]).unwrap();
        writer.write_record(0x0091, &[]).unwrap();
        writer.write_record(0x0092, &[]).unwrap();
        let mut begin = 1u32.to_le_bytes().to_vec();
        begin.extend_from_slice(&0u32.to_le_bytes());
        begin.extend_from_slice(&1u32.to_le_bytes());
        begin.extend_from_slice(&[0; 16]);
        writer
            .write_record(record_types::BEGIN_COND_FORMATTING, &begin)
            .unwrap();
        writer
            .write_record(record_types::END_COND_FORMATTING, &[])
            .unwrap();
        writer.write_record(0x0082, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = RecordIter::new(std::io::Cursor::new(worksheet));
        let mut reader = XlsbCellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(matches!(
            reader.next_cell(),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn rejects_incomplete_color_scale_collection() {
        let mut worksheet = Vec::new();
        let mut writer = RecordWriter::new(&mut worksheet);
        writer.write_record(0x0094, &[0; 16]).unwrap();
        writer.write_record(0x0091, &[]).unwrap();
        writer.write_record(0x0092, &[]).unwrap();
        let mut begin = 1u32.to_le_bytes().to_vec();
        begin.extend_from_slice(&0u32.to_le_bytes());
        begin.extend_from_slice(&1u32.to_le_bytes());
        begin.extend_from_slice(&[0; 16]);
        writer
            .write_record(record_types::BEGIN_COND_FORMATTING, &begin)
            .unwrap();

        let mut rule = 3u32.to_le_bytes().to_vec();
        rule.extend_from_slice(&2u32.to_le_bytes());
        rule.extend_from_slice(&u32::MAX.to_le_bytes());
        rule.extend_from_slice(&1u32.to_le_bytes());
        rule.extend_from_slice(&[0; 10]);
        rule.extend_from_slice(&[0; 12]);
        rule.extend_from_slice(&u32::MAX.to_le_bytes());
        writer
            .write_record(record_types::BEGIN_CF_RULE, &rule)
            .unwrap();
        writer
            .write_record(record_types::BEGIN_COLOR_SCALE, &[])
            .unwrap();
        writer
            .write_record(record_types::END_COLOR_SCALE, &[])
            .unwrap();
        writer.write_record(record_types::END_CF_RULE, &[]).unwrap();
        writer
            .write_record(record_types::END_COND_FORMATTING, &[])
            .unwrap();
        writer.write_record(0x0082, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = RecordIter::new(std::io::Cursor::new(worksheet));
        let mut reader = XlsbCellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(reader.next_cell().is_err());
    }
}
