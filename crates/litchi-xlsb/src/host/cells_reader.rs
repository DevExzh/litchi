//! XLSB cells reader implementation

use crate::conditional_formatting::{
    AxisPosition14, Bar, Bar14, Color, Formatting, Icon, IconSet, IconSet14, RecordKind, Rule,
    RuleType, Scale, Value, icon_count14, parse_classic_header, parse_rule_extension_guid,
};
use crate::package::cell::{Cell, CellHeader};
use crate::package::data_validation::{
    DataValidationSettings, Validation, parse_collection_settings, parse_dval_list,
};
use crate::package::error::{Error, Result};
use crate::package::formula::{
    CellParsedFormula, FormulaGroup, FormulaResolutionContext, GroupKind,
};
use crate::package::hyperlinks::Hyperlink;
use crate::package::merged_cells::MergedCell;
use crate::package::records::Stream;
use crate::package::shared_strings::SharedString;
use crate::package::web_extension_bindings::Binding;
use crate::raw::{Cursor, kind};
use crate::sheet::{AutoFilter, ColumnInfo, RowInfo, SheetProtection, StrongProtection};
use litchi_core::binary;
use litchi_core::sheet::CellValue;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};
use std::sync::Arc;

struct ParsedFormulaCell {
    header: CellHeader,
    cached_value: CellValue,
    formula: CellParsedFormula,
    flags: u16,
}

fn build_color_scale(
    cfvos: Vec<Value>,
    colors: Vec<Color>,
    record: &'static str,
    extension14: bool,
) -> Result<Scale> {
    if !(cfvos.len() == 2 || cfvos.len() == 3) || colors.len() != cfvos.len() {
        return Err(Error::Unrecognized {
            typ: record.to_string(),
            val: format!("{} thresholds and {} colors", cfvos.len(), colors.len()),
        });
    }
    if (extension14 && cfvos.iter().any(|cfvo| matches!(cfvo.cfvo_type, 8 | 9)))
        || cfvos[0].cfvo_type == 3
        || cfvos[cfvos.len() - 1].cfvo_type == 2
        || (cfvos.len() == 3 && matches!(cfvos[1].cfvo_type, 2 | 3))
    {
        return Err(Error::Unrecognized {
            typ: record.to_string(),
            val: "invalid min/mid/max threshold type".to_string(),
        });
    }
    let has_middle = colors.len() == 3;
    let mut cfvos = cfvos.into_iter();
    let min_cfvo = cfvos.next().expect("validated threshold count");
    let middle_cfvo = if has_middle { cfvos.next() } else { None };
    let max_cfvo = cfvos.next().expect("validated threshold count");
    let mut colors = colors.into_iter();
    let min_color_record = colors.next().expect("validated color count");
    let mid_color_record = if has_middle { colors.next() } else { None };
    let max_color_record = colors.next().expect("validated color count");
    Ok(Scale {
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
pub struct CellsReader<'a, RS>
where
    RS: Read + Seek,
{
    iter: Stream<RS>,
    shared_strings: &'a [SharedString],
    formula_context: &'a FormulaResolutionContext,
    cell_xf_count: usize,
    dimensions: Dimensions,
    current_row: u32,
    last_row: Option<u32>,
    buf: Vec<u8>,
    pending_record: Option<(crate::raw::Kind, Vec<u8>)>,
    formula_groups: Vec<Arc<FormulaGroup>>,
    /// Merged cells found in the worksheet
    pub merged_cells: Vec<MergedCell>,
    /// Hyperlinks found in the worksheet
    pub hyperlinks: Vec<Hyperlink>,
    /// Column formatting records found before sheet data.
    pub column_infos: Vec<ColumnInfo>,
    /// Row header metadata found within sheet data.
    pub row_infos: Vec<RowInfo>,
    /// Worksheet AutoFilter range.
    pub auto_filter: Option<AutoFilter>,
    /// Worksheet protection settings.
    pub sheet_protection: Option<SheetProtection>,
    /// ISO strong password-verifier metadata.
    pub strong_sheet_protection: Option<StrongProtection>,
    /// Classic worksheet data-validation rules.
    pub data_validations: Vec<Validation>,
    /// UI settings from the classic validation collection.
    pub data_validation_settings: Option<DataValidationSettings>,
    /// UI settings from the Office 2013 validation collection.
    pub data_validation14_settings: Option<DataValidationSettings>,
    /// Classic and Office 2013 conditional-formatting blocks in stream order.
    pub conditional_formattings: Vec<Formatting>,
    /// Inert Office Add-in bindings from the worksheet WEBEXTENSIONS collection.
    pub web_extension_bindings: Vec<Binding>,
    /// Sheet views from the worksheet WSVIEWS collection.
    pub sheet_views: Vec<crate::package::sheet_view::SheetView>,
    saw_web_extension_collection: bool,
}

impl<'a, RS> CellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub fn new(
        mut iter: Stream<RS>,
        shared_strings: &'a [SharedString],
        formula_context: &'a FormulaResolutionContext,
        cell_xf_count: usize,
    ) -> Result<Self> {
        let mut buf = Vec::with_capacity(1024);

        // Walk the worksheet preamble up to BrtWsDim (worksheet dimensions),
        // capturing the sheet-view collection while skipping everything else.
        let mut sheet_views = Vec::new();
        loop {
            let typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;
            if typ == kind::WS_DIM {
                break;
            }
            if typ == kind::BEGIN_WS_VIEWS {
                sheet_views = crate::package::sheet_view::read_sheet_views(&mut iter, &mut buf)?;
            }
        }
        if buf.len() != 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: buf.len(),
            });
        }
        let dimensions = Self::parse_dimensions(&buf);

        // Read worksheet preamble through BrtBeginSheetData, retaining column
        // formatting and sheet views while safely ignoring unrelated and
        // future records. Producers disagree on whether the WSVIEWS block
        // precedes or follows BrtWsDim, so both phases capture it.
        let mut column_infos = Vec::new();
        loop {
            let typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;
            if typ == kind::BEGIN_SHEET_DATA {
                break;
            }
            if typ == kind::COL_INFO {
                column_infos.push(Self::parse_column_info(&buf, cell_xf_count)?);
            } else if typ == kind::BEGIN_WS_VIEWS {
                sheet_views = crate::package::sheet_view::read_sheet_views(&mut iter, &mut buf)?;
            }
        }

        Ok(CellsReader {
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
            web_extension_bindings: Vec::new(),
            sheet_views,
            saw_web_extension_collection: false,
        })
    }

    #[allow(dead_code)]
    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell>> {
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

            if typ == kind::END_SHEET_DATA {
                // BrtEndSheetData - continue to read advanced features
                self.read_advanced_features()?;
                return Ok(None);
            }

            let cell_header = if (kind::CELL_BLANK.get()..=kind::FMLA_ERROR.get())
                .contains(&typ.get())
                || typ == kind::CELL_R_STRING
            {
                Some(self.parse_cell_header()?)
            } else {
                None
            };

            match typ {
                kind::ROW_HDR => {
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
                kind::CELL_BLANK
                    // BrtCellBlank
                    if self.buf.len() >= 8 => {
                        let header = cell_header.unwrap();
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Empty,
                        )));
                    },
                kind::CELL_RK
                    // BrtCellRk
                    if self.buf.len() >= 12 => {
                        let header = cell_header.unwrap();
                        let mut cursor = Cursor::new(&self.buf[8..], "BrtCellRk");
                        let value = Self::cell_value_from_number(cursor.read_rk()?);
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            value,
                        )));
                    },
                kind::CELL_ERROR
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
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Error(error_msg.to_string()),
                        )));
                    },
                kind::CELL_BOOL
                    // BrtCellBool
                    if self.buf.len() >= 9 => {
                        let header = cell_header.unwrap();
                        let mut cursor = Cursor::new(&self.buf[8..], "BrtCellBool");
                        let value = cursor.read_bool8()?;
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Bool(value),
                        )));
                    },
                kind::CELL_REAL
                    // BrtCellReal
                    if self.buf.len() >= 16 => {
                        let header = cell_header.unwrap();
                        let value = binary::read_f64_le_at(&self.buf, 8)?;
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            CellValue::Float(value),
                        )));
                    },
                kind::CELL_ST
                    // BrtCellSt
                    if self.buf.len() >= 8 => {
                        let header = cell_header.unwrap();
                        let (string, _) = super::records::decode_string(&self.buf[8..])?;
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            CellValue::String(string),
                        )));
                    },
                kind::CELL_ISST
                    // BrtCellIsst
                    if self.buf.len() >= 12 => {
                        let header = cell_header.unwrap();
                        let idx = binary::read_u32_le_at(&self.buf, 8)? as usize;
                        let value = if idx < self.shared_strings.len() {
                            CellValue::String(self.shared_strings[idx].text.clone())
                        } else {
                            CellValue::Error("Invalid SST index".to_string())
                        };
                        return Ok(Some(Cell::new_styled(
                            self.current_row,
                            header,
                            value,
                        )));
                    },
                kind::CELL_R_STRING => {
                    if self.buf.len() < 9 {
                        return Err(Error::InvalidLength {
                            expected: 9,
                            found: self.buf.len(),
                        });
                    }
                    let header = cell_header.unwrap();
                    let rich_string = SharedString::parse(&self.buf[8..])?;
                    return Ok(Some(Cell::new_rich_string(
                        self.current_row,
                        header,
                        rich_string,
                    )));
                },
                kind::FMLA_STRING => {
                    // BrtFmlaString - formula with string result
                    if self.buf.len() < 12 {
                        return Err(Error::InvalidLength {
                            expected: 12,
                            found: self.buf.len(),
                        });
                    }
                    let header = cell_header.unwrap();
                    let (string, consumed) =
                        super::records::decode_string(&self.buf[8..])?;
                    let parsed =
                        self.parse_formula_cell(header, CellValue::String(string), 8 + consumed)?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                kind::FMLA_NUM => {
                    // BrtFmlaNum - formula with numeric result
                    if self.buf.len() < 18 {
                        return Err(Error::InvalidLength {
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
                kind::FMLA_BOOL => {
                    // BrtFmlaBool - formula with boolean result
                    if self.buf.len() < 11 {
                        return Err(Error::InvalidLength {
                            expected: 11,
                            found: self.buf.len(),
                        });
                    }
                    let header = cell_header.unwrap();
                    let mut cursor = Cursor::new(&self.buf[8..], "BrtFmlaBool cached value");
                    let bool_value = cursor.read_bool8()?;
                    let parsed = self.parse_formula_cell(
                        header,
                        CellValue::Bool(bool_value),
                        9,
                    )?;
                    return self.resolve_formula_record(parsed).map(Some);
                },
                kind::FMLA_ERROR => {
                    // BrtFmlaError - formula with error result
                    if self.buf.len() < 11 {
                        return Err(Error::InvalidLength {
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
                kind::ARR_FMLA | kind::SHR_FMLA => {
                    return Err(Error::InvalidFormula(
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

    fn parse_column_info(data: &[u8], cell_xf_count: usize) -> Result<ColumnInfo> {
        if data.len() != 18 {
            return Err(Error::InvalidLength {
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
            return Err(Error::Unrecognized {
                typ: "BrtColInfo range".to_string(),
                val: format!("{first_column}..={last_column}"),
            });
        }
        if width_raw > 0xFFFF {
            return Err(Error::Unrecognized {
                typ: "BrtColInfo coldx".to_string(),
                val: width_raw.to_string(),
            });
        }
        if style_id as usize >= cell_xf_count {
            return Err(Error::Unrecognized {
                typ: "BrtColInfo ixfe".to_string(),
                val: format!("{style_id} (cell XF count {cell_xf_count})"),
            });
        }
        if flags & !0x170F != 0 {
            return Err(Error::Unrecognized {
                typ: "BrtColInfo flags".to_string(),
                val: format!("0x{flags:04X}"),
            });
        }
        Ok(ColumnInfo {
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
    ) -> Result<RowInfo> {
        if data.len() < 17 {
            return Err(Error::InvalidLength {
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
            return Err(Error::Unrecognized {
                typ: "BrtRowHdr ccolspan".to_string(),
                val: span_count.to_string(),
            });
        }
        let expected = span_count
            .checked_mul(8)
            .and_then(|size| size.checked_add(17))
            .ok_or_else(|| Error::Encoding("BrtRowHdr size overflow".to_string()))?;
        if data.len() != expected {
            return Err(Error::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        if row >= 0x10_0000 || previous_row.is_some_and(|previous| row <= previous) {
            return Err(Error::Unrecognized {
                typ: "BrtRowHdr rw".to_string(),
                val: row.to_string(),
            });
        }
        if height_twips > 0x2000 {
            return Err(Error::Unrecognized {
                typ: "BrtRowHdr miyRw".to_string(),
                val: height_twips.to_string(),
            });
        }
        if flags1 & 0xFC != 0 || flags2 & 0x80 != 0 || phonetic_flags & 0xFE != 0 {
            return Err(Error::Unrecognized {
                typ: "BrtRowHdr flags".to_string(),
                val: format!("0x{flags1:02X}/0x{flags2:02X}/0x{phonetic_flags:02X}"),
            });
        }
        let style_applied = flags2 & 0x40 != 0;
        if style_applied && raw_style_id as usize >= cell_xf_count {
            return Err(Error::Unrecognized {
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
                return Err(Error::Unrecognized {
                    typ: "BrtRowHdr BrtColSpan".to_string(),
                    val: format!("{first}..={last}"),
                });
            }
            previous_segment = Some(segment);
            column_spans.push((first, last));
        }

        Ok(RowInfo {
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

    fn parse_cell_header(&self) -> Result<CellHeader> {
        Self::decode_cell_header(&self.buf, self.cell_xf_count)
    }

    fn decode_cell_header(data: &[u8], cell_xf_count: usize) -> Result<CellHeader> {
        if data.len() < 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let flags = data[7];
        if flags & 0xFE != 0 {
            return Err(Error::Unrecognized {
                typ: "Cell flags".to_string(),
                val: format!("0x{flags:02X}"),
            });
        }
        let col = binary::read_u32_le_at(data, 0)?;
        let style_id = u32::from(data[4]) | (u32::from(data[5]) << 8) | (u32::from(data[6]) << 16);
        if style_id as usize >= cell_xf_count {
            return Err(Error::Unrecognized {
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
    ) -> Result<ParsedFormulaCell> {
        let formula_offset = flags_offset
            .checked_add(2)
            .ok_or_else(|| Error::InvalidFormula("formula record offset overflow".to_string()))?;
        if self.buf.len() < formula_offset {
            return Err(Error::InvalidLength {
                expected: formula_offset,
                found: self.buf.len(),
            });
        }
        let flags = binary::read_u16_le_at(&self.buf, flags_offset)?;
        if flags & !0x0002 != 0 {
            return Err(Error::InvalidFormula(format!(
                "invalid GrbitFmla flags 0x{flags:04X}"
            )));
        }
        let (formula, consumed) = CellParsedFormula::parse(&self.buf[formula_offset..])?;
        if formula_offset + consumed != self.buf.len() {
            return Err(Error::InvalidFormula(format!(
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

    fn resolve_formula_record(&mut self, parsed: ParsedFormulaCell) -> Result<Cell> {
        let position = (self.current_row, parsed.header.col);
        let exp_cell = parsed.formula.exp_cell()?;

        let next_type = self.iter.read_type()?;
        let mut next_data = Vec::new();
        let _ = self.iter.fill_buffer(&mut next_data)?;
        let new_group = match next_type {
            kind::ARR_FMLA => Some(FormulaGroup::parse_array(&next_data)?),
            kind::SHR_FMLA => Some(FormulaGroup::parse_shared(&next_data)?),
            _ => {
                self.pending_record = Some((next_type, next_data));
                None
            },
        };

        let group = if let Some(group) = new_group {
            if exp_cell.is_none() {
                return Err(Error::InvalidFormula(format!(
                    "{:?} formula definition is not preceded by PtgExp",
                    group.kind
                )));
            }
            if exp_cell != Some(position) {
                return Err(Error::InvalidFormula(format!(
                    "group definition at ({}, {}) is referenced by PtgExp ({}, {})",
                    position.0,
                    position.1,
                    exp_cell.map_or(u32::MAX, |target| target.0),
                    exp_cell.map_or(u32::MAX, |target| target.1)
                )));
            }
            if group.formula.rgce.first() == Some(&crate::package::formula::ptg_types::PTG_EXP) {
                return Err(Error::InvalidFormula(
                    "array/shared formula definition cannot contain PtgExp".to_string(),
                ));
            }
            match group.kind {
                GroupKind::Array if group.range.top_left() != position => {
                    return Err(Error::InvalidFormula(format!(
                        "BrtArrFmla range {} is not anchored at {}",
                        group.range.to_a1(),
                        crate::package::utils::cell_reference(position.0, position.1)
                    )));
                },
                GroupKind::Shared if group.range.top_left() != position => {
                    return Err(Error::InvalidFormula(format!(
                        "BrtShrFmla range {} is not anchored at {}",
                        group.range.to_a1(),
                        crate::package::utils::cell_reference(position.0, position.1)
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
            Ok(Cell::new_grouped_formula(
                position.0,
                parsed.header,
                parsed.cached_value,
                parsed.formula,
                parsed.flags,
                group,
                self.formula_context,
            ))
        } else if exp_cell.is_some() {
            Err(Error::InvalidFormula(format!(
                "PtgExp cell {} has no array/shared formula definition",
                crate::package::utils::cell_reference(position.0, position.1)
            )))
        } else {
            Ok(Cell::new_formula_binary(
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

    fn cell_value_from_number(value: f64) -> CellValue {
        if value == value.round() && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
            CellValue::Int(value as i64)
        } else {
            CellValue::Float(value)
        }
    }

    /// Read advanced features after sheet data
    ///
    /// This reads merged cells, hyperlinks, and other advanced features
    /// that appear after the sheet data section.
    fn read_advanced_features(&mut self) -> Result<()> {
        loop {
            self.buf.clear();

            // Try to read next record type
            let typ = match self.iter.read_type() {
                Ok(t) => t,
                Err(_) => break, // End of stream
            };

            let _ = self.iter.fill_buffer(&mut self.buf)?;

            match typ {
                kind::BEGIN_MERGE_CELLS => {
                    // BrtBeginMergeCells - start of merged cells section
                    self.read_merged_cells()?;
                },
                kind::MERGE_CELL => {
                    // BrtMergeCell - single merged cell
                    if let Ok(merged) = MergedCell::parse(&self.buf) {
                        self.merged_cells.push(merged);
                    }
                },
                kind::H_LINK => {
                    // BrtHLink - hyperlink
                    if let Ok(hyperlink) = Hyperlink::parse(&self.buf) {
                        self.hyperlinks.push(hyperlink);
                    }
                },
                kind::BEGIN_A_FILTER => {
                    if self.auto_filter.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginAFilter".to_string(),
                            val: "duplicate worksheet AutoFilter".to_string(),
                        });
                    }
                    self.auto_filter = Some(Self::parse_auto_filter(&self.buf)?);
                    self.consume_auto_filter_records()?;
                },
                kind::SHEET_PROTECTION => {
                    if self.sheet_protection.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtSheetProtection".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    self.sheet_protection = Some(Self::parse_sheet_protection(&self.buf)?);
                },
                kind::SHEET_PROTECTION_ISO => {
                    if self.sheet_protection.is_some() || self.strong_sheet_protection.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "duplicate protection record".to_string(),
                        });
                    }
                    let (strong, iso_flags) = Self::parse_strong_sheet_protection(&self.buf)?;
                    self.buf.clear();
                    let next_type = self.iter.read_type()?;
                    let _ = self.iter.fill_buffer(&mut self.buf)?;
                    if next_type != kind::SHEET_PROTECTION {
                        return Err(Error::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "not immediately followed by BrtSheetProtection".to_string(),
                        });
                    }
                    let base = Self::parse_sheet_protection(&self.buf)?;
                    if base.password_hash.is_some()
                        || Self::sheet_protection_flags(&base) != iso_flags
                    {
                        return Err(Error::Unrecognized {
                            typ: "BrtSheetProtectionIso".to_string(),
                            val: "following protection record does not match".to_string(),
                        });
                    }
                    self.sheet_protection = Some(base);
                    self.strong_sheet_protection = Some(strong);
                },
                kind::BEGIN_D_VALS => {
                    if self.data_validation_settings.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginDVals".to_string(),
                            val: "duplicate collection".to_string(),
                        });
                    }
                    let (settings, count) = parse_collection_settings(&self.buf, false)?;
                    self.data_validation_settings = Some(settings);
                    self.consume_classic_data_validations(count)?;
                },
                kind::BEGIN_D_VALS14 => {
                    if self.data_validation14_settings.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginDVals14".to_string(),
                            val: "duplicate collection".to_string(),
                        });
                    }
                    let (settings, count) = parse_collection_settings(&self.buf, true)?;
                    self.data_validation14_settings = Some(settings);
                    self.consume_extension_data_validations(count)?;
                },
                kind::BEGIN_COND_FORMATTING => {
                    let (mut formatting, count, base) = parse_classic_header(&self.buf)?;
                    self.consume_conditional_formatting(&mut formatting, count, base)?;
                    self.conditional_formattings.push(formatting);
                },
                kind::BEGIN_COND_FORMATTING14 => {
                    let (mut formatting, count, base) =
                        Formatting::parse_extension14_header_with_base(&self.buf)?;
                    self.consume_extension_conditional_formatting(&mut formatting, count, base)?;
                    self.conditional_formattings.push(formatting);
                },
                kind::BEGIN_WEB_EXTENSIONS => {
                    if self.saw_web_extension_collection {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginWebExtensions".to_string(),
                            val: "duplicate collection".to_string(),
                        });
                    }
                    if !self.buf.is_empty() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginWebExtensions".to_string(),
                            val: "begin record must be empty".to_string(),
                        });
                    }
                    self.saw_web_extension_collection = true;
                    self.consume_web_extension_bindings()?;
                },
                kind::END_SHEET => {
                    // BrtEndSheet - end of worksheet
                    self.resolve_conditional_formatting_links()?;
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }

        Ok(())
    }

    fn consume_web_extension_bindings(&mut self) -> Result<()> {
        let mut app_refs = HashSet::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::WEB_EXTENSION => {
                    if self.web_extension_bindings.len() == 65_536 {
                        return Err(Error::Unrecognized {
                            typ: "WEBEXTENSIONS".to_string(),
                            val: "binding count exceeds 65,536".to_string(),
                        });
                    }
                    let binding = Binding::parse_payload(&self.buf, |index| {
                        self.formula_context.is_internal_single_sheet_xti(index)
                    })?;
                    if !app_refs.insert(binding.application_reference.clone()) {
                        return Err(Error::Unrecognized {
                            typ: "WEBEXTENSIONS".to_string(),
                            val: "duplicate binding appRef".to_string(),
                        });
                    }
                    self.web_extension_bindings.push(binding);
                },
                kind::END_WEB_EXTENSIONS => {
                    if !self.buf.is_empty() {
                        return Err(Error::Unrecognized {
                            typ: "BrtEndWebExtensions".to_string(),
                            val: "end record must be empty".to_string(),
                        });
                    }
                    if self.web_extension_bindings.is_empty() {
                        return Err(Error::Unrecognized {
                            typ: "WEBEXTENSIONS".to_string(),
                            val: "collection requires at least one binding".to_string(),
                        });
                    }
                    return Ok(());
                },
                other => {
                    return Err(Error::Unrecognized {
                        typ: "WEBEXTENSIONS".to_string(),
                        val: format!("unexpected record 0x{other:04X}"),
                    });
                },
            }
        }
    }

    fn resolve_conditional_formatting_links(&mut self) -> Result<()> {
        let mut classic = HashMap::new();
        for formatting in &self.conditional_formattings {
            if formatting.record_kind != RecordKind::Classic {
                continue;
            }
            for rule in &formatting.rules {
                let Some(guid) = rule.classic_extension_guid else {
                    continue;
                };
                let bar = rule.data_bar.as_ref().ok_or_else(|| Error::Unrecognized {
                    typ: "BrtCFRuleExt".to_string(),
                    val: "is attached to a non-data-bar rule".to_string(),
                })?;
                if classic
                    .insert(guid, (rule.priority, bar.min_length, bar.max_length))
                    .is_some()
                {
                    return Err(Error::Unrecognized {
                        typ: "BrtCFRuleExt".to_string(),
                        val: "duplicate GUID".to_string(),
                    });
                }
            }
        }

        let mut matched = HashSet::new();
        for formatting in &mut self.conditional_formattings {
            if formatting.record_kind != RecordKind::Extension14 {
                continue;
            }
            for rule in &mut formatting.rules {
                let Some(metadata) = rule.extension14.as_mut() else {
                    continue;
                };
                if metadata.priority != -1 {
                    continue;
                }
                metadata.linked_classic_priority = None;
                if !metadata.guid_present {
                    continue;
                }
                let Some(&(priority, classic_min, classic_max)) = classic.get(&metadata.guid)
                else {
                    continue;
                };
                if !matched.insert(metadata.guid) {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginCFRule14".to_string(),
                        val: "multiple data-bar extensions use the same GUID".to_string(),
                    });
                }
                let bar = rule
                    .data_bar14
                    .as_ref()
                    .expect("validated CFRule14 visualization");
                let expected_lengths = if bar.min_length == 0 && bar.max_length == 100 {
                    (10, 90)
                } else {
                    (bar.min_length, bar.max_length)
                };
                if (classic_min, classic_max) != expected_lengths {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginDatabar14".to_string(),
                        val: "widths do not agree with the linked classic data bar".to_string(),
                    });
                }
                metadata.linked_classic_priority = Some(priority);
            }
        }
        if let Some(orphan) = classic.keys().find(|guid| !matched.contains(*guid)) {
            return Err(Error::Unrecognized {
                typ: "BrtCFRuleExt".to_string(),
                val: format!("GUID {orphan:02X?} has no matching data-bar extension"),
            });
        }
        Ok(())
    }

    fn consume_classic_data_validations(&mut self, expected_count: u32) -> Result<()> {
        let start = self.data_validations.len();
        let mut pending_list = None;
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::D_VAL_LIST => {
                    if pending_list.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtDValList".to_string(),
                            val: "consecutive list overrides".to_string(),
                        });
                    }
                    pending_list = Some(parse_dval_list(&self.buf)?);
                },
                kind::D_VAL => {
                    let rule = Validation::parse_classic(
                        &self.buf,
                        pending_list.take(),
                        self.formula_context,
                    )?;
                    self.data_validations.push(rule);
                },
                kind::END_D_VALS => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    if pending_list.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtDValList".to_string(),
                            val: "not followed by BrtDVal".to_string(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginDVals collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let found = self.data_validations.len() - start;
        if found != expected_count as usize {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDVals count".to_string(),
                val: format!("declared {expected_count}, found {found}"),
            });
        }
        Ok(())
    }

    fn consume_extension_data_validations(&mut self, expected_count: u32) -> Result<()> {
        let start = self.data_validations.len();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::D_VAL14 => self.data_validations.push(Validation::parse_extension14(
                    &self.buf,
                    self.formula_context,
                )?),
                kind::END_D_VALS14 => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginDVals14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let found = self.data_validations.len() - start;
        if found != expected_count as usize {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDVals14 count".to_string(),
                val: format!("declared {expected_count}, found {found}"),
            });
        }
        Ok(())
    }

    fn consume_conditional_formatting(
        &mut self,
        formatting: &mut Formatting,
        expected_count: u32,
        base: (u32, u32),
    ) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::BEGIN_CF_RULE => {
                    let mut rule = Rule::parse_with_context(&self.buf, base, self.formula_context)?;
                    if formatting
                        .rules
                        .iter()
                        .chain(self.conditional_formattings.iter().flat_map(|cf| &cf.rules))
                        .any(|existing| existing.priority == rule.priority)
                    {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginCFRule priority".to_string(),
                            val: format!("duplicate {}", rule.priority),
                        });
                    }
                    self.consume_conditional_rule(&mut rule, base)?;
                    formatting.rules.push(rule);
                },
                kind::END_COND_FORMATTING => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginConditionalFormatting collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if formatting.rules.len() != expected_count as usize {
            return Err(Error::Unrecognized {
                typ: "BrtBeginConditionalFormatting count".to_string(),
                val: format!(
                    "declared {expected_count}, found {}",
                    formatting.rules.len()
                ),
            });
        }
        Ok(())
    }

    fn consume_extension_conditional_formatting(
        &mut self,
        formatting: &mut Formatting,
        expected_count: u32,
        base: (u32, u32),
    ) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::BEGIN_CF_RULE14 => {
                    let mut rule = Rule::parse_extension14_with_context(
                        &self.buf,
                        base,
                        self.formula_context,
                    )?;
                    let signed_priority = rule
                        .extension14
                        .as_ref()
                        .expect("extension parser supplies metadata")
                        .priority;
                    if signed_priority > 0
                        && formatting
                            .rules
                            .iter()
                            .chain(self.conditional_formattings.iter().flat_map(|cf| &cf.rules))
                            .any(|existing| existing.priority == rule.priority)
                    {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginCFRule14 priority".to_string(),
                            val: format!("duplicate {}", rule.priority),
                        });
                    }
                    self.consume_extension_conditional_rule(&mut rule, base)?;
                    formatting.rules.push(rule);
                },
                kind::END_COND_FORMATTING14 => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginConditionalFormatting14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if formatting.rules.len() != expected_count as usize {
            return Err(Error::Unrecognized {
                typ: "BrtBeginConditionalFormatting14 count".to_string(),
                val: format!(
                    "declared {expected_count}, found {}",
                    formatting.rules.len()
                ),
            });
        }
        Ok(())
    }

    fn consume_extension_conditional_rule(
        &mut self,
        rule: &mut Rule,
        base: (u32, u32),
    ) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::BEGIN_COLOR_SCALE14 => {
                    if rule.color_scale14.is_some() || !self.buf.is_empty() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginColorScale14".to_string(),
                            val: "duplicate record or nonempty payload".to_string(),
                        });
                    }
                    rule.color_scale14 = Some(self.consume_color_scale14(base)?);
                },
                kind::BEGIN_DATABAR14 => {
                    if rule.data_bar14.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginDatabar14".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    let priority = rule
                        .extension14
                        .as_ref()
                        .expect("extension parser supplies metadata")
                        .priority;
                    rule.data_bar14 = Some(self.consume_data_bar14(&begin, base, priority)?);
                },
                kind::BEGIN_ICON_SET14 => {
                    if rule.icon_set14.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginIconSet14".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    rule.icon_set14 = Some(self.consume_icon_set14(&begin, base)?);
                },
                kind::END_CF_RULE14 => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginCFRule14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let valid_visual = match rule.rule_type {
            RuleType::ColorScale => {
                rule.color_scale14.is_some()
                    && rule.data_bar14.is_none()
                    && rule.icon_set14.is_none()
            },
            RuleType::DataBar => {
                rule.color_scale14.is_none()
                    && rule.data_bar14.is_some()
                    && rule.icon_set14.is_none()
            },
            RuleType::IconSet => {
                rule.color_scale14.is_none()
                    && rule.data_bar14.is_none()
                    && rule.icon_set14.is_some()
            },
            _ => {
                rule.color_scale14.is_none()
                    && rule.data_bar14.is_none()
                    && rule.icon_set14.is_none()
            },
        };
        if !valid_visual {
            return Err(Error::Unrecognized {
                typ: "BrtBeginCFRule14 collection".to_string(),
                val: "visualization records do not match rule type".to_string(),
            });
        }
        Ok(())
    }

    fn consume_conditional_rule(&mut self, rule: &mut Rule, base: (u32, u32)) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::BEGIN_COLOR_SCALE => {
                    if rule.color_scale.is_some() || !self.buf.is_empty() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginColorScale".to_string(),
                            val: "duplicate record or nonempty payload".to_string(),
                        });
                    }
                    rule.color_scale = Some(self.consume_color_scale(base)?);
                },
                kind::BEGIN_DATABAR => {
                    if rule.data_bar.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginDatabar".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    rule.data_bar = Some(self.consume_data_bar(&begin, base)?);
                },
                kind::BEGIN_ICON_SET => {
                    if rule.icon_set.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtBeginIconSet".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    let begin = self.buf.clone();
                    rule.icon_set = Some(self.consume_icon_set(&begin, base)?);
                },
                kind::CF_RULE_EXT => {
                    if rule.classic_extension_guid.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "BrtCFRuleExt".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    rule.classic_extension_guid = Some(parse_rule_extension_guid(&self.buf)?);
                },
                kind::END_CF_RULE => {
                    if !self.buf.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: self.buf.len(),
                        });
                    }
                    break;
                },
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginCFRule collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        let valid_visual = match rule.rule_type {
            RuleType::ColorScale => {
                rule.color_scale.is_some() && rule.data_bar.is_none() && rule.icon_set.is_none()
            },
            RuleType::DataBar => {
                rule.color_scale.is_none() && rule.data_bar.is_some() && rule.icon_set.is_none()
            },
            RuleType::IconSet => {
                rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_some()
            },
            _ => rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_none(),
        };
        if !valid_visual {
            return Err(Error::Unrecognized {
                typ: "BrtBeginCFRule collection".to_string(),
                val: "visualization records do not match rule type".to_string(),
            });
        }
        Ok(())
    }

    fn consume_color_scale(&mut self, base: (u32, u32)) -> Result<Scale> {
        let mut cfvos = Vec::new();
        let mut colors = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::CFVO if colors.is_empty() => cfvos.push(Value::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                kind::COLOR => colors.push(Color::parse(&self.buf)?),
                kind::END_COLOR_SCALE if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginColorScale collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if !(cfvos.len() == 2 || cfvos.len() == 3) || colors.len() != cfvos.len() {
            return Err(Error::Unrecognized {
                typ: "BrtBeginColorScale collection".to_string(),
                val: format!("{} thresholds and {} colors", cfvos.len(), colors.len()),
            });
        }
        if cfvos.first().is_some_and(|cfvo| cfvo.cfvo_type == 3)
            || cfvos.last().is_some_and(|cfvo| cfvo.cfvo_type == 2)
            || (cfvos.len() == 3 && matches!(cfvos[1].cfvo_type, 2 | 3))
        {
            return Err(Error::Unrecognized {
                typ: "BrtBeginColorScale collection".to_string(),
                val: "invalid min/mid/max threshold type".to_string(),
            });
        }
        let has_middle = colors.len() == 3;
        let mut cfvos = cfvos.into_iter();
        let min_cfvo = cfvos.next().ok_or_else(|| Error::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing minimum threshold".to_string(),
        })?;
        let middle_cfvo = if has_middle { cfvos.next() } else { None };
        let max_cfvo = cfvos.next().ok_or_else(|| Error::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing maximum threshold".to_string(),
        })?;
        let mut colors = colors.into_iter();
        let min_color_record = colors.next().ok_or_else(|| Error::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing minimum color".to_string(),
        })?;
        let mid_color_record = if has_middle { colors.next() } else { None };
        let max_color_record = colors.next().ok_or_else(|| Error::Unrecognized {
            typ: "BrtBeginColorScale collection".to_string(),
            val: "missing maximum color".to_string(),
        })?;
        Ok(Scale {
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

    fn consume_color_scale14(&mut self, base: (u32, u32)) -> Result<Scale> {
        let mut cfvos = Vec::new();
        let mut colors = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::CFVO14 if colors.is_empty() => cfvos.push(
                    Value::parse_extension14_with_context(&self.buf, base, self.formula_context)?,
                ),
                kind::COLOR14 => colors.push(Color::parse_extension14(&self.buf)?),
                kind::END_COLOR_SCALE14 if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginColorScale14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        build_color_scale(cfvos, colors, "BrtBeginColorScale14 collection", true)
    }

    fn consume_data_bar(&mut self, begin: &[u8], base: (u32, u32)) -> Result<Bar> {
        if begin.len() != 3 || begin[0] > begin[1] || begin[1] > 100 || begin[2] > 1 {
            return Err(Error::Unrecognized {
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
                kind::CFVO if color.is_none() => cfvos.push(Value::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                kind::COLOR if color.is_none() => color = Some(Color::parse(&self.buf)?),
                kind::END_DATABAR if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginDatabar collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != 2 || color.is_none() {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: format!("{} thresholds, color={}", cfvos.len(), color.is_some()),
            });
        }
        if cfvos[0].cfvo_type == 3 || cfvos[1].cfvo_type == 2 {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: "invalid minimum/maximum threshold type".to_string(),
            });
        }
        let [min_cfvo, max_cfvo]: [Value; 2] =
            cfvos.try_into().map_err(|_| Error::Unrecognized {
                typ: "BrtBeginDatabar collection".to_string(),
                val: "invalid threshold count".to_string(),
            })?;
        let color_record = color.ok_or_else(|| Error::Unrecognized {
            typ: "BrtBeginDatabar collection".to_string(),
            val: "missing color".to_string(),
        })?;
        Ok(Bar {
            min_cfvo,
            max_cfvo,
            color: color_record.argb.unwrap_or(0),
            show_value: begin[2] != 0,
            min_length: begin[0],
            max_length: begin[1],
            color_record,
        })
    }

    fn consume_data_bar14(
        &mut self,
        begin: &[u8],
        base: (u32, u32),
        priority: i32,
    ) -> Result<Bar14> {
        let header = Bar14::parse_header(begin)?;
        let mut cfvos = Vec::new();
        let mut colors = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::CFVO14 if colors.is_empty() => cfvos.push(
                    Value::parse_extension14_with_context(&self.buf, base, self.formula_context)?,
                ),
                kind::COLOR14 => colors.push(Color::parse_extension14(&self.buf)?),
                kind::END_DATABAR14 if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginDatabar14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != 2
            || matches!(cfvos[0].cfvo_type, 3 | 9)
            || matches!(cfvos[1].cfvo_type, 2 | 8)
        {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDatabar14 collection".to_string(),
                val: "invalid minimum/maximum thresholds".to_string(),
            });
        }
        let expected_colors = usize::from(priority != -1)
            + usize::from(header.border)
            + usize::from(header.custom_negative_fill)
            + usize::from(header.custom_negative_border && header.border)
            + usize::from(header.axis_position != AxisPosition14::None);
        if colors.len() != expected_colors {
            return Err(Error::Unrecognized {
                typ: "BrtBeginDatabar14 collection".to_string(),
                val: format!("expected {expected_colors} colors, found {}", colors.len()),
            });
        }
        let [min_cfvo, max_cfvo]: [Value; 2] = cfvos.try_into().expect("validated two thresholds");
        let mut colors = colors.into_iter();
        let positive_color = (priority != -1).then(|| colors.next()).flatten();
        let border_color = header.border.then(|| colors.next()).flatten();
        let negative_color = header.custom_negative_fill.then(|| colors.next()).flatten();
        let negative_border_color = (header.custom_negative_border && header.border)
            .then(|| colors.next())
            .flatten();
        let axis_color = (header.axis_position != AxisPosition14::None)
            .then(|| colors.next())
            .flatten();
        Ok(Bar14 {
            min_cfvo,
            max_cfvo,
            positive_color,
            border_color,
            negative_color,
            negative_border_color,
            axis_color,
            min_length: header.min_length,
            max_length: header.max_length,
            show_value: header.show_value,
            direction: header.direction,
            axis_position: header.axis_position,
            border: header.border,
            gradient: header.gradient,
            custom_negative_fill: header.custom_negative_fill,
            custom_negative_border: header.custom_negative_border,
            unused_flags: header.unused_flags,
        })
    }

    fn consume_icon_set(&mut self, begin: &[u8], base: (u32, u32)) -> Result<IconSet> {
        if begin.len() != 6 {
            return Err(Error::InvalidLength {
                expected: 6,
                found: begin.len(),
            });
        }
        let icon_set = binary::read_u32_le_at(begin, 0)?;
        let flags = binary::read_u16_le_at(begin, 4)?;
        if icon_set > 16 || flags & !0x7e != 0 {
            return Err(Error::Unrecognized {
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
                kind::CFVO => cfvos.push(Value::parse_with_context(
                    &self.buf,
                    base,
                    self.formula_context,
                )?),
                kind::END_ICON_SET if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginIconSet collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != expected {
            return Err(Error::Unrecognized {
                typ: "BrtBeginIconSet collection".to_string(),
                val: format!("expected {expected} thresholds, found {}", cfvos.len()),
            });
        }
        if cfvos
            .iter()
            .any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9) || !cfvo.save_greater_than_or_equal)
        {
            return Err(Error::Unrecognized {
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

    fn consume_icon_set14(&mut self, begin: &[u8], base: (u32, u32)) -> Result<IconSet14> {
        let header = IconSet14::parse_header(begin)?;
        let expected = icon_count14(header.icon_set_type);
        let mut cfvos = Vec::new();
        let mut icons = Vec::new();
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            match typ {
                kind::CFVO14 if icons.is_empty() => cfvos.push(
                    Value::parse_extension14_with_context(&self.buf, base, self.formula_context)?,
                ),
                kind::CF_ICON => icons.push(Icon::parse(&self.buf)?),
                kind::END_ICON_SET14 if self.buf.is_empty() => break,
                _ => {
                    return Err(Error::Unrecognized {
                        typ: "BrtBeginIconSet14 collection".to_string(),
                        val: format!("unexpected record 0x{typ:04X}"),
                    });
                },
            }
        }
        if cfvos.len() != expected
            || cfvos.iter().any(|cfvo| {
                matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9) || !cfvo.save_greater_than_or_equal
            })
        {
            return Err(Error::Unrecognized {
                typ: "BrtBeginIconSet14 collection".to_string(),
                val: format!("invalid set of {} thresholds", cfvos.len()),
            });
        }
        if (header.custom && icons.len() != expected) || (!header.custom && !icons.is_empty()) {
            return Err(Error::Unrecognized {
                typ: "BrtBeginIconSet14 collection".to_string(),
                val: format!("invalid set of {} custom icons", icons.len()),
            });
        }
        Ok(IconSet14 {
            icon_set_type: header.icon_set_type,
            cfvos,
            custom_icons: header.custom.then_some(icons),
            show_value: header.show_value,
            reverse: header.reverse,
            unused_flags: header.unused_flags,
        })
    }

    fn parse_auto_filter(data: &[u8]) -> Result<AutoFilter> {
        if data.len() != 16 {
            return Err(Error::InvalidLength {
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
            return Err(Error::Unrecognized {
                typ: "BrtBeginAFilter rfx".to_string(),
                val: format!(
                    "rows {first_row}..={last_row}, columns {first_column}..={last_column}"
                ),
            });
        }
        Ok(AutoFilter {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }

    fn consume_auto_filter_records(&mut self) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            if typ == kind::END_A_FILTER {
                if !self.buf.is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: self.buf.len(),
                    });
                }
                return Ok(());
            }
            if typ == kind::BEGIN_A_FILTER {
                return Err(Error::Unrecognized {
                    typ: "BrtBeginAFilter".to_string(),
                    val: "nested AutoFilter".to_string(),
                });
            }
        }
    }

    fn parse_sheet_protection(data: &[u8]) -> Result<SheetProtection> {
        if data.len() != 66 {
            return Err(Error::InvalidLength {
                expected: 66,
                found: data.len(),
            });
        }
        let password = binary::read_u16_le_at(data, 0)?;
        let mut flags = [false; 16];
        for (index, flag) in flags.iter_mut().enumerate() {
            let value = binary::read_u32_le_at(data, 2 + index * 4)?;
            if value > 1 {
                return Err(Error::Unrecognized {
                    typ: "BrtSheetProtection Boolean".to_string(),
                    val: format!("field {index}: {value}"),
                });
            }
            *flag = value != 0;
        }
        Ok(SheetProtection {
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

    fn sheet_protection_flags(protection: &SheetProtection) -> [bool; 16] {
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

    fn parse_strong_sheet_protection(data: &[u8]) -> Result<(StrongProtection, [bool; 16])> {
        if data.len() < 83 {
            return Err(Error::InvalidLength {
                expected: 83,
                found: data.len(),
            });
        }
        let spin_count = binary::read_u32_le_at(data, 0)?;
        if spin_count > 10_000_000 {
            return Err(Error::Unrecognized {
                typ: "BrtSheetProtectionIso dwSpinCount".to_string(),
                val: spin_count.to_string(),
            });
        }
        let mut flags = [false; 16];
        for (index, flag) in flags.iter_mut().enumerate() {
            let value = binary::read_u32_le_at(data, 4 + index * 4)?;
            if value > 1 {
                return Err(Error::Unrecognized {
                    typ: "BrtSheetProtectionIso Boolean".to_string(),
                    val: format!("field {index}: {value}"),
                });
            }
            *flag = value != 0;
        }
        let mut offset = 68;
        let hash = Self::read_length_prefixed_bytes(data, &mut offset, "rgbHash")?;
        if hash.is_empty() {
            return Err(Error::Unrecognized {
                typ: "BrtSheetProtectionIso rgbHash".to_string(),
                val: "empty".to_string(),
            });
        }
        let salt = Self::read_length_prefixed_bytes(data, &mut offset, "rgbSalt")?;
        let algorithm = Self::read_nullable_wide_string(data, &mut offset)?
            .filter(|value| !value.is_empty())
            .ok_or_else(|| Error::Unrecognized {
                typ: "BrtSheetProtectionIso szAlgName".to_string(),
                val: "null or empty".to_string(),
            })?;
        if offset != data.len() {
            return Err(Error::Unrecognized {
                typ: "BrtSheetProtectionIso".to_string(),
                val: format!("{} trailing bytes", data.len() - offset),
            });
        }
        Ok((
            StrongProtection {
                spin_count,
                hash,
                salt,
                algorithm,
            },
            flags,
        ))
    }

    fn read_length_prefixed_bytes(data: &[u8], offset: &mut usize, field: &str) -> Result<Vec<u8>> {
        let data_offset = offset.checked_add(4).ok_or_else(|| {
            Error::Encoding(format!("BrtSheetProtectionIso {field} offset overflow"))
        })?;
        if data_offset > data.len() {
            return Err(Error::InvalidLength {
                expected: data_offset,
                found: data.len(),
            });
        }
        let count = binary::read_u32_le_at(data, *offset)? as usize;
        let end = data_offset.checked_add(count).ok_or_else(|| {
            Error::Encoding(format!("BrtSheetProtectionIso {field} size overflow"))
        })?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        *offset = end;
        Ok(data[data_offset..end].to_vec())
    }

    fn read_nullable_wide_string(data: &[u8], offset: &mut usize) -> Result<Option<String>> {
        let text_offset = offset
            .checked_add(4)
            .ok_or_else(|| Error::Encoding("ISO algorithm offset overflow".to_string()))?;
        if text_offset > data.len() {
            return Err(Error::InvalidLength {
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
            .ok_or_else(|| Error::Encoding("ISO algorithm size overflow".to_string()))?;
        let end = text_offset
            .checked_add(byte_count)
            .ok_or_else(|| Error::Encoding("ISO algorithm end overflow".to_string()))?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let units = data[text_offset..end]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let value = String::from_utf16(&units)
            .map_err(|error| Error::Encoding(format!("invalid ISO algorithm: {error}")))?;
        *offset = end;
        Ok(Some(value))
    }

    /// Read merged cells section
    fn read_merged_cells(&mut self) -> Result<()> {
        loop {
            self.buf.clear();
            let typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;

            if typ == kind::END_MERGE_CELLS {
                // BrtEndMergeCells
                break;
            }

            if typ == kind::MERGE_CELL {
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
    use crate::conditional_formatting::{AxisPosition14, RecordKind, RuleMetadata};
    use crate::raw::Writer;
    use litchi_core::sheet::Cell;

    fn rich_string_worksheet() -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_COL_INFOS, &[]).unwrap();
        let mut column = 2u32.to_le_bytes().to_vec();
        column.extend_from_slice(&4u32.to_le_bytes());
        column.extend_from_slice(&4096u32.to_le_bytes());
        column.extend_from_slice(&0u32.to_le_bytes());
        column.extend_from_slice(&0x130Fu16.to_le_bytes());
        writer.write_record(kind::COL_INFO, &column).unwrap();
        writer.write_record(kind::END_COL_INFOS, &[]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();

        let mut row = 4u32.to_le_bytes().to_vec();
        row.extend_from_slice(&0u32.to_le_bytes());
        row.extend_from_slice(&500u16.to_le_bytes());
        row.extend_from_slice(&[3, 0x7A, 1]);
        row.extend_from_slice(&1u32.to_le_bytes());
        row.extend_from_slice(&2u32.to_le_bytes());
        row.extend_from_slice(&2u32.to_le_bytes());
        writer.write_record(kind::ROW_HDR, &row).unwrap();

        let mut cell = 2u32.to_le_bytes().to_vec();
        cell.extend_from_slice(&[0, 0, 0, 1]);
        cell.push(1);
        cell.extend_from_slice(&2u32.to_le_bytes());
        cell.extend_from_slice(&[b'A', 0, b'B', 0]);
        cell.extend_from_slice(&2u32.to_le_bytes());
        cell.extend_from_slice(&[0, 0, 3, 0, 1, 0, 5, 0]);
        writer.write_record(kind::CELL_R_STRING, &cell).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();
        bytes
    }

    #[test]
    fn decodes_and_validates_cell_style_header() {
        type Reader<'a> = CellsReader<'a, std::io::Cursor<&'a [u8]>>;

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
            Err(Error::Unrecognized { .. })
        ));

        let reserved = [0, 0, 0, 0, 0, 0, 0, 2];
        assert!(matches!(
            Reader::decode_cell_header(&reserved, 1),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn parses_iso_sheet_protection_metadata_and_matching_flags() {
        type Reader<'a> = CellsReader<'a, std::io::Cursor<&'a [u8]>>;
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
        let mut writer = Writer::new(&mut worksheet);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        writer
            .write_record(kind::SHEET_PROTECTION_ISO, &iso)
            .unwrap();
        writer.write_record(kind::SHEET_PROTECTION, &base).unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();
        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(worksheet));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(reader.next_cell().unwrap().is_none());
        assert_eq!(reader.sheet_protection.unwrap(), protection);
        assert_eq!(reader.strong_sheet_protection.unwrap(), strong);
    }

    #[test]
    fn reads_inline_rich_string_cells() {
        let bytes = rich_string_worksheet();
        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(bytes));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();

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
        let mut writer = Writer::new(&mut worksheet);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        let mut begin = vec![0; 14];
        begin.extend_from_slice(&1u32.to_le_bytes());
        writer.write_record(kind::BEGIN_D_VALS, &begin).unwrap();
        writer.write_record(kind::END_D_VALS, &[]).unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(worksheet));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(matches!(
            reader.next_cell(),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn rejects_conditional_formatting_with_a_mismatched_rule_count() {
        let mut worksheet = Vec::new();
        let mut writer = Writer::new(&mut worksheet);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        let mut begin = 1u32.to_le_bytes().to_vec();
        begin.extend_from_slice(&0u32.to_le_bytes());
        begin.extend_from_slice(&1u32.to_le_bytes());
        begin.extend_from_slice(&[0; 16]);
        writer
            .write_record(kind::BEGIN_COND_FORMATTING, &begin)
            .unwrap();
        writer.write_record(kind::END_COND_FORMATTING, &[]).unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(worksheet));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(matches!(
            reader.next_cell(),
            Err(Error::Unrecognized { .. })
        ));
    }

    #[test]
    fn reads_office_2013_conditional_formatting_visualizations() {
        fn metadata(priority: i32) -> RuleMetadata {
            RuleMetadata {
                priority,
                unused: priority as u32,
                guid: [priority as u8; 16],
                guid_present: true,
                linked_classic_priority: None,
            }
        }

        let mut color_rule = Rule::new(RuleType::ColorScale, 1);
        color_rule.extension14 = Some(metadata(1));
        let color_thresholds = [Value::new(2, None), Value::new(3, None)];
        let color_records = [Color::from_argb(0xffff_0000), Color::from_argb(0xff00_ff00)];

        let mut bar_rule = Rule::new(RuleType::DataBar, 2);
        bar_rule.extension14 = Some(metadata(2));
        let bar = Bar14::new(
            Value::new(8, None),
            Value::new(9, None),
            Color::from_argb(0xff44_72c4),
        );

        let mut icon_rule = Rule::new(RuleType::IconSet, 3);
        icon_rule.extension14 = Some(metadata(3));
        let mut icon_thresholds = vec![
            Value::new(1, Some("0".to_string())),
            Value::new(1, Some("33".to_string())),
            Value::new(1, Some("67".to_string())),
        ];
        for threshold in &mut icon_thresholds {
            threshold.save_greater_than_or_equal = true;
        }
        let icon_set = IconSet14::new(18, icon_thresholds.clone());

        let mut formatting = Formatting::new(vec!["A1:A10".to_string()]);
        formatting.record_kind = RecordKind::Extension14;
        formatting.rules = vec![color_rule.clone(), bar_rule.clone(), icon_rule.clone()];

        let mut worksheet = Vec::new();
        let mut writer = Writer::new(&mut worksheet);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        writer
            .write_record(
                kind::BEGIN_COND_FORMATTING14,
                &formatting.serialize_extension14_header().unwrap(),
            )
            .unwrap();

        writer
            .write_record(
                kind::BEGIN_CF_RULE14,
                &color_rule.serialize_extension14().unwrap(),
            )
            .unwrap();
        writer.write_record(kind::BEGIN_COLOR_SCALE14, &[]).unwrap();
        for threshold in &color_thresholds {
            writer
                .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
                .unwrap();
        }
        for color in color_records {
            writer
                .write_record(kind::COLOR14, &color.serialize_extension14().unwrap())
                .unwrap();
        }
        writer.write_record(kind::END_COLOR_SCALE14, &[]).unwrap();
        writer.write_record(kind::END_CF_RULE14, &[]).unwrap();

        writer
            .write_record(
                kind::BEGIN_CF_RULE14,
                &bar_rule.serialize_extension14().unwrap(),
            )
            .unwrap();
        writer
            .write_record(kind::BEGIN_DATABAR14, &bar.serialize_header().unwrap())
            .unwrap();
        for threshold in [&bar.min_cfvo, &bar.max_cfvo] {
            writer
                .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
                .unwrap();
        }
        for color in [bar.positive_color, bar.axis_color].into_iter().flatten() {
            writer
                .write_record(kind::COLOR14, &color.serialize_extension14().unwrap())
                .unwrap();
        }
        writer.write_record(kind::END_DATABAR14, &[]).unwrap();
        writer.write_record(kind::END_CF_RULE14, &[]).unwrap();

        writer
            .write_record(
                kind::BEGIN_CF_RULE14,
                &icon_rule.serialize_extension14().unwrap(),
            )
            .unwrap();
        writer
            .write_record(
                kind::BEGIN_ICON_SET14,
                &icon_set.serialize_header().unwrap(),
            )
            .unwrap();
        for threshold in &icon_thresholds {
            writer
                .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
                .unwrap();
        }
        writer.write_record(kind::END_ICON_SET14, &[]).unwrap();
        writer.write_record(kind::END_CF_RULE14, &[]).unwrap();
        writer
            .write_record(kind::END_COND_FORMATTING14, &[])
            .unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(worksheet));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(reader.next_cell().unwrap().is_none());
        let parsed = &reader.conditional_formattings[0];
        assert_eq!(parsed.record_kind, RecordKind::Extension14);
        assert_eq!(parsed.rules.len(), 3);
        assert_eq!(
            parsed.rules[0]
                .color_scale14
                .as_ref()
                .unwrap()
                .min_cfvo
                .cfvo_type,
            2
        );
        let parsed_bar = parsed.rules[1].data_bar14.as_ref().unwrap();
        assert_eq!(parsed_bar.axis_position, AxisPosition14::Automatic);
        assert_eq!(parsed_bar.positive_color.unwrap().argb, Some(0xff44_72c4));
        let parsed_icons = parsed.rules[2].icon_set14.as_ref().unwrap();
        assert_eq!(parsed_icons.icon_set_type, 18);
        assert_eq!(parsed_icons.cfvos.len(), 3);
    }

    #[test]
    fn rejects_incomplete_color_scale_collection() {
        let mut worksheet = Vec::new();
        let mut writer = Writer::new(&mut worksheet);
        writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
        writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
        writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
        let mut begin = 1u32.to_le_bytes().to_vec();
        begin.extend_from_slice(&0u32.to_le_bytes());
        begin.extend_from_slice(&1u32.to_le_bytes());
        begin.extend_from_slice(&[0; 16]);
        writer
            .write_record(kind::BEGIN_COND_FORMATTING, &begin)
            .unwrap();

        let mut rule = 3u32.to_le_bytes().to_vec();
        rule.extend_from_slice(&2u32.to_le_bytes());
        rule.extend_from_slice(&u32::MAX.to_le_bytes());
        rule.extend_from_slice(&1u32.to_le_bytes());
        rule.extend_from_slice(&[0; 10]);
        rule.extend_from_slice(&[0; 12]);
        rule.extend_from_slice(&u32::MAX.to_le_bytes());
        writer.write_record(kind::BEGIN_CF_RULE, &rule).unwrap();
        writer.write_record(kind::BEGIN_COLOR_SCALE, &[]).unwrap();
        writer.write_record(kind::END_COLOR_SCALE, &[]).unwrap();
        writer.write_record(kind::END_CF_RULE, &[]).unwrap();
        writer.write_record(kind::END_COND_FORMATTING, &[]).unwrap();
        writer.write_record(kind::END_SHEET, &[]).unwrap();

        let formula_context = FormulaResolutionContext::default();
        let iter = Stream::new(std::io::Cursor::new(worksheet));
        let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
        assert!(reader.next_cell().is_err());
    }
}
