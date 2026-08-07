//! BIFF12 stream decoding and worksheet feature orchestration.

use super::model::{CellsReader, Dimensions, ParsedFormulaCell};
use crate::conditional_formatting::{Formatting, parse_classic_header};
use crate::hyperlinks::Hyperlink;
use crate::merged_cells::MergedCell;
use crate::package::cell::{Cell, CellHeader};
use crate::package::data_validation::parse_collection_settings;
use crate::package::error::{Error, Result};
use crate::package::formula::{Context, Group, GroupKind, ParsedFormula};
use crate::package::records::Stream;
use crate::package::shared_strings::SharedString;
use crate::package::web_extension_bindings::Binding;
use crate::raw::{Cursor, kind};
use litchi_core::binary;
use litchi_core::sheet::CellValue;
use std::collections::HashSet;
use std::io::{Read, Seek};
use std::sync::Arc;

impl<'a, RS> CellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub fn new(
        mut iter: Stream<RS>,
        shared_strings: &'a [SharedString],
        formula_context: &'a Context,
        cell_xf_count: usize,
    ) -> Result<Self> {
        let mut buf = Vec::with_capacity(1024);

        // Walk the worksheet preamble up to BrtWsDim (worksheet dimensions),
        // capturing the sheet-view collection while skipping everything else.
        let mut views = Vec::new();
        loop {
            let typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;
            if typ == kind::WS_DIM {
                break;
            }
            if typ == kind::BEGIN_WS_VIEWS {
                views = crate::package::sheet_view::read_views(&mut iter, &mut buf)?;
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
                views = crate::package::sheet_view::read_views(&mut iter, &mut buf)?;
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
            views,
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
                        let value = super::semantic::cell_value_from_number(cursor.read_rk()?);
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
                        let (string, _) = crate::package::records::decode_string(&self.buf[8..])?;
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
                        crate::package::records::decode_string(&self.buf[8..])?;
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
                    let error_msg = super::semantic::error_text(self.buf[8]);
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
        let (formula, consumed) = ParsedFormula::parse(&self.buf[formula_offset..])?;
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
            kind::ARR_FMLA => Some(Group::parse_array(&next_data)?),
            kind::SHR_FMLA => Some(Group::parse_shared(&next_data)?),
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
