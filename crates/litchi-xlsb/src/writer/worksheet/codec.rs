//! BIFF12 worksheet record encoding.

use crate::package::error::{Error, Result};
use crate::package::formula::{Group, GroupKind, ParsedFormula, Range};
use crate::package::sheet_view::{
    MAX_SHEET_VIEW_SELECTIONS, SheetPane, SheetPanePosition, SheetPaneState, SheetSelection,
};
use crate::raw::{Writer, kind};
use litchi_core::sheet::CellValue;
use std::collections::BTreeMap;
use std::io::Write;

use super::model::{CellData, MutableWorksheet};

impl MutableWorksheet {
    /// Write worksheet properties (BrtWsProp) - REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.864 + spec example 3.7.21: 23 bytes total
    /// Structure: flags (3 bytes) + brtcolorTab (8 bytes) + rwSync (4) + colSync (4) + strName (4)
    pub(super) fn write_ws_properties<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Flags (3 bytes per spec example 3.7.21):
        // Byte 0-1 (USHORT): flags A-O
        // Byte 2 (BYTE): flags P-Q + reserved
        //
        // From spec example: 0xC9, 0x04, 0x02
        // 0xC9 = fShowAutoBreaks(1) + fPublish(1) + fRowSumsBelow(1) + fColSumsRight(1) + fShowOutlineSymbols(1)
        // 0x04 = remaining bits
        // 0x02 = fCondFmtCalc(1) at bit 1
        temp_writer.write_u8(0xC9)?;
        temp_writer.write_u8(0x04)?;
        temp_writer.write_u8(0x02)?; // Third byte - fCondFmtCalc flag

        // brtcolorTab (8 bytes) - BrtColor structure
        // From spec example: xColorType=0x00 (auto), index=0x40
        temp_writer.write_u8(0x00)?; // fValidRGB(0) + xColorType(0x00)
        temp_writer.write_u8(0x40)?; // index
        temp_writer.write_u16(0)?; // nTintAndShade
        temp_writer.write_u8(0)?; // bRed
        temp_writer.write_u8(0)?; // bGreen
        temp_writer.write_u8(0)?; // bBlue
        temp_writer.write_u8(0)?; // bAlpha

        // rwSync (4 bytes) - RwNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // colSync (4 bytes) - ColNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // strName - CodeName (XLWideString): empty string
        temp_writer.write_u32(0)?;

        writer.write_record(kind::WS_PROP, &data)?;
        Ok(())
    }

    /// Write worksheet views (REQUIRED by Excel)
    ///
    /// [MS-XLSB] 2.4.307/2.4.308: BrtBeginWsViews / BrtBeginWsView, optionally
    /// followed by BrtPane (2.4.723) and BrtSel (2.4.790) records.
    pub(super) fn write_ws_views<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut view = self.sheet_view.clone().unwrap_or_default();
        let configured = self.sheet_view.is_some() || self.freeze_panes.is_some();

        if let Some(freeze) = &self.freeze_panes {
            if view.pane.is_some() || !view.selections.is_empty() {
                return Err(Error::Unrecognized {
                    typ: "worksheet sheet view".to_string(),
                    val: "freeze panes and explicit sheet-view pane selections cannot both be set"
                        .to_string(),
                });
            }
            let y_split = freeze.freeze_rows;
            let x_split = freeze.freeze_cols;
            let active_pane = match (x_split > 0, y_split > 0) {
                (true, true) => SheetPanePosition::BottomRight,
                (true, false) => SheetPanePosition::TopRight,
                (false, true) => SheetPanePosition::BottomLeft,
                (false, false) => unreachable!("freeze panes require a nonzero split"),
            };
            let top_left_cell = crate::package::utils::cell_reference(y_split, x_split);
            view.pane = Some(SheetPane {
                x_split: (x_split > 0).then_some(f64::from(x_split)),
                y_split: (y_split > 0).then_some(f64::from(y_split)),
                top_left_cell: Some(top_left_cell.clone()),
                active_pane: Some(active_pane),
                state: Some(SheetPaneState::Frozen),
            });
            view.selections.push(SheetSelection {
                pane: Some(active_pane),
                active_cell: Some(top_left_cell.clone()),
                active_cell_id: None,
                sqref: Some(top_left_cell),
            });
        }

        if view.selections.len() > MAX_SHEET_VIEW_SELECTIONS {
            return Err(Error::Unrecognized {
                typ: "worksheet sheet view".to_string(),
                val: "a sheet view cannot contain more than four selections".to_string(),
            });
        }

        writer.write_record(kind::BEGIN_WS_VIEWS, &[])?;

        // BrtBeginWsView (30 bytes according to spec)
        let view_data =
            crate::package::sheet_view::write_ws_view_payload(configured.then_some(&view))?;
        writer.write_record(kind::BEGIN_WS_VIEW, &view_data)?;

        if let Some(pane) = view.pane.as_ref() {
            let pane_data = crate::package::sheet_view::write_pane_payload(pane)?;
            writer.write_record(kind::PANE, &pane_data)?;
        }
        for selection in &view.selections {
            let selection_data = crate::package::sheet_view::write_selection_payload(selection)?;
            writer.write_record(kind::SEL, &selection_data)?;
        }

        writer.write_record(kind::END_WS_VIEW, &[])?;
        writer.write_record(kind::END_WS_VIEWS, &[])?;

        Ok(())
    }

    /// Write SHEET_FORMAT_PR record (0x01E5) - sheet formatting properties
    /// REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.862 + spec example 3.7.28: 12 bytes total
    pub(super) fn write_sheet_format_pr<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // dxGCol (4 bytes) - 0xFFFFFFFF = use cchDefColWidth instead
        temp_writer.write_u32(0xFFFFFFFF)?;

        // cchDefColWidth (2 bytes) - default column width in characters
        // Spec example 3.7.28: 0x0008 (8 characters)
        temp_writer.write_u16(8)?;

        // miyDefRwHeight (2 bytes) - default row height in twips
        // Spec example 3.7.28: 0x012C (300 twips = 15 points)
        temp_writer.write_u16(300)?;

        // Flags (4 bytes): all zeros per spec example
        // fUnsynced=0, fDyZero=0, fExAsc=0, fExDesc=0, reserved=0, iOutLevelRw=0, iOutLevelCol=0
        temp_writer.write_u32(0)?;

        writer.write_record(kind::WS_FMT_INFO, &data)?;
        Ok(())
    }

    /// Write worksheet dimensions record
    pub(super) fn write_dimensions<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        if let Some((min_row, min_col, max_row, max_col)) = self.dimensions() {
            let (max_row, max_col) = self
                .formula_groups_for_write()?
                .iter()
                .fold((max_row, max_col), |(row, col), group| {
                    (row.max(group.range.row_last), col.max(group.range.col_last))
                });
            temp_writer.write_u32(min_row)?;
            temp_writer.write_u32(max_row)?;
            temp_writer.write_u32(min_col)?;
            temp_writer.write_u32(max_col)?;
        } else {
            // Empty worksheet
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
        }

        writer.write_record(kind::WS_DIM, &data)?;
        Ok(())
    }

    /// Write all cells
    pub(super) fn write_cells<W: Write>(
        &self,
        writer: &mut Writer<W>,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
    ) -> Result<()> {
        let formula_groups = self.formula_groups_for_write()?;
        if formula_groups.is_empty() {
            return self.write_cells_from(writer, shared_strings, &self.cells, &formula_groups);
        }

        // Grouped formulas require a formula cell record at every position in
        // their range. Materialize only for this uncommon path, keeping the
        // ordinary worksheet writer allocation-free.
        let mut cells = self.cells.clone();
        for group in &formula_groups {
            for row in group.range.row_first..=group.range.row_last {
                for col in group.range.col_first..=group.range.col_last {
                    cells.entry((row, col)).or_insert_with(|| CellData {
                        value: CellValue::Empty,
                        style: 0,
                        formula_binary: None,
                        formula_flags: 0,
                    });
                }
            }
        }
        self.write_cells_from(writer, shared_strings, &cells, &formula_groups)
    }

    pub(super) fn write_cells_from<W: Write>(
        &self,
        writer: &mut Writer<W>,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
        cells: &BTreeMap<(u32, u32), CellData>,
        formula_groups: &[Group],
    ) -> Result<()> {
        let mut current_row: Option<u32> = None;

        for ((row, col), cell_data) in cells {
            // Write row header if row changed
            if current_row != Some(*row) {
                self.write_row_header(writer, *row, cells)?;
                current_row = Some(*row);
            }

            if let Some(group) = Self::formula_group_for_cell(formula_groups, *row, *col) {
                self.write_grouped_formula_cell(writer, *row, *col, cell_data, group)?;
            } else {
                self.write_cell(writer, *row, *col, cell_data, shared_strings)?;
            }
        }

        Ok(())
    }

    pub(super) fn formula_groups_for_write(&self) -> Result<Vec<Group>> {
        let mut groups = self.formula_groups.clone();
        for (&position, cell) in &self.cells {
            let CellValue::Formula {
                formula,
                cached_value,
                is_array: true,
                array_range,
            } = &cell.value
            else {
                continue;
            };
            let range_text = array_range.as_deref().ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "array formula at {} has no array range",
                    crate::package::utils::cell_reference(position.0, position.1)
                ))
            })?;
            let range = Range::parse_a1(range_text)?;
            if range.top_left() != position {
                continue;
            }
            if groups
                .iter()
                .any(|group| group.kind == GroupKind::Array && group.range == range)
            {
                continue;
            }
            groups.push(Group {
                kind: GroupKind::Array,
                range,
                formula: if let Some(formula) = &cell.formula_binary {
                    formula.clone()
                } else {
                    crate::package::formula::text::Compiler::compile(formula)?
                },
                always_calculate: cached_value.is_none(),
            });
        }

        for (index, group) in groups.iter().enumerate() {
            if group.formula.exp_cell()?.is_some() {
                return Err(Error::InvalidFormula(
                    "array/shared formula definition cannot contain PtgExp".to_string(),
                ));
            }
            if groups[..index]
                .iter()
                .any(|existing| existing.range.top_left() == group.range.top_left())
            {
                return Err(Error::InvalidFormula(format!(
                    "multiple formula definitions cannot share anchor {}",
                    crate::package::utils::cell_reference(
                        group.range.row_first,
                        group.range.col_first
                    )
                )));
            }
        }
        Ok(groups)
    }

    pub(super) fn formula_group_for_cell(groups: &[Group], row: u32, col: u32) -> Option<&Group> {
        groups
            .iter()
            .find(|group| group.range.top_left() == (row, col))
            .or_else(|| {
                groups
                    .iter()
                    .filter(|group| group.range.contains(row, col))
                    .min_by_key(|group| {
                        u64::from(group.range.row_last - group.range.row_first + 1)
                            * u64::from(group.range.col_last - group.range.col_first + 1)
                    })
            })
    }

    pub(super) fn write_grouped_formula_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        row: u32,
        col: u32,
        cell_data: &CellData,
        group: &Group,
    ) -> Result<()> {
        let cached_value = match &cell_data.value {
            CellValue::Formula { cached_value, .. } => cached_value.as_deref(),
            CellValue::Empty => None,
            value => Some(value),
        };
        let placeholder = ParsedFormula::exp(group.range.row_first, group.range.col_first)?;
        self.write_formula_cell(
            writer,
            col,
            cell_data.style,
            "",
            cached_value,
            false,
            Some(&placeholder),
            cell_data.formula_flags,
        )?;

        if group.range.top_left() == (row, col) {
            let record_type = match group.kind {
                GroupKind::Array => kind::ARR_FMLA,
                GroupKind::Shared => kind::SHR_FMLA,
            };
            writer.write_record(record_type, &group.to_record_data()?)?;
        }
        Ok(())
    }

    /// Write row header record with BrtColSpan elements
    ///
    /// BrtRowHdr structure (2.4.761):
    /// - rw (4 bytes): Row index
    /// - ixfe (4 bytes): Style index
    /// - miyRw (2 bytes): Row height in twips (1/20 of a point)
    /// - flags1 (1 byte): fExtraAsc | fExtraDsc | reserved
    /// - flags2 (1 byte): outline/visibility flags
    /// - phonetic (1 byte): phonetic guide flags
    /// - ccolspan (4 bytes): number of BrtColSpan elements
    /// - rgBrtColspan (variable): array of BrtColSpan, each 8 bytes
    ///   (colFirst (u32) + colLast (u32))
    pub(super) fn write_row_header<W: Write>(
        &self,
        writer: &mut Writer<W>,
        row: u32,
        cells: &BTreeMap<(u32, u32), CellData>,
    ) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Fixed part
        temp_writer.write_u32(row)?; // rw: Row index
        temp_writer.write_u32(0)?; // ixfe: Style index (0 = default)

        // Row height in twips (1/20 of a point). When no explicit height is
        // configured, use Excel's default of 15 points (300 twips).
        let miy_rw: u16 = if let Some(info) = self.rows.get(&row) {
            if let Some(height_pts) = info.height {
                (height_pts * 20.0).round() as u16
            } else {
                0x012C
            }
        } else {
            0x012C
        };
        temp_writer.write_u16(miy_rw)?;

        // flags1: extra ascender/descender padding (unused here).
        temp_writer.write_u8(0)?;

        // flags2: outline / visibility / custom height flags.
        // Bits 0-2: outline level, 0x10: hidden, 0x20: custom height.
        let mut flags2: u8 = 0;
        if let Some(info) = self.rows.get(&row) {
            if info.hidden {
                flags2 |= 0x10;
            }
            if info.height.is_some() {
                flags2 |= 0x20;
            }
        }
        temp_writer.write_u8(flags2)?;

        // phonetic guide: 0 = no phonetic information
        temp_writer.write_u8(0)?;

        // Collect all columns that have cells in this row (BTreeMap preserves sorted order)
        let cells_in_row: Vec<u32> = cells
            .keys()
            .filter(|(r, _)| *r == row)
            .map(|(_, c)| *c)
            .collect();

        if cells_in_row.is_empty() {
            // No cells in row - write 0 colspans
            temp_writer.write_u32(0)?;
        } else {
            // Group columns by 1024-wide segments, as in [MS-XLSB] BrtColSpan and SheetJS
            let mut spans: Vec<(u32, u32)> = Vec::new();
            let mut current_segment = cells_in_row[0] / 1024;
            let mut segment_first = cells_in_row[0];
            let mut segment_last = cells_in_row[0];

            for &col in &cells_in_row[1..] {
                let segment = col / 1024;
                if segment == current_segment {
                    segment_last = col;
                } else {
                    spans.push((segment_first, segment_last));
                    current_segment = segment;
                    segment_first = col;
                    segment_last = col;
                }
            }
            spans.push((segment_first, segment_last));

            // Number of spans
            temp_writer.write_u32(spans.len() as u32)?;

            // Each span is a BrtColSpan: colFirst (u32) + colLast (u32)
            for (first, last) in spans {
                temp_writer.write_u32(first)?; // colFirst
                temp_writer.write_u32(last)?; // colLast
            }
        }

        writer.write_record(kind::ROW_HDR, &data)?;
        Ok(())
    }

    /// Write a single cell record
    pub(super) fn write_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        _row: u32,
        col: u32,
        cell_data: &CellData,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
    ) -> Result<()> {
        match &cell_data.value {
            CellValue::Empty => self.write_blank_cell(writer, col, cell_data.style)?,
            CellValue::String(s) => {
                self.write_shared_string_cell(writer, col, s, cell_data.style, shared_strings)?
            },
            CellValue::Int(i) => self.write_number_cell(writer, col, *i as f64, cell_data.style)?,
            CellValue::Float(f) => self.write_number_cell(writer, col, *f, cell_data.style)?,
            CellValue::Bool(b) => self.write_bool_cell(writer, col, *b, cell_data.style)?,
            CellValue::Error(e) => self.write_error_cell(writer, col, e, cell_data.style)?,
            CellValue::DateTime(dt) => {
                // Excel DateTime is already stored as serial number (days since epoch)
                // CellValue::DateTime stores the Excel serial number directly
                self.write_number_cell(writer, col, *dt, cell_data.style)?;
            },
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                ..
            } => self.write_formula_cell(
                writer,
                col,
                cell_data.style,
                formula,
                cached_value.as_deref(),
                *is_array,
                cell_data.formula_binary.as_ref(),
                cell_data.formula_flags,
            )?,
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn write_formula_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        style: u32,
        formula_text: &str,
        cached_value: Option<&CellValue>,
        is_array: bool,
        encoded: Option<&ParsedFormula>,
        flags: u16,
    ) -> Result<()> {
        if is_array {
            return Err(Error::UnsupportedFeature(
                "XLSB array formula writing requires BrtArrFmla".to_string(),
            ));
        }
        if flags & !0x0002 != 0 {
            return Err(Error::InvalidFormula(format!(
                "invalid GrbitFmla flags 0x{flags:04X}"
            )));
        }
        let effective_flags = if cached_value.is_none() {
            flags | 0x0002
        } else {
            flags
        };

        let compiled;
        let parsed = if let Some(encoded) = encoded {
            encoded
        } else {
            compiled = crate::package::formula::text::Compiler::compile(formula_text)?;
            &compiled
        };
        let formula_bytes = parsed.to_bytes()?;
        let cached = cached_value.unwrap_or(&CellValue::Float(0.0));

        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);
        Self::write_cell_structure(&mut temp_writer, col, style)?;

        let record_type = match cached {
            CellValue::String(value) => {
                temp_writer.write_wide_string(value)?;
                kind::FMLA_STRING
            },
            CellValue::Bool(value) => {
                temp_writer.write_u8(u8::from(*value))?;
                kind::FMLA_BOOL
            },
            CellValue::Error(error) => {
                temp_writer.write_u8(Self::error_code(error))?;
                kind::FMLA_ERROR
            },
            CellValue::Int(value) => {
                temp_writer.write_f64(*value as f64)?;
                kind::FMLA_NUM
            },
            CellValue::Float(value) | CellValue::DateTime(value) => {
                temp_writer.write_f64(*value)?;
                kind::FMLA_NUM
            },
            CellValue::Empty => {
                temp_writer.write_f64(0.0)?;
                kind::FMLA_NUM
            },
            CellValue::Formula { .. } => {
                return Err(Error::InvalidFormula(
                    "formula cached value cannot itself be a formula".to_string(),
                ));
            },
        };
        temp_writer.write_u16(effective_flags)?;
        data.extend_from_slice(&formula_bytes);
        Ok(writer.write_record(record_type, &data)?)
    }

    /// Write the Cell structure (2.5.10) - 8 bytes
    ///
    /// Cell structure:
    /// - column (4 bytes): Column index
    /// - iStyleRef (3 bytes, 24-bit): Style XF index
    /// - fPhShow (1 bit): Phonetic info flag
    /// - reserved (7 bits): Reserved
    pub(super) fn write_cell_structure<W: Write>(
        temp_writer: &mut Writer<W>,
        col: u32,
        style: u32,
    ) -> Result<()> {
        // Column (4 bytes)
        temp_writer.write_u32(col)?;

        // iStyleRef (3 bytes) + flags (1 byte) = 4 bytes total
        temp_writer.write_u8((style & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 8) & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 16) & 0xFF) as u8)?;
        temp_writer.write_u8(0)?; // fPhShow=0, reserved=0

        Ok(())
    }

    /// Write a blank cell (BrtCellBlank - 8 bytes)
    pub(super) fn write_blank_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        style: u32,
    ) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        Self::write_cell_structure(&mut temp_writer, col, style)?;

        writer.write_record(kind::CELL_BLANK, &data)?;
        Ok(())
    }

    /// Write a shared string cell (BrtCellIsst - Cell + u32 = 12 bytes)
    pub(super) fn write_shared_string_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        value: &str,
        style: u32,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
    ) -> Result<()> {
        // Add string to shared strings table and get index
        let string_index = shared_strings.add_string(value.to_string());

        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Cell structure (8 bytes) + isst index (4 bytes) = 12 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u32(string_index)?;

        writer.write_record(kind::CELL_ISST, &data)?;
        Ok(())
    }

    /// Write a number cell (BrtCellReal - Cell + f64 = 16 bytes)
    pub(super) fn write_number_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        value: f64,
        style: u32,
    ) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Cell structure (8 bytes) + Xnum value (8 bytes) = 16 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_f64(value)?;

        writer.write_record(kind::CELL_REAL, &data)?;
        Ok(())
    }

    /// Write a boolean cell (BrtCellBool - Cell + u8 = 9 bytes)
    pub(super) fn write_bool_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        value: bool,
        style: u32,
    ) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Cell structure (8 bytes) + fBool (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(if value { 1 } else { 0 })?;

        writer.write_record(kind::CELL_BOOL, &data)?;
        Ok(())
    }

    /// Write an error cell (BrtCellError - Cell + u8 = 9 bytes)
    pub(super) fn write_error_cell<W: Write>(
        &self,
        writer: &mut Writer<W>,
        col: u32,
        error: &str,
        style: u32,
    ) -> Result<()> {
        let error_code = Self::error_code(error);

        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Cell structure (8 bytes) + bError (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(error_code)?;

        writer.write_record(kind::CELL_ERROR, &data)?;
        Ok(())
    }

    pub(super) fn error_code(error: &str) -> u8 {
        match error {
            "#NULL!" => 0x00,
            "#DIV/0!" => 0x07,
            "#VALUE!" => 0x0F,
            "#REF!" => 0x17,
            "#NAME?" => 0x1D,
            "#NUM!" => 0x24,
            "#N/A" => 0x2A,
            "#GETTING_DATA" => 0x2B,
            _ => 0x2A, // Default to #N/A
        }
    }

    /// Write merged cells
    pub(super) fn write_merged_cells<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // BrtBeginMergeCells (0x00B1) payload is a single DWORD count of BrtMergeCell
        // records that follow. SheetJS writes this as write_BrtBeginMergeCells(cnt).
        let mut header = Vec::new();
        let mut temp_writer = Writer::new(&mut header);
        temp_writer.write_u32(self.merged_cells.len() as u32)?;

        writer.write_record(kind::BEGIN_MERGE_CELLS, &header)?;

        for merged in &self.merged_cells {
            let data = merged.serialize();
            writer.write_record(kind::MERGE_CELL, &data)?;
        }

        writer.write_record(kind::END_MERGE_CELLS, &[])?;
        Ok(())
    }

    /// Write column information records.
    pub(super) fn write_col_infos<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.columns.is_empty() {
            return Ok(());
        }

        writer.write_record(kind::BEGIN_COL_INFOS, &[])?;

        for (col, info) in &self.columns {
            let mut data = Vec::new();
            let mut temp_writer = Writer::new(&mut data);

            // firstCol / lastCol (both 0-based inclusive).
            temp_writer.write_u32(*col)?;
            temp_writer.write_u32(*col)?;

            // Width is stored as 256ths of a character, mirroring SheetJS
            // write_BrtColInfo and [MS-XLSB] 2.4.323.
            let width_chars = info.width.unwrap_or(10.0);
            let width_raw = (width_chars * 256.0).round() as u32;
            temp_writer.write_u32(width_raw)?;

            // Style XF index (we currently do not support per-column styles).
            temp_writer.write_u32(0)?;

            // Flags (2 bytes): 0x0001 = hidden, 0x0002 = custom width,
            // 0x0004 = best fit.
            let mut flags: u16 = 0;
            if info.hidden {
                flags |= 0x0001;
            }
            if info.width.is_some() {
                flags |= 0x0002;
            }
            if info.best_fit {
                flags |= 0x0004;
            }
            temp_writer.write_u16(flags)?;

            writer.write_record(kind::COL_INFO, &data)?;
        }

        writer.write_record(kind::END_COL_INFOS, &[])?;
        Ok(())
    }

    /// Write hyperlinks
    pub(super) fn write_hyperlinks<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        for hyperlink in &self.hyperlinks {
            let data = hyperlink.try_serialize()?;
            writer.write_record(kind::H_LINK, &data)?;
        }
        Ok(())
    }

    /// Write sheet protection if configured.
    pub(super) fn write_sheet_protection<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let Some(ref prot) = self.sheet_protection else {
            return Ok(());
        };

        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Password hash (Method 1). When absent, write 0.
        temp_writer.write_u16(prot.password_hash.unwrap_or(0))?;

        // Guard DWORD: this record should not be written if no protection.
        temp_writer.write_u32(1)?;

        fn flag(default_true: bool, value: Option<bool>) -> u32 {
            if default_true {
                if let Some(v) = value {
                    if !v { 1 } else { 0 }
                } else {
                    0
                }
            } else if let Some(v) = value {
                if v { 0 } else { 1 }
            } else {
                1
            }
        }

        temp_writer.write_u32(flag(false, prot.objects))?;
        temp_writer.write_u32(flag(false, prot.scenarios))?;
        temp_writer.write_u32(flag(true, prot.format_cells))?;
        temp_writer.write_u32(flag(true, prot.format_columns))?;
        temp_writer.write_u32(flag(true, prot.format_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_columns))?;
        temp_writer.write_u32(flag(true, prot.insert_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_hyperlinks))?;
        temp_writer.write_u32(flag(true, prot.delete_columns))?;
        temp_writer.write_u32(flag(true, prot.delete_rows))?;
        temp_writer.write_u32(flag(false, prot.select_locked_cells))?;
        temp_writer.write_u32(flag(true, prot.sort))?;
        temp_writer.write_u32(flag(true, prot.auto_filter))?;
        temp_writer.write_u32(flag(true, prot.pivot_tables))?;
        temp_writer.write_u32(flag(false, prot.select_unlocked_cells))?;

        writer.write_record(kind::SHEET_PROTECTION, &data)?;
        Ok(())
    }

    /// Write basic auto-filter range if configured.
    pub(super) fn write_auto_filter<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let Some(ref af) = self.auto_filter else {
            return Ok(());
        };
        if af.row_first > af.row_last
            || af.row_last >= 0x10_0000
            || af.col_first > af.col_last
            || af.col_last >= 0x4000
        {
            return Err(Error::Encoding(format!(
                "invalid AutoFilter range: rows {}..={}, columns {}..={}",
                af.row_first, af.row_last, af.col_first, af.col_last
            )));
        }

        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // UncheckedRfX: row_first, row_last, col_first, col_last
        temp_writer.write_u32(af.row_first)?;
        temp_writer.write_u32(af.row_last)?;
        temp_writer.write_u32(af.col_first)?;
        temp_writer.write_u32(af.col_last)?;

        writer.write_record(kind::BEGIN_A_FILTER, &data)?;
        writer.write_record(kind::END_A_FILTER, &[])?;
        Ok(())
    }
}
