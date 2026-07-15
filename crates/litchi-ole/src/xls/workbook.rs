//! Workbook implementation for XLS files

use crate::xls::cell::XlsCell;
use crate::xls::error::{XlsError, XlsResult};
use crate::xls::formula::{FormulaContext, ptg_exp_anchor, render_formula, render_shared_formula};
use crate::xls::pivot_table::PivotTable;
use crate::xls::records::{
    BiffVersion, BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, FormulaValue,
    RecordIter, SharedStringProperties, SharedStringTable, XlsEncoding,
};
use crate::xls::worksheet::XlsWorksheet;
use crate::xls::{autofilter, comments, hyperlinks, merged_cells, pivot_table, protection, utils};
use litchi_cfb::OleFile;
use litchi_core::sheet::{Result, Worksheet as SheetTrait, WorksheetIterator};
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

#[derive(Debug)]
struct SharedFormulaTemplate {
    first_row: u16,
    last_row: u16,
    first_col: u16,
    last_col: u16,
    tokens: Vec<u8>,
    relative: bool,
}

impl SharedFormulaTemplate {
    fn contains(&self, row: u16, col: u16) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_col..=self.last_col).contains(&col)
    }

    fn render(&self, context: Option<&FormulaContext>, row: u16, col: u16) -> Option<String> {
        if !self.contains(row, col) {
            return None;
        }
        if self.relative {
            render_shared_formula(&self.tokens, context, row, col)
        } else {
            render_formula(&self.tokens, context)
        }
    }
}

fn parse_shared_formula_template(
    record_type: u16,
    data: &[u8],
) -> XlsResult<SharedFormulaTemplate> {
    let (fixed_size, length_offset, relative) = match record_type {
        0x04bc => (10usize, 8usize, true),
        0x0221 => (14usize, 12usize, false),
        _ => {
            return Err(XlsError::UnexpectedRecordType {
                expected: 0x04bc,
                found: record_type,
            });
        },
    };
    if data.len() < fixed_size {
        return Err(XlsError::InvalidLength {
            expected: fixed_size,
            found: data.len(),
        });
    }
    let first_row = u16::from_le_bytes([data[0], data[1]]);
    let last_row = u16::from_le_bytes([data[2], data[3]]);
    let first_col = u16::from(data[4]);
    let last_col = u16::from(data[5]);
    if first_row > last_row || first_col > last_col {
        return Err(XlsError::InvalidRecord {
            record_type,
            message: "shared formula range is reversed".to_string(),
        });
    }
    let token_len = usize::from(u16::from_le_bytes([
        data[length_offset],
        data[length_offset + 1],
    ]));
    let end = fixed_size
        .checked_add(token_len)
        .ok_or_else(|| XlsError::InvalidRecord {
            record_type,
            message: "shared formula token length overflows".to_string(),
        })?;
    let tokens = data
        .get(fixed_size..end)
        .ok_or(XlsError::InvalidLength {
            expected: end,
            found: data.len(),
        })?
        .to_vec();
    if tokens.is_empty() {
        return Err(XlsError::InvalidRecord {
            record_type,
            message: "shared formula token stream is empty".to_string(),
        });
    }
    Ok(SharedFormulaTemplate {
        first_row,
        last_row,
        first_col,
        last_col,
        tokens,
        relative,
    })
}

/// XLS workbook implementation
#[derive(Debug)]
pub struct XlsWorkbook<R: Read + Seek> {
    ole_file: OleFile<R>,
    worksheets: Vec<XlsWorksheet>,
    worksheet_names: Vec<String>,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    shared_strings: Option<Arc<Vec<String>>>,
    /// Sparse rich-text and phonetic properties parallel to `shared_strings`.
    shared_string_properties: Option<Arc<Vec<Option<Box<SharedStringProperties>>>>>,
    shared_string_reference_count: u32,
    biff_version: BiffVersion,
    is_1904_date_system: bool,
    formula_context: FormulaContext,
}

impl<R: Read + Seek> XlsWorkbook<R> {
    /// Open an XLS workbook from a reader
    pub fn new(reader: R) -> XlsResult<Self> {
        let ole_file = OleFile::open(reader)?;

        let mut workbook = XlsWorkbook {
            ole_file,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            shared_strings: None,
            shared_string_properties: None,
            shared_string_reference_count: 0,
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
        };

        workbook.parse_workbook()?;
        Ok(workbook)
    }

    /// Create an XLS workbook from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - An already-parsed OLE file
    pub fn from_ole_file(ole_file: OleFile<R>) -> XlsResult<Self> {
        let mut workbook = XlsWorkbook {
            ole_file,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            shared_strings: None,
            shared_string_properties: None,
            shared_string_reference_count: 0,
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
        };

        workbook.parse_workbook()?;
        Ok(workbook)
    }

    /// Parse the workbook stream
    fn parse_workbook(&mut self) -> XlsResult<()> {
        // Find and read the Workbook stream
        let workbook_data = self
            .ole_file
            .open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;

        let mut record_iter = RecordIter::new(std::io::Cursor::new(&workbook_data))?;
        let mut encoding = XlsEncoding::from_codepage(1252)?; // Default codepage
        let mut bound_sheets = Vec::new();
        let mut strings = Vec::new();
        let mut string_properties = Vec::new();

        // Parse workbook globals
        self.parse_workbook_globals(
            &mut record_iter,
            &mut encoding,
            &mut bound_sheets,
            &mut strings,
            &mut string_properties,
        )?;

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.shared_string_properties = Some(Arc::new(string_properties));
        self.worksheet_names = bound_sheets.iter().map(|s| s.name.clone()).collect();
        self.formula_context
            .set_sheet_names(self.worksheet_names.clone());

        // Parse worksheets from positions in the workbook stream
        for bound_sheet in &bound_sheets {
            match self.parse_worksheet_from_position(bound_sheet, &encoding, &mut record_iter) {
                Ok(worksheet) => {
                    self.worksheets.push(worksheet);
                },
                Err(_e) => {
                    // Failed to parse worksheet, continue with next
                },
            }
        }

        Ok(())
    }

    /// Parse workbook globals (SST, bound sheets, etc.)
    fn parse_workbook_globals<Reader: Read + Seek>(
        &mut self,
        record_iter: &mut RecordIter<Reader>,
        encoding: &mut XlsEncoding,
        bound_sheets: &mut Vec<BoundSheetRecord>,
        strings: &mut Vec<String>,
        string_properties: &mut Vec<Option<Box<SharedStringProperties>>>,
    ) -> XlsResult<()> {
        // Collect all records first for easier processing
        let mut records = Vec::new();
        for record_result in record_iter.by_ref() {
            records.push(record_result?);
        }

        let mut i = 0;
        while i < records.len() {
            let record = &records[i];

            match record.header.record_type {
                0x0809 => {
                    // BOF
                    let bof = BofRecord::parse(&record.data)?;
                    self.biff_version = bof.version;
                    self.is_1904_date_system = bof.is_1904_date_system;
                },
                0x0042
                    // CodePage
                    if record.data.len() >= 2 => {
                        let codepage = litchi_core::binary::read_u16_le_at(&record.data, 0)?;
                        *encoding = XlsEncoding::from_codepage(codepage)?;
                    },
                0x0022
                    // Date1904
                    if record.data.len() >= 2 => {
                        let flag = litchi_core::binary::read_u16_le_at(&record.data, 0)?;
                        self.is_1904_date_system = flag == 1;
                    },
                0x0085 => {
                    // BoundSheet8
                    let sheet = BoundSheetRecord::parse(&record.data, encoding)?;
                    bound_sheets.push(sheet);
                },
                0x01AE => {
                    // SUPBOOK: retain its position and whether it references
                    // this workbook so EXTERNSHEET indices remain stable.
                    self.formula_context.add_sup_book(&record.data);
                },
                0x0017 => {
                    // EXTERNSHEET: XTI entries used by PtgRef3d/PtgArea3d.
                    self.formula_context
                        .add_extern_sheet(&record.data)
                        .map_err(|message| XlsError::InvalidRecord {
                            record_type: 0x0017,
                            message: message.to_string(),
                        })?;
                },
                0x00FC => {
                    // SST
                    // SST may span multiple records, collect them all
                    let mut sst_records = vec![record.clone()];
                    let mut sst_idx = i + 1;

                    // Collect all following CONTINUE records
                    while sst_idx < records.len() && records[sst_idx].header.record_type == 0x003C {
                        sst_records.push(records[sst_idx].clone());
                        sst_idx += 1;
                    }

                    let sst = SharedStringTable::parse_from_records(&sst_records, encoding)?;
                    self.shared_string_reference_count = sst.total_count;
                    strings.extend(sst.strings);
                    string_properties.extend(sst.properties);

                    // Skip the CONTINUE records we consumed
                    i = sst_idx - 1;
                },
                0x000A => {
                    // EOF - End of workbook globals
                    break;
                },
                _ => {
                    // Skip other records for now
                },
            }
            i += 1;
        }

        Ok(())
    }

    /// Parse a worksheet from its position in the workbook stream
    fn parse_worksheet_from_position<Reader: Read + Seek>(
        &self,
        bound_sheet: &BoundSheetRecord,
        encoding: &XlsEncoding,
        record_iter: &mut RecordIter<Reader>,
    ) -> XlsResult<XlsWorksheet> {
        // Seek to the worksheet position
        record_iter.seek(bound_sheet.position as u64)?;

        // Skip the BOF record at the beginning of the worksheet
        if let Some(record_result) = record_iter.next() {
            let record = record_result?;
            if record.header.record_type != 0x0809 {
                // BOF
                return Err(XlsError::UnexpectedRecordType {
                    expected: 0x0809,
                    found: record.header.record_type,
                });
            }
        } else {
            return Err(XlsError::Eof("Expected BOF record for worksheet"));
        }

        // Parse worksheet records (clone Arc is cheap - just increments ref count)
        let shared_strings = self
            .shared_strings
            .clone()
            .unwrap_or_else(|| Arc::new(Vec::new()));
        Self::parse_worksheet_records(
            record_iter,
            encoding,
            &bound_sheet.name,
            shared_strings,
            Some(&self.formula_context),
        )
    }

    /// Parse worksheet records sequentially
    fn parse_worksheet_records<Reader: Read + Seek>(
        record_iter: &mut RecordIter<Reader>,
        encoding: &XlsEncoding,
        name: &str,
        shared_strings: Arc<Vec<String>>,
        formula_context: Option<&FormulaContext>,
    ) -> XlsResult<XlsWorksheet> {
        let mut worksheet = XlsWorksheet::with_shared_strings(name.to_string(), shared_strings);

        // Accumulator for pivot table records: we collect SX* records in order
        // and assemble complete PivotTable structs when SXVIEW boundaries are hit.
        let mut current_pivot: Option<PivotTable> = None;

        // Collector for TXO comment text: tracks OBJ→TXO→CONTINUE sequences.
        let mut txo_collector = comments::TxoCollector::new();
        let mut pending_string_formula: Option<CellRecord> = None;
        let mut shared_formulas = HashMap::<(u16, u16), SharedFormulaTemplate>::new();

        for record_result in record_iter.by_ref() {
            let record = record_result?;

            if let Some(mut formula) = pending_string_formula.take() {
                if record.header.record_type != 0x0207 {
                    return Err(XlsError::InvalidRecord {
                        record_type: record.header.record_type,
                        message: "String-valued Formula must be followed by a String record"
                            .to_string(),
                    });
                }
                let text = utils::parse_string_record(&record.data, encoding)?;
                if let CellRecord::Formula { value, .. } = &mut formula {
                    *value = FormulaValue::String(text);
                }
                if let Some(mut cell) = XlsCell::from_record_with_formula_context(
                    &formula,
                    worksheet.shared_strings(),
                    formula_context,
                ) {
                    if let CellRecord::Formula {
                        row, col, formula, ..
                    } = &formula
                    {
                        if let Some(anchor) = ptg_exp_anchor(formula) {
                            if let Some(rendered) = shared_formulas
                                .get(&anchor)
                                .and_then(|template| template.render(formula_context, *row, *col))
                            {
                                cell.set_rendered_formula(Some(rendered));
                            }
                        }
                    }
                    worksheet.add_cell(cell);
                }
                continue;
            }

            match record.header.record_type {
                0x0809 => { // BOF - Beginning of worksheet
                    // This marks the start of a worksheet
                }
                0x000A => { // EOF - End of worksheet
                    // Flush any in-progress pivot table
                    if let Some(pt) = current_pivot.take() {
                        worksheet.add_pivot_table(pt);
                    }
                    break;
                }
                0x0200 => { // Dimensions
                    if let Ok(dimensions) = DimensionsRecord::parse(&record.data) {
                        worksheet.set_dimensions(dimensions.first_row, dimensions.last_row,
                                               dimensions.first_col, dimensions.last_col);
                    }
                }
                // Cell records
                0x0201 | // Blank
                0x0203 | // Number
                0x0204 | // Label
                0x0205 | // BoolErr
                0x027E | // RK
                0x00FD | // LabelSst
                0x0006   // Formula
                => {
                    let cell_record = CellRecord::parse(record.header.record_type, &record.data, encoding)?;
                    if matches!(
                        &cell_record,
                        CellRecord::Formula {
                            value: FormulaValue::StringPending,
                            ..
                        }
                    ) {
                        pending_string_formula = Some(cell_record);
                    } else if let Some(mut cell) = XlsCell::from_record_with_formula_context(
                        &cell_record,
                        worksheet.shared_strings(),
                        formula_context,
                    ) {
                        if let CellRecord::Formula {
                            row, col, formula, ..
                        } = &cell_record
                        {
                            if let Some(anchor) = ptg_exp_anchor(formula) {
                                if let Some(rendered) = shared_formulas
                                    .get(&anchor)
                                    .and_then(|template| {
                                        template.render(formula_context, *row, *col)
                                    })
                                {
                                    cell.set_rendered_formula(Some(rendered));
                                }
                            }
                        }
                        worksheet.add_cell(cell);
                    }
                }

                0x04BC | 0x0221 => { // ShrFmla or Array
                    let template = parse_shared_formula_template(
                        record.header.record_type,
                        &record.data,
                    )?;
                    let anchor = (template.first_row, template.first_col);
                    let rendered = template.render(formula_context, anchor.0, anchor.1);
                    shared_formulas.insert(anchor, template);
                    if let Some(cell) = worksheet.get_cell_mut(
                        u32::from(anchor.0),
                        u32::from(anchor.1),
                    ) {
                        cell.set_rendered_formula(rendered);
                    }
                }

                0x00BD => { // MulRk
                    for cell_record in CellRecord::parse_mul_rk(&record.data)? {
                        if let Some(cell) = XlsCell::from_record_with_formula_context(
                            &cell_record,
                            worksheet.shared_strings(),
                            formula_context,
                        ) {
                            worksheet.add_cell(cell);
                        }
                    }
                }

                0x00BE => { // MulBlank
                    for cell_record in CellRecord::parse_mul_blank(&record.data)? {
                        if let Some(cell) = XlsCell::from_record_with_formula_context(
                            &cell_record,
                            worksheet.shared_strings(),
                            formula_context,
                        ) {
                            worksheet.add_cell(cell);
                        }
                    }
                }

                // --- Merged cells (MERGECELLS 0x00E5) ---
                rt if rt == merged_cells::RECORD_TYPE => {
                    let mut ranges = Vec::new();
                    if merged_cells::parse_mergecells_record(&record.data, &mut ranges).is_ok() {
                        worksheet.add_merged_cells(&ranges);
                    }
                }

                // --- Hyperlinks (HLINK 0x01B8) ---
                rt if rt == hyperlinks::RECORD_TYPE => {
                    if let Ok(link) = hyperlinks::parse_hlink_record(&record.data) {
                        worksheet.add_hyperlink(link);
                    }
                }

                // --- Comments (NOTE 0x001C) ---
                rt if rt == comments::RECORD_TYPE => {
                    if let Ok(comment) = comments::parse_note_record(&record.data) {
                        worksheet.add_comment(comment);
                    }
                }

                // --- OBJ record (0x005D) — extract object ID for TXO linking ---
                rt if rt == comments::OBJ_TYPE => {
                    txo_collector.feed_obj(&record.data);
                }

                // --- TXO record (0x01B6) — text object header ---
                rt if rt == comments::TXO_TYPE => {
                    txo_collector.feed_txo(&record.data);
                }

                // --- CONTINUE record (0x003C) — may carry TXO text data ---
                rt if rt == comments::CONTINUE_TYPE => {
                    txo_collector.feed_continue(&record.data);
                }

                // --- AutoFilter (AUTOFILTERINFO 0x009D) ---
                rt if rt == autofilter::AUTOFILTERINFO_TYPE => {
                    if let Ok(count) = autofilter::parse_autofilterinfo(&record.data) {
                        worksheet.set_autofilter_info(count);
                    }
                }

                // --- AutoFilter column (AUTOFILTER 0x009E) ---
                rt if rt == autofilter::AUTOFILTER_TYPE => {
                    if let Ok(col) = autofilter::parse_autofilter(&record.data) {
                        worksheet.add_autofilter_column(col);
                    }
                }

                // --- Sort (SORT 0x0090) ---
                rt if rt == autofilter::SORT_TYPE => {
                    if let Ok(info) = autofilter::parse_sort(&record.data) {
                        worksheet.set_sort_info(info);
                    }
                }

                // --- Sheet protection records ---
                rt if rt == protection::PROTECT_TYPE => {
                    if let Ok(val) = protection::parse_protect_bool(&record.data) {
                        worksheet.protection_mut().sheet_protected = val;
                    }
                }
                rt if rt == protection::OBJECTPROTECT_TYPE => {
                    if let Ok(val) = protection::parse_protect_bool(&record.data) {
                        worksheet.protection_mut().objects_protected = val;
                    }
                }
                rt if rt == protection::SCENPROTECT_TYPE => {
                    if let Ok(val) = protection::parse_protect_bool(&record.data) {
                        worksheet.protection_mut().scenarios_protected = val;
                    }
                }
                rt if rt == protection::PASSWORD_TYPE => {
                    if let Ok(hash) = protection::parse_password(&record.data) {
                        worksheet.protection_mut().password_hash = hash;
                    }
                }

                // --- Pivot table records ---
                rt if rt == pivot_table::SXVIEW_TYPE => {
                    // New SXVIEW starts a new pivot table; flush previous if any
                    if let Some(pt) = current_pivot.take() {
                        worksheet.add_pivot_table(pt);
                    }
                    if let Ok(view) = pivot_table::parse_sxview(&record.data) {
                        current_pivot = Some(PivotTable::new(view));
                    }
                }
                rt if rt == pivot_table::SXVD_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(field) = pivot_table::parse_sxvd(&record.data)
                    {
                        pt.fields.push(field);
                    }
                }
                rt if rt == pivot_table::SXVI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(item) = pivot_table::parse_sxvi(&record.data)
                    {
                        pt.items.push(item);
                    }
                }
                rt if rt == pivot_table::SXDI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(di) = pivot_table::parse_sxdi(&record.data)
                    {
                        pt.data_items.push(di);
                    }
                }
                rt if rt == pivot_table::SXVS_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(src) = pivot_table::parse_sxvs(&record.data)
                    {
                        pt.source_type = src;
                    }
                }
                rt if rt == pivot_table::SXPI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(entries) = pivot_table::parse_sxpi(&record.data)
                    {
                        pt.page_entries.extend(entries);
                    }
                }

                _ => {
                    // Skip other records
                }
            }
        }

        // Resolve comment texts from TXO data collected during parsing.
        txo_collector.resolve_comment_texts(worksheet.comments_mut());

        Ok(worksheet)
    }

    /// Access the typed `XlsWorksheet` at the given index.
    ///
    /// This provides access to XLS-specific data (protection, comments,
    /// autofilter, pivot tables) that is not exposed through the generic
    /// `WorkbookTrait` / `Worksheet` trait.
    pub fn xls_worksheet(&self, index: usize) -> XlsResult<&XlsWorksheet> {
        self.worksheets
            .get(index)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {}", index)))
    }

    /// Total number of cell references represented by the workbook SST.
    pub fn shared_string_reference_count(&self) -> u32 {
        self.shared_string_reference_count
    }

    /// Rich-text and phonetic properties for a shared-string index.
    ///
    /// Returns `None` for an out-of-range index and for an ordinary string
    /// without either optional BIFF8 payload.
    pub fn shared_string_properties(&self, index: u32) -> Option<&SharedStringProperties> {
        self.shared_string_properties
            .as_ref()?
            .get(index as usize)?
            .as_deref()
    }

    /// Rich-text and phonetic properties for a cell backed by `LabelSst`.
    pub fn shared_string_properties_for_cell(
        &self,
        cell: &XlsCell,
    ) -> Option<&SharedStringProperties> {
        self.shared_string_properties(cell.shared_string_index()?)
    }
}

impl<R: Read + Seek + std::fmt::Debug + Send + Sync> litchi_core::sheet::WorkbookTrait
    for XlsWorkbook<R>
{
    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>> {
        if self.worksheets.is_empty() {
            return Err(Box::new(XlsError::WorksheetNotFound(
                "No worksheets found".to_string(),
            )));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[0]))
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn SheetTrait + '_>> {
        for worksheet in &self.worksheets {
            if worksheet.name() == name {
                // Return reference instead of clone - zero-copy!
                return Ok(Box::new(worksheet));
            }
        }
        Err(Box::new(XlsError::WorksheetNotFound(name.to_string())))
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn SheetTrait + '_>> {
        if index >= self.worksheets.len() {
            return Err(Box::new(XlsError::WorksheetNotFound(format!(
                "Index {} out of bounds",
                index
            ))));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[index]))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(XlsWorksheetIterator {
            worksheets: self.worksheets.iter().collect(),
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        0 // Default to first sheet
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904_date_system
    }
}

/// Worksheet iterator for XLS workbooks
struct XlsWorksheetIterator<'a> {
    worksheets: Vec<&'a XlsWorksheet>,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>>> {
        if self.index >= self.worksheets.len() {
            None
        } else {
            let worksheet = self.worksheets[self.index];
            self.index += 1;
            // Return reference instead of clone - zero-copy!
            Some(Ok(Box::new(worksheet)))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::sheet::Cell;
    use std::io::Cursor;

    fn push_record(stream: &mut Vec<u8>, record_type: u16, data: &[u8]) {
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&(data.len() as u16).to_le_bytes());
        stream.extend_from_slice(data);
    }

    fn string_formula_data(row: u16, col: u16) -> Vec<u8> {
        let mut data = formula_data(row, col, &[]);
        data[6] = 0;
        data
    }

    fn formula_data(row: u16, col: u16, tokens: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&row.to_le_bytes());
        data.extend_from_slice(&col.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[3, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
        data.extend_from_slice(tokens);
        data
    }

    #[test]
    fn worksheet_expands_packed_numeric_and_blank_cells() {
        let mut mul_rk = Vec::new();
        mul_rk.extend_from_slice(&2u16.to_le_bytes());
        mul_rk.extend_from_slice(&4u16.to_le_bytes());
        mul_rk.extend_from_slice(&0u16.to_le_bytes());
        mul_rk.extend_from_slice(&((7u32 << 2) | 0x02).to_le_bytes());
        mul_rk.extend_from_slice(&1u16.to_le_bytes());
        mul_rk.extend_from_slice(&((250u32 << 2) | 0x03).to_le_bytes());
        mul_rk.extend_from_slice(&5u16.to_le_bytes());

        let mut mul_blank = Vec::new();
        mul_blank.extend_from_slice(&3u16.to_le_bytes());
        mul_blank.extend_from_slice(&1u16.to_le_bytes());
        mul_blank.extend_from_slice(&2u16.to_le_bytes());
        mul_blank.extend_from_slice(&3u16.to_le_bytes());
        mul_blank.extend_from_slice(&2u16.to_le_bytes());

        let mut stream = Vec::new();
        push_record(&mut stream, 0x00BD, &mul_rk);
        push_record(&mut stream, 0x00BE, &mul_blank);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            None,
        )
        .unwrap();

        assert!(matches!(
            worksheet.get_cell(2, 4).unwrap().value(),
            litchi_core::sheet::CellValue::Float(value) if *value == 7.0
        ));
        assert!(matches!(
            worksheet.get_cell(2, 5).unwrap().value(),
            litchi_core::sheet::CellValue::Float(value) if *value == 2.5
        ));
        assert!(worksheet.get_cell(3, 1).unwrap().is_empty());
        assert!(worksheet.get_cell(3, 2).unwrap().is_empty());
    }

    #[test]
    fn worksheet_resolves_formula_value_from_following_string_record() {
        let mut string_data = Vec::new();
        string_data.extend_from_slice(&3u16.to_le_bytes());
        string_data.push(1);
        for code_unit in "文😀".encode_utf16() {
            string_data.extend_from_slice(&code_unit.to_le_bytes());
        }

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(4, 5));
        push_record(&mut stream, 0x0207, &string_data);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            None,
        )
        .unwrap();
        let cell = worksheet.get_cell(4, 5).unwrap();

        assert!(cell.is_formula());
        assert!(matches!(
            cell.value(),
            litchi_core::sheet::CellValue::String(value) if value == "文😀"
        ));
    }

    #[test]
    fn worksheet_rejects_formula_missing_its_string_record() {
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(0, 0));
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let result = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            None,
        );

        assert!(result.is_err());
    }

    #[test]
    fn worksheet_expands_shared_formula_relative_references() {
        let anchor = [0x01, 0, 0, 1, 0];
        let template = [
            0x4c, 0, 0, 0xff, 0xc0, // same row, previous column
            0x1e, 2, 0, 0x05, // * 2
        ];
        let mut shared = Vec::new();
        shared.extend_from_slice(&0u16.to_le_bytes());
        shared.extend_from_slice(&1u16.to_le_bytes());
        shared.extend_from_slice(&[1, 1, 0, 2]); // columns, reserved, cUse
        shared.extend_from_slice(&(template.len() as u16).to_le_bytes());
        shared.extend_from_slice(&template);

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 1, &anchor));
        push_record(&mut stream, 0x04bc, &shared);
        push_record(&mut stream, 0x0006, &formula_data(1, 1, &anchor));
        push_record(&mut stream, 0x000a, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            None,
        )
        .unwrap();

        let first = worksheet.get_cell(0, 1).unwrap();
        let second = worksheet.get_cell(1, 1).unwrap();
        assert_eq!(first.formula(), Some("=(A1*2)"));
        assert_eq!(second.formula(), Some("=(A2*2)"));
        assert_eq!(first.formula_bytes(), Some(anchor.as_slice()));
        assert_eq!(second.formula_bytes(), Some(anchor.as_slice()));
    }

    #[test]
    fn worksheet_resolves_array_formula_for_every_cell() {
        let anchor = [0x01, 0, 0, 2, 0];
        let template = [0x1e, 7, 0];
        let mut array = Vec::new();
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&1u16.to_le_bytes());
        array.extend_from_slice(&[2, 2]);
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&0u32.to_le_bytes());
        array.extend_from_slice(&(template.len() as u16).to_le_bytes());
        array.extend_from_slice(&template);

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x0006, &formula_data(1, 2, &anchor));
        push_record(&mut stream, 0x000a, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            None,
        )
        .unwrap();

        assert_eq!(worksheet.get_cell(0, 2).unwrap().formula(), Some("=7"));
        assert_eq!(worksheet.get_cell(1, 2).unwrap().formula(), Some("=7"));
    }
}
