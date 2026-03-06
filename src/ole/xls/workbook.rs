//! Workbook implementation for XLS files

use crate::ole::file::OleFile;
use crate::ole::xls::cell::XlsCell;
use crate::ole::xls::error::{XlsError, XlsResult};
use crate::ole::xls::pivot_table::PivotTable;
use crate::ole::xls::records::{
    BiffVersion, BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, RecordIter,
    SharedStringTable, XlsEncoding,
};
use crate::ole::xls::worksheet::XlsWorksheet;
use crate::ole::xls::{autofilter, comments, hyperlinks, merged_cells, pivot_table, protection};
use crate::sheet::{Result, Worksheet as SheetTrait, WorksheetIterator};
use std::io::{Read, Seek};
use std::sync::Arc;

/// XLS workbook implementation
#[derive(Debug)]
pub struct XlsWorkbook<R: Read + Seek> {
    ole_file: OleFile<R>,
    worksheets: Vec<XlsWorksheet>,
    worksheet_names: Vec<String>,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    shared_strings: Option<Arc<Vec<String>>>,
    biff_version: BiffVersion,
    is_1904_date_system: bool,
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
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
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
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
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

        // Parse workbook globals
        self.parse_workbook_globals(
            &mut record_iter,
            &mut encoding,
            &mut bound_sheets,
            &mut strings,
        )?;

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.worksheet_names = bound_sheets.iter().map(|s| s.name.clone()).collect();

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
                0x0042 => {
                    // CodePage
                    if record.data.len() >= 2 {
                        let codepage = crate::common::binary::read_u16_le_at(&record.data, 0)?;
                        *encoding = XlsEncoding::from_codepage(codepage)?;
                    }
                },
                0x0022 => {
                    // Date1904
                    if record.data.len() >= 2 {
                        let flag = crate::common::binary::read_u16_le_at(&record.data, 0)?;
                        self.is_1904_date_system = flag == 1;
                    }
                },
                0x0085 => {
                    // BoundSheet8
                    let sheet = BoundSheetRecord::parse(&record.data, encoding)?;
                    bound_sheets.push(sheet);
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
                    strings.extend(sst.strings);

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
        Self::parse_worksheet_records(record_iter, encoding, &bound_sheet.name, shared_strings)
    }

    /// Parse worksheet records sequentially
    fn parse_worksheet_records<Reader: Read + Seek>(
        record_iter: &mut RecordIter<Reader>,
        encoding: &XlsEncoding,
        name: &str,
        shared_strings: Arc<Vec<String>>,
    ) -> XlsResult<XlsWorksheet> {
        let mut worksheet = XlsWorksheet::with_shared_strings(name.to_string(), shared_strings);

        // Accumulator for pivot table records: we collect SX* records in order
        // and assemble complete PivotTable structs when SXVIEW boundaries are hit.
        let mut current_pivot: Option<PivotTable> = None;

        // Collector for TXO comment text: tracks OBJ→TXO→CONTINUE sequences.
        let mut txo_collector = comments::TxoCollector::new();

        for record_result in record_iter.by_ref() {
            let record = record_result?;

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
                0x00BD | // MulRk
                0x0006   // Formula
                => {
                    let cell_record = CellRecord::parse(record.header.record_type, &record.data, encoding)?;
                    if let Some(cell) = XlsCell::from_record(&cell_record, worksheet.shared_strings()) {
                        worksheet.add_cell(cell);
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
}

impl<R: Read + Seek + std::fmt::Debug + Send + Sync> crate::sheet::WorkbookTrait
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
