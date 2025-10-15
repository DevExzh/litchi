//! Workbook implementation for XLS files

use std::io::{Read, Seek};
use crate::sheet::{Workbook, Worksheet as SheetTrait, WorksheetIterator};
use crate::ole::xls::error::{XlsError, XlsResult};
use crate::ole::xls::records::{RecordIter, BofRecord, BoundSheetRecord, SharedStringTable, XlsEncoding, BiffVersion, CellRecord, DimensionsRecord};
use crate::ole::xls::worksheet::XlsWorksheet;
use crate::ole::xls::cell::XlsCell;
use crate::ole::file::OleFile;

/// XLS workbook implementation
#[derive(Debug)]
pub struct XlsWorkbook<R: Read + Seek> {
    ole_file: OleFile<R>,
    worksheets: Vec<XlsWorksheet>,
    worksheet_names: Vec<String>,
    shared_strings: Option<Vec<String>>,
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

    /// Parse the workbook stream
    fn parse_workbook(&mut self) -> XlsResult<()> {
        // Find and read the Workbook stream
        let workbook_data = self.ole_file.open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;

        let mut record_iter = RecordIter::new(std::io::Cursor::new(&workbook_data))?;
        let mut encoding = XlsEncoding::from_codepage(1252)?; // Default codepage
        let mut bound_sheets = Vec::new();
        let mut strings = Vec::new();

        // Parse workbook globals
        self.parse_workbook_globals(&mut record_iter, &mut encoding, &mut bound_sheets, &mut strings)?;

        self.shared_strings = Some(strings);
        self.worksheet_names = bound_sheets.iter().map(|s| s.name.clone()).collect();

        // Parse worksheets from positions in the workbook stream
        println!("DEBUG: Found {} bound sheets", bound_sheets.len());
        for bound_sheet in &bound_sheets {
            println!("DEBUG: Parsing worksheet '{}' at position {}", bound_sheet.name, bound_sheet.position);
            match self.parse_worksheet_from_position(bound_sheet, &encoding, &mut record_iter) {
                Ok(worksheet) => {
                    println!("DEBUG: Successfully parsed worksheet");
                    self.worksheets.push(worksheet);
                }
                Err(e) => {
                    println!("DEBUG: Failed to parse worksheet {}: {}", bound_sheet.name, e);
                }
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
                0x0809 => { // BOF
                    let bof = BofRecord::parse(&record.data)?;
                    self.biff_version = bof.version;
                    self.is_1904_date_system = bof.is_1904_date_system;
                }
                0x0042 => { // CodePage
                    if record.data.len() >= 2 {
                        let codepage = crate::ole::binary::read_u16_le_at(&record.data, 0)?;
                        *encoding = XlsEncoding::from_codepage(codepage)?;
                    }
                }
                0x0022 => { // Date1904
                    if record.data.len() >= 2 {
                        let flag = crate::ole::binary::read_u16_le_at(&record.data, 0)?;
                        self.is_1904_date_system = flag == 1;
                    }
                }
                0x0085 => { // BoundSheet8
                    let sheet = BoundSheetRecord::parse(&record.data, encoding)?;
                    bound_sheets.push(sheet);
                }
                0x00FC => { // SST
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
                }
                0x000A => { // EOF - End of workbook globals
                    break;
                }
                _ => {
                    // Skip other records for now
                }
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
            if record.header.record_type != 0x0809 { // BOF
                return Err(XlsError::UnexpectedRecordType {
                    expected: 0x0809,
                    found: record.header.record_type,
                });
            }
        } else {
            return Err(XlsError::Eof("Expected BOF record for worksheet"));
        }

        // Parse worksheet records
        let shared_strings = self.shared_strings.as_ref().unwrap_or(&Vec::new()).clone();
        Self::parse_worksheet_records(record_iter, encoding, &bound_sheet.name, shared_strings)
    }

    /// Parse worksheet records sequentially
    fn parse_worksheet_records<Reader: Read + Seek>(
        record_iter: &mut RecordIter<Reader>,
        encoding: &XlsEncoding,
        name: &str,
        shared_strings: Vec<String>
    ) -> XlsResult<XlsWorksheet> {
        let mut worksheet = XlsWorksheet::with_shared_strings(name.to_string(), shared_strings);

        for record_result in record_iter.by_ref() {
            let record = record_result?;

            match record.header.record_type {
                0x0809 => { // BOF - Beginning of worksheet
                    // This marks the start of a worksheet
                }
                0x000A => { // EOF - End of worksheet
                    break;
                }
                0x0200 => { // Dimensions
                    // Parse dimensions to understand worksheet bounds
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
                _ => {
                    // Skip other records for now
                }
            }
        }

        Ok(worksheet)
    }
}

impl<R: Read + Seek> Workbook for XlsWorkbook<R> {
    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        if self.worksheets.is_empty() {
            return Err(Box::new(XlsError::WorksheetNotFound("No worksheets found".to_string())));
        }
        Ok(Box::new(self.worksheets[0].clone()))
    }

    fn worksheet_names(&self) -> Vec<String> {
        self.worksheet_names.clone()
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        for worksheet in &self.worksheets {
            if worksheet.name() == name {
                return Ok(Box::new(worksheet.clone()));
            }
        }
        Err(Box::new(XlsError::WorksheetNotFound(name.to_string())))
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        if index >= self.worksheets.len() {
            return Err(Box::new(XlsError::WorksheetNotFound(format!("Index {} out of bounds", index))));
        }
        Ok(Box::new(self.worksheets[index].clone()))
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
}

/// Worksheet iterator for XLS workbooks
struct XlsWorksheetIterator<'a> {
    worksheets: Vec<&'a XlsWorksheet>,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>, Box<dyn std::error::Error>>> {
        if self.index >= self.worksheets.len() {
            None
        } else {
            let worksheet = self.worksheets[self.index];
            self.index += 1;
            Some(Ok(Box::new(worksheet.clone())))
        }
    }
}
