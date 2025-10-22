//! Workbook implementation for XLSB files

use crate::ooxml::opc::OpcPackage;
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::{XlsbRecordIter, record_types};
use crate::ooxml::xlsb::worksheet::XlsbWorksheet;
use crate::sheet::{Worksheet as SheetTrait, WorksheetIterator};
use std::io::{BufReader, Cursor, Read, Seek};

/// XLSB workbook implementation
#[allow(dead_code)]
pub struct XlsbWorkbook {
    package: OpcPackage,
    worksheets: Vec<XlsbWorksheet>,
    worksheet_names: Vec<String>,
    shared_strings: Vec<String>,
    is_1904: bool,
}

impl std::fmt::Debug for XlsbWorkbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("XlsbWorkbook")
            .field("worksheet_names", &self.worksheet_names)
            .field("shared_strings_count", &self.shared_strings.len())
            .field("is_1904", &self.is_1904)
            .finish()
    }
}

impl XlsbWorkbook {
    /// Open an XLSB workbook from a reader
    pub fn new<R: Read + Seek>(reader: R) -> XlsbResult<Self> {
        let package = OpcPackage::from_reader(reader)?;
        let mut workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            shared_strings: Vec::new(),
            is_1904: false,
        };

        workbook.load_workbook_info()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    /// Load workbook information from workbook.bin
    fn load_workbook_info(&mut self) -> XlsbResult<()> {
        let workbook_uri = crate::ooxml::opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;

        let blob = workbook_part.blob();
        let mut iter = XlsbRecordIter::new(BufReader::new(blob));
        Self::read_workbook(&mut iter, &mut self.worksheet_names, &mut self.is_1904)?;

        Ok(())
    }

    /// Load shared strings from xl/sharedStrings.bin
    fn load_shared_strings(&mut self) -> XlsbResult<()> {
        let shared_strings_uri = crate::ooxml::opc::PackURI::new("/xl/sharedStrings.bin")?;
        if let Ok(shared_strings_part) = self.package.get_part(&shared_strings_uri) {
            let blob = shared_strings_part.blob();
            let mut iter = XlsbRecordIter::new(BufReader::new(blob));
            Self::read_shared_strings(&mut iter, &mut self.shared_strings)?;
        }

        Ok(())
    }

    /// Get a worksheet by index (lazy loading)
    fn get_worksheet(&self, index: usize) -> XlsbResult<XlsbWorksheet> {
        if index >= self.worksheet_names.len() {
            return Err(crate::ooxml::error::OoxmlError::InvalidFormat(format!(
                "Worksheet index {} out of bounds",
                index
            ))
            .into());
        }

        let name = &self.worksheet_names[index];
        // For now, assume worksheets are at xl/worksheets/sheet1.bin, sheet2.bin, etc.
        let sheet_path = format!("/xl/worksheets/sheet{}.bin", index + 1);
        let sheet_uri = crate::ooxml::opc::PackURI::new(&sheet_path)?;

        let sheet_part = self.package.get_part(&sheet_uri)?;
        let blob = sheet_part.blob();
        let cursor = Cursor::new(blob);
        Self::read_worksheet(cursor, name.clone(), &self.shared_strings)
    }

    /// Read shared strings from SST
    fn read_shared_strings(
        iter: &mut XlsbRecordIter<impl Read>,
        strings: &mut Vec<String>,
    ) -> XlsbResult<()> {
        for record in iter.by_ref() {
            let record = record?;
            // println!("DEBUG SST: Record type 0x{:04X}, data len {}", record.header.record_type, record.data.len());
            match record.header.record_type {
                record_types::BEGIN_SST => {
                    // println!("DEBUG SST: Found BEGIN_SST");
                    // SST header, continue reading
                },
                record_types::SST_ITEM => {
                    // println!("DEBUG SST: Found SST_ITEM");
                    if let Ok(sst_item) =
                        crate::ooxml::xlsb::records::SstItemRecord::parse(&record.data)
                    {
                        // println!("DEBUG SST: Parsed string: '{}'", sst_item.string);
                        strings.push(sst_item.string);
                    } /* else {
                    println!("DEBUG SST: Failed to parse SST_ITEM");
                    }*/
                },
                record_types::END_SST => {
                    // println!("DEBUG SST: Found END_SST, breaking");
                    break;
                },
                _ => {
                    // Skip other records
                    // println!("DEBUG SST: Skipping record type 0x{:04X}", record.header.record_type);
                },
            }
        }
        // println!("DEBUG SST: Total strings parsed: {}", strings.len());
        Ok(())
    }

    /// Read workbook structure
    fn read_workbook(
        iter: &mut XlsbRecordIter<impl Read>,
        worksheet_names: &mut Vec<String>,
        is_1904: &mut bool,
    ) -> XlsbResult<()> {
        for record in iter.by_ref() {
            let record = record?;
            // println!("DEBUG: Record type 0x{:04X}, data len {}", record.header.record_type, record.data.len());
            match record.header.record_type {
                record_types::WORKBOOK_PROP => {
                    // println!("DEBUG: Found WORKBOOK_PROP");
                    if let Ok(prop) =
                        crate::ooxml::xlsb::records::WorkbookPropRecord::parse(&record.data)
                    {
                        *is_1904 = prop.is_date1904;
                    }
                },
                record_types::BUNDLE_SH => {
                    // println!("DEBUG: Found BUNDLE_SH");
                    match crate::ooxml::xlsb::records::BundleSheetRecord::parse(&record.data) {
                        Ok(bundle_sh) => {
                            // println!("DEBUG: Parsed sheet name: {}", bundle_sh.name);
                            worksheet_names.push(bundle_sh.name);
                        },
                        Err(_e) => {
                            // println!("DEBUG: Failed to parse BundleSheetRecord: {:?}", e);
                        },
                    }
                },
                record_types::END_BUNDLE_SHS => {
                    // println!("DEBUG: Found END_BUNDLE_SHS, breaking");
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }
        Ok(())
    }

    /// Read a worksheet
    fn read_worksheet(
        cursor: Cursor<&[u8]>,
        name: String,
        shared_strings: &[String],
    ) -> XlsbResult<XlsbWorksheet> {
        let mut worksheet = XlsbWorksheet::new(name);
        let iter =
            crate::ooxml::xlsb::records::RecordIter::<std::io::Cursor<&[u8]>>::from_cursor(cursor);
        let mut cells_reader =
            crate::ooxml::xlsb::cells_reader::XlsbCellsReader::new(iter, shared_strings.to_vec())?;

        while let Some(cell) = cells_reader.next_cell()? {
            worksheet.add_cell(cell);
        }

        Ok(worksheet)
    }
}

impl crate::sheet::WorkbookTrait for XlsbWorkbook {
    fn active_sheet_index(&self) -> usize {
        0
    }

    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        self.worksheet_by_index(0)
    }

    fn worksheet_count(&self) -> usize {
        self.worksheet_names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_index(
        &self,
        index: usize,
    ) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        let worksheet = self.get_worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_by_name(
        &self,
        name: &str,
    ) -> Result<Box<dyn SheetTrait + '_>, Box<dyn std::error::Error>> {
        for (i, ws_name) in self.worksheet_names.iter().enumerate() {
            if ws_name == name {
                return self.worksheet_by_index(i);
            }
        }
        Err(format!("Worksheet '{}' not found", name).into())
    }

    fn worksheets<'a>(&'a self) -> Box<dyn WorksheetIterator<'a> + 'a> {
        Box::new(XlsbWorksheetIterator {
            workbook: self,
            index: 0,
        })
    }
}

pub struct XlsbWorksheetIterator<'a> {
    workbook: &'a XlsbWorkbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsbWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>, Box<dyn std::error::Error>>> {
        if self.index < self.workbook.worksheet_names.len() {
            match self.workbook.get_worksheet(self.index) {
                Ok(worksheet) => {
                    self.index += 1;
                    Some(Ok(Box::new(worksheet)))
                },
                Err(e) => {
                    self.index += 1; // Continue to next worksheet even on error
                    Some(Err(Box::new(e)))
                },
            }
        } else {
            None
        }
    }
}
