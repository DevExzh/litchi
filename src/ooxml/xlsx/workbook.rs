//! Excel Workbook implementation.
//!
//! This module provides the concrete implementation of the Workbook trait
//! for Excel (.xlsx) files using the Office Open XML format.

use crate::ooxml::opc::{OpcPackage, PackURI};
use crate::sheet::{Workbook as WorkbookTrait, Worksheet as WorksheetTrait, WorksheetIterator, Result as SheetResult};
use crate::ooxml::xlsx::{SharedStrings, Styles};

use super::worksheet::{Worksheet, WorksheetInfo, WorksheetIterator as XlsxWorksheetIterator};
use super::parsers::{workbook_parser, styles_parser};

/// Concrete implementation of a Workbook for Excel files.
pub struct Workbook {
    /// The underlying OPC package
    package: OpcPackage,
    /// Cached worksheet information
    worksheets: Vec<WorksheetInfo>,
    /// Active worksheet index (0-based)
    active_sheet_index: usize,
    /// Shared strings table for efficient string storage
    shared_strings: SharedStrings,
    /// Styles information
    styles: Styles,
}

impl Workbook {
    /// Create a new workbook from an OPC package.
    pub fn new(package: OpcPackage) -> SheetResult<Self> {
        let mut workbook = Workbook {
            package,
            worksheets: Vec::new(),
            active_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            styles: Styles::new(),
        };

        workbook.load_workbook_info()?;
        workbook.load_shared_strings()?;
        workbook.load_styles()?;

        Ok(workbook)
    }

    /// Load workbook information from workbook.xml
    fn load_workbook_info(&mut self) -> SheetResult<()> {
        let workbook_uri = PackURI::new("/xl/workbook.xml")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;

        // Parse the workbook XML to extract sheet information
        let content = std::str::from_utf8(workbook_part.blob())?;

        // Extract sheets from workbook.xml
        let (worksheets, active_sheet_index) = workbook_parser::parse_workbook_xml(content)?;
        self.worksheets = worksheets;
        self.active_sheet_index = active_sheet_index;

        Ok(())
    }

    /// Load shared strings from xl/sharedStrings.xml
    fn load_shared_strings(&mut self) -> SheetResult<()> {
        let shared_strings_uri = PackURI::new("/xl/sharedStrings.xml")?;
        if let Ok(shared_strings_part) = self.package.get_part(&shared_strings_uri) {
            let content = std::str::from_utf8(shared_strings_part.blob())?;
            self.shared_strings = SharedStrings::parse(content)?;
        }

        Ok(())
    }

    /// Load styles from xl/styles.xml
    fn load_styles(&mut self) -> SheetResult<()> {
        let styles_uri = PackURI::new("/xl/styles.xml")?;
        if let Ok(styles_part) = self.package.get_part(&styles_uri) {
            let content = std::str::from_utf8(styles_part.blob())?;
            self.styles = styles_parser::parse_styles_xml(content)?;
        }
        Ok(())
    }

    /// Get a worksheet by index
    fn get_worksheet(&self, index: usize) -> SheetResult<Worksheet<'_>> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        let info = &self.worksheets[index];
        let mut worksheet = Worksheet::new(self, info.clone());

        // Load worksheet data
        worksheet.load_data()?;

        Ok(worksheet)
    }

    /// Get the OPC package (for internal use by worksheet)
    pub(crate) fn package(&self) -> &OpcPackage {
        &self.package
    }

    /// Get the shared strings table (for internal use by worksheet)
    pub(crate) fn shared_strings(&self) -> &SharedStrings {
        &self.shared_strings
    }
}

impl WorkbookTrait for Workbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        let worksheet = self.get_worksheet(self.active_sheet_index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_names(&self) -> Vec<String> {
        self.worksheets.iter().map(|ws| ws.name.clone()).collect()
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        for (index, ws_info) in self.worksheets.iter().enumerate() {
            if ws_info.name == name {
                let worksheet = self.get_worksheet(index)?;
                return Ok(Box::new(worksheet));
            }
        }
        Err(format!("Worksheet '{}' not found", name).into())
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        let worksheet = self.get_worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(XlsxWorksheetIterator::new(self.worksheets.clone(), self))
    }

    fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        self.active_sheet_index
    }
}