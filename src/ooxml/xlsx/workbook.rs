//! Excel Workbook implementation.
//!
//! This module provides the concrete implementation of the Workbook trait
//! for Excel (.xlsx) files using the Office Open XML format.

use crate::ooxml::common::DocumentProperties;
use crate::ooxml::opc::{OpcPackage, PackURI};
use crate::ooxml::xlsx::writer::{MutableWorkbookData, MutableWorksheet};
use crate::ooxml::xlsx::{SharedStrings, Styles};
use crate::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as WorksheetTrait, WorksheetIterator,
};

use super::parsers::workbook_parser;
use super::worksheet::{Worksheet, WorksheetInfo, WorksheetIterator as XlsxWorksheetIterator};

/// Concrete implementation of a Workbook for Excel files.
#[derive(Debug)]
pub struct Workbook {
    /// The underlying OPC package
    package: OpcPackage,
    /// Cached worksheet information
    worksheets: Vec<WorksheetInfo>,
    /// Cached worksheet names for zero-copy returns
    worksheet_names: Vec<String>,
    /// Active worksheet index (0-based)
    active_sheet_index: usize,
    /// Shared strings table for efficient string storage
    shared_strings: SharedStrings,
    /// Styles information
    styles: Styles,
    /// Mutable workbook data for writing (cached)
    mutable_data: Option<MutableWorkbookData>,
    /// Document properties (metadata)
    properties: DocumentProperties,
}

impl Workbook {
    /// Create a new empty workbook.
    ///
    /// Creates a minimal valid Excel workbook with one default worksheet.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let workbook = Workbook::create()?;
    /// // Add data to worksheets...
    /// workbook.save("new_workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn create() -> SheetResult<Self> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;
        use crate::ooxml::xlsx::template;

        let mut package = OpcPackage::new();

        // Create workbook.xml
        let workbook_uri = PackURI::new("/xl/workbook.xml")?;
        let workbook_part = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.to_string(),
            template::default_workbook_xml().as_bytes().to_vec(),
        );
        // Use relative path for package-level relationship
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));

        // Create worksheet
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let worksheet_part = BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.to_string(),
            template::default_worksheet_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to(
                "worksheets/sheet1.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            );
        }
        package.add_part(Box::new(worksheet_part));

        // Create styles.xml
        let styles_uri = PackURI::new("/xl/styles.xml")?;
        let styles_part = BlobPart::new(
            styles_uri,
            ct::SML_STYLES.to_string(),
            template::default_styles_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to("styles.xml", rt::STYLES);
        }
        package.add_part(Box::new(styles_part));

        // Create sharedStrings.xml
        let shared_strings_uri = PackURI::new("/xl/sharedStrings.xml")?;
        let shared_strings_part = BlobPart::new(
            shared_strings_uri,
            ct::SML_SHARED_STRINGS.to_string(),
            template::default_shared_strings_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to(
                "sharedStrings.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            );
        }
        package.add_part(Box::new(shared_strings_part));

        // Create theme
        let theme_uri = PackURI::new("/xl/theme/theme1.xml")?;
        let theme_part = BlobPart::new(
            theme_uri,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to("theme/theme1.xml", rt::THEME);
        }
        package.add_part(Box::new(theme_part));

        // Create core.xml
        let core_props_uri = PackURI::new("/docProps/core.xml")?;
        let core_props_part = BlobPart::new(
            core_props_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );
        package.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        package.add_part(Box::new(core_props_part));

        // Create app.xml
        let app_props_uri = PackURI::new("/docProps/app.xml")?;
        let app_props_part = BlobPart::new(
            app_props_uri,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );
        package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        package.add_part(Box::new(app_props_part));

        Self::new(package)
    }

    /// Create a new workbook from an OPC package.
    pub fn new(package: OpcPackage) -> SheetResult<Self> {
        let mut workbook = Workbook {
            package,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            active_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            styles: Styles::new(),
            mutable_data: Some(MutableWorkbookData::new()),
            properties: DocumentProperties::new(),
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

        // Cache worksheet names for zero-copy returns
        self.worksheet_names = worksheets.iter().map(|ws| ws.name.clone()).collect();
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
            self.styles = Styles::parse(content)
                .map_err(|e| -> Box<dyn std::error::Error> { Box::new(e) })?;
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

    fn worksheet_names(&self) -> &[String] {
        // Return cached slice - zero-copy!
        &self.worksheet_names
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

impl Workbook {
    /// Open a workbook from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> SheetResult<Self> {
        let package = OpcPackage::open(path)?;
        Self::new(package)
    }

    /// Get a mutable worksheet for writing and modification.
    ///
    /// # Arguments
    ///
    /// * `index` - Worksheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// ws.set_cell_value(1, 1, "Hello");
    /// ws.set_cell_value(1, 2, "World");
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn worksheet_mut(&mut self, index: usize) -> SheetResult<&mut MutableWorksheet> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().worksheet_mut(index)
    }

    /// Add a new worksheet to the workbook.
    ///
    /// # Arguments
    ///
    /// * `name` - Name for the new worksheet
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Data");
    /// wb.add_worksheet("Summary");
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn add_worksheet(&mut self, name: &str) -> &mut MutableWorksheet {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .add_worksheet(name.to_string())
    }

    /// Define a named range.
    ///
    /// Named ranges allow you to refer to cells or ranges by meaningful names.
    ///
    /// # Arguments
    /// * `name` - Name for the range (e.g., "TaxRate", "SalesData")
    /// * `reference` - Reference formula (e.g., "Sheet1!$A$1:$B$10", "Sheet1!$C$5")
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.define_name("TaxRate", "Sheet1!$A$1");
    /// wb.define_name("SalesData", "Sheet1!$A$1:$D$100");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn define_name(&mut self, name: &str, reference: &str) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name(name, reference);
    }

    /// Define a sheet-scoped named range.
    ///
    /// # Arguments
    /// * `name` - Name for the range
    /// * `reference` - Reference formula
    /// * `sheet_id` - 1-based sheet ID
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet_id: u32) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name_local(name, reference, sheet_id);
    }

    /// Define a named range with a comment.
    pub fn define_name_with_comment(&mut self, name: &str, reference: &str, comment: &str) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name_with_comment(name, reference, comment);
    }

    /// Remove a named range by name.
    pub fn remove_name(&mut self, name: &str) -> bool {
        self.mutable_data
            .as_mut()
            .map(|d| d.remove_name(name))
            .unwrap_or(false)
    }

    /// Get a reference to the workbook properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::create()?;
    /// let props = wb.properties();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the workbook properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.properties_mut().title = Some("My Workbook".to_string());
    /// wb.properties_mut().creator = Some("John Doe".to_string());
    /// wb.save("workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
    }

    /// Save the workbook to a file.
    ///
    /// Writes the complete Excel workbook including all worksheets, styles,
    /// and shared strings to an .xlsx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .xlsx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut workbook = Workbook::create()?;
    /// // Modify workbook...
    /// workbook.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> SheetResult<()> {
        // If we have mutable data, update the workbook parts
        let should_update = self
            .mutable_data
            .as_ref()
            .map(|d| d.is_modified())
            .unwrap_or(false);

        if should_update {
            // Take mutable_data temporarily to avoid borrow issues
            if let Some(mut mutable_data) = self.mutable_data.take() {
                self.update_workbook_parts(&mut mutable_data)?;
                self.mutable_data = Some(mutable_data);
            }
        }

        // Update core properties
        self.update_core_properties()?;

        self.package.save(path)?;
        Ok(())
    }

    /// Update workbook parts with modified data.
    fn update_workbook_parts(&mut self, data: &mut MutableWorkbookData) -> SheetResult<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::{BlobPart, Part};

        let workbook_uri = PackURI::new("/xl/workbook.xml")?;

        // Create temporary workbook part to manage relationships
        let mut temp_wb_part = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.to_string(),
            Vec::new(),
        );

        // Build styles from all worksheets FIRST
        let (styles_builder, worksheet_style_indices) = data.build_styles()?;

        // Generate and write styles.xml
        let styles_xml = styles_builder.to_xml()?;
        let styles_uri = PackURI::new("/xl/styles.xml")?;
        let styles_part = BlobPart::new(
            styles_uri,
            ct::SML_STYLES.to_string(),
            styles_xml.into_bytes(),
        );
        self.package.add_part(Box::new(styles_part));

        // Create styles relationship
        temp_wb_part.relate_to("styles.xml", rt::STYLES);

        // Track worksheet relationship IDs for workbook.xml generation
        let mut worksheet_rel_ids: Vec<String> = Vec::new();

        // Update worksheet parts and create relationships
        // IMPORTANT: Create relationships for ALL worksheets, not just modified ones
        for (index, ws) in data.worksheets.iter().enumerate() {
            // Get style indices for this worksheet
            let style_indices = worksheet_style_indices
                .get(index)
                .cloned()
                .unwrap_or_default();

            // Generate XML with proper style indices
            let ws_xml = ws.to_xml(&mut data.shared_strings, &style_indices)?;
            let ws_uri = PackURI::new(format!("/xl/worksheets/sheet{}.xml", ws.sheet_id()))?;
            let ws_part = BlobPart::new(ws_uri, ct::SML_WORKSHEET.to_string(), ws_xml.into_bytes());
            self.package.add_part(Box::new(ws_part));

            // Create relationship and track the ID (for ALL sheets)
            let rel_target = format!("worksheets/sheet{}.xml", ws.sheet_id());
            let rid = temp_wb_part.relate_to(
                &rel_target,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            );
            worksheet_rel_ids.push(rid);
        }

        // Update shared strings
        let ss_xml = data.shared_strings.to_xml()?;
        let ss_uri = PackURI::new("/xl/sharedStrings.xml")?;
        let ss_part = BlobPart::new(
            ss_uri,
            ct::SML_SHARED_STRINGS.to_string(),
            ss_xml.into_bytes(),
        );
        self.package.add_part(Box::new(ss_part));

        // Create shared strings relationship
        temp_wb_part.relate_to(
            "sharedStrings.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        );

        // Now generate workbook XML with actual relationship IDs
        let workbook_xml = data.generate_workbook_xml_with_rels(&worksheet_rel_ids)?;
        temp_wb_part.set_blob(workbook_xml.into_bytes());

        // Add the workbook part to the package
        self.package.add_part(Box::new(temp_wb_part));

        Ok(())
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> SheetResult<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::part::BlobPart;

        let core_uri = PackURI::new("/docProps/core.xml")?;

        // Generate XML from properties
        let xml = self.properties.to_xml();

        // Create or update the core properties part
        let core_part = BlobPart::new(
            core_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.package.add_part(Box::new(core_part));

        Ok(())
    }
}
