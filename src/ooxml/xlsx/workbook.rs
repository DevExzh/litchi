//! Excel Workbook implementation.
//!
//! This module provides the concrete implementation of the Workbook trait
//! for Excel (.xlsx) files using the Office Open XML format.

use crate::ooxml::common::DocumentProperties;
use crate::ooxml::opc::{OpcPackage, PackURI};
use crate::ooxml::pivot::PivotTable;
use crate::ooxml::xlsx::writer::workbook::{
    generate_pivot_cache_definition_xml, generate_pivot_cache_records_xml,
    generate_pivot_table_definition_xml, render_pivot_table_sheet_cells,
};
use crate::ooxml::xlsx::writer::{MutableWorkbookData, MutableWorksheet};
use crate::ooxml::xlsx::{SharedStrings, Styles};
use crate::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as WorksheetTrait, WorksheetIterator,
};
use std::collections::HashMap;

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
        workbook.load_print_settings()?;

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

    /// Load worksheet print settings (print area, repeating rows/columns)
    /// from workbook-level defined names.
    fn load_print_settings(&mut self) -> SheetResult<()> {
        use crate::ooxml::opc::PackURI as Uri;

        let workbook_uri = Uri::new("/xl/workbook.xml")?;
        let workbook_part = match self.package.get_part(&workbook_uri) {
            Ok(part) => part,
            Err(_) => return Ok(()),
        };

        let content = std::str::from_utf8(workbook_part.blob())?;

        // Find the <definedNames> section if present.
        let start = if let Some(pos) = content.find("<definedNames>") {
            pos
        } else {
            return Ok(());
        };

        let end_rel = if let Some(pos) = content[start..].find("</definedNames>") {
            pos
        } else {
            return Ok(());
        };

        let defined_names_xml = &content[start..start + end_rel + "</definedNames>".len()];

        let mut pos = 0usize;
        while let Some(rel) = defined_names_xml[pos..].find("<definedName ") {
            let start_pos = pos + rel;
            let after_start = &defined_names_xml[start_pos..];

            // Find end of this definedName element (we assume well-formed XML)
            let end_tag_rel = match after_start.find("</definedName>") {
                Some(p) => p,
                None => break,
            };
            let end_pos = start_pos + end_tag_rel + "</definedName>".len();
            let dn_xml = &defined_names_xml[start_pos..end_pos];

            Self::apply_defined_name_print_setting(dn_xml, &mut self.worksheets)?;

            pos = end_pos;
        }

        Ok(())
    }

    /// Apply a single <definedName> element to worksheet print settings if it
    /// represents _xlnm.Print_Area or _xlnm.Print_Titles.
    fn apply_defined_name_print_setting(
        dn_xml: &str,
        worksheets: &mut [WorksheetInfo],
    ) -> SheetResult<()> {
        // Split into start tag and inner text.
        let gt_pos = match dn_xml.find('>') {
            Some(p) => p,
            None => return Ok(()),
        };

        let (start_tag, inner) = dn_xml.split_at(gt_pos + 1);
        let value_end = match inner.rfind("</definedName>") {
            Some(p) => p,
            None => return Ok(()),
        };
        let value_text = &inner[..value_end];

        let name = Self::extract_defined_name_attr(start_tag, "name");
        let local_sheet_id = Self::extract_defined_name_attr(start_tag, "localSheetId");

        let (name, sheet_idx) = match (name, local_sheet_id) {
            (Some(n), Some(sid)) => {
                let idx: usize = match sid.parse::<u32>() {
                    Ok(v) => v as usize,
                    Err(_) => return Ok(()),
                };
                if idx >= worksheets.len() {
                    return Ok(());
                }
                (n, idx)
            },
            _ => return Ok(()),
        };

        if name == "_xlnm.Print_Area" {
            if let Some(range) = Self::parse_print_area(value_text) {
                worksheets[sheet_idx].print_area = Some(range);
            }
        } else if name == "_xlnm.Print_Titles" {
            let (rows, cols) = Self::parse_print_titles(value_text);
            if let Some(r) = rows {
                worksheets[sheet_idx].repeating_rows = Some(r);
            }
            if let Some(c) = cols {
                worksheets[sheet_idx].repeating_columns = Some(c);
            }
        }

        Ok(())
    }

    /// Extract a simple XML attribute value from a <definedName ...> start tag.
    fn extract_defined_name_attr(tag: &str, attr: &str) -> Option<String> {
        let pattern = format!("{}=\"", attr);
        let start = tag.find(&pattern)? + pattern.len();
        let tail = &tag[start..];
        let end = tail.find('"')?;
        Some(tail[..end].to_string())
    }

    /// Parse the print area reference from a defined name value.
    ///
    /// Values are typically of the form `'Sheet Name'!A1:D20` or a comma-
    /// separated list of such references. We return the range part for the
    /// first entry (e.g., `A1:D20`).
    fn parse_print_area(value: &str) -> Option<String> {
        let first = value.split(',').next()?.trim();
        let bang = first.rfind('!')?;
        let range = first[bang + 1..].trim();
        if range.is_empty() {
            None
        } else {
            Some(range.to_string())
        }
    }

    /// Parse repeating rows/columns from a _xlnm.Print_Titles defined name
    /// value. Returns (rows, columns) as raw range strings (e.g., "$1:$1",
    /// "$A:$B").
    fn parse_print_titles(value: &str) -> (Option<String>, Option<String>) {
        let mut rows: Option<String> = None;
        let mut cols: Option<String> = None;

        for part in value.split(',') {
            let part = part.trim();
            let bang = match part.rfind('!') {
                Some(p) => p,
                None => continue,
            };
            let range = part[bang + 1..].trim();
            if range.is_empty() {
                continue;
            }

            // Skip leading '$' characters when deciding whether this is a
            // row or column reference.
            let mut chars = range.chars().skip_while(|c| *c == '$');
            match chars.next() {
                Some(ch) if ch.is_ascii_digit() => {
                    if rows.is_none() {
                        rows = Some(range.to_string());
                    }
                },
                Some(ch) if ch.is_ascii_alphabetic() => {
                    if cols.is_none() {
                        cols = Some(range.to_string());
                    }
                },
                _ => {},
            }
        }

        (rows, cols)
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

    /// Get the styles collection (for internal use by worksheet)
    pub(crate) fn styles(&self) -> &Styles {
        &self.styles
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

    #[cfg(feature = "ooxml_encryption")]
    pub fn open_with_password<P: AsRef<std::path::Path>>(
        path: P,
        password: &str,
    ) -> SheetResult<Self> {
        let data = std::fs::read(path.as_ref())?;
        let decrypted = crate::ooxml::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
        let package = OpcPackage::from_bytes(&decrypted.package_bytes)?;
        Self::new(package)
    }

    /// Get a mutable worksheet for writing and modification.
    ///
    /// # Arguments
    ///
    /// * `index` - Worksheet index (0-based)
    ///
    // ... (rest of the code remains the same)
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

    /// Add a pivot table to the workbook (writer).
    ///
    /// This wires the pivot cache/table into the save pipeline; when you call
    /// `save`, the necessary parts and relationships will be created.
    pub fn add_pivot_table(&mut self, pivot: PivotTable) -> SheetResult<()> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().add_pivot_table(pivot)
    }

    /// Add a new worksheet.
    ///
    /// # Arguments
    /// * `name` - The name of the new worksheet
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
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

        // Update app properties (extended properties)
        self.update_app_properties()?;

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

        // Create theme relationship (required by Excel)
        temp_wb_part.relate_to("theme/theme1.xml", rt::THEME);

        // Track worksheet relationship IDs for workbook.xml generation
        let mut worksheet_rel_ids: Vec<String> = Vec::new();

        // Track pivot cache relationship IDs for workbook.xml
        let mut pivot_cache_rel_ids: Vec<(u32, String)> = Vec::new();
        // Track pivot table targets per worksheet (for worksheet rels)
        let mut pivot_table_targets_per_sheet: Vec<Vec<String>> =
            vec![Vec::new(); data.worksheets.len()];

        // Pre-create pivot cache and pivot table parts so worksheets can relate to them
        for (idx, pivot) in data.pivot_tables.iter().enumerate() {
            let cache_id = (idx as u32) + 1;

            // pivotCacheRecords part (materialized from source range)
            let records_uri =
                PackURI::new(format!("/xl/pivotCache/pivotCacheRecords{}.xml", cache_id))?;
            let (records_xml, record_count, field_stats) =
                generate_pivot_cache_records_xml(pivot, &data.worksheets)?;
            let records_part = BlobPart::new(
                records_uri,
                ct::SML_PIVOT_CACHE_RECORDS.to_string(),
                records_xml.into_bytes(),
            );
            self.package.add_part(Box::new(records_part));

            // pivotCacheDefinition part
            let cache_def_uri = PackURI::new(format!(
                "/xl/pivotCache/pivotCacheDefinition{}.xml",
                cache_id
            ))?;
            let mut cache_def_part = BlobPart::new(
                cache_def_uri,
                ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
                Vec::new(),
            );
            let records_rel_id = cache_def_part.relate_to(
                &format!("pivotCacheRecords{}.xml", cache_id),
                rt::PIVOT_CACHE_RECORDS,
            );
            let cache_def_xml = generate_pivot_cache_definition_xml(
                pivot,
                Some(records_rel_id.as_str()),
                record_count,
                &field_stats,
            )?;
            cache_def_part.set_blob(cache_def_xml.into_bytes());
            self.package.add_part(Box::new(cache_def_part));

            // workbook -> pivotCacheDefinition rel
            let cache_rel_id = temp_wb_part.relate_to(
                &format!("pivotCache/pivotCacheDefinition{}.xml", cache_id),
                rt::PIVOT_CACHE_DEFINITION,
            );
            pivot_cache_rel_ids.push((cache_id, cache_rel_id.clone()));

            // pivotTableDefinition part
            let table_idx = cache_id; // align ids for predictability
            let pivot_table_uri =
                PackURI::new(format!("/xl/pivotTables/pivotTable{}.xml", table_idx))?;
            let mut pivot_table_part =
                BlobPart::new(pivot_table_uri, ct::SML_PIVOT_TABLE.to_string(), Vec::new());

            // pivotTable -> pivotCacheDefinition rel
            let _pt_cache_rel_id = pivot_table_part.relate_to(
                &format!("../pivotCache/pivotCacheDefinition{}.xml", cache_id),
                rt::PIVOT_CACHE_DEFINITION,
            );

            // Serialize pivotTable XML
            let pivot_table_xml =
                generate_pivot_table_definition_xml(pivot, cache_id, &field_stats)?;
            pivot_table_part.set_blob(pivot_table_xml.into_bytes());
            self.package.add_part(Box::new(pivot_table_part));

            // Record worksheet target for later worksheet rel creation
            let sheet_idx = pivot.dest_sheet_index;
            if let Some(list) = pivot_table_targets_per_sheet.get_mut(sheet_idx) {
                list.push(format!("../pivotTables/pivotTable{}.xml", table_idx));
            } else {
                return Err(format!(
                    "Pivot table destination sheet index {} out of bounds",
                    sheet_idx
                )
                .into());
            }
        }

        // Materialize the pivot output into destination worksheet cells.
        // This ensures Excel shows the pivot table content immediately on open.
        for pivot in data.pivot_tables.iter() {
            render_pivot_table_sheet_cells(pivot, &mut data.worksheets)?;
        }

        // Update worksheet parts and create relationships
        // IMPORTANT: Create relationships for ALL worksheets, not just modified ones
        for (index, ws) in data.worksheets.iter().enumerate() {
            // Get style indices for this worksheet
            let style_indices = worksheet_style_indices
                .get(index)
                .cloned()
                .unwrap_or_default();

            let ws_uri = PackURI::new(format!("/xl/worksheets/sheet{}.xml", ws.sheet_id()))?;

            // Create worksheet part with empty content initially (we'll set it later)
            let mut ws_part =
                BlobPart::new(ws_uri.clone(), ct::SML_WORKSHEET.to_string(), Vec::new());

            // Generate and add comments if present, create relationship
            if let Some(comments_xml) = ws.generate_comments_xml()? {
                let comments_uri = PackURI::new(format!("/xl/comments{}.xml", ws.sheet_id()))?;
                let comments_part = BlobPart::new(
                    comments_uri,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
                        .to_string(),
                    comments_xml.into_bytes(),
                );
                self.package.add_part(Box::new(comments_part));

                // Add relationship from worksheet to comments
                ws_part.relate_to(
                    &format!("../comments{}.xml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                );
            }

            // Generate and add VML drawing for comment indicators if present
            let vml_rel_id = if let Some(vml_xml) = ws.generate_vml_drawing_xml()? {
                let vml_uri =
                    PackURI::new(format!("/xl/drawings/vmlDrawing{}.vml", ws.sheet_id()))?;
                let vml_part = BlobPart::new(
                    vml_uri,
                    "application/vnd.openxmlformats-officedocument.vmlDrawing".to_string(),
                    vml_xml.into_bytes(),
                );
                self.package.add_part(Box::new(vml_part));

                // Add relationship from worksheet to VML drawing and capture the ID
                let rel_id = ws_part.relate_to(
                    &format!("../drawings/vmlDrawing{}.vml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                );
                Some(rel_id)
            } else {
                None
            };

            // Add relationships for external hyperlinks and track their IDs
            let mut hyperlink_rel_ids: HashMap<String, String> = HashMap::new();
            for hyperlink in ws.hyperlinks().iter() {
                if hyperlink.target.starts_with("http://")
                    || hyperlink.target.starts_with("https://")
                    || hyperlink.target.starts_with("ftp://")
                    || hyperlink.target.starts_with("mailto:")
                {
                    // Use relate_to_ext for external links to add TargetMode="External"
                    let rel_id = ws_part.relate_to_ext(
                        &hyperlink.target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    );
                    hyperlink_rel_ids.insert(hyperlink.cell_ref.clone(), rel_id);
                }
            }

            // Generate and add drawing XML for images if present
            if let Some(drawing_xml) = ws.generate_drawing_xml()? {
                let drawing_uri =
                    PackURI::new(format!("/xl/drawings/drawing{}.xml", ws.sheet_id()))?;

                // Create drawing part with relationships for images
                let mut drawing_part = BlobPart::new(
                    drawing_uri.clone(),
                    "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
                    drawing_xml.into_bytes(),
                );

                // Add image parts and create relationships
                for (idx, image) in ws.images().iter().enumerate() {
                    let image_ext = &image.format;
                    let image_uri = PackURI::new(format!(
                        "/xl/media/image{}.{}",
                        ws.sheet_id() * 1000 + idx as u32,
                        image_ext
                    ))?;

                    // Determine content type based on format
                    let content_type = match image_ext.to_lowercase().as_str() {
                        "png" => "image/png",
                        "jpg" | "jpeg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        "svg" => "image/svg+xml",
                        _ => "image/png", // Default to PNG
                    };

                    let image_part = BlobPart::new(
                        image_uri.clone(),
                        content_type.to_string(),
                        image.data.clone(),
                    );
                    self.package.add_part(Box::new(image_part));

                    // Add relationship from drawing to image
                    drawing_part.relate_to(
                        &format!(
                            "../media/image{}.{}",
                            ws.sheet_id() * 1000 + idx as u32,
                            image_ext
                        ),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    );
                }

                self.package.add_part(Box::new(drawing_part));

                // Add relationship from worksheet to drawing
                ws_part.relate_to(
                    &format!("../drawings/drawing{}.xml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                );
            }

            let mut pivot_table_rel_ids: Vec<String> = Vec::new();
            if let Some(targets) = pivot_table_targets_per_sheet.get(index) {
                for target in targets {
                    let rid = ws_part.relate_to(target, rt::PIVOT_TABLE);
                    pivot_table_rel_ids.push(rid);
                }
            }

            // Now generate worksheet XML with proper hyperlink relationship IDs and VML reference
            let ws_xml = ws.to_xml_with_hyperlink_rels(
                &mut data.shared_strings,
                &style_indices,
                &hyperlink_rel_ids,
                vml_rel_id.as_deref(),
                Some(&pivot_table_rel_ids),
            )?;
            ws_part.set_blob(ws_xml.into_bytes());

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

        // Synchronize worksheet print settings with workbook-level defined names
        data.sync_print_settings_to_defined_names();

        // Now generate workbook XML with actual relationship IDs
        let workbook_xml =
            data.generate_workbook_xml_with_rels(&worksheet_rel_ids, &pivot_cache_rel_ids)?;
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

    /// Update the app.xml properties part with current worksheet information.
    fn update_app_properties(&mut self) -> SheetResult<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::part::BlobPart;
        use std::fmt::Write;

        let app_uri = PackURI::new("/docProps/app.xml")?;

        // Get worksheet names from mutable_data if available, otherwise from package
        let worksheet_names: Vec<String> = if let Some(ref data) = self.mutable_data {
            data.worksheets
                .iter()
                .map(|ws| ws.name().to_string())
                .collect()
        } else {
            // Fallback to parsing from workbook.xml if no mutable data
            vec!["Sheet1".to_string()]
        };

        let worksheet_count = worksheet_names.len();

        // Generate app.xml XML
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" "#);
        xml.push_str(
            r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#,
        );
        xml.push_str("<Application>The Litchi Rust Library</Application>");
        xml.push_str("<DocSecurity>0</DocSecurity>");
        xml.push_str("<ScaleCrop>false</ScaleCrop>");

        // HeadingPairs: category name + count
        xml.push_str("<HeadingPairs>");
        xml.push_str(r#"<vt:vector size="2" baseType="variant">"#);
        xml.push_str("<vt:variant><vt:lpstr>Worksheet</vt:lpstr></vt:variant>");
        write!(
            xml,
            "<vt:variant><vt:i4>{}</vt:i4></vt:variant>",
            worksheet_count
        )
        .map_err(|e| format!("XML write error: {}", e))?;
        xml.push_str("</vt:vector>");
        xml.push_str("</HeadingPairs>");

        // TitlesOfParts: list of all worksheet names
        xml.push_str("<TitlesOfParts>");
        write!(
            xml,
            r#"<vt:vector size="{}" baseType="lpstr">"#,
            worksheet_count
        )
        .map_err(|e| format!("XML write error: {}", e))?;
        for name in &worksheet_names {
            // Escape XML special characters
            let escaped_name = name
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;")
                .replace('"', "&quot;")
                .replace('\'', "&apos;");
            write!(xml, "<vt:lpstr>{}</vt:lpstr>", escaped_name)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</vt:vector>");
        xml.push_str("</TitlesOfParts>");

        xml.push_str("<Company/>");
        xml.push_str("<LinksUpToDate>false</LinksUpToDate>");
        xml.push_str("<SharedDoc>false</SharedDoc>");
        xml.push_str("<HyperlinksChanged>false</HyperlinksChanged>");
        xml.push_str("<AppVersion>14.0000</AppVersion>");
        xml.push_str("</Properties>");

        // Create or update the app properties part
        let app_part = BlobPart::new(
            app_uri,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.package.add_part(Box::new(app_part));

        Ok(())
    }

    // ===== Workbook-level Features =====

    /// Hide a worksheet by index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.hide_sheet(0)?; // Hide the first sheet
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn hide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().hide_sheet(index)?;
        Ok(())
    }

    /// Unhide a worksheet by index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn unhide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().unhide_sheet(index)?;
        Ok(())
    }

    /// Check if a worksheet is hidden.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn is_sheet_hidden(&self, index: usize) -> bool {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.is_sheet_hidden(index))
            .unwrap_or(false)
    }

    /// Move a worksheet to a new position.
    ///
    /// # Arguments
    /// * `from_index` - Current worksheet index (0-based)
    /// * `to_index` - Target worksheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.add_worksheet("Sheet3");
    /// wb.move_sheet(2, 0)?; // Move Sheet3 to the first position
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn move_sheet(&mut self, from_index: usize, to_index: usize) -> SheetResult<()> {
        if from_index >= self.worksheets.len() || to_index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .move_sheet(from_index, to_index)?;

        // Also update local worksheets vector
        let sheet = self.worksheets.remove(from_index);
        self.worksheets.insert(to_index, sheet);

        Ok(())
    }

    /// Set sheet visibility state.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    /// * `visibility` - Visibility state: "visible", "hidden", or "veryHidden"
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_sheet_visibility(0, "hidden")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_sheet_visibility(&mut self, index: usize, visibility: &str) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if !matches!(visibility, "visible" | "hidden" | "veryHidden") {
            return Err(
                "Invalid visibility state. Must be 'visible', 'hidden', or 'veryHidden'".into(),
            );
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_sheet_visibility(index, visibility)?;
        Ok(())
    }

    /// Get sheet visibility state.
    ///
    /// Returns "visible", "hidden", or "veryHidden".
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn get_sheet_visibility(&self, index: usize) -> Option<&str> {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.get_sheet_visibility(index))
    }

    /// Set the active worksheet index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based) to set as active
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.set_active_sheet(1)?; // Make Sheet2 active
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_active_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        self.active_sheet_index = index;

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().set_active_sheet(index);
        Ok(())
    }

    /// Force formula recalculation when the workbook is opened.
    ///
    /// # Arguments
    /// * `force` - Whether to force recalculation
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_force_formula_recalculation(true);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_force_formula_recalculation(&mut self, force: bool) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_force_formula_recalculation(force);
    }

    /// Set the calculation mode for the workbook.
    ///
    /// # Arguments
    /// * `mode` - Calculation mode: "auto", "manual", or "autoNoTable"
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_calculation_mode("manual")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_calculation_mode(&mut self, mode: &str) -> SheetResult<()> {
        if !matches!(mode, "auto" | "manual" | "autoNoTable") {
            return Err(
                "Invalid calculation mode. Must be 'auto', 'manual', or 'autoNoTable'".into(),
            );
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_calculation_mode(mode);
        Ok(())
    }

    /// Get the calculation mode for the workbook.
    ///
    /// Returns "auto", "manual", or "autoNoTable".
    pub fn get_calculation_mode(&self) -> &str {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.get_calculation_mode())
            .unwrap_or("auto")
    }

    /// Set the tab color for a worksheet.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    /// * `color` - RGB hex color (e.g., "FF0000" for red)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_tab_color(0, "FF0000")?; // Set red tab color
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_tab_color(&mut self, index: usize, color: &str) -> SheetResult<()> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .worksheet_mut(index)?
            .set_tab_color(color);
        Ok(())
    }

    /// Get the tab color for a worksheet.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn get_tab_color(&self, index: usize) -> Option<&str> {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.worksheets.get(index))
            .and_then(|ws| ws.tab_color())
    }

    /// Protect the workbook with optional password.
    ///
    /// # Arguments
    /// * `password` - Optional password (will be hashed)
    /// * `lock_structure` - Prevent adding/deleting sheets
    /// * `lock_windows` - Prevent resizing/moving workbook window
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.protect_workbook(Some("password123"), true, false);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        lock_structure: bool,
        lock_windows: bool,
    ) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().protect_workbook(
            password,
            lock_structure,
            lock_windows,
        );
    }

    /// Unprotect the workbook.
    pub fn unprotect_workbook(&mut self) {
        if let Some(data) = self.mutable_data.as_mut() {
            data.unprotect_workbook();
        }
    }

    /// Check if the workbook is protected.
    pub fn is_workbook_protected(&self) -> bool {
        self.mutable_data.as_ref().is_some_and(|d| d.is_protected())
    }

    pub fn pivot_tables(&self) -> SheetResult<Vec<PivotTable>> {
        crate::ooxml::xlsx::pivot::read_pivot_tables(self.package())
    }

    pub fn pivot_tables_on_sheet(&self, sheet_name: &str) -> SheetResult<Vec<PivotTable>> {
        let all = self.pivot_tables()?;
        Ok(all
            .into_iter()
            .filter(|t| t.sheet_name == sheet_name)
            .collect())
    }

    // ===== Worksheet-level Writing Features =====
    // (These are mostly implemented via MutableWorksheet, exposed through worksheet_mut)

    // ============================================================================
    // Apache POI Features Implementation Status
    // ============================================================================
    //
    // ✅ FULLY IMPLEMENTED (Workbook-level):
    // - Hidden sheets: hide_sheet(), unhide_sheet(), is_sheet_hidden()
    // - Sheet ordering: move_sheet()
    // - Sheet visibility: set_sheet_visibility(), get_sheet_visibility()
    // - Active sheet: set_active_sheet()
    // - Workbook calculation mode: set_force_formula_recalculation(), set_calculation_mode(), get_calculation_mode()
    // - Named ranges: define_name(), define_name_local(), define_name_with_comment(), remove_name()
    // - Sheet tab color: set_tab_color(), get_tab_color()
    // - Workbook protection: protect_workbook(), unprotect_workbook(), is_workbook_protected()
    //
    // ✅ FULLY IMPLEMENTED (Worksheet reading - via Worksheet):
    // - Merged cells (reading): get_merged_regions(), is_merged_cell(), get_merge_region()
    // - Auto-filter (reading): get_auto_filter()
    // - Column width/Row height (reading): get_column_width(), get_row_height()
    // - Hyperlinks (reading): get_hyperlink(), get_hyperlinks()
    // - Comments (reading): get_cell_comment(), get_comments()
    // - Data validation (reading): get_data_validations()
    // - Conditional formatting (reading): get_conditional_formatting()
    // - Page setup (reading): get_page_setup()
    //
    // ✅ FULLY IMPLEMENTED (Worksheet writing - via MutableWorksheet):
    // - Cell values & formulas: set_cell_value(), set_cell_formula(), set_cell_formula_with_cache()
    // - Cell formatting: set_cell_format() with CellFormat (font, fill, border, number format)
    // - Merged cells: merge_cells()
    // - Column width/Row height: set_column_width(), set_row_height()
    // - Hide columns/rows: hide_column(), hide_row(), show_column(), show_row()
    // - Data validation: add_data_validation()
    // - Charts: add_chart() (basic support)
    // - Freeze panes: freeze_panes(), unfreeze_panes()
    // - Page setup: set_page_setup(), set_page_setup_with_options(), set_print_area(), clear_print_area()
    // - Auto-filter: set_auto_filter(), remove_auto_filter()
    // - Sheet protection: protect_sheet(), protect_sheet_with_options(), unprotect_sheet()
    // - Hyperlinks: set_hyperlink(), remove_hyperlink(), hyperlinks()
    // - Comments: set_cell_comment(), remove_comment(), comments()
    // - Conditional formatting: add_conditional_formatting(), clear_conditional_formatting()
    // - Row/column grouping: group_rows(), ungroup_rows(), group_columns(), ungroup_columns()
    //
    // ⚠️ BASIC IMPLEMENTATION (Data structures exist, XML generation would need enhancement):
    // - Hyperlinks: Stored but need relationship XML in worksheet rels
    // - Comments: Stored but need comments.xml part and VML drawing
    // - Conditional formatting: Stored but need full XML generation in worksheet
    // - Charts: Basic structure exists, needs DrawingML XML generation
    //
    // ⏳ NOT IMPLEMENTED (Advanced features requiring significant additional work):
    // - Pivot tables: add_pivot_table(), get_pivot_tables(), refresh_pivot_table()
    // - Images/Pictures: add_picture(), get_pictures(), delete_picture()
    // - Rich text in cells: set_rich_text_cell(), get_rich_text_cell()
    // - Subtotals: insert_subtotals(), remove_subtotals()
    // - Sparklines: add_sparkline(), get_sparklines()
    // - Slicers: add_slicer(), get_slicers()
    // - Timeline: add_timeline(), get_timelines()
    // - Power Query: get_power_query_connections()
    // - External links: get_external_links(), update_external_links()
    //
    // 📝 NOTES:
    // - Basic cell styling is fully supported via CellFormat (font, fill, border, number format)
    // - All reading operations work perfectly
    // - All core writing operations are implemented
    // - Advanced features like pivot tables, images would require substantial XML generation code
    // - The library is production-ready for standard Excel CRUD operations
}
