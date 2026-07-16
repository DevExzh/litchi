//! Excel Workbook implementation.
//!
//! This module provides the concrete implementation of the Workbook trait
//! for Excel (.xlsx) files using the Office Open XML format.

use crate::common::DocumentProperties;
use crate::pivot::PivotTable;
use crate::xlsx::writer::workbook::{
    generate_pivot_cache_definition_xml, generate_pivot_cache_records_xml,
    generate_pivot_table_definition_xml, render_pivot_table_sheet_cells,
};
use crate::xlsx::writer::{MutableWorkbookData, MutableWorksheet};
use crate::xlsx::{Cell, SharedStrings, Styles};
use crate::xlsx::external_links::{ExternalLinkEntry, load_external_link};
use crate::xlsx::calculation_properties::{
    WorkbookCalculationProperties, parse_workbook_calculation_properties,
};
use litchi_core::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as WorksheetTrait, WorksheetIterator,
};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::HashMap;

use super::parsers::workbook_parser;
use super::worksheet::{Worksheet, WorksheetInfo, WorksheetIterator as XlsxWorksheetIterator};

/// Concrete implementation of a Workbook for Excel files.
#[derive(Debug)]
pub struct Workbook {
    /// The underlying OPC package
    package: OpcPackage,
    /// Actual main-workbook part location resolved from the package relationship.
    workbook_uri: PackURI,
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
    /// Whether the workbook uses the 1904 date system
    is_1904_date_system: bool,
    /// Effective workbook formula calculation policy.
    calculation_properties: Option<WorkbookCalculationProperties>,
    external_links: Vec<ExternalLinkEntry>,
}

impl Workbook {
    /// Discover inert embedded-object and embedded-package relationships.
    pub fn embedded_parts(&self) -> crate::error::Result<Vec<crate::EmbeddedPart<'_>>> {
        crate::embedded_object::discover_embedded_parts(&self.package)
    }

    /// Create a new empty workbook.
    ///
    /// Creates a minimal valid Excel workbook with one default worksheet.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut workbook = Workbook::create()?;
    /// // Add data to worksheets...
    /// workbook.save("new_workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn create() -> SheetResult<Self> {
        use crate::xlsx::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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

        let mut workbook = Self::new(package)?;
        workbook.mutable_data = Some(MutableWorkbookData::new());
        Ok(workbook)
    }

    /// Create a new workbook from an OPC package.
    pub fn new(package: OpcPackage) -> SheetResult<Self> {
        use litchi_opc::constants::content_type as ct;

        let workbook_part = package.main_document_part()?;
        if !matches!(
            workbook_part.content_type(),
            ct::SML_SHEET_MAIN
                | ct::SML_TEMPLATE_MAIN
                | ct::SML_SHEET_MACRO_MAIN
                | ct::SML_TEMPLATE_MACRO_MAIN
        ) {
            return Err(format!(
                "main document part '{}' is not an XML workbook",
                workbook_part.partname()
            )
            .into());
        }
        let workbook_uri = workbook_part.partname().clone();
        let mut workbook = Workbook {
            package,
            workbook_uri,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            active_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            styles: Styles::new(),
            mutable_data: None,
            properties: DocumentProperties::new(),
            is_1904_date_system: false,
            calculation_properties: None,
            external_links: Vec::new(),
        };

        workbook.load_workbook_info()?;
        workbook.load_external_links()?;
        workbook.load_shared_strings()?;
        workbook.load_styles()?;

        Ok(workbook)
    }

    /// Load workbook information from workbook.xml
    fn load_workbook_info(&mut self) -> SheetResult<()> {
        let workbook_part = self.package.get_part(&self.workbook_uri)?;

        // Parse the workbook XML to extract sheet information
        let content = std::str::from_utf8(workbook_part.blob())?;

        // Extract sheets from workbook.xml
        let mut details = workbook_parser::parse_workbook_details(content)?;
        let calculation_properties = parse_workbook_calculation_properties(content.as_bytes())?;
        Self::apply_defined_name_print_settings(&details.defined_names, &mut details.sheets);

        // Cache worksheet names for zero-copy returns
        self.worksheet_names = details.sheets.iter().map(|ws| ws.name.clone()).collect();
        self.worksheets = details.sheets;
        self.active_sheet_index = details.active_sheet_index;
        self.is_1904_date_system = details.uses_1904_date_system;
        self.calculation_properties = calculation_properties;

        Ok(())
    }

    fn load_external_links(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};
        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let content = std::str::from_utf8(workbook_part.blob())?;
        let details = workbook_parser::parse_workbook_details(content)?;
        let mut links = Vec::with_capacity(details.external_reference_ids.len());
        for (offset, relationship_id) in details.external_reference_ids.into_iter().enumerate() {
            let relationship = workbook_part.rels().get(&relationship_id).ok_or_else(|| format!("workbook external reference '{relationship_id}' has no relationship"))?;
            if relationship.is_external() || !matches!(relationship.reltype(), rt::EXTERNAL_LINK | rt::STRICT_EXTERNAL_LINK) {
                return Err(format!("workbook external reference '{relationship_id}' has an invalid relationship").into());
            }
            let uri = relationship.target_partname()?;
            let part = self.package.get_part(&uri)?;
            if part.content_type() != ct::SML_EXTERNAL_LINK {
                return Err(format!("external-link part '{uri}' has invalid content type '{}', expected '{}'", part.content_type(), ct::SML_EXTERNAL_LINK).into());
            }
            let index = u32::try_from(offset + 1).map_err(|_| "external-link index overflow")?;
            links.push(load_external_link(part, relationship_id, index)?);
        }
        self.external_links = links;
        Ok(())
    }

    pub fn external_links(&self) -> &[ExternalLinkEntry] {
        &self.external_links
    }

    /// Effective workbook formula calculation policy, when `calcPr` is present.
    pub fn calculation_properties(&self) -> Option<&WorkbookCalculationProperties> {
        self.calculation_properties.as_ref()
    }

    pub fn external_link(&self, one_based_index: u32) -> Option<&ExternalLinkEntry> {
        one_based_index.checked_sub(1).and_then(|index| self.external_links.get(index as usize))
    }

    /// Load shared strings from xl/sharedStrings.xml
    fn load_shared_strings(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        let Some(shared_strings_uri) = self.related_workbook_part_uri(
            &[rt::SHARED_STRINGS, rt::STRICT_SHARED_STRINGS],
            "shared strings",
        )?
        else {
            return Ok(());
        };
        let shared_strings_part = self.package.get_part(&shared_strings_uri)?;
        Self::require_part_content_type(
            &shared_strings_uri,
            shared_strings_part.content_type(),
            ct::SML_SHARED_STRINGS,
        )?;
        let content = std::str::from_utf8(shared_strings_part.blob())?;
        self.shared_strings = SharedStrings::parse(content)?;

        Ok(())
    }

    /// Load styles from xl/styles.xml
    fn load_styles(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        let Some(styles_uri) =
            self.related_workbook_part_uri(&[rt::STYLES, rt::STRICT_STYLES], "styles")?
        else {
            return Ok(());
        };
        let styles_part = self.package.get_part(&styles_uri)?;
        Self::require_part_content_type(&styles_uri, styles_part.content_type(), ct::SML_STYLES)?;
        let content = std::str::from_utf8(styles_part.blob())?;
        self.styles = Styles::parse(content)
            .map_err(|e| -> Box<dyn std::error::Error + Send + Sync> { Box::new(e) })?;
        Ok(())
    }

    fn related_workbook_part_uri(
        &self,
        relationship_types: &[&str],
        description: &str,
    ) -> SheetResult<Option<PackURI>> {
        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let mut matching = workbook_part
            .rels()
            .iter()
            .filter(|relationship| relationship_types.contains(&relationship.reltype()));
        let Some(relationship) = matching.next() else {
            return Ok(None);
        };
        if matching.next().is_some() {
            return Err(format!("workbook has multiple {description} relationships").into());
        }
        if relationship.is_external() {
            return Err(format!("workbook {description} relationship cannot be external").into());
        }
        Ok(Some(relationship.target_partname()?))
    }

    fn require_part_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
        if actual != expected {
            return Err(
                format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
            );
        }
        Ok(())
    }

    /// Apply sheet-scoped print-area and print-title defined names.
    fn apply_defined_name_print_settings(
        defined_names: &[workbook_parser::DefinedName],
        worksheets: &mut [WorksheetInfo],
    ) {
        for defined_name in defined_names {
            let Some(sheet_idx) = defined_name
                .local_sheet_id
                .and_then(|index| usize::try_from(index).ok())
            else {
                continue;
            };
            let worksheet = &mut worksheets[sheet_idx];
            if defined_name.name == "_xlnm.Print_Area" {
                worksheet.print_area = Self::parse_print_area(&defined_name.value);
            } else if defined_name.name == "_xlnm.Print_Titles" {
                let (rows, columns) = Self::parse_print_titles(&defined_name.value);
                worksheet.repeating_rows = rows;
                worksheet.repeating_columns = columns;
            }
        }
    }

    /// Parse the print area reference from a defined name value.
    ///
    /// Values are typically of the form `'Sheet Name'!A1:D20` or a comma-
    /// separated list of such references. We return the range part for the
    /// first entry (e.g., `A1:D20`).
    fn parse_print_area(value: &str) -> Option<String> {
        let first = Self::split_defined_name_areas(value).into_iter().next()?;
        let bang = first.rfind('!')?;
        let range = first[bang + 1..].trim();
        Self::is_valid_print_cell_range(range).then(|| range.to_string())
    }

    /// Parse repeating rows/columns from a _xlnm.Print_Titles defined name
    /// value. Returns (rows, columns) as raw range strings (e.g., "$1:$1",
    /// "$A:$B").
    fn parse_print_titles(value: &str) -> (Option<String>, Option<String>) {
        let mut rows: Option<String> = None;
        let mut cols: Option<String> = None;

        for part in Self::split_defined_name_areas(value) {
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
                Some(ch)
                    if ch.is_ascii_digit()
                        && rows.is_none()
                        && Self::is_valid_print_title_range(range, true) =>
                {
                    rows = Some(range.to_string());
                },
                Some(ch)
                    if ch.is_ascii_alphabetic()
                        && cols.is_none()
                        && Self::is_valid_print_title_range(range, false) =>
                {
                    cols = Some(range.to_string());
                },
                _ => {},
            }
        }

        (rows, cols)
    }

    fn split_defined_name_areas(value: &str) -> Vec<&str> {
        let bytes = value.as_bytes();
        let mut areas = Vec::new();
        let mut start = 0;
        let mut quoted = false;
        let mut index = 0;
        while index < bytes.len() {
            match bytes[index] {
                b'\'' if quoted && bytes.get(index + 1) == Some(&b'\'') => index += 1,
                b'\'' => quoted = !quoted,
                b',' if !quoted => {
                    areas.push(value[start..index].trim());
                    start = index + 1;
                },
                _ => {},
            }
            index += 1;
        }
        areas.push(value[start..].trim());
        areas
    }

    fn is_valid_print_cell_range(range: &str) -> bool {
        let mut references = range.split(':');
        let valid = references
            .by_ref()
            .take(2)
            .all(|reference| Cell::reference_to_coords(&reference.replace('$', "")).is_ok());
        valid && references.next().is_none()
    }

    fn is_valid_print_title_range(range: &str, rows: bool) -> bool {
        let mut endpoints = range.split(':');
        let mut count = 0;
        let valid = endpoints.by_ref().take(2).all(|endpoint| {
            count += 1;
            let endpoint = endpoint.trim_matches('$');
            let reference = if rows {
                format!("A{endpoint}")
            } else {
                format!("{endpoint}1")
            };
            Cell::reference_to_coords(&reference).is_ok()
        });
        valid && count == 2 && endpoints.next().is_none()
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

    /// Resolve a worksheet through the relationship declared by workbook.xml.
    pub(crate) fn worksheet_part_uri(&self, worksheet: &WorksheetInfo) -> SheetResult<PackURI> {
        use litchi_opc::constants::relationship_type as rt;

        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let relationship = workbook_part
            .rels()
            .get(&worksheet.relationship_id)
            .ok_or_else(|| {
                format!(
                    "Worksheet '{}' references missing relationship '{}'",
                    worksheet.name, worksheet.relationship_id
                )
            })?;

        if relationship.reltype() != rt::WORKSHEET && relationship.reltype() != rt::STRICT_WORKSHEET
        {
            return Err(format!(
                "Relationship '{}' for worksheet '{}' has invalid type '{}'",
                worksheet.relationship_id,
                worksheet.name,
                relationship.reltype()
            )
            .into());
        }

        if relationship.is_external() {
            return Err(format!(
                "Relationship '{}' for worksheet '{}' has an external target",
                worksheet.relationship_id, worksheet.name
            )
            .into());
        }

        Ok(relationship.target_partname()?)
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

    fn is_1904_date_system(&self) -> bool {
        self.is_1904_date_system
    }
}

impl Workbook {
    /// Open a workbook from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> SheetResult<Self> {
        let package = OpcPackage::open(path)?;
        Self::new(package)
    }

    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<std::path::Path>>(
        path: P,
        password: &str,
    ) -> SheetResult<Self> {
        let data = std::fs::read(path.as_ref())?;
        let decrypted = crate::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// ws.set_cell_value(1, 1, "Hello");
    /// ws.set_cell_value(1, 2, "World");
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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

    /// Read the person list from the workbook.
    ///
    /// Persons are used in threaded comments to identify comment authors.
    /// Returns `None` if no person list is present in the workbook.
    pub fn persons(&self) -> SheetResult<Option<crate::xlsx::PersonList>> {
        crate::xlsx::read_persons(self.package())
    }

    /// Read threaded comments for a specific worksheet by index.
    ///
    /// Threaded comments are the modern comment format in Excel with support for
    /// conversation threads, @mentions, and timestamps.
    ///
    /// # Arguments
    /// * `sheet_index` - Zero-based worksheet index
    ///
    /// Returns `None` if the worksheet has no threaded comments.
    pub fn threaded_comments(
        &self,
        sheet_index: usize,
    ) -> SheetResult<Option<crate::xlsx::ThreadedComments>> {
        let ws_info = self
            .worksheets
            .get(sheet_index)
            .ok_or("Worksheet index out of bounds")?;

        let worksheet_uri = self.worksheet_part_uri(ws_info)?;
        let comments = crate::xlsx::read_threaded_comments(self.package(), &worksheet_uri)?;
        if let Some(comments) = comments.as_ref() {
            let people = self.persons()?;
            validate_threaded_comment_people(comments.comments.iter(), people.as_ref())?;
        }
        Ok(comments)
    }

    /// Read threaded comments for a worksheet by name.
    ///
    /// # Arguments
    /// * `sheet_name` - Name of the worksheet
    ///
    /// Returns `None` if the worksheet has no threaded comments.
    pub fn threaded_comments_by_name(
        &self,
        sheet_name: &str,
    ) -> SheetResult<Option<crate::xlsx::ThreadedComments>> {
        let sheet_index = self
            .worksheets
            .iter()
            .position(|ws| ws.name == sheet_name)
            .ok_or_else(|| format!("Worksheet '{}' not found", sheet_name))?;

        self.threaded_comments(sheet_index)
    }

    /// Add a new worksheet.
    ///
    /// # Arguments
    /// * `name` - The name of the new worksheet
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.define_name("TaxRate", "Sheet1!$A$1");
    /// wb.define_name("SalesData", "Sheet1!$A$1:$D$100");
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::create()?;
    /// let props = wb.properties();
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the workbook properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.properties_mut().title = Some("My Workbook".to_string());
    /// wb.properties_mut().creator = Some("John Doe".to_string());
    /// wb.save("workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
    }

    /// Set the person list for threaded comments.
    ///
    /// Persons are used to identify authors of threaded comments.
    ///
    /// # Arguments
    /// * `person_list` - List of persons who can author comments
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, Person, PersonList};
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut person_list = PersonList::default();
    /// person_list.persons.push(Person {
    ///     display_name: "John Doe".to_string(),
    ///     id: "{11111111-2222-3333-4444-555555555555}".to_string(),
    ///     user_id: None,
    ///     provider_id: None,
    /// });
    /// wb.set_person_list(person_list);
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_person_list(&mut self, person_list: crate::xlsx::PersonList) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        if let Some(ref mut data) = self.mutable_data {
            data.person_list = Some(person_list);
        }
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut workbook = Workbook::create()?;
    /// // Modify workbook...
    /// workbook.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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

    /// Generate metadata.xml for threaded comments support.
    fn generate_metadata_xml() -> String {
        xml_minifier::minified_xml!("resources/metadata.xml").to_string()
    }

    /// Generate bridge comments.xml for threaded comments backwards compatibility.
    fn generate_bridge_comments_xml(threaded_comments: &crate::xlsx::ThreadedComments) -> String {
        use litchi_core::xml::escape::escape_xml;
        use std::collections::HashMap;

        let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xml.push_str(
            "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
        );

        // Build author list and comment grouping
        let mut authors = Vec::new();
        let mut author_map: HashMap<String, usize> = HashMap::new();
        let mut comments_by_cell: HashMap<String, Vec<&crate::xlsx::ThreadedComment>> =
            HashMap::new();

        // Group comments by cell
        for comment in &threaded_comments.comments {
            if let Some(ref cell_ref) = comment.cell_ref {
                comments_by_cell
                    .entry(cell_ref.clone())
                    .or_default()
                    .push(comment);
            }
        }

        // Build authors list with tc={id} entries
        for cell_comments in comments_by_cell.values() {
            for comment in cell_comments {
                if !author_map.contains_key(&comment.id) {
                    let author_entry = format!("tc={}", comment.id);
                    author_map.insert(comment.id.clone(), authors.len());
                    authors.push(author_entry);
                }
            }
        }

        // Write authors
        xml.push_str("<authors>");
        for author in &authors {
            xml.push_str(&format!("<author>{}</author>", escape_xml(author)));
        }
        xml.push_str("</authors>");

        // Write comment list
        xml.push_str("<commentList>");
        for (cell_ref, cell_comments) in comments_by_cell {
            if let Some(first_comment) = cell_comments.first() {
                let author_id = author_map.get(&first_comment.id).unwrap_or(&0);

                // Build combined text from all comments in thread
                let mut combined_text = String::new();
                for (i, comment) in cell_comments.iter().enumerate() {
                    if i > 0 {
                        combined_text.push_str("\nReply:\n    ");
                    } else {
                        combined_text.push_str("Comment:\n    ");
                    }
                    if let Some(ref text) = comment.text {
                        combined_text.push_str(text);
                    }
                }

                xml.push_str(&format!(
                    "<comment ref=\"{}\" authorId=\"{}\"><text><r><t>{}</t></r></text></comment>",
                    escape_xml(&cell_ref),
                    author_id,
                    escape_xml(&combined_text)
                ));
            }
        }
        xml.push_str("</commentList>");
        xml.push_str("</comments>");
        xml
    }

    /// Update workbook parts with modified data.
    fn update_workbook_parts(&mut self, data: &mut MutableWorkbookData) -> SheetResult<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::{BlobPart, Part};

        validate_workbook_tables(data)?;

        let preserved_external_relationships = {
            let workbook_part = self.package.get_part(&self.workbook_uri)?;
            self.external_links.iter().map(|link| {
                let relationship = workbook_part.rels().get(&link.relationship_id).ok_or_else(|| format!("missing preserved external-link relationship '{}'", link.relationship_id))?;
                Ok((relationship.reltype().to_string(), relationship.target_ref().to_string(), relationship.r_id().to_string(), relationship.is_external()))
            }).collect::<SheetResult<Vec<_>>>()?
        };

        let workbook_uri = PackURI::new("/xl/workbook.xml")?;

        // Create temporary workbook part to manage relationships
        let mut temp_wb_part = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.to_string(),
            Vec::new(),
        );
        for (relationship_type, target, relationship_id, external) in preserved_external_relationships {
            temp_wb_part.rels_mut().add_relationship(relationship_type, target, relationship_id, external);
        }

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

        // Write person list and metadata if present (for threaded comments)
        let has_threaded_comments = data
            .worksheets
            .iter()
            .any(|ws| !ws.threaded_comments().is_empty());
        validate_threaded_comment_people(
            data.worksheets
                .iter()
                .flat_map(|worksheet| worksheet.threaded_comments()),
            data.person_list.as_ref(),
        )?;

        if let Some(person_list) = data.person_list.as_ref()
            && !person_list.persons.is_empty()
        {
            let persons_xml = crate::xlsx::write_persons(person_list)?;
            let persons_uri = PackURI::new("/xl/persons/person.xml")?;
            let persons_part = BlobPart::new(
                persons_uri,
                ct::SML_PERSONS.to_string(),
                persons_xml.into_bytes(),
            );
            self.package.add_part(Box::new(persons_part));

            // Add relationship from workbook to persons part
            temp_wb_part.relate_to("persons/person.xml", rt::PERSONS);
        }

        // Write metadata.xml if there are threaded comments
        if has_threaded_comments {
            let metadata_xml = Self::generate_metadata_xml();
            // ...
            let metadata_uri = PackURI::new("/xl/metadata.xml")?;
            let metadata_part = BlobPart::new(
                metadata_uri,
                ct::SML_SHEET_METADATA.to_string(),
                metadata_xml.into_bytes(),
            );
            self.package.add_part(Box::new(metadata_part));

            // Create relationship from workbook to metadata
            temp_wb_part.relate_to("metadata.xml", rt::SHEET_METADATA);
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

            // Generate and add threaded comments if present
            if !ws.threaded_comments().is_empty() {
                let threaded_comments = crate::xlsx::ThreadedComments {
                    comments: ws.threaded_comments().to_vec(),
                };
                let tc_xml = crate::xlsx::write_threaded_comments(&threaded_comments)?;
                let tc_uri = PackURI::new(format!(
                    "/xl/threadedComments/threadedComment{}.xml",
                    ws.sheet_id()
                ))?;
                let tc_part = BlobPart::new(
                    tc_uri,
                    ct::SML_THREADED_COMMENTS.to_string(),
                    tc_xml.into_bytes(),
                );
                self.package.add_part(Box::new(tc_part));

                // Add relationship from worksheet to threaded comments
                ws_part.relate_to(
                    &format!("../threadedComments/threadedComment{}.xml", ws.sheet_id()),
                    rt::THREADED_COMMENTS,
                );

                // Generate bridge comments.xml for backwards compatibility
                let bridge_xml = Self::generate_bridge_comments_xml(&threaded_comments);
                let bridge_uri = PackURI::new(format!("/xl/comments{}.xml", ws.sheet_id()))?;
                let bridge_part = BlobPart::new(
                    bridge_uri,
                    ct::SML_COMMENTS.to_string(),
                    bridge_xml.into_bytes(),
                );
                self.package.add_part(Box::new(bridge_part));

                // Add relationship from worksheet to bridge comments
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

            // Generate and add table XML files if present and track relationship IDs
            let mut table_rel_ids: Vec<String> = Vec::new();
            for table in ws.tables().iter() {
                use crate::xlsx::writer::table::serialize_table;

                let table_xml = serialize_table(table)?;
                let table_uri = PackURI::new(format!("/xl/tables/table{}.xml", table.id))?;
                let table_part =
                    BlobPart::new(table_uri, ct::SML_TABLE.to_string(), table_xml.into_bytes());
                self.package.add_part(Box::new(table_part));

                // Add relationship from worksheet to table and capture the ID
                let rel_id =
                    ws_part.relate_to(&format!("../tables/table{}.xml", table.id), rt::TABLE);
                table_rel_ids.push(rel_id);
            }

            // Generate and add drawing XML for images if present
            let drawing_rel_id = if let Some(drawing_xml) = ws.generate_drawing_xml()? {
                let drawing_uri =
                    PackURI::new(format!("/xl/drawings/drawing{}.xml", ws.sheet_id()))?;

                // Create drawing part with relationships for images
                let mut drawing_part = BlobPart::new(
                    drawing_uri.clone(),
                    ct::OFC_DRAWING.to_string(),
                    drawing_xml.into_bytes(),
                );

                // Add image parts and create relationships
                for (idx, image) in ws.images().iter().enumerate() {
                    let image_ext = &image.format;
                    let image_uri = PackURI::new(format!(
                        "/xl/media/image{}_{}.{}",
                        ws.sheet_id(),
                        idx + 1,
                        image_ext
                    ))?;

                    // Determine content type based on format
                    let content_type = match image_ext.to_lowercase().as_str() {
                        "png" => "image/png",
                        "jpg" | "jpeg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        "svg" => "image/svg+xml",
                        _ => {
                            return Err(format!(
                                "Unsupported worksheet image format '{}'",
                                image.format
                            )
                            .into());
                        },
                    };

                    let image_part = BlobPart::new(
                        image_uri.clone(),
                        content_type.to_string(),
                        image.data.clone(),
                    );
                    self.package.add_part(Box::new(image_part));

                    // Add relationship from drawing to image
                    drawing_part.relate_to(
                        &format!("../media/image{}_{}.{}", ws.sheet_id(), idx + 1, image_ext),
                        rt::IMAGE,
                    );
                }

                // Add chart parts and create relationships
                for (idx, chart) in ws.charts().iter().enumerate() {
                    let chart_name = format!("chart{}_{}", ws.sheet_id(), idx + 1);
                    let chart_uri = PackURI::new(format!("/xl/charts/{chart_name}.xml"))?;
                    if chart.chart.external_data.is_some() != chart.external_data_part.is_some() {
                        return Err(format!(
                            "Worksheet chart {} external-data metadata and package payload disagree",
                            idx + 1
                        )
                        .into());
                    }
                    if chart.chart.user_shapes.is_some() != chart.user_shapes_part.is_some() {
                        return Err(format!(
                            "Worksheet chart {} user-shapes metadata and package payload disagree",
                            idx + 1
                        )
                        .into());
                    }

                    let mut chart_part =
                        BlobPart::new(chart_uri.clone(), ct::DML_CHART.to_string(), Vec::new());
                    let mut chart_related_resources = Vec::new();
                    let mut additional_relationship_ids = std::collections::HashSet::new();
                    for (relationship_index, relationship) in
                        chart.additional_relationships.iter().enumerate()
                    {
                        if relationship.relationship_id.is_empty()
                            || relationship.relationship_type.is_empty()
                            || !additional_relationship_ids
                                .insert(relationship.relationship_id.as_str())
                        {
                            return Err(format!(
                                "Worksheet chart {} has invalid additional relationship metadata",
                                idx + 1
                            )
                            .into());
                        }
                        let (target, external) = match &relationship.target {
                            crate::xlsx::ChartRelationshipTarget::Embedded {
                                data,
                                content_type,
                                extension,
                            } => {
                                if content_type.is_empty()
                                    || extension.is_empty()
                                    || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                                {
                                    return Err(format!(
                                        "Worksheet chart {} has invalid embedded related resource",
                                        idx + 1
                                    )
                                    .into());
                                }
                                let resource_name = format!(
                                    "chartResource{}_{}_{}.{}",
                                    ws.sheet_id(),
                                    idx + 1,
                                    relationship_index + 1,
                                    extension.to_ascii_lowercase()
                                );
                                let resource_uri =
                                    PackURI::new(format!("/xl/chartResources/{resource_name}"))?;
                                chart_related_resources.push(BlobPart::new(
                                    resource_uri,
                                    content_type.clone(),
                                    data.clone(),
                                ));
                                (format!("../chartResources/{resource_name}"), false)
                            },
                            crate::xlsx::ChartRelationshipTarget::External { target } => {
                                if target.is_empty() {
                                    return Err(format!(
                                        "Worksheet chart {} has an empty external related target",
                                        idx + 1
                                    )
                                    .into());
                                }
                                (target.clone(), true)
                            },
                        };
                        chart_part.rels_mut().add_relationship(
                            relationship.relationship_type.clone(),
                            target,
                            relationship.relationship_id.clone(),
                            external,
                        );
                    }
                    let mut embedded_external_part = None;
                    let external_data_relationship_id = if let Some(external_data) =
                        chart.external_data_part.as_ref()
                    {
                        if !crate::xlsx::chart::is_chart_external_data_relationship_type(
                            &external_data.relationship_type,
                        ) {
                            return Err(format!(
                                    "Worksheet chart {} has invalid external-data relationship type '{}'",
                                    idx + 1,
                                    external_data.relationship_type
                                )
                                .into());
                        }
                        let (target, external) = match &external_data.target {
                            crate::xlsx::ChartExternalDataTarget::Embedded {
                                data,
                                content_type,
                                extension,
                            } => {
                                let expected_content_type =
                                    crate::xlsx::chart::chart_external_data_content_type(
                                        &external_data.relationship_type,
                                    )
                                    .expect("relationship type was validated above");
                                if content_type.is_empty()
                                    || content_type != expected_content_type
                                    || extension.is_empty()
                                    || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                                {
                                    return Err(format!(
                                            "Worksheet chart {} has invalid embedded external-data metadata",
                                            idx + 1
                                        )
                                        .into());
                                }
                                let external_name = format!(
                                    "chartData{}_{}.{}",
                                    ws.sheet_id(),
                                    idx + 1,
                                    extension.to_ascii_lowercase()
                                );
                                let external_uri =
                                    PackURI::new(format!("/xl/embeddings/{external_name}"))?;
                                embedded_external_part = Some(BlobPart::new(
                                    external_uri,
                                    content_type.clone(),
                                    data.clone(),
                                ));
                                (format!("../embeddings/{external_name}"), false)
                            },
                            crate::xlsx::ChartExternalDataTarget::Linked { target } => {
                                if target.is_empty() {
                                    return Err(format!(
                                            "Worksheet chart {} has an empty linked external-data target",
                                            idx + 1
                                        )
                                        .into());
                                }
                                (target.clone(), true)
                            },
                        };
                        let relationship_id = if let Some(relationship_id) = chart
                            .chart
                            .external_data
                            .as_ref()
                            .and_then(|metadata| metadata.relationship_id.as_deref())
                        {
                            if relationship_id.is_empty()
                                || chart_part.rels().get(relationship_id).is_some()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has a conflicting external-data relationship ID",
                                    idx + 1
                                )
                                .into());
                            }
                            chart_part.rels_mut().add_relationship(
                                external_data.relationship_type.clone(),
                                target,
                                relationship_id.to_string(),
                                external,
                            );
                            relationship_id.to_string()
                        } else if external {
                            chart_part.relate_to_ext(&target, &external_data.relationship_type)
                        } else {
                            chart_part.relate_to(&target, &external_data.relationship_type)
                        };
                        Some(relationship_id)
                    } else {
                        None
                    };

                    let mut user_shapes_part_to_add = None;
                    let mut user_shape_resources = Vec::new();
                    let user_shapes_relationship_id = if let Some(user_shapes) =
                        chart.user_shapes_part.as_ref()
                    {
                        let referenced_ids =
                            crate::xlsx::chart::chart_user_shapes_relationship_ids(
                                &user_shapes.xml,
                            )?;
                        let declared_ids: std::collections::HashSet<&str> = user_shapes
                            .relationships
                            .iter()
                            .map(|relationship| relationship.relationship_id.as_str())
                            .collect();
                        if declared_ids.len() != user_shapes.relationships.len()
                            || referenced_ids.len() != declared_ids.len()
                            || !referenced_ids
                                .iter()
                                .all(|id| declared_ids.contains(id.as_str()))
                        {
                            return Err(format!(
                                "Worksheet chart {} user-shapes relationship declarations do not match its XML",
                                idx + 1
                            )
                            .into());
                        }

                        let user_shapes_name =
                            format!("chartDrawing{}_{}.xml", ws.sheet_id(), idx + 1);
                        let user_shapes_uri =
                            PackURI::new(format!("/xl/drawings/{user_shapes_name}"))?;
                        let mut part = BlobPart::new(
                            user_shapes_uri,
                            ct::DML_CHARTSHAPES.to_string(),
                            user_shapes.xml.clone(),
                        );
                        for (relationship_index, relationship) in
                            user_shapes.relationships.iter().enumerate()
                        {
                            if relationship.relationship_id.is_empty()
                                || relationship.relationship_type.is_empty()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has invalid user-shapes relationship metadata",
                                    idx + 1
                                )
                                .into());
                            }
                            let (target, external) = match &relationship.target {
                                crate::xlsx::ChartUserShapesRelationshipTarget::Embedded {
                                    data,
                                    content_type,
                                    extension,
                                } => {
                                    if content_type.is_empty()
                                        || extension.is_empty()
                                        || !extension
                                            .bytes()
                                            .all(|byte| byte.is_ascii_alphanumeric())
                                    {
                                        return Err(format!(
                                            "Worksheet chart {} has invalid embedded user-shapes resource",
                                            idx + 1
                                        )
                                        .into());
                                    }
                                    let resource_name = format!(
                                        "chartShape{}_{}_{}.{}",
                                        ws.sheet_id(),
                                        idx + 1,
                                        relationship_index + 1,
                                        extension.to_ascii_lowercase()
                                    );
                                    let resource_uri =
                                        PackURI::new(format!("/xl/media/{resource_name}"))?;
                                    user_shape_resources.push(BlobPart::new(
                                        resource_uri,
                                        content_type.clone(),
                                        data.clone(),
                                    ));
                                    (format!("../media/{resource_name}"), false)
                                },
                                crate::xlsx::ChartUserShapesRelationshipTarget::External {
                                    target,
                                } => {
                                    if target.is_empty() {
                                        return Err(format!(
                                            "Worksheet chart {} has an empty external user-shapes target",
                                            idx + 1
                                        )
                                        .into());
                                    }
                                    (target.clone(), true)
                                },
                            };
                            part.rels_mut().add_relationship(
                                relationship.relationship_type.clone(),
                                target,
                                relationship.relationship_id.clone(),
                                external,
                            );
                        }
                        let target = format!("../drawings/{user_shapes_name}");
                        let relationship_id = if let Some(relationship_id) = chart
                            .chart
                            .user_shapes
                            .as_ref()
                            .and_then(|metadata| metadata.relationship_id.as_deref())
                        {
                            if relationship_id.is_empty()
                                || chart_part.rels().get(relationship_id).is_some()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has a conflicting user-shapes relationship ID",
                                    idx + 1
                                )
                                .into());
                            }
                            chart_part.rels_mut().add_relationship(
                                rt::CHART_USER_SHAPES.to_string(),
                                target,
                                relationship_id.to_string(),
                                false,
                            );
                            relationship_id.to_string()
                        } else {
                            chart_part.relate_to(&target, rt::CHART_USER_SHAPES)
                        };
                        user_shapes_part_to_add = Some(part);
                        Some(relationship_id)
                    } else {
                        None
                    };

                    for relationship_id in
                        crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)?
                    {
                        if relationship_id.is_empty()
                            || chart_part.rels().get(&relationship_id).is_none()
                        {
                            return Err(format!(
                                "Worksheet chart {} fragment references missing relationship '{relationship_id}'",
                                idx + 1
                            )
                            .into());
                        }
                    }

                    let chart_xml = crate::xlsx::chart::generate_chart_xml_with_external_data_id(
                        &chart.chart,
                        external_data_relationship_id.as_deref(),
                        user_shapes_relationship_id.as_deref(),
                    )
                    .map_err(|e| format!("Failed to generate chart XML: {e}"))?;
                    chart_part.set_blob(chart_xml);
                    if let Some(external_part) = embedded_external_part {
                        self.package.add_part(Box::new(external_part));
                    }
                    for resource in chart_related_resources {
                        self.package.add_part(Box::new(resource));
                    }
                    for resource in user_shape_resources {
                        self.package.add_part(Box::new(resource));
                    }
                    if let Some(user_shapes_part) = user_shapes_part_to_add {
                        self.package.add_part(Box::new(user_shapes_part));
                    }
                    self.package.add_part(Box::new(chart_part));

                    // Add relationship from drawing to chart
                    drawing_part.relate_to(&format!("../charts/{chart_name}.xml"), rt::CHART);
                }

                self.package.add_part(Box::new(drawing_part));

                // Add relationship from worksheet to drawing
                Some(ws_part.relate_to(
                    &format!("../drawings/drawing{}.xml", ws.sheet_id()),
                    rt::DRAWING,
                ))
            } else {
                None
            };

            let mut pivot_table_rel_ids: Vec<String> = Vec::new();
            if let Some(targets) = pivot_table_targets_per_sheet.get(index) {
                for target in targets {
                    let rid = ws_part.relate_to(target, rt::PIVOT_TABLE);
                    pivot_table_rel_ids.push(rid);
                }
            }

            // Now generate worksheet XML with proper hyperlink relationship IDs and VML reference
            let ws_xml = ws.to_xml_with_part_rels(
                &mut data.shared_strings,
                &style_indices,
                crate::xlsx::writer::sheet::WorksheetPartRelationships {
                    hyperlinks: Some(&hyperlink_rel_ids),
                    vml_drawing: vml_rel_id.as_deref(),
                    pivot_tables: Some(&pivot_table_rel_ids),
                    tables: Some(&table_rel_ids),
                    drawing: drawing_rel_id.as_deref(),
                },
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
        let external_reference_ids: Vec<String> = self.external_links.iter().map(|link| link.relationship_id.clone()).collect();
        let workbook_xml = data.generate_workbook_xml_with_external_rels(
            &worksheet_rel_ids,
            &pivot_cache_rel_ids,
            &external_reference_ids,
        )?;
        temp_wb_part.set_blob(workbook_xml.into_bytes());

        // Add the workbook part to the package
        self.package.add_part(Box::new(temp_wb_part));

        Ok(())
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::part::BlobPart;

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
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::part::BlobPart;
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.hide_sheet(0)?; // Hide the first sheet
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.add_worksheet("Sheet3");
    /// wb.move_sheet(2, 0)?; // Move Sheet3 to the first position
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_sheet_visibility(0, "hidden")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.set_active_sheet(1)?; // Make Sheet2 active
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_force_formula_recalculation(true);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_calculation_mode("manual")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_tab_color(0, "FF0000")?; // Set red tab color
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.protect_workbook(Some("password123"), true, false);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
        crate::xlsx::pivot::read_pivot_tables(self.package())
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

fn validate_workbook_tables(data: &MutableWorkbookData) -> SheetResult<()> {
    use std::collections::HashSet;

    let table_count = data
        .worksheets
        .iter()
        .map(|worksheet| worksheet.tables().len())
        .sum();
    let mut ids = HashSet::with_capacity(table_count);
    let mut names = HashSet::with_capacity(table_count);
    let mut display_names = HashSet::with_capacity(table_count);
    for worksheet in &data.worksheets {
        for table in worksheet.tables() {
            if !ids.insert(table.id) {
                return Err(format!("duplicate workbook table ID {}", table.id).into());
            }
            if !names.insert(table.name.to_ascii_lowercase()) {
                return Err(format!("duplicate workbook table name '{}'", table.name).into());
            }
            if !display_names.insert(table.display_name.to_ascii_lowercase()) {
                return Err(format!(
                    "duplicate workbook table display name '{}'",
                    table.display_name
                )
                .into());
            }
        }
    }
    Ok(())
}

fn validate_threaded_comment_people<'a>(
    comments: impl IntoIterator<Item = &'a crate::xlsx::ThreadedComment>,
    person_list: Option<&crate::xlsx::PersonList>,
) -> SheetResult<()> {
    use std::collections::HashSet;

    let person_ids: HashSet<&str> = person_list
        .map(|people| {
            people
                .persons
                .iter()
                .map(|person| person.id.as_str())
                .collect()
        })
        .unwrap_or_default();
    for comment in comments {
        if !person_ids.contains(comment.person_id.as_str()) {
            return Err(format!(
                "threaded comment '{}' references missing person '{}'",
                comment.id, comment.person_id
            )
            .into());
        }
        for mention in &comment.mentions {
            if !person_ids.contains(mention.mention_person_id.as_str()) {
                return Err(format!(
                    "mention '{}' references missing person '{}'",
                    mention.mention_id, mention.mention_person_id
                )
                .into());
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{Workbook, validate_threaded_comment_people};
    use crate::charts::{ChartExtensionList, ChartShapeProperties, plot_area::TypeGroup};
    use litchi_core::sheet::{CellValue, WorkbookTrait, Worksheet as _};
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

    use crate::xlsx::{
        ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship,
        ChartRelationshipTarget, ChartUserShapesPart, ChartUserShapesRelationship,
        ChartUserShapesRelationshipTarget, Mention, Person, PersonList, Table, TableColumn,
        ThreadedComment, WorksheetChart,
    };

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sales" sheetId="42" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sales!$A$1:$D$20</definedName>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sales!$1:$2,Sales!$A:$B</definedName>
  </definedNames>
</workbook>"#;

    #[test]
    fn saves_and_reloads_worksheet_tables_and_drawings() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("tables.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut table = Table::new(7, "SalesTable", "A1:B4");
        table.columns = vec![
            TableColumn::new(1, "Region"),
            TableColumn::new(2, "Revenue"),
        ];
        table.auto_filter_range = Some("A1:B4".to_string());
        let worksheet = workbook.worksheet_mut(0).unwrap();
        worksheet.add_table(table);
        worksheet
            .add_image(
                b"\x89PNG\r\n\x1a\nfixture".to_vec(),
                "png",
                1,
                1,
                4,
                2,
                Some("Logo"),
            )
            .unwrap();
        let mut worksheet_chart = WorksheetChart::bar_chart(
            "Revenue",
            "Sheet1!$A$2:$A$4",
            "Sheet1!$B$2:$B$4",
            ChartAnchor::new(3, 1, 9, 14),
        )
        .unwrap();
        worksheet_chart.chart.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId9"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        worksheet_chart.chart.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="shared"><x:reference r:id="rId1" r:link="rId10"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        worksheet.add_chart(
            worksheet_chart
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId9".to_string(),
                relationship_type: rt::IMAGE.to_string(),
                target: ChartRelationshipTarget::Embedded {
                    data: b"chart background".to_vec(),
                    content_type: ct::PNG.to_string(),
                    extension: "png".to_string(),
                },
            })
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId10".to_string(),
                relationship_type: rt::HYPERLINK.to_string(),
                target: ChartRelationshipTarget::External {
                    target: "https://example.test/chart".to_string(),
                },
            })
            .with_external_data_part(
                ChartExternalDataPart::embedded_workbook(b"PK chart workbook".to_vec()),
                Some(false),
            )
            .with_user_shapes_part(ChartUserShapesPart {
                xml: br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>1</cdr:y></cdr:to><cdr:pic><a:blip r:embed="rId5"/></cdr:pic></cdr:relSizeAnchor></c:userShapes>"#.to_vec(),
                relationships: vec![ChartUserShapesRelationship {
                    relationship_id: "rId5".to_string(),
                    relationship_type: rt::IMAGE.to_string(),
                    target: ChartUserShapesRelationshipTarget::Embedded {
                        data: b"shape image".to_vec(),
                        content_type: ct::PNG.to_string(),
                        extension: "png".to_string(),
                    },
                }],
            }),
        );

        workbook.save(&path).unwrap();
        let package = OpcPackage::open(&path).unwrap();
        let sheet_part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        let sheet_xml = std::str::from_utf8(sheet_part.blob()).unwrap();
        assert!(sheet_xml.contains(r#"<drawing r:id="rId2"/>"#));
        let drawing_relationship = sheet_part.rels().get("rId2").unwrap();
        assert_eq!(drawing_relationship.reltype(), rt::DRAWING);
        assert_eq!(
            drawing_relationship.target_partname().unwrap().as_str(),
            "/xl/drawings/drawing1.xml"
        );
        let drawing_part = package
            .get_part(&drawing_relationship.target_partname().unwrap())
            .unwrap();
        assert_eq!(drawing_part.content_type(), ct::OFC_DRAWING);
        assert_eq!(
            drawing_part.rels().get("rId1").unwrap().reltype(),
            rt::IMAGE
        );
        assert_eq!(
            drawing_part.rels().get("rId2").unwrap().reltype(),
            rt::CHART
        );
        let chart_relationship = drawing_part.rels().get("rId2").unwrap();
        let chart_part = package
            .get_part(&chart_relationship.target_partname().unwrap())
            .unwrap();
        let external_relationship = chart_part.rels().get("rId1").unwrap();
        assert_eq!(external_relationship.reltype(), rt::PACKAGE);
        assert_eq!(
            package
                .get_part(&external_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"PK chart workbook"
        );
        let chart_xml = std::str::from_utf8(chart_part.blob()).unwrap();
        assert!(chart_xml.contains(r#"r:embed="rId9""#));
        let background_relationship = chart_part.rels().get("rId9").unwrap();
        assert_eq!(background_relationship.reltype(), rt::IMAGE);
        assert_eq!(
            package
                .get_part(&background_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"chart background"
        );
        let link_relationship = chart_part.rels().get("rId10").unwrap();
        assert_eq!(link_relationship.reltype(), rt::HYPERLINK);
        assert!(link_relationship.is_external());
        assert_eq!(link_relationship.target_ref(), "https://example.test/chart");
        assert!(chart_xml.contains(r#"<c:externalData r:id="rId1">"#));
        assert!(chart_xml.contains(r#"<c:autoUpdate val="0"/>"#));
        let user_shapes_relationship = chart_part.rels().get("rId2").unwrap();
        assert_eq!(user_shapes_relationship.reltype(), rt::CHART_USER_SHAPES);
        let user_shapes_part = package
            .get_part(&user_shapes_relationship.target_partname().unwrap())
            .unwrap();
        assert_eq!(user_shapes_part.content_type(), ct::DML_CHARTSHAPES);
        let shape_image_relationship = user_shapes_part.rels().get("rId5").unwrap();
        assert_eq!(shape_image_relationship.reltype(), rt::IMAGE);
        assert_eq!(
            package
                .get_part(&shape_image_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"shape image"
        );

        let mut reopened = Workbook::open(&path).unwrap();
        let worksheet = reopened.get_worksheet(0).unwrap();

        assert_eq!(worksheet.tables().len(), 1);
        let table = &worksheet.tables()[0];
        assert_eq!(table.id, 7);
        assert_eq!(table.name, "SalesTable");
        assert_eq!(table.ref_range, "A1:B4");
        assert_eq!(table.column_names(), vec!["Region", "Revenue"]);
        assert_eq!(worksheet.images().len(), 1);
        assert_eq!(worksheet.images()[0].format, "png");
        assert_eq!(worksheet.images()[0].position, (0, 0, 3, 1));
        assert_eq!(worksheet.images()[0].description.as_deref(), Some("Logo"));
        assert_eq!(worksheet.charts().len(), 1);
        assert_eq!(worksheet.charts()[0].anchor.from_col, 3);
        assert_eq!(worksheet.charts()[0].anchor.to_row, 14);
        assert_eq!(
            worksheet.charts()[0]
                .chart
                .external_data
                .as_ref()
                .unwrap()
                .auto_update,
            Some(false)
        );
        let ChartExternalDataTarget::Embedded {
            data,
            content_type,
            extension,
        } = &worksheet.charts()[0]
            .external_data_part
            .as_ref()
            .unwrap()
            .target
        else {
            panic!("expected embedded chart workbook");
        };
        assert_eq!(data, b"PK chart workbook");
        assert_eq!(content_type, ct::OFC_PACKAGE);
        assert_eq!(extension, "xlsx");
        let user_shapes = worksheet.charts()[0].user_shapes_part.as_ref().unwrap();
        assert_eq!(user_shapes.relationships.len(), 1);
        assert_eq!(user_shapes.relationships[0].relationship_id, "rId5");
        let ChartUserShapesRelationshipTarget::Embedded { data, .. } =
            &user_shapes.relationships[0].target
        else {
            panic!("expected embedded chart user-shape resource");
        };
        assert_eq!(data, b"shape image");
        assert_eq!(worksheet.charts()[0].additional_relationships.len(), 2);
        let background = worksheet.charts()[0]
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId9")
            .unwrap();
        let ChartRelationshipTarget::Embedded { data, .. } = &background.target else {
            panic!("expected embedded chart relationship resource");
        };
        assert_eq!(data, b"chart background");
        let link = worksheet.charts()[0]
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId10")
            .unwrap();
        let ChartRelationshipTarget::External { target } = &link.target else {
            panic!("expected external chart relationship target");
        };
        assert_eq!(target, "https://example.test/chart");
        let TypeGroup::Bar(group) = &worksheet.charts()[0].chart.plot_area.type_groups[0] else {
            panic!("expected reopened bar chart");
        };
        let series = &group.common.series[0];
        assert_eq!(
            series
                .categories
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$A$2:$A$4"
        );
        assert_eq!(
            series
                .values
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$B$2:$B$4"
        );

        let second_path = directory.path().join("tables-roundtrip.xlsx");
        reopened.save(&second_path).unwrap();
        let second_reopen = Workbook::open(&second_path).unwrap();
        assert_eq!(
            second_reopen.get_worksheet(0).unwrap().charts()[0]
                .chart
                .external_data
                .as_ref()
                .unwrap()
                .relationship_id
                .as_deref(),
            Some("rId1")
        );
    }

    #[test]
    fn rejects_dangling_chart_fragment_relationships() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("dangling-chart-relationship.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut chart = WorksheetChart::bar_chart(
            "Revenue",
            "Sheet1!$A$2:$A$4",
            "Sheet1!$B$2:$B$4",
            ChartAnchor::new(3, 1, 9, 14),
        )
        .unwrap();
        let TypeGroup::Bar(group) = &mut chart.chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        let series = &mut group.common.series[0];
        series.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId403"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut point = crate::charts::DataPoint::new(0);
        point.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="dangling"><x:reference r:id="rId404"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        series.data_points.push(point);
        chart.chart.chart_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="chart"><x:reference r:id="rId405"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId406"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.type_groups[0]
            .common_mut()
            .extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="group"><x:reference r:id="rId407"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.data_table = Some(crate::charts::DataTable {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId408"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::DataTable::default()
        });
        let TypeGroup::Bar(group) = &mut chart.chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        group.series_lines.push(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId409"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        let mut line =
            crate::charts::LineTypeGroup::new(crate::charts::types::BarGrouping::Standard);
        line.up_down_bars = Some(crate::charts::UpDownBars {
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="bars"><x:reference r:id="rId410"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::UpDownBars::default()
        });
        let mut line_series = crate::charts::Series::new(0);
        line_series.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId424"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="series-marker"><x:reference r:id="rId425"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut line_point = crate::charts::DataPoint::new(0);
        line_point.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId426"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_point.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="point-marker"><x:reference r:id="rId427"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.data_points.push(line_point);
        let mut line_labels = crate::charts::DataLabels::new();
        line_labels.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId430"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId431"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.leader_lines = Some(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId432"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        line_labels.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="labels"><x:reference r:id="rId433"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut line_label = crate::charts::DataLabel::new(0);
        line_label.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId434"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_label.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId435"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_label.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="label"><x:reference r:id="rId436"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.labels.push(line_label);
        line_series.data_labels = Some(line_labels);
        line_series.error_bars.push(crate::charts::series::ErrorBar {
            direction: crate::charts::series::ErrorBarDirection::Y,
            error_type: crate::charts::series::ErrorBarType::Both,
            value_type: crate::charts::series::ErrorBarValueType::Fixed,
            value: Some(1.0),
            plus_values: None,
            minus_values: None,
            no_end_cap: false,
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId437"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="error-bars"><x:reference r:id="rId438"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        let mut line_trendline = crate::charts::series::Trendline::linear();
        line_trendline.show_label = true;
        line_trendline.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId439"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId440"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId441"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="trendline-label"><x:reference r:id="rId442"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="trendline"><x:reference r:id="rId443"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.trendlines.push(line_trendline);
        line.common.series.push(line_series);
        chart
            .chart
            .plot_area
            .type_groups
            .push(TypeGroup::Line(line));
        let mut surface = crate::charts::plot_area::SurfaceTypeGroup::new();
        let mut band = crate::charts::plot_area::BandFormat::new(0);
        band.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId444"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        surface.band_formats = Some(vec![band]);
        chart
            .chart
            .plot_area
            .type_groups
            .push(TypeGroup::Surface(surface));
        let mut legend = crate::charts::Legend {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId411"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::Legend::default()
        };
        let mut legend_entry = crate::charts::legend::LegendEntry::new(0);
        legend_entry.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="entry"><x:reference r:id="rId412"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        legend.entries.push(legend_entry);
        chart.chart.legend = Some(legend);
        chart.chart.title_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="title"><x:reference r:id="rId413"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let axis_common = chart.chart.plot_area.axes[0].common_mut();
        axis_common.title = Some(crate::charts::TitleText::from_string("Axis"));
        axis_common.title_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId414"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        axis_common.major_gridlines = Some(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId415"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        axis_common.scaling_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="scaling"><x:reference r:id="rId416"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        axis_common.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="axis"><x:reference r:id="rId417"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let crate::charts::Axis::Value(value_axis) = &mut chart.chart.plot_area.axes[1] else {
            panic!("expected value axis");
        };
        let mut display_units =
            crate::charts::axis::DisplayUnits::built_in(crate::charts::axis::BuiltInUnit::Millions);
        display_units.show_label = true;
        display_units.label_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId418"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        display_units.label_text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId419"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        display_units.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="display-units"><x:reference r:id="rId420"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        value_axis.display_units = Some(Box::new(display_units));
        let mut pivot_format = crate::charts::PivotFormat::new(0);
        pivot_format.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId421"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId422"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="pivot"><x:reference r:id="rId423"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.marker = Some(crate::charts::Marker {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId428"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="pivot-marker"><x:reference r:id="rId429"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::Marker::default()
        });
        chart.chart.pivot_formats = Some(vec![pivot_format]);
        let fragment_ids = crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)
            .expect("relationship-bearing fragments should be valid XML");
        assert_eq!(
            fragment_ids,
            [
                "rId403".to_string(),
                "rId404".to_string(),
                "rId405".to_string(),
                "rId406".to_string(),
                "rId407".to_string(),
                "rId408".to_string(),
                "rId409".to_string(),
                "rId410".to_string(),
                "rId411".to_string(),
                "rId412".to_string(),
                "rId413".to_string(),
                "rId414".to_string(),
                "rId415".to_string(),
                "rId416".to_string(),
                "rId417".to_string(),
                "rId418".to_string(),
                "rId419".to_string(),
                "rId420".to_string(),
                "rId421".to_string(),
                "rId422".to_string(),
                "rId423".to_string(),
                "rId424".to_string(),
                "rId425".to_string(),
                "rId426".to_string(),
                "rId427".to_string(),
                "rId428".to_string(),
                "rId429".to_string(),
                "rId430".to_string(),
                "rId431".to_string(),
                "rId432".to_string(),
                "rId433".to_string(),
                "rId434".to_string(),
                "rId435".to_string(),
                "rId436".to_string(),
                "rId437".to_string(),
                "rId438".to_string(),
                "rId439".to_string(),
                "rId440".to_string(),
                "rId441".to_string(),
                "rId442".to_string(),
                "rId443".to_string(),
                "rId444".to_string(),
            ]
            .into()
        );
        workbook.worksheet_mut(0).unwrap().add_chart(chart);

        let error = workbook.save(&path).unwrap_err().to_string();
        assert!(error.contains("fragment references missing relationship"));
        assert!(
            [
                "rId403", "rId404", "rId405", "rId406", "rId407", "rId408", "rId409", "rId410",
                "rId411", "rId412", "rId413", "rId414", "rId415", "rId416", "rId417", "rId418",
                "rId419", "rId420", "rId421", "rId422", "rId423", "rId424", "rId425", "rId426",
                "rId427", "rId428", "rId429", "rId430", "rId431", "rId432", "rId433", "rId434",
                "rId435", "rId436", "rId437", "rId438", "rId439", "rId440", "rId441", "rId442",
                "rId443", "rId444",
            ]
            .iter()
            .any(|relationship_id| error.contains(relationship_id))
        );
    }

    #[test]
    fn rejects_duplicate_table_identity_across_worksheets() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("duplicate-tables.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut first = Table::new(1, "Sales", "A1:A2");
        first.columns.push(TableColumn::new(1, "Value"));
        workbook.worksheet_mut(0).unwrap().add_table(first);

        let worksheet = workbook.add_worksheet("Other");
        let mut duplicate = Table::new(1, "OtherTable", "A1:A2");
        duplicate.columns.push(TableColumn::new(1, "Value"));
        worksheet.add_table(duplicate);

        let error = workbook.save(path).unwrap_err().to_string();
        assert!(error.contains("duplicate workbook table ID 1"));
    }

    #[test]
    fn validates_threaded_comment_people_references() {
        let person_id = "{11111111-1111-1111-1111-111111111111}";
        let missing_id = "{22222222-2222-2222-2222-222222222222}";
        let people = PersonList {
            persons: vec![Person {
                id: person_id.into(),
                ..Default::default()
            }],
        };
        let comment = ThreadedComment {
            id: "{33333333-3333-3333-3333-333333333333}".into(),
            person_id: person_id.into(),
            mentions: vec![Mention {
                mention_id: "{44444444-4444-4444-4444-444444444444}".into(),
                mention_person_id: person_id.into(),
                ..Default::default()
            }],
            ..Default::default()
        };

        validate_threaded_comment_people([&comment], Some(&people)).unwrap();
        assert!(validate_threaded_comment_people([&comment], None).is_err());

        let mut invalid = comment;
        invalid.mentions[0].mention_person_id = missing_id.into();
        assert!(validate_threaded_comment_people([&invalid], Some(&people)).is_err());
    }

    #[test]
    fn parses_quoted_and_bounded_print_defined_names() {
        assert_eq!(
            Workbook::parse_print_area("'Sales, West'!$A$1:$D$20,'Other'!$A$1"),
            Some("$A$1:$D$20".to_string())
        );
        assert_eq!(
            Workbook::parse_print_titles("'O''Brien, West'!$1:$2,'O''Brien, West'!$A:$B"),
            (Some("$1:$2".to_string()), Some("$A:$B".to_string()))
        );
        assert_eq!(Workbook::parse_print_area("Sales!XFE1"), None);
        assert_eq!(
            Workbook::parse_print_titles("Sales!$0:$1,Sales!$A:$XFE"),
            (None, None)
        );
    }

    const WORKSHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0" showGridLines="1"/>
    <sheetView workbookViewId="2" view="pageBreakPreview" topLeftCell="B2"
               zoomScale="90" zoomScaleNormal="110" showOutlineSymbols="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="B3" sqref="B3:C4"/>
    </sheetView>
  </sheetViews>
  <cols><col min="2" max="3" width="12.5" hidden="1" customWidth="1"/></cols>
  <sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>
  <autoFilter ref="A1:E10">
    <sortState ref="A2:E10" caseSensitive="1" sortMethod="stroke">
      <sortCondition ref="E2:E10" descending="true" sortBy="value"/>
    </sortState>
  </autoFilter>
  <mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>
  <hyperlinks>
    <hyperlink ref="B2:C3" r:id="rId1" location="Section 1"
               display="Example &amp; Co" tooltip="Open example"/>
    <hyperlink ref="D4" location="&apos;Other Sheet&apos;!A1"/>
  </hyperlinks>
  <dataValidations count="1">
    <dataValidation sqref="E1:E2" type="whole" operator="between" allowBlank="1"
                    showErrorMessage="true" errorTitle="Out of range">
      <formula1>1</formula1><formula2>10</formula2>
    </dataValidation>
  </dataValidations>
  <conditionalFormatting sqref="E1:E2">
    <cfRule type="expression" priority="1"><formula>E1&gt;0</formula></cfRule>
  </conditionalFormatting>
  <pageSetup paperSize="9" orientation="landscape" scale="110"
             fitToWidth="1" fitToHeight="2"/>
  <rowBreaks count="1" manualBreakCount="1">
    <brk id="20" min="0" max="16383" man="1" pt="true"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="0"><brk id="3"/></colBreaks>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
         xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="stacked" dateAxis="1">
          <x14:colorSeries rgb="FF123456"/>
          <x14:sparklines>
            <x14:sparkline><xm:f>Sales!A1:A3</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline>
            <x14:sparkline><xm:f>Sales!B1:B3</xm:f><xm:sqref>G2</xm:sqref></x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;

    const COMMENTS_XML: &str = r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice &amp; Bob</author></authors>
  <commentList>
    <comment ref="C3" authorId="0" guid="{comment-guid}" shapeId="5">
      <text><r><t>Hello </t></r><r><t>world</t></r></text>
    </comment>
  </commentList>
</comments>"#;

    fn package_with_worksheet_relationship(reltype: &str, external: bool) -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            WORKBOOK_XML.as_bytes().to_vec(),
        );

        let relationship_id = if external {
            workbook_part.relate_to_ext("https://example.invalid/sheet.xml", reltype)
        } else {
            workbook_part.relate_to("custom/sales-data.xml", reltype)
        };
        assert_eq!(relationship_id, "rId1");
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));

        if !external {
            let mut worksheet_part = BlobPart::new(
                PackURI::new("/xl/custom/sales-data.xml").unwrap(),
                ct::SML_WORKSHEET.to_string(),
                WORKSHEET_XML.as_bytes().to_vec(),
            );
            let hyperlink_reltype = if reltype == rt::STRICT_WORKSHEET {
                rt::STRICT_HYPERLINK
            } else {
                rt::HYPERLINK
            };
            assert_eq!(
                worksheet_part.relate_to_ext("https://example.com/report", hyperlink_reltype),
                "rId1"
            );
            let comments_reltype = if reltype == rt::STRICT_WORKSHEET {
                rt::STRICT_COMMENTS
            } else {
                rt::COMMENTS
            };
            assert_eq!(
                worksheet_part.relate_to("../comments/custom-comments.xml", comments_reltype),
                "rId2"
            );
            package.add_part(Box::new(worksheet_part));
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/comments/custom-comments.xml").unwrap(),
                ct::SML_COMMENTS.to_string(),
                COMMENTS_XML.as_bytes().to_vec(),
            )));
        }

        package
    }

    fn package_with_custom_workbook_parts() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book/main.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                    <sheets><sheet name="Custom" sheetId="1" r:id="rId1"/></sheets>
                </workbook>"#
                .to_vec(),
        );
        assert_eq!(
            workbook_part.relate_to("../sheets/data.xml", rt::STRICT_WORKSHEET),
            "rId1"
        );
        assert_eq!(
            workbook_part.relate_to("../assets/strings.xml", rt::STRICT_SHARED_STRINGS),
            "rId2"
        );
        assert_eq!(
            workbook_part.relate_to("../assets/styles.xml", rt::STRICT_STYLES),
            "rId3"
        );
        package.relate_to("custom/book/main.xml", rt::STRICT_OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/sheets/data.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
                    <sheetData><row r="1"><c r="A1" t="s" s="0"><v>0</v></c></row></sheetData>
                </worksheet>"#
                .to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/strings.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_string(),
            br#"<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="1" uniqueCount="1">
                    <si><t>Relationship resolved</t></si>
                </sst>"#
                .to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/styles.xml").unwrap(),
            ct::SML_STYLES.to_string(),
            crate::xlsx::template::default_styles_xml()
                .as_bytes()
                .to_vec(),
        )));
        package
    }

    fn replace_hyperlink_relationship(package: &mut OpcPackage, reltype: &str, external: bool) {
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let relationships = package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut();
        relationships.remove("rId1").unwrap();
        relationships.add_relationship(
                reltype.to_string(),
                "https://example.com/replaced".to_string(),
                "rId1".to_string(),
                external,
            );
    }

    #[test]
    fn loads_worksheet_from_relationship_target() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::WORKSHEET, false)).unwrap();
        let worksheet = workbook.get_worksheet(0).unwrap();

        assert_eq!(
            worksheet.cell_value(1, 1).unwrap().as_ref(),
            &CellValue::Int(7)
        );
        assert_eq!(worksheet.get_column_width(2), Some(12.5));
        assert!(worksheet.is_column_hidden(3));
        assert_eq!(worksheet.get_merged_regions(), &[(2, 2, 3, 3)]);
        let external_link = worksheet.get_hyperlink(3, 3).unwrap();
        assert_eq!(external_link.target, "https://example.com/report");
        assert_eq!(external_link.location.as_deref(), Some("Section 1"));
        assert_eq!(external_link.display.as_deref(), Some("Example & Co"));
        assert_eq!(external_link.tooltip.as_deref(), Some("Open example"));
        assert_eq!(
            worksheet.get_hyperlink(4, 4).unwrap().target,
            "'Other Sheet'!A1"
        );
        let validation = &worksheet.get_data_validations()[0];
        assert_eq!(validation.range, "E1:E2");
        assert_eq!(validation.operator.as_deref(), Some("between"));
        assert_eq!(validation.formula.as_deref(), Some("1"));
        assert_eq!(validation.formula2.as_deref(), Some("10"));
        assert!(validation.allow_blank);
        assert!(validation.show_error_message);
        assert_eq!(validation.error_title.as_deref(), Some("Out of range"));
        assert_eq!(worksheet.get_conditional_formatting().len(), 1);
        assert_eq!(
            worksheet.get_conditional_formatting()[0].rule_type,
            "expression"
        );
        assert_eq!(worksheet.get_page_setup().paper_size, Some(9));
        assert!(worksheet.get_page_setup().landscape);
        assert_eq!(worksheet.get_page_setup().scale, Some(110));
        assert_eq!(worksheet.row_breaks().len(), 1);
        assert!(worksheet.row_breaks()[0].manual);
        assert!(worksheet.row_breaks()[0].pivot);
        assert_eq!(worksheet.col_breaks()[0].max, 1_048_575);
        assert!(!worksheet.col_breaks()[0].manual);
        let filter = worksheet.get_auto_filter().unwrap();
        assert_eq!(filter.range.as_deref(), Some("A1:E10"));
        let sort = filter.sort_state.as_ref().unwrap();
        assert_eq!(sort.ref_range, "A2:E10");
        assert_eq!(sort.case_sensitive, Some(true));
        assert_eq!(sort.sort_method, Some(crate::xlsx::SortMethod::Stroke));
        assert_eq!(sort.conditions.len(), 1);
        assert_eq!(sort.conditions[0].descending, Some(true));
        assert_eq!(worksheet.sheet_views().len(), 2);
        let view = worksheet.sheet_view().unwrap();
        assert_eq!(view.workbook_view_id, Some(2));
        assert_eq!(
            view.view_type,
            Some(crate::xlsx::SheetViewType::PageBreakPreview)
        );
        assert_eq!(view.top_left_cell.as_deref(), Some("B2"));
        assert_eq!(view.zoom_scale, Some(90));
        assert_eq!(view.zoom_scale_normal, Some(110));
        assert_eq!(view.show_outline_symbols, Some(false));
        assert_eq!(view.pane.as_ref().unwrap().x_split, Some(1.0));
        assert_eq!(
            view.pane.as_ref().unwrap().state,
            Some(crate::xlsx::SheetPaneState::Frozen)
        );
        assert_eq!(view.selections[0].sqref.as_deref(), Some("B3:C4"));
        let sparkline = &worksheet.sparkline_groups()[0];
        assert_eq!(
            sparkline.sparkline_type,
            crate::xlsx::SparklineType::WinLoss
        );
        assert!(sparkline.options.date_axis);
        assert_eq!(sparkline.sparklines.len(), 2);
        assert_eq!(sparkline.sparklines[0].location, "F2");
        assert_eq!(worksheet.get_print_area(), Some("$A$1:$D$20"));
        assert_eq!(worksheet.get_repeating_rows(), Some("$1:$2"));
        assert_eq!(worksheet.get_repeating_columns(), Some("$A:$B"));
        let comment = worksheet.get_cell_comment(3, 3).unwrap();
        assert_eq!(comment.author.as_deref(), Some("Alice & Bob"));
        assert_eq!(comment.author_id, 0);
        assert_eq!(comment.text, "Hello world");
        assert_eq!(comment.guid.as_deref(), Some("{comment-guid}"));
        assert_eq!(comment.shape_id, Some(5));
    }

    #[test]
    fn resolves_strict_custom_workbook_related_parts() {
        let workbook = Workbook::new(package_with_custom_workbook_parts()).unwrap();
        let worksheet = workbook.get_worksheet(0).unwrap();

        assert_eq!(worksheet.name(), "Custom");
        assert_eq!(
            worksheet.cell_value(1, 1).unwrap().as_str(),
            Some("Relationship resolved")
        );
    }

    #[test]
    fn rejects_invalid_workbook_related_parts() {
        let mut package = package_with_custom_workbook_parts();
        let workbook_uri = PackURI::new("/custom/book/main.xml").unwrap();
        package
            .get_part_mut(&workbook_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::STRICT_SHARED_STRINGS.to_string(),
                "../assets/other-strings.xml".to_string(),
                "rId4".to_string(),
                false,
            );
        assert!(Workbook::new(package).is_err());

        let mut package = package_with_custom_workbook_parts();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/styles.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_string(),
            Vec::new(),
        )));
        assert!(Workbook::new(package).is_err());
    }

    #[test]
    fn accepts_strict_worksheet_relationship_type() {
        let workbook = Workbook::new(package_with_worksheet_relationship(
            rt::STRICT_WORKSHEET,
            false,
        ))
        .unwrap();

        assert_eq!(workbook.worksheet_by_index(0).unwrap().row_count(), 1);
    }

    #[test]
    fn rejects_external_comments_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let relationships = package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut();
        relationships.remove("rId2").unwrap();
        relationships.add_relationship(
                rt::COMMENTS.to_string(),
                "https://example.com/comments.xml".to_string(),
                "rId2".to_string(),
                true,
            );
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_duplicate_comments_relationships() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::COMMENTS.to_string(),
                "../comments/other-comments.xml".to_string(),
                "rId3".to_string(),
                false,
            );
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_wrong_comments_content_type() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/comments/custom-comments.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            COMMENTS_XML.as_bytes().to_vec(),
        )));
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_non_worksheet_relationship_type() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::CHART, false)).unwrap();
        let error = workbook
            .worksheet_by_index(0)
            .err()
            .expect("non-worksheet relationship must fail")
            .to_string();

        assert!(error.contains("invalid type"), "unexpected error: {error}");
    }

    #[test]
    fn rejects_external_worksheet_target() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::WORKSHEET, true)).unwrap();
        let error = workbook
            .worksheet_by_index(0)
            .err()
            .expect("external worksheet relationship must fail")
            .to_string();

        assert!(
            error.contains("external target"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn rejects_missing_hyperlink_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let worksheet_part = package.get_part_mut(&worksheet_uri).unwrap();
        let xml = std::str::from_utf8(worksheet_part.blob())
            .unwrap()
            .replace("r:id=\"rId1\"", "r:id=\"missingLink\"");
        worksheet_part.set_blob(xml.into_bytes());
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("missing hyperlink relationship must fail")
            .to_string();

        assert!(
            error.contains("missing relationship 'missingLink'"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn rejects_invalid_hyperlink_relationship_type() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        replace_hyperlink_relationship(&mut package, rt::CHART, true);
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("non-hyperlink relationship must fail")
            .to_string();

        assert!(error.contains("invalid type"), "unexpected error: {error}");
    }

    #[test]
    fn rejects_internal_hyperlink_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        replace_hyperlink_relationship(&mut package, rt::HYPERLINK, false);
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("internal hyperlink relationship must fail")
            .to_string();

        assert!(error.contains("not external"), "unexpected error: {error}");
    }
}
