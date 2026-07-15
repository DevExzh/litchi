//! Main Spreadsheet structure and implementation.

use super::{
    CalculationSettings, Consolidation, ContentValidation, DatabaseRange, DdeLink, LabelRange,
    NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, Sheet, SheetProtection,
    SpreadsheetProtection,
    calculation::parse_calculation_settings,
    consolidation::parse_consolidation,
    data_validation::parse_content_validations,
    database_range::parse_database_ranges,
    dde::parse_dde_links,
    label_range::parse_label_ranges,
    parser::OdsParser,
    protection::parse_protection,
    style_protection::{CellStyleProtection, CellStyleRegistry},
};
use crate::core::{Content, Meta, OwnedPackage, Styles};
use litchi_core::{Error, Metadata, Result};
use std::path::Path;

/// An OpenDocument spreadsheet (.ods).
///
/// This struct represents a complete ODS spreadsheet and provides methods to access
/// its sheets, cells, and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::Spreadsheet;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut spreadsheet = Spreadsheet::open("data.ods")?;
///
/// // Get sheet count
/// println!("Sheets: {}", spreadsheet.sheet_count()?);
///
/// // Access first sheet
/// if let Some(sheet) = spreadsheet.sheet_by_index(0)? {
///     println!("Sheet: {}", sheet.name()?);
///     println!("Rows: {}, Columns: {}", sheet.row_count()?, sheet.column_count()?);
/// }
///
/// // Export to CSV
/// let csv = spreadsheet.to_csv()?;
/// # Ok(())
/// # }
/// ```
pub struct Spreadsheet {
    package: OwnedPackage,
    #[allow(dead_code)]
    content: Content,
    #[allow(dead_code)]
    styles: Option<Styles>,
    meta: Option<Meta>,
    named_definitions: Vec<NamedDefinition>,
    content_validations: Vec<ContentValidation>,
    database_ranges: Vec<DatabaseRange>,
    calculation_settings: Option<CalculationSettings>,
    label_ranges: Vec<LabelRange>,
    consolidation: Option<Consolidation>,
    dde_links: Vec<DdeLink>,
    protection: SpreadsheetProtection,
    sheet_protections: Vec<SheetProtection>,
    cell_styles: CellStyleRegistry,
}

impl Spreadsheet {
    pub(crate) fn into_package(self) -> OwnedPackage {
        self.package
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    /// Open an ODS spreadsheet from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ods file
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Create a Spreadsheet from a byte buffer.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODS file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("data.ods")?;
    /// let spreadsheet = Spreadsheet::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned_package = OwnedPackage::from_bytes(bytes)?;
        let package = owned_package.package()?;

        // Verify this is a spreadsheet
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.spreadsheet") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODS file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;
        let named_definitions = OdsParser::parse_named_definitions(content.xml_content())?;
        let content_validations = parse_content_validations(content.xml_content())?;
        let database_ranges = parse_database_ranges(content.xml_content())?;
        let calculation_settings = parse_calculation_settings(content.xml_content())?;
        let label_ranges = parse_label_ranges(content.xml_content())?;
        let consolidation = parse_consolidation(content.xml_content())?;
        let dde_links = parse_dde_links(content.xml_content())?;
        let (protection, sheet_protections) = parse_protection(content.xml_content())?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };
        let cell_styles = CellStyleRegistry::parse(
            styles.as_ref().map(Styles::xml_content),
            content.xml_content(),
        )?;

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package: owned_package,
            content,
            styles,
            meta,
            named_definitions,
            content_validations,
            database_ranges,
            calculation_settings,
            label_ranges,
            consolidation,
            dde_links,
            protection,
            sheet_protections,
            cell_styles,
        })
    }

    /// Return spreadsheet-wide formula calculation settings.
    pub fn calculation_settings(&self) -> Option<&CalculationSettings> {
        self.calculation_settings.as_ref()
    }

    /// Return spreadsheet row and column label ranges in document order.
    pub fn label_ranges(&self) -> &[LabelRange] {
        &self.label_ranges
    }

    /// Return the inert spreadsheet consolidation declaration.
    pub fn consolidation(&self) -> Option<&Consolidation> {
        self.consolidation.as_ref()
    }

    /// Return inert DDE declarations and their document-stored cached tables.
    pub fn dde_links(&self) -> &[DdeLink] {
        &self.dde_links
    }

    /// Create an ODS spreadsheet from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Get the number of sheets in the spreadsheet.
    pub fn sheet_count(&mut self) -> Result<usize> {
        let sheets = self.sheets()?;
        Ok(sheets.len())
    }

    /// Get all sheets in the spreadsheet.
    ///
    /// Returns a vector of `Sheet` objects representing all sheets in the document.
    pub fn sheets(&mut self) -> Result<Vec<Sheet>> {
        let package = self.package.package()?;
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let mut sheets = OdsParser::parse_sheets(content.xml_content())?;
        if sheets.len() != self.sheet_protections.len() {
            return Err(Error::InvalidFormat(format!(
                "sheet protection count {} does not match sheet count {}",
                self.sheet_protections.len(),
                sheets.len()
            )));
        }
        for (sheet, protection) in sheets.iter_mut().zip(&self.sheet_protections) {
            sheet.protection = protection.clone();
        }
        for cell in sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
        {
            if let Some(name) = cell.validation_name.as_deref()
                && self.content_validation(name).is_none()
            {
                return Err(Error::InvalidFormat(format!(
                    "cell references missing content validation '{name}'"
                )));
            }
        }
        Ok(sheets)
    }

    /// Return all named ranges and expressions in document order.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Return document-level spreadsheet content validations in document order.
    pub fn content_validations(&self) -> &[ContentValidation] {
        &self.content_validations
    }

    /// Return spreadsheet database ranges, filters, sort keys, and subtotal rules.
    ///
    /// External database sources are inert metadata and are never executed.
    pub fn database_ranges(&self) -> &[DatabaseRange] {
        &self.database_ranges
    }

    /// Find a content-validation definition by name.
    pub fn content_validation(&self, name: &str) -> Option<&ContentValidation> {
        self.content_validations
            .iter()
            .find(|validation| validation.name == name)
    }

    /// Return document-structure protection metadata.
    pub fn protection(&self) -> &SpreadsheetProtection {
        &self.protection
    }

    /// Resolve the inherited `style:cell-protect` value for a cell.
    pub fn cell_style_protection(&self, cell: &super::Cell) -> Result<Option<CellStyleProtection>> {
        self.cell_styles.resolve(cell.style_name())
    }

    /// Return all named ranges, including global and sheet-local ranges.
    pub fn named_ranges(&self) -> impl Iterator<Item = &NamedRange> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Range(range) => Some(range),
                NamedDefinition::Expression(_) => None,
            })
    }

    /// Find a named range by name and scope.
    pub fn named_range(&self, name: &str, scope: &NamedDefinitionScope) -> Option<&NamedRange> {
        self.named_definitions
            .iter()
            .find_map(|definition| match definition {
                NamedDefinition::Range(range) if range.name == name && &range.scope == scope => {
                    Some(range)
                },
                _ => None,
            })
    }

    /// Return all named expressions, including global and sheet-local expressions.
    pub fn named_expressions(&self) -> impl Iterator<Item = &NamedExpression> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Expression(expression) => Some(expression),
                NamedDefinition::Range(_) => None,
            })
    }

    /// Find a named expression by name and scope.
    pub fn named_expression(
        &self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<&NamedExpression> {
        self.named_definitions
            .iter()
            .find_map(|definition| match definition {
                NamedDefinition::Expression(expression)
                    if expression.name == name && &expression.scope == scope =>
                {
                    Some(expression)
                },
                _ => None,
            })
    }

    /// Get a sheet by name.
    ///
    /// Returns `Some(sheet)` if a sheet with the given name exists, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the sheet to find
    pub fn sheet_by_name(&mut self, name: &str) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().find(|sheet| sheet.name == name))
    }

    /// Get a sheet by index.
    ///
    /// Returns `Some(sheet)` if a sheet exists at the given index, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `index` - 0-based index of the sheet
    pub fn sheet_by_index(&mut self, index: usize) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().nth(index))
    }

    /// Extract all text content from the spreadsheet.
    ///
    /// Returns text from all cells, separated by newlines.
    pub fn text(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut all_text = Vec::new();

        for sheet in sheets {
            for row in sheet.rows {
                for cell in row.cells {
                    if !cell.text.trim().is_empty() {
                        all_text.push(cell.text.trim().to_string());
                    }
                }
            }
        }

        Ok(all_text.join("\n"))
    }

    /// Export spreadsheet data as CSV.
    ///
    /// Converts all sheets to CSV format, with sheets separated by double newlines.
    /// Properly escapes CSV special characters.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let csv = spreadsheet.to_csv()?;
    /// std::fs::write("output.csv", csv)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_csv(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut csv_output = String::new();

        for (sheet_index, sheet) in sheets.iter().enumerate() {
            if sheet_index > 0 {
                csv_output.push_str("\n\n"); // Separate sheets with double newline
            }

            for (row_index, row) in sheet.rows.iter().enumerate() {
                if row_index > 0 {
                    csv_output.push('\n');
                }

                for (col_index, cell) in row.cells.iter().enumerate() {
                    if col_index > 0 {
                        csv_output.push(',');
                    }

                    // Escape CSV special characters and wrap in quotes if needed
                    let cell_text = &cell.text;
                    if cell_text.contains(',')
                        || cell_text.contains('"')
                        || cell_text.contains('\n')
                    {
                        let escaped = cell_text.replace('"', "\"\"");
                        csv_output.push('"');
                        csv_output.push_str(&escaped);
                        csv_output.push('"');
                    } else {
                        csv_output.push_str(cell_text);
                    }
                }
            }
        }

        Ok(csv_output)
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }

    // Note: For spreadsheet modification operations, see `MutableSpreadsheet` which provides
    // full CRUD operations on sheets, rows, and cells including set_cell, clear_cell, add/remove
    // rows and sheets.

    /// Save the spreadsheet to a new file.
    ///
    /// This method saves the current spreadsheet state to a new file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODS file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("input.ods")?;
    /// spreadsheet.save("output.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full spreadsheet modification support is planned for future releases. For now,
    /// to modify a spreadsheet, use `SpreadsheetBuilder` to create a new one with
    /// the desired content.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the spreadsheet to bytes.
    ///
    /// This method serializes the spreadsheet to an ODF-compliant ZIP archive.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let bytes = spreadsheet.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.package.as_bytes().to_vec())
    }

    // Note: DELETE operations are available via `MutableSpreadsheet`. To modify this spreadsheet:
    //   1. Convert: `let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?`
    //   2. Modify: `mutable.remove_sheet(0)?`, `mutable.set_cell(0, 0, 0, value)?`, etc.
    //   3. Save: `mutable.save("output.ods")?`
    // Available methods: remove_sheet, remove_row, set_cell, clear_cell, clear_sheet, etc.
}
