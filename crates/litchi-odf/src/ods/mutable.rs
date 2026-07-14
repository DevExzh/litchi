//! Mutable spreadsheet structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODS spreadsheets that allows
//! for in-place modification of sheets, rows, and cells.

use crate::core::{OdfStructure, OwnedPackage, PackageWriter};
use crate::ods::{
    Cell, CellAnnotation, CellRangeSource, CellValue, Column, ContentValidation, DatabaseRange,
    NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, Row, Sheet,
    SheetPrintSettings, SheetScenario, SheetStyle, SheetTableSource, Spreadsheet,
    SpreadsheetProtection, TableStructure, TableVisibility,
    cell::{merge_cell_range, unmerge_cell_range},
    data_validation::{validate_collection, write_content_validations},
    database_range::write_database_ranges,
    named_expression::{ensure_unique, write_named_definitions},
    protection::{
        has_extensions as has_protection_extensions, write_sheet_attributes, write_sheet_options,
        write_spreadsheet_attributes,
    },
    scenario::{validate_scenario, write_sheet_preamble},
    source::validate_table_source,
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, TableStructureAxis,
        validate_sheet_print_settings, validate_table_structure, write_columns,
        write_row_attributes, write_sheet_formatting_attributes, write_table_structure,
    },
    style_protection::{PreservedXmlFragment, extract_automatic_styles, extract_font_face_decls},
};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::path::Path;

/// A mutable ODS spreadsheet that supports in-place modifications.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Spreadsheet, MutableSpreadsheet};
///
/// # fn main() -> litchi_core::Result<()> {
/// let spreadsheet = Spreadsheet::open("input.ods")?;
/// let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?;
///
/// // Modify the spreadsheet
/// mutable.add_sheet("NewSheet")?;
/// mutable.save("output.ods")?;
/// # Ok(())
/// # }
/// ```
pub struct MutableSpreadsheet {
    /// Mutable sheets
    sheets: Vec<Sheet>,
    /// Document metadata
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML
    styles_xml: Option<String>,
    /// Original content automatic styles, including unsupported style properties.
    automatic_styles: Option<PreservedXmlFragment>,
    /// Original content font-face declarations referenced by preserved styles.
    font_face_decls: Option<PreservedXmlFragment>,
    /// Global and sheet-local named ranges and expressions.
    named_definitions: Vec<NamedDefinition>,
    content_validations: Vec<ContentValidation>,
    database_ranges: Vec<DatabaseRange>,
    protection: SpreadsheetProtection,
    /// Original package retained for copying auxiliary package parts.
    source_package: Option<OwnedPackage>,
}

impl MutableSpreadsheet {
    fn sheet_max_cols(sheet: &Sheet) -> usize {
        sheet.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    }

    fn has_formulas(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|s| s.rows.iter())
            .flat_map(|r| r.cells.iter())
            .any(|c| c.formula.is_some())
            || self
                .named_definitions
                .iter()
                .any(|definition| matches!(definition, NamedDefinition::Expression(_)))
    }

    fn has_annotations(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .any(Cell::has_annotation)
    }

    fn push_table_columns(out: &mut String, max_cols: usize) {
        if max_cols <= 1 {
            out.push_str("<table:table-column/>");
        } else {
            out.push_str(&format!(
                r#"<table:table-column table:number-columns-repeated="{}"/>"#,
                max_cols
            ));
        }
    }

    fn push_cell(out: &mut String, cell: &Cell) {
        super::cell::write_cell_xml(out, cell);
    }

    fn push_row(out: &mut String, row: &Row) {
        out.push_str("<table:table-row");
        write_row_attributes(
            out,
            row.style_name.as_deref(),
            row.default_cell_style_name.as_deref(),
            row.visibility,
        );
        out.push('>');
        for cell in &row.cells {
            Self::push_cell(out, cell);
        }
        out.push_str("</table:table-row>");
    }

    /// Create a mutable spreadsheet from an existing Spreadsheet.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{Spreadsheet, MutableSpreadsheet};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?;
    /// mutable.add_sheet("NewSheet")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_spreadsheet(mut spreadsheet: Spreadsheet) -> Result<Self> {
        let styles_xml = spreadsheet.styles_xml().map(str::to_owned);
        let automatic_styles = extract_automatic_styles(spreadsheet.content_xml())?;
        let font_face_decls = extract_font_face_decls(spreadsheet.content_xml())?;
        let sheets = spreadsheet.sheets()?;
        let metadata = spreadsheet.metadata()?;
        let named_definitions = spreadsheet.named_definitions().to_vec();
        let content_validations = spreadsheet.content_validations().to_vec();
        let database_ranges = spreadsheet.database_ranges().to_vec();
        let protection = spreadsheet.protection().clone();
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet".to_string();
        let source_package = Some(spreadsheet.into_package());

        Ok(Self {
            sheets,
            metadata,
            mimetype,
            styles_xml,
            automatic_styles,
            font_face_decls,
            named_definitions,
            content_validations,
            database_ranges,
            protection,
            source_package,
        })
    }

    /// Create a new empty mutable spreadsheet.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.spreadsheet".to_string(),
            styles_xml: None,
            automatic_styles: None,
            font_face_decls: None,
            named_definitions: Vec::new(),
            content_validations: Vec::new(),
            database_ranges: Vec::new(),
            protection: SpreadsheetProtection::default(),
            source_package: None,
        }
    }

    /// Get all sheets.
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Get mutable reference to sheets.
    pub fn sheets_mut(&mut self) -> &mut Vec<Sheet> {
        &mut self.sheets
    }

    /// Get metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get mutable reference to metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Return all global and sheet-local named definitions.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Return document-level content validation definitions.
    pub fn content_validations(&self) -> &[ContentValidation] {
        &self.content_validations
    }

    /// Return database ranges and their filter/sort metadata.
    pub fn database_ranges(&self) -> &[DatabaseRange] {
        &self.database_ranges
    }

    /// Mutably access database ranges.
    pub fn database_ranges_mut(&mut self) -> &mut Vec<DatabaseRange> {
        &mut self.database_ranges
    }

    /// Add a validated database range.
    pub fn add_database_range(&mut self, range: DatabaseRange) -> Result<()> {
        range.validate()?;
        self.database_ranges.push(range);
        Ok(())
    }

    /// Remove a database range by index.
    pub fn remove_database_range(&mut self, index: usize) -> Option<DatabaseRange> {
        (index < self.database_ranges.len()).then(|| self.database_ranges.remove(index))
    }

    /// Return document-structure protection metadata.
    pub fn protection(&self) -> &SpreadsheetProtection {
        &self.protection
    }

    /// Mutably access document-structure protection metadata.
    pub fn protection_mut(&mut self) -> &mut SpreadsheetProtection {
        &mut self.protection
    }

    /// Return protection metadata for a sheet by index.
    pub fn sheet_protection(&self, sheet_index: usize) -> Result<&crate::ods::SheetProtection> {
        self.sheets
            .get(sheet_index)
            .map(|sheet| &sheet.protection)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "Sheet index {sheet_index} out of bounds"
                ))
            })
    }

    /// Mutably access protection metadata for a sheet by index.
    pub fn sheet_protection_mut(
        &mut self,
        sheet_index: usize,
    ) -> Result<&mut crate::ods::SheetProtection> {
        self.sheets
            .get_mut(sheet_index)
            .map(|sheet| &mut sheet.protection)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "Sheet index {sheet_index} out of bounds"
                ))
            })
    }

    /// Add a uniquely named content-validation definition.
    pub fn add_content_validation(&mut self, validation: ContentValidation) -> Result<()> {
        validation.validate()?;
        if self
            .content_validations
            .iter()
            .any(|existing| existing.name == validation.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate content validation '{}'",
                validation.name
            )));
        }
        self.content_validations.push(validation);
        Ok(())
    }

    /// Remove and return a content-validation definition.
    pub fn remove_content_validation(&mut self, name: &str) -> Option<ContentValidation> {
        let index = self
            .content_validations
            .iter()
            .position(|validation| validation.name == name)?;
        Some(self.content_validations.remove(index))
    }

    /// Add a named range.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<()> {
        self.add_named_definition(range.into())
    }

    /// Add a named OpenFormula expression.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<()> {
        self.add_named_definition(expression.into())
    }

    /// Add either kind of named definition.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<()> {
        definition.validate()?;
        self.validate_scope(definition.scope())?;
        ensure_unique(&self.named_definitions, &definition)?;
        self.named_definitions.push(definition);
        Ok(())
    }

    /// Remove a named definition and return it if it exists.
    pub fn remove_named_definition(
        &mut self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<NamedDefinition> {
        let index = self
            .named_definitions
            .iter()
            .position(|definition| definition.name() == name && definition.scope() == scope)?;
        Some(self.named_definitions.remove(index))
    }

    fn validate_scope(&self, scope: &NamedDefinitionScope) -> Result<()> {
        if let NamedDefinitionScope::Sheet(sheet_name) = scope
            && !self.sheets.iter().any(|sheet| sheet.name == *sheet_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "named definition refers to missing sheet '{sheet_name}'"
            )));
        }
        Ok(())
    }

    fn validate_named_definitions(&self) -> Result<()> {
        for (index, definition) in self.named_definitions.iter().enumerate() {
            definition.validate()?;
            self.validate_scope(definition.scope())?;
            if self.named_definitions[..index].iter().any(|existing| {
                existing.name() == definition.name() && existing.scope() == definition.scope()
            }) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "duplicate named definition '{}' in {:?}",
                    definition.name(),
                    definition.scope()
                )));
            }
        }
        Ok(())
    }

    fn validate_annotations(&self) -> Result<()> {
        for annotation in self
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .filter_map(Cell::annotation)
        {
            annotation.validate()?;
        }
        Ok(())
    }

    fn validate_content_validations(&self) -> Result<()> {
        validate_collection(&self.content_validations)?;
        let names = self
            .content_validations
            .iter()
            .map(|validation| validation.name.as_str())
            .collect::<std::collections::HashSet<_>>();
        for cell in self
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
        {
            if let Some(name) = cell.validation_name.as_deref()
                && !names.contains(name)
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cell references missing content validation '{name}'"
                )));
            }
        }
        Ok(())
    }

    fn validate_database_ranges(&self) -> Result<()> {
        self.database_ranges
            .iter()
            .try_for_each(DatabaseRange::validate)
    }

    fn has_validation_event_listeners(&self) -> bool {
        self.content_validations.iter().any(|validation| {
            validation
                .error_macro
                .as_ref()
                .is_some_and(|error_macro| !error_macro.event_listeners.is_empty())
        })
    }

    /// Add a new sheet.
    pub fn add_sheet(&mut self, name: &str) -> Result<()> {
        let sheet = Sheet {
            name: name.to_string(),
            rows: Vec::new(),
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            style: Default::default(),
            print_settings: Default::default(),
            title: None,
            description: None,
            table_source: None,
            scenario: None,
            protection: crate::ods::SheetProtection::default(),
        };
        self.sheets.push(sheet);
        Ok(())
    }

    /// Remove a sheet at index.
    pub fn remove_sheet(&mut self, index: usize) -> Result<Sheet> {
        if index < self.sheets.len() {
            let sheet = self.sheets.remove(index);
            self.named_definitions
                .retain(|definition| definition.scope().sheet_name() != Some(sheet.name.as_str()));
            Ok(sheet)
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                index
            )))
        }
    }

    /// Add a row to a sheet.
    pub fn add_row(&mut self, sheet_index: usize, cells: Vec<Cell>) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let row_index = self.sheets[sheet_index].rows.len();
            let row = Row {
                cells,
                index: row_index,
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            };
            self.sheets[sheet_index].rows.push(row);
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Remove a row from a sheet.
    pub fn remove_row(&mut self, sheet_index: usize, row_index: usize) -> Result<Row> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];
            if row_index < sheet.rows.len() {
                Ok(sheet.rows.remove(row_index))
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row_index
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Set a cell value.
    ///
    /// # Arguments
    ///
    /// * `sheet_index` - Index of the sheet
    /// * `row` - Row index
    /// * `col` - Column index  
    /// * `value` - New cell value
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{MutableSpreadsheet, CellValue};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.set_cell(0, 0, 0, CellValue::Text("Hello".to_string()))?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        value: CellValue,
    ) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];

            // Ensure row exists
            while sheet.rows.len() <= row {
                let row_index = sheet.rows.len();
                sheet.rows.push(Row {
                    cells: Vec::new(),
                    index: row_index,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
                });
            }

            let row_data = &mut sheet.rows[row];

            // Ensure cell exists
            while row_data.cells.len() <= col {
                let col_index = row_data.cells.len();
                row_data.cells.push(Cell {
                    value: CellValue::Empty,
                    text: String::new(),
                    formula: None,
                    annotation: None,
                    range_source: None,
                    validation_name: None,
                    style_name: None,
                    matrix_span: None,
                    merge: Default::default(),
                    protect: None,
                    protected: None,
                    row,
                    col: col_index,
                });
            }

            // Set the cell value
            row_data.cells[col].value = value.clone();
            row_data.cells[col].text = match value {
                CellValue::Empty => String::new(),
                CellValue::Text(ref s) => s.clone(),
                CellValue::Number(n) => n.to_string(),
                CellValue::Boolean(b) => b.to_string(),
                CellValue::Date(ref d) => d.clone(),
                CellValue::Currency(n, ref currency) => format!("{} {}", n, currency),
                CellValue::Percentage(n) => format!("{}%", n * 100.0),
                CellValue::Time(ref t) => t.clone(),
            };

            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Attach or replace an annotation on a cell.
    pub fn set_cell_annotation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        annotation: CellAnnotation,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                range_source: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].annotation = Some(annotation);
        Ok(())
    }

    /// Remove and return an annotation from a cell.
    pub fn remove_cell_annotation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<Option<CellAnnotation>> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_annotation))
    }

    /// Attach or replace inert external-range metadata on a cell.
    pub fn set_cell_range_source(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        source: CellRangeSource,
    ) -> Result<()> {
        let sheet = self.sheets.get(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        let exists = sheet
            .rows
            .get(row)
            .is_some_and(|row| row.cells.get(col).is_some());
        if !exists {
            self.set_cell(sheet_index, row, col, CellValue::Empty)?;
        }
        self.sheets[sheet_index].rows[row].cells[col].set_range_source(source);
        Ok(())
    }

    /// Remove and return external-range metadata from a cell.
    pub fn remove_cell_range_source(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<Option<CellRangeSource>> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_range_source))
    }

    /// Apply a named content validation to a cell.
    pub fn set_cell_validation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        validation_name: &str,
    ) -> Result<()> {
        if !self
            .content_validations
            .iter()
            .any(|validation| validation.name == validation_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "missing content validation '{validation_name}'"
            )));
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                range_source: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].set_validation_name(validation_name);
        Ok(())
    }

    /// Remove and return the content-validation name applied to a cell.
    pub fn clear_cell_validation(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<Option<String>> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .and_then(|cell| cell.validation_name.take()))
    }

    /// Set both ODF cell-protection attributes.
    pub fn set_cell_protection(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        protect: Option<bool>,
        protected: Option<bool>,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                range_source: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].set_protection(protect, protected);
        Ok(())
    }

    /// Apply a table-cell style name to a cell.
    pub fn set_cell_style_name(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                range_source: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].set_style_name(style_name);
        Ok(())
    }

    /// Remove and return a cell's directly applied table-cell style name.
    pub fn clear_cell_style_name(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<Option<String>> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .and_then(|cell| cell.style_name.take()))
    }

    /// Set matrix formula result dimensions on a cell.
    pub fn set_cell_matrix_span(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        row_span: usize,
        column_span: usize,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        let matrix_span = crate::ods::CellMatrixSpan::new(row_span, column_span)?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: Default::default(),
            });
        }
        let row_data = &mut sheet.rows[row];
        while row_data.cells.len() <= col {
            row_data.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                range_source: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row,
                col: row_data.cells.len(),
            });
        }
        row_data.cells[col].matrix_span = Some(matrix_span);
        Ok(())
    }

    /// Remove matrix formula result dimensions from a cell.
    pub fn clear_cell_matrix_span(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
    ) -> Result<bool> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(sheet
            .rows
            .get_mut(row)
            .and_then(|row| row.cells.get_mut(col))
            .is_some_and(|cell| cell.matrix_span.take().is_some()))
    }

    /// Set structural metadata for a row.
    pub fn set_row_metadata(
        &mut self,
        sheet_index: usize,
        row: usize,
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        visibility: TableVisibility,
    ) -> Result<()> {
        if row >= MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "row index {row} exceeds the spreadsheet safety limit"
            )));
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: TableVisibility::Visible,
            });
        }
        let item = &mut sheet.rows[row];
        item.style_name = style_name;
        item.default_cell_style_name = default_cell_style_name;
        item.visibility = visibility;
        Ok(())
    }

    /// Set structural metadata for a logical column.
    pub fn set_column_metadata(
        &mut self,
        sheet_index: usize,
        column: usize,
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        visibility: TableVisibility,
    ) -> Result<()> {
        if column >= MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "column index {column} exceeds the spreadsheet safety limit"
            )));
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.columns.len() <= column {
            sheet.columns.push(Column {
                index: sheet.columns.len(),
                ..Column::default()
            });
        }
        let item = &mut sheet.columns[column];
        item.style_name = style_name;
        item.default_cell_style_name = default_cell_style_name;
        item.visibility = visibility;
        Ok(())
    }

    /// Replace the nested row grouping and header structure in a sheet.
    pub fn set_row_structure(
        &mut self,
        sheet_index: usize,
        structure: Vec<TableStructure>,
    ) -> Result<()> {
        let required = validate_table_structure(&structure, TableStructureAxis::Rows)?;
        if required > MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "row structure exceeds the {MAX_EXPANDED_ROWS_PER_SHEET} row safety limit"
            )));
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.rows.len() < required {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: TableVisibility::Visible,
            });
        }
        sheet.row_structure = structure;
        Ok(())
    }

    /// Replace the nested column grouping and header structure in a sheet.
    pub fn set_column_structure(
        &mut self,
        sheet_index: usize,
        structure: Vec<TableStructure>,
    ) -> Result<()> {
        let required = validate_table_structure(&structure, TableStructureAxis::Columns)?;
        if required > MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "column structure exceeds the {MAX_EXPANDED_COLUMNS_PER_SHEET} column safety limit"
            )));
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        while sheet.columns.len() < required {
            sheet.columns.push(Column {
                index: sheet.columns.len(),
                ..Column::default()
            });
        }
        sheet.column_structure = structure;
        Ok(())
    }

    /// Replace sheet-level table style and template settings.
    pub fn set_sheet_style(&mut self, sheet_index: usize, style: SheetStyle) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.style = style;
        Ok(())
    }

    /// Replace sheet printing controls and ranges.
    pub fn set_sheet_print_settings(
        &mut self,
        sheet_index: usize,
        settings: SheetPrintSettings,
    ) -> Result<()> {
        validate_sheet_print_settings(&settings)?;
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.print_settings = settings;
        Ok(())
    }

    /// Set or clear a sheet's human-readable title.
    pub fn set_sheet_title(&mut self, sheet_index: usize, title: Option<String>) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.title = title;
        Ok(())
    }

    /// Set or clear a sheet's human-readable description.
    pub fn set_sheet_description(
        &mut self,
        sheet_index: usize,
        description: Option<String>,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.description = description;
        Ok(())
    }

    /// Set or clear a sheet's inert external linked-table metadata.
    pub fn set_sheet_table_source(
        &mut self,
        sheet_index: usize,
        table_source: Option<SheetTableSource>,
    ) -> Result<()> {
        if let Some(source) = &table_source {
            validate_table_source(source)?;
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.table_source = table_source;
        Ok(())
    }

    /// Set or clear a sheet's what-if scenario metadata.
    pub fn set_sheet_scenario(
        &mut self,
        sheet_index: usize,
        scenario: Option<SheetScenario>,
    ) -> Result<()> {
        if let Some(scenario) = &scenario {
            validate_scenario(scenario)?;
        }
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        sheet.scenario = scenario;
        Ok(())
    }

    /// Merge a rectangular cell range in a sheet.
    pub fn merge_cells(
        &mut self,
        sheet_index: usize,
        start_row: usize,
        start_col: usize,
        row_span: usize,
        column_span: usize,
    ) -> Result<()> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        merge_cell_range(&mut sheet.rows, start_row, start_col, row_span, column_span)
    }

    /// Remove a merge anchored at a cell.
    pub fn unmerge_cells(
        &mut self,
        sheet_index: usize,
        start_row: usize,
        start_col: usize,
    ) -> Result<bool> {
        let sheet = self.sheets.get_mut(sheet_index).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("Sheet index {sheet_index} out of bounds"))
        })?;
        Ok(unmerge_cell_range(&mut sheet.rows, start_row, start_col))
    }

    /// Clear a cell value.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.clear_cell(0, 0, 0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_cell(&mut self, sheet_index: usize, row: usize, col: usize) -> Result<()> {
        if sheet_index < self.sheets.len() {
            let sheet = &mut self.sheets[sheet_index];
            if row < sheet.rows.len() {
                let row_data = &mut sheet.rows[row];
                if col < row_data.cells.len() {
                    row_data.cells[col].value = CellValue::Empty;
                    row_data.cells[col].text = String::new();
                    Ok(())
                } else {
                    Err(litchi_core::Error::InvalidFormat(format!(
                        "Column index {} out of bounds",
                        col
                    )))
                }
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Row index {} out of bounds",
                    row
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Clear all content from a sheet.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableSpreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = MutableSpreadsheet::new();
    /// spreadsheet.add_sheet("Sheet1")?;
    /// spreadsheet.clear_sheet(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_sheet(&mut self, sheet_index: usize) -> Result<()> {
        if sheet_index < self.sheets.len() {
            self.sheets[sheet_index].rows.clear();
            self.sheets[sheet_index].row_structure.clear();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Sheet index {} out of bounds",
                sheet_index
            )))
        }
    }

    /// Generate content.xml from current state.
    fn generate_content_xml(&self) -> Result<String> {
        let mut body = String::new();
        write_content_validations(&mut body, &self.content_validations);

        for sheet in &self.sheets {
            let escaped_name = escape_xml(&sheet.name);
            body.push_str("<table:table table:name=\"");
            body.push_str(&escaped_name);
            body.push('"');
            write_sheet_formatting_attributes(&mut body, &sheet.style, &sheet.print_settings)?;
            write_sheet_attributes(&mut body, &sheet.protection);
            body.push('>');
            write_sheet_preamble(
                &mut body,
                sheet.title.as_deref(),
                sheet.description.as_deref(),
                sheet.table_source.as_ref(),
                sheet.scenario.as_ref(),
            )?;
            write_sheet_options(&mut body, &sheet.protection.options);

            let total_columns = Self::sheet_max_cols(sheet).max(sheet.columns.len()).max(1);
            write_table_structure(
                &mut body,
                &sheet.column_structure,
                total_columns,
                TableStructureAxis::Columns,
                |out, range| {
                    let explicit_end = range.end.min(sheet.columns.len());
                    if range.start < explicit_end {
                        write_columns(out, &sheet.columns[range.start..explicit_end]);
                    }
                    let default_start = range.start.max(sheet.columns.len());
                    if default_start < range.end {
                        Self::push_table_columns(out, range.end - default_start);
                    }
                },
            )?;

            write_table_structure(
                &mut body,
                &sheet.row_structure,
                sheet.rows.len(),
                TableStructureAxis::Rows,
                |out, range| {
                    for row in &sheet.rows[range] {
                        Self::push_row(out, row);
                    }
                },
            )?;

            write_named_definitions(
                &mut body,
                self.named_definitions.iter().filter(|definition| {
                    definition.scope().sheet_name() == Some(sheet.name.as_str())
                }),
            );

            body.push_str("</table:table>");
        }

        write_named_definitions(
            &mut body,
            self.named_definitions
                .iter()
                .filter(|definition| definition.scope() == &NamedDefinitionScope::Global),
        );

        write_database_ranges(&mut body, &self.database_ranges)?;

        let of_ns = if self.has_formulas() {
            " xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\""
        } else {
            ""
        };

        let mut out = String::with_capacity(body.len() + 256);
        out.push_str(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
        );
        out.push_str(of_ns);
        let has_annotations = self.has_annotations();
        let has_validation_event_listeners = self.has_validation_event_listeners();
        let has_table_sources = self.sheets.iter().any(|sheet| {
            sheet.table_source.is_some()
                || sheet
                    .rows
                    .iter()
                    .flat_map(|row| &row.cells)
                    .any(|cell| cell.range_source.is_some())
        });
        let has_protection_extensions = has_protection_extensions(
            &self.protection,
            self.sheets.iter().map(|sheet| &sheet.protection),
        );
        if has_annotations {
            out.push_str(r#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#);
        }
        if has_validation_event_listeners {
            out.push_str(r#" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0""#);
        }
        if (has_validation_event_listeners || has_table_sources) && !has_annotations {
            out.push_str(r#" xmlns:xlink="http://www.w3.org/1999/xlink""#);
        }
        if has_protection_extensions && !has_annotations {
            out.push_str(r#" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#);
        }
        if self.automatic_styles.is_some() || self.font_face_decls.is_some() {
            let mut declared = vec!["office", "table", "text"];
            if self.has_formulas() {
                declared.push("of");
            }
            if has_annotations {
                declared.extend(["dc", "meta", "draw", "svg", "xlink", "fo", "style", "loext"]);
            }
            if has_validation_event_listeners {
                declared.extend(["script", "presentation"]);
            }
            if (has_validation_event_listeners || has_table_sources) && !has_annotations {
                declared.push("xlink");
            }
            if has_protection_extensions && !has_annotations {
                declared.push("loext");
            }
            if let Some(font_face_decls) = &self.font_face_decls {
                font_face_decls.write_missing_namespaces(&mut out, declared.iter().copied());
                declared.extend(font_face_decls.namespace_prefixes());
            }
            if let Some(automatic_styles) = &self.automatic_styles {
                automatic_styles.write_missing_namespaces(&mut out, declared.iter().copied());
            }
        }
        out.push_str(r#" office:version="1.3">"#);
        if let Some(font_face_decls) = &self.font_face_decls {
            out.push_str(&font_face_decls.xml);
        } else {
            out.push_str("<office:font-face-decls/>");
        }
        if let Some(automatic_styles) = &self.automatic_styles {
            out.push_str(&automatic_styles.xml);
        } else {
            out.push_str("<office:automatic-styles/>");
        }
        out.push_str("<office:body><office:spreadsheet");
        write_spreadsheet_attributes(&mut out, &self.protection);
        out.push('>');
        out.push_str(&body);
        out.push_str(r#"</office:spreadsheet></office:body></office:document-content>"#);
        Ok(out)
    }

    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date></office:meta></office:document-meta>"#,
            now
        )
    }

    /// Save the modified spreadsheet.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate_named_definitions()?;
        self.validate_annotations()?;
        self.validate_content_validations()?;
        self.validate_database_ranges()?;
        let mut writer = PackageWriter::new();

        writer.set_mimetype(&self.mimetype)?;

        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

impl Default for MutableSpreadsheet {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ods::{
        CellStyleProtection, NamedRangeUsage, SpreadsheetBuilder, ValidationDisplayList,
        ValidationErrorMacro, ValidationEventListener, ValidationScriptEventListener,
    };

    fn package_with_cell_styles() -> Vec<u8> {
        let content = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:v="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><o:font-face-decls><s:font-face s:name="Test Font" v:font-family="'Test Font'"/></o:font-face-decls><o:automatic-styles><s:style s:name="Auto&amp;Locked" s:family="table-cell" s:parent-style-name="Named&amp;Locked"><s:table-cell-properties f:background-color="#fff" s:font-name="Test Font"/></s:style></o:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Auto&amp;Locked"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"##;
        let styles = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.3"><office:styles><style:style style:name="Named&amp;Locked" style:family="table-cell"><style:table-cell-properties style:cell-protect="protected"/></style:style></office:styles></office:document-styles>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
        writer.add_file("settings.xml", b"sheet settings").unwrap();
        writer
            .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type(
                "custom/data.bin",
                b"spreadsheet custom data",
                "application/x-ods-test",
            )
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn package_with_repeated_merged_cells() -> Vec<u8> {
        let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Merged"><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2" office:value-type="string"><text:p>anchor</text:p></table:table-cell><table:covered-table-cell/><table:table-cell office:value-type="string"><text:p>C</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn mutable_spreadsheet_preserves_and_edits_named_definitions() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder
            .add_named_range(
                NamedRange::new(
                    "LocalPrintArea",
                    "$Sheet1.$A$1:.$C$9",
                    NamedDefinitionScope::sheet("Sheet1"),
                )
                .unwrap()
                .with_usage(NamedRangeUsage::PrintRange),
            )
            .unwrap();
        let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        assert_eq!(mutable.named_definitions().len(), 1);
        mutable
            .add_named_expression(
                NamedExpression::new("GlobalValue", "of:=42", NamedDefinitionScope::Global)
                    .unwrap(),
            )
            .unwrap();

        let output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert!(
            output
                .named_range("LocalPrintArea", &NamedDefinitionScope::sheet("Sheet1"))
                .is_some()
        );
        assert!(
            output
                .named_expression("GlobalValue", &NamedDefinitionScope::Global)
                .is_some()
        );
    }

    #[test]
    fn removing_sheet_removes_its_local_definitions_only() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Sheet1").unwrap();
        mutable
            .add_named_range(
                NamedRange::new(
                    "Local",
                    "$Sheet1.$A$1",
                    NamedDefinitionScope::sheet("Sheet1"),
                )
                .unwrap(),
            )
            .unwrap();
        mutable
            .add_named_expression(
                NamedExpression::new("Global", "of:=1", NamedDefinitionScope::Global).unwrap(),
            )
            .unwrap();

        mutable.remove_sheet(0).unwrap();
        assert_eq!(mutable.named_definitions().len(), 1);
        assert_eq!(mutable.named_definitions()[0].name(), "Global");
    }

    #[test]
    fn mutable_spreadsheet_adds_edits_and_removes_annotations() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Sheet1").unwrap();
        let mut annotation = CellAnnotation::new("review this");
        annotation.set_creator(Some("Reviewer"));
        mutable.set_cell_annotation(0, 3, 4, annotation).unwrap();

        mutable.sheets_mut()[0].rows[3].cells[4]
            .annotation_mut()
            .unwrap()
            .push_paragraph("follow-up");
        let mut round_trip = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = round_trip.sheets().unwrap();
        let annotation = sheets[0].rows[3].cells[4].annotation().unwrap();
        assert_eq!(annotation.creator().as_deref(), Some("Reviewer"));
        assert_eq!(annotation.text(), "review this\nfollow-up");

        assert!(mutable.remove_cell_annotation(0, 3, 4).unwrap().is_some());
        assert!(mutable.remove_cell_annotation(0, 3, 4).unwrap().is_none());
    }

    #[test]
    fn mutable_spreadsheet_round_trips_content_validations() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Sheet1").unwrap();
        let mut validation = ContentValidation::new("list").unwrap();
        validation
            .set_condition("of:cell-content-is-in-list(\"red\";\"green\")")
            .unwrap();
        validation.display_list = Some(ValidationDisplayList::SortAscending);
        validation.error_macro = Some(ValidationErrorMacro {
            execute: Some(false),
            event_listeners: vec![ValidationEventListener::Script(
                ValidationScriptEventListener {
                    event_name: "dom:change".to_string(),
                    language: "ooo:script".to_string(),
                    macro_name: Some("Standard.Module1.Invalid".to_string()),
                    href: None,
                    actuate: None,
                },
            )],
        });
        mutable.add_content_validation(validation.clone()).unwrap();
        mutable.set_cell_validation(0, 4, 5, "list").unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(output.content_validation("list"), Some(&validation));
        assert_eq!(
            output.sheets().unwrap()[0].rows[4].cells[5].validation_name(),
            Some("list")
        );

        assert!(mutable.remove_content_validation("list").is_some());
        assert!(mutable.to_bytes().is_err());
        assert_eq!(
            mutable.clear_cell_validation(0, 4, 5).unwrap().as_deref(),
            Some("list")
        );
        assert!(mutable.to_bytes().is_ok());
    }

    #[test]
    fn mutable_spreadsheet_preserves_protection_metadata() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Protected").unwrap();
        mutable.protection_mut().structure_protected = Some(true);
        mutable.protection_mut().key = crate::ods::ProtectionKey {
            value: Some("document-key".to_string()),
            digest_algorithm: Some("urn:sha256".to_string()),
            secondary_digest_algorithm: None,
        };
        let sheet = mutable.sheet_protection_mut(0).unwrap();
        sheet.protected = Some(true);
        sheet.options.delete_rows = Some(false);
        sheet.options.use_pivot = Some(true);
        mutable
            .set_cell_protection(0, 1, 1, Some(true), Some(false))
            .unwrap();

        let spreadsheet = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(output.protection().structure_protected, Some(true));
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].protection.options.delete_rows, Some(false));
        assert_eq!(sheets[0].protection.options.use_pivot, Some(true));
        assert_eq!(sheets[0].rows[1].cells[1].protect(), Some(true));
        assert_eq!(sheets[0].rows[1].cells[1].protected(), Some(false));
    }

    #[test]
    fn mutable_spreadsheet_preserves_and_resolves_cell_styles() {
        let source_bytes = package_with_cell_styles();
        let mut spreadsheet = Spreadsheet::from_bytes(source_bytes.clone()).unwrap();
        assert_eq!(spreadsheet.to_bytes().unwrap(), source_bytes);
        let sheets = spreadsheet.sheets().unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert_eq!(cell.style_name(), Some("Auto&Locked"));
        assert_eq!(
            spreadsheet.cell_style_protection(cell).unwrap(),
            Some(CellStyleProtection::Protected)
        );

        let mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = crate::core::OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        let styles = String::from_utf8(package.get_file("styles.xml").unwrap()).unwrap();
        assert!(content.contains(r##"f:background-color="#fff""##));
        assert!(content.contains(r#"v:font-family="'Test Font'""#));
        assert_eq!(content.matches("xmlns:s=").count(), 1);
        assert!(content.contains(r#"table:style-name="Auto&amp;Locked""#));
        assert!(styles.contains(r#"style:name="Named&amp;Locked""#));
        assert_eq!(package.get_file("settings.xml").unwrap(), b"sheet settings");
        assert_eq!(
            package.get_file("custom/data.bin").unwrap(),
            b"spreadsheet custom data"
        );
        let borrowed = package.package().unwrap();
        assert_eq!(
            borrowed.manifest().get_media_type("Object 1/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(
            borrowed.manifest().get_media_type("custom/data.bin"),
            Some("application/x-ods-test")
        );

        let mut output = Spreadsheet::from_bytes(bytes).unwrap();
        let sheets = output.sheets().unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert_eq!(cell.style_name(), Some("Auto&Locked"));
        assert_eq!(
            output.cell_style_protection(cell).unwrap(),
            Some(CellStyleProtection::Protected)
        );
    }

    #[test]
    fn mutable_spreadsheet_preserves_repeated_rows_and_merged_cells() {
        let spreadsheet = Spreadsheet::from_bytes(package_with_repeated_merged_cells()).unwrap();
        let mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = crate::core::OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r#"table:number-rows-spanned="2""#));
        assert!(content.contains("<table:covered-table-cell"));

        let mut output = Spreadsheet::from_bytes(bytes).unwrap();
        let sheets = output.sheets().unwrap();
        let rows = &sheets[0].rows;
        assert_eq!(rows.len(), 4);
        assert_eq!(rows[1].cells[0].coordinates(), (1, 0));
        assert_eq!(rows[2].cells[0].span(), Some((2, 2)));
        assert_eq!(rows[2].cells[1].merge(), crate::ods::CellMerge::Covered);
        assert_eq!(rows[2].cells[2].coordinates(), (2, 2));
        assert_eq!(rows[3].cells.len(), 2);
    }

    #[test]
    fn mutable_spreadsheet_creates_and_removes_merged_ranges() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Merged").unwrap();
        mutable
            .set_cell(0, 1, 1, CellValue::Text("anchor".to_string()))
            .unwrap();
        mutable.merge_cells(0, 1, 1, 2, 2).unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].rows[1].cells[1].span(), Some((2, 2)));
        assert_eq!(
            sheets[0].rows[2].cells[2].merge(),
            crate::ods::CellMerge::Covered
        );

        assert!(mutable.unmerge_cells(0, 1, 1).unwrap());
        assert!(!mutable.unmerge_cells(0, 1, 1).unwrap());
        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert!(
            output.sheets().unwrap()[0]
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .all(|cell| cell.merge() == crate::ods::CellMerge::None)
        );
    }

    #[test]
    fn mutable_spreadsheet_round_trips_matrix_formula_spans() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Matrix").unwrap();
        mutable.set_cell(0, 0, 0, CellValue::Number(1.0)).unwrap();
        mutable.sheets_mut()[0].rows[0].cells[0].formula = Some("of:=SEQUENCE(4;3)".to_string());
        assert!(mutable.set_cell_matrix_span(0, 0, 0, 0, 3).is_err());
        mutable.set_cell_matrix_span(0, 0, 0, 4, 3).unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let cell = &output.sheets().unwrap()[0].rows[0].cells[0];
        assert_eq!(
            cell.matrix_span().map(|span| (span.rows(), span.columns())),
            Some((4, 3))
        );
        assert!(mutable.clear_cell_matrix_span(0, 0, 0).unwrap());
        assert!(!mutable.clear_cell_matrix_span(0, 0, 0).unwrap());
    }

    #[test]
    fn mutable_spreadsheet_round_trips_row_and_column_metadata() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Structure").unwrap();
        assert!(
            mutable
                .set_row_metadata(
                    0,
                    MAX_EXPANDED_ROWS_PER_SHEET,
                    None,
                    None,
                    TableVisibility::Visible,
                )
                .is_err()
        );
        assert!(
            mutable
                .set_column_metadata(
                    0,
                    MAX_EXPANDED_COLUMNS_PER_SHEET,
                    None,
                    None,
                    TableVisibility::Visible,
                )
                .is_err()
        );
        mutable
            .set_cell(0, 0, 0, CellValue::Text("value".to_string()))
            .unwrap();
        for column in 0..2 {
            mutable
                .set_column_metadata(
                    0,
                    column,
                    Some("Col&Style".to_string()),
                    Some("DefaultCell".to_string()),
                    crate::ods::TableVisibility::Collapse,
                )
                .unwrap();
        }
        mutable
            .set_row_metadata(
                0,
                0,
                Some("RowStyle".to_string()),
                Some("RowCell".to_string()),
                crate::ods::TableVisibility::Filter,
            )
            .unwrap();

        let bytes = mutable.to_bytes().unwrap();
        let package = crate::core::OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r#"table:number-columns-repeated="2""#));
        assert!(content.contains(r#"table:visibility="filter""#));

        let mut output = Spreadsheet::from_bytes(bytes).unwrap();
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].columns.len(), 2);
        assert_eq!(
            sheets[0].columns[0].style_name.as_deref(),
            Some("Col&Style")
        );
        assert_eq!(
            sheets[0].columns[0].visibility,
            crate::ods::TableVisibility::Collapse
        );
        assert_eq!(
            sheets[0].rows[0].visibility,
            crate::ods::TableVisibility::Filter
        );
        assert_eq!(
            sheets[0].rows[0].default_cell_style_name.as_deref(),
            Some("RowCell")
        );
    }

    #[test]
    fn mutable_spreadsheet_round_trips_table_structure() {
        let structure = vec![TableStructure::Group(crate::ods::TableGroup {
            display: false,
            children: vec![
                TableStructure::Header(crate::ods::TableRange::new(0, 1).unwrap()),
                TableStructure::Range(crate::ods::TableRange::new(1, 2).unwrap()),
            ],
        })];
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Outline").unwrap();
        mutable.set_column_structure(0, structure.clone()).unwrap();
        mutable.set_row_structure(0, structure.clone()).unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].column_structure, structure);
        assert_eq!(sheets[0].row_structure, sheets[0].column_structure);
    }

    #[test]
    fn mutable_spreadsheet_round_trips_sheet_style_and_print_settings() {
        let style = SheetStyle {
            style_name: Some("MutableStyle".to_string()),
            template_name: None,
            usage: crate::SheetStyleUsage {
                use_banding_row_styles: Some(true),
                ..crate::SheetStyleUsage::default()
            },
        };
        let print =
            SheetPrintSettings::new(false, vec!["'Q1 Sales'.$A$1:$C$9".to_string()]).unwrap();
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Print").unwrap();
        mutable.set_sheet_style(0, style.clone()).unwrap();
        mutable.set_sheet_print_settings(0, print.clone()).unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].style, style);
        assert_eq!(sheets[0].print_settings, print);
    }

    #[test]
    fn mutable_spreadsheet_round_trips_sheet_text_and_scenario() {
        let mut scenario = SheetScenario::new(vec![".A1:.C3".to_string()], false).unwrap();
        scenario.display_border = Some(true);
        scenario.comment = Some("Mutable & safe".to_string());
        let mut source = SheetTableSource::new("../mutable.ods");
        source.filter_options = Some("A&B".to_string());
        source.actuate_on_request = true;
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Scenario").unwrap();
        mutable
            .set_sheet_title(0, Some("Mutable title".to_string()))
            .unwrap();
        mutable
            .set_sheet_description(0, Some("Mutable description".to_string()))
            .unwrap();
        mutable
            .set_sheet_table_source(0, Some(source.clone()))
            .unwrap();
        mutable
            .set_sheet_scenario(0, Some(scenario.clone()))
            .unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let sheets = output.sheets().unwrap();
        assert_eq!(sheets[0].title.as_deref(), Some("Mutable title"));
        assert_eq!(
            sheets[0].description.as_deref(),
            Some("Mutable description")
        );
        assert_eq!(sheets[0].table_source.as_ref(), Some(&source));
        assert_eq!(sheets[0].scenario.as_ref(), Some(&scenario));
    }

    #[test]
    fn mutable_spreadsheet_edits_cell_range_sources() {
        let mut mutable = MutableSpreadsheet::new();
        mutable.add_sheet("Imports").unwrap();
        mutable.set_cell(0, 0, 1, CellValue::Number(42.0)).unwrap();
        let source = CellRangeSource::new("Data", "../data.ods", 2, 3).unwrap();
        mutable
            .set_cell_range_source(0, 0, 1, source.clone())
            .unwrap();

        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            output.sheets().unwrap()[0].rows[0].cells[1].range_source(),
            Some(&source)
        );
        assert_eq!(
            output.sheets().unwrap()[0].rows[0].cells[1].value,
            CellValue::Number(42.0)
        );
        assert_eq!(
            mutable.remove_cell_range_source(0, 0, 1).unwrap(),
            Some(source)
        );
        let mut output = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert!(
            output.sheets().unwrap()[0].rows[0].cells[1]
                .range_source()
                .is_none()
        );
    }
}
