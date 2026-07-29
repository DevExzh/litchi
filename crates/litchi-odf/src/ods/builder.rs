//! OpenDocument Spreadsheet builder.
//!
//! This module provides a builder pattern for creating new ODS spreadsheets from scratch.

use crate::core::{OdfStructure, PackageWriter};
use crate::ods::{
    CalculationSettings, Cell, CellAnnotation, CellDetective, CellHyperlink, CellRangeSource,
    CellValue, Column,
    ConditionalCellStyle, ConditionalFormat, Consolidation, ContentValidation, DataPilotTable,
    DatabaseRange, DdeLink, LabelRange,
    NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, Row, Sheet,
    SheetPrintSettings, SheetScenario, SheetStyle, SheetTableSource, SpreadsheetProtection,
    TableCellProtectionStyle, TableStructure, TableVisibility,
    calculation::write_calculation_settings,
    cell::{merge_cell_range, unmerge_cell_range},
    conditional_format::{
        CALCEXT_NAMESPACE_DECLARATION, MAX_CONDITIONAL_FORMATS_PER_SHEET,
        validate_conditional_format, validate_conditional_formats, write_conditional_formats,
    },
    consolidation::write_consolidation,
    data_pilot::write_data_pilot_tables,
    data_validation::{validate_collection, write_content_validations},
    database_range::write_database_ranges,
    dde::write_dde_links,
    label_range::write_label_ranges,
    named_expression::{ensure_unique, write_named_definitions},
    protection::{
        has_extensions as has_protection_extensions, write_sheet_attributes, write_sheet_options,
        write_spreadsheet_attributes,
    },
    scenario::{validate_scenario, write_sheet_preamble},
    sheet_image::{
        MAX_IMAGES_PER_SHEET, append_sheet_image_alternative, insert_sheet_image_alternative,
        normalize_sheet_image, remove_sheet_image_alternative,
    },
    source::validate_table_source,
    style_protection::{
        rewrite_managed_cell_styles, validate_conditional_style_collection,
        validate_protection_style_collection, validate_style_name,
    },
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, TableStructureAxis,
        validate_sheet_print_settings, validate_table_structure, write_columns,
        write_row_attributes, write_sheet_formatting_attributes, write_table_structure,
    },
};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::{collections::HashSet, ops::Range, path::Path};

/// Builder for creating new ODS spreadsheets.
///
/// This builder allows you to create ODS spreadsheets programmatically by adding
/// sheets, rows, and cells, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::SpreadsheetBuilder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = SpreadsheetBuilder::new();
/// builder.add_sheet("Sheet1")?;
/// builder.add_row_with_values(&["Name", "Age", "City"])?;
/// builder.add_row_with_values(&["Alice", "30", "New York"])?;
/// builder.save("spreadsheet.ods")?;
/// # Ok(())
/// # }
/// ```
pub struct SpreadsheetBuilder {
    sheets: Vec<Sheet>,
    metadata: Metadata,
    named_definitions: Vec<NamedDefinition>,
    content_validations: Vec<ContentValidation>,
    database_ranges: Vec<DatabaseRange>,
    data_pilot_tables: Vec<DataPilotTable>,
    calculation_settings: Option<CalculationSettings>,
    label_ranges: Vec<LabelRange>,
    consolidation: Option<Consolidation>,
    dde_links: Vec<DdeLink>,
    protection: SpreadsheetProtection,
    common_table_cell_styles: Vec<String>,
    conditional_cell_styles: Vec<ConditionalCellStyle>,
    table_cell_protection_styles: Vec<TableCellProtectionStyle>,
}

impl Default for SpreadsheetBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl SpreadsheetBuilder {
    /// Create a new spreadsheet builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// let builder = SpreadsheetBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            metadata: Metadata::default(),
            named_definitions: Vec::new(),
            content_validations: Vec::new(),
            database_ranges: Vec::new(),
            data_pilot_tables: Vec::new(),
            calculation_settings: None,
            label_ranges: Vec::new(),
            consolidation: None,
            dde_links: Vec::new(),
            protection: SpreadsheetProtection::default(),
            common_table_cell_styles: Vec::new(),
            conditional_cell_styles: Vec::new(),
            table_cell_protection_styles: Vec::new(),
        }
    }

    /// Return common table-cell styles declared by this builder.
    pub fn common_table_cell_styles(&self) -> &[String] {
        &self.common_table_cell_styles
    }

    /// Declare an empty common table-cell style usable as a conditional rule target.
    pub fn add_common_table_cell_style(
        &mut self,
        style_name: impl Into<String>,
    ) -> Result<&mut Self> {
        let style_name = style_name.into();
        validate_style_name(&style_name, "common table-cell style name")?;
        if self.common_table_cell_styles.iter().any(|name| name == &style_name) {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate common table-cell style name '{style_name}'"
            )));
        }
        if self.common_table_cell_styles.len() >= 65_536 {
            return Err(litchi_core::Error::InvalidFormat(
                "common table-cell style limit exceeded".to_string(),
            ));
        }
        self.common_table_cell_styles.push(style_name);
        Ok(self)
    }

    /// Return authored conditional table-cell styles in deterministic write order.
    pub fn conditional_cell_styles(&self) -> &[ConditionalCellStyle] {
        &self.conditional_cell_styles
    }

    /// Create a uniquely named conditional table-cell style atomically.
    pub fn create_conditional_cell_style(
        &mut self,
        style: ConditionalCellStyle,
    ) -> Result<&mut Self> {
        if self
            .conditional_cell_styles
            .iter()
            .any(|existing| existing.style_name == style.style_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate conditional style name '{}'",
                style.style_name
            )));
        }
        let mut candidate = self.conditional_cell_styles.clone();
        candidate.push(style);
        self.validate_conditional_styles(&candidate)?;
        self.conditional_cell_styles = candidate;
        Ok(self)
    }

    /// Replace an existing conditional style atomically and return its old value.
    pub fn replace_conditional_cell_style(
        &mut self,
        style: ConditionalCellStyle,
    ) -> Result<ConditionalCellStyle> {
        let index = self
            .conditional_cell_styles
            .iter()
            .position(|existing| existing.style_name == style.style_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "missing conditional style '{}'",
                    style.style_name
                ))
            })?;
        let mut candidate = self.conditional_cell_styles.clone();
        let previous = std::mem::replace(&mut candidate[index], style);
        self.validate_conditional_styles(&candidate)?;
        self.conditional_cell_styles = candidate;
        Ok(previous)
    }

    /// Remove and return a conditional style by name.
    pub fn remove_conditional_cell_style(
        &mut self,
        style_name: &str,
    ) -> Option<ConditionalCellStyle> {
        self.conditional_cell_styles
            .iter()
            .position(|style| style.style_name == style_name)
            .map(|index| self.conditional_cell_styles.remove(index))
    }

    fn validate_conditional_styles(&self, styles: &[ConditionalCellStyle]) -> Result<()> {
        let common = self
            .common_table_cell_styles
            .iter()
            .cloned()
            .collect::<HashSet<_>>();
        validate_conditional_style_collection(styles, &common)?;
        rewrite_managed_cell_styles(None, styles, &self.table_cell_protection_styles)?;
        Ok(())
    }

    /// Return automatic table-cell protection styles in write order.
    pub fn table_cell_protection_styles(&self) -> &[TableCellProtectionStyle] {
        &self.table_cell_protection_styles
    }

    /// Create a uniquely named automatic protection style atomically.
    pub fn create_table_cell_protection_style(
        &mut self,
        style: TableCellProtectionStyle,
    ) -> Result<&mut Self> {
        if self.table_cell_protection_styles.iter().any(|existing| {
            existing.style_name == style.style_name
        }) || self
            .conditional_cell_styles
            .iter()
            .any(|existing| existing.style_name == style.style_name
                && existing.parent_style_name != style.parent_style_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate or incompatible protection style name '{}'",
                style.style_name
            )));
        }
        let mut candidate = self.table_cell_protection_styles.clone();
        candidate.push(style);
        self.validate_protection_styles(&candidate)?;
        self.table_cell_protection_styles = candidate;
        Ok(self)
    }

    /// Replace an automatic protection style atomically.
    pub fn replace_table_cell_protection_style(
        &mut self,
        style: TableCellProtectionStyle,
    ) -> Result<TableCellProtectionStyle> {
        let index = self
            .table_cell_protection_styles
            .iter()
            .position(|existing| existing.style_name == style.style_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "missing protection style '{}'",
                    style.style_name
                ))
            })?;
        let mut candidate = self.table_cell_protection_styles.clone();
        let previous = std::mem::replace(&mut candidate[index], style);
        self.validate_protection_styles(&candidate)?;
        self.table_cell_protection_styles = candidate;
        Ok(previous)
    }

    /// Remove and return an automatic protection style.
    pub fn remove_table_cell_protection_style(
        &mut self,
        style_name: &str,
    ) -> Option<TableCellProtectionStyle> {
        self.table_cell_protection_styles
            .iter()
            .position(|style| style.style_name == style_name)
            .map(|index| self.table_cell_protection_styles.remove(index))
    }

    fn validate_protection_styles(&self, styles: &[TableCellProtectionStyle]) -> Result<()> {
        validate_protection_style_collection(styles)?;
        let known = self
            .common_table_cell_styles
            .iter()
            .cloned()
            .chain(self.conditional_cell_styles.iter().map(|style| style.style_name.clone()))
            .chain(styles.iter().map(|style| style.style_name.clone()))
            .collect::<HashSet<_>>();
        for style in styles {
            if let Some(parent) = &style.parent_style_name
                && !known.contains(parent)
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "protection style '{}' references missing parent style '{parent}'",
                    style.style_name
                )));
            }
            let mut visited = HashSet::new();
            let mut current = style;
            while let Some(parent) = &current.parent_style_name {
                if !visited.insert(current.style_name.as_str()) {
                    return Err(litchi_core::Error::InvalidFormat(format!(
                        "cyclic protection style inheritance at '{}'",
                        style.style_name
                    )));
                }
                let Some(next) = styles.iter().find(|candidate| &candidate.style_name == parent)
                else {
                    break;
                };
                current = next;
            }
        }
        Ok(())
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Return the named ranges and expressions added to this spreadsheet.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Return document-level cell validation definitions.
    pub fn content_validations(&self) -> &[ContentValidation] {
        &self.content_validations
    }

    /// Return spreadsheet-wide formula calculation settings.
    pub fn calculation_settings(&self) -> Option<&CalculationSettings> {
        self.calculation_settings.as_ref()
    }

    /// Set or clear validated spreadsheet-wide calculation settings.
    pub fn set_calculation_settings(
        &mut self,
        settings: Option<CalculationSettings>,
    ) -> Result<&mut Self> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        self.calculation_settings = settings;
        Ok(self)
    }

    /// Return row and column label ranges in document order.
    pub fn label_ranges(&self) -> &[LabelRange] {
        &self.label_ranges
    }

    /// Add a validated row or column label range.
    pub fn add_label_range(&mut self, range: LabelRange) -> Result<&mut Self> {
        range.validate()?;
        self.label_ranges.push(range);
        Ok(self)
    }

    /// Remove a label range by index.
    pub fn remove_label_range(&mut self, index: usize) -> Option<LabelRange> {
        (index < self.label_ranges.len()).then(|| self.label_ranges.remove(index))
    }

    fn validate_label_ranges(&self) -> Result<()> {
        self.label_ranges.iter().try_for_each(LabelRange::validate)
    }

    /// Return the inert spreadsheet consolidation declaration.
    pub fn consolidation(&self) -> Option<&Consolidation> {
        self.consolidation.as_ref()
    }

    /// Set or clear a validated spreadsheet consolidation declaration.
    pub fn set_consolidation(&mut self, consolidation: Option<Consolidation>) -> Result<&mut Self> {
        if let Some(consolidation) = &consolidation {
            consolidation.validate()?;
        }
        self.consolidation = consolidation;
        Ok(self)
    }

    /// Return inert DDE declarations and their cached tables.
    pub fn dde_links(&self) -> &[DdeLink] {
        &self.dde_links
    }

    /// Add a validated inert DDE declaration and cached table.
    pub fn add_dde_link(&mut self, link: DdeLink) -> Result<&mut Self> {
        link.validate()?;
        self.dde_links.push(link);
        Ok(self)
    }

    /// Remove a DDE declaration by index.
    pub fn remove_dde_link(&mut self, index: usize) -> Option<DdeLink> {
        (index < self.dde_links.len()).then(|| self.dde_links.remove(index))
    }

    /// Return database ranges added to this spreadsheet.
    pub fn database_ranges(&self) -> &[DatabaseRange] {
        &self.database_ranges
    }

    /// Add a database range with optional filter, sort, source, and subtotal metadata.
    pub fn add_database_range(&mut self, range: DatabaseRange) -> Result<&mut Self> {
        range.validate()?;
        self.database_ranges.push(range);
        Ok(self)
    }

    /// Remove a database range by index.
    pub fn remove_database_range(&mut self, index: usize) -> Option<DatabaseRange> {
        (index < self.database_ranges.len()).then(|| self.database_ranges.remove(index))
    }

    fn validate_database_ranges(&self) -> Result<()> {
        crate::ods::database_range::validate_database_range_collection(&self.database_ranges)
    }

    /// Return data-pilot (pivot-table) declarations.
    pub fn data_pilot_tables(&self) -> &[DataPilotTable] {
        &self.data_pilot_tables
    }

    /// Add a validated data-pilot table.
    pub fn add_data_pilot_table(&mut self, table: DataPilotTable) -> Result<&mut Self> {
        let mut tables = self.data_pilot_tables.clone();
        tables.push(table.clone());
        crate::ods::data_pilot::validate_data_pilot_tables(&tables)?;
        self.data_pilot_tables.push(table);
        Ok(self)
    }

    /// Remove a data-pilot table by index.
    pub fn remove_data_pilot_table(&mut self, index: usize) -> Option<DataPilotTable> {
        (index < self.data_pilot_tables.len()).then(|| self.data_pilot_tables.remove(index))
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
    pub fn add_content_validation(&mut self, validation: ContentValidation) -> Result<&mut Self> {
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
        Ok(self)
    }

    /// Add a named range.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<&mut Self> {
        self.add_named_definition(range.into())
    }

    /// Add a named OpenFormula expression.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<&mut Self> {
        self.add_named_definition(expression.into())
    }

    /// Add either kind of named definition.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<&mut Self> {
        definition.validate()?;
        self.validate_scope(definition.scope())?;
        ensure_unique(&self.named_definitions, &definition)?;
        self.named_definitions.push(definition);
        Ok(self)
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
        crate::ods::named_expression::validate_named_definition_collection(&self.named_definitions)?;
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

    fn validate_hyperlinks(&self) -> Result<()> {
        for cell in self
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
        {
            cell.validate_hyperlinks()?;
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

    fn has_validation_event_listeners(&self) -> bool {
        self.content_validations.iter().any(|validation| {
            validation
                .error_macro
                .as_ref()
                .is_some_and(|error_macro| !error_macro.event_listeners.is_empty())
        })
    }

    /// Add a new sheet to the spreadsheet
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the sheet
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_sheet(&mut self, name: &str) -> Result<&mut Self> {
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
            dde_source: None,
            scenario: None,
            conditional_formats: Vec::new(),
            images: Vec::new(),
            shapes: Vec::new(),
            protection: crate::ods::SheetProtection::default(),
        };
        self.sheets.push(sheet);
        Ok(self)
    }

    /// Add a row to the current sheet with string values
    ///
    /// # Arguments
    ///
    /// * `values` - String values for the cells in the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_values(&["A", "B", "C"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_values(&mut self, values: &[&str]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, &value)| Cell {
                text: value.to_string(),
                value: CellValue::Text(value.to_string()),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row: row_index,
                col,
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a row with numbers
    ///
    /// # Arguments
    ///
    /// * `values` - Numeric values for the cells in the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_numbers(&[1.0, 2.5, 3.14])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_numbers(&mut self, values: &[f64]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, &value)| Cell {
                text: value.to_string(),
                value: CellValue::Number(value),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row: row_index,
                col,
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a row with mixed values (numbers, text, booleans)
    ///
    /// # Arguments
    ///
    /// * `values` - Cell values for the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::{SpreadsheetBuilder, CellValue};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_row_with_cell_values(&[
    ///     CellValue::Text("Product".to_string()),
    ///     CellValue::Number(99.99),
    ///     CellValue::Boolean(true),
    /// ])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row_with_cell_values(&mut self, values: &[CellValue]) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let cells: Vec<Cell> = values
            .iter()
            .enumerate()
            .map(|(col, value)| {
                let text = match value {
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Text(t) => t.clone(),
                    CellValue::Boolean(b) => b.to_string(),
                    CellValue::Date(d) => d.clone(),
                    CellValue::Currency(n, code) => format!("{} {}", n, code),
                    CellValue::Percentage(n) => format!("{}%", n),
                    CellValue::Time(t) => t.clone(),
                    CellValue::Empty => String::new(),
                };
                Cell {
                    text,
                    value: value.clone(),
                    formula: None,
                    annotation: None,
                    hyperlinks: Vec::new(),
                    rich_text: None,
                    range_source: None,
                    detective: None,
                    validation_name: None,
                    style_name: None,
                    matrix_span: None,
                    merge: Default::default(),
                    protect: None,
                    protected: None,
                    row: row_index,
                    col,
                }
            })
            .collect();

        let row = Row {
            cells,
            index: row_index,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Set a cell value at a specific position in the current sheet
    ///
    /// # Arguments
    ///
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Cell value
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::{SpreadsheetBuilder, CellValue};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.set_cell(0, 0, CellValue::Number(42.0))?;
    /// builder.set_cell(0, 1, CellValue::Text("Hello".to_string()))?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        if let Some(sheet) = self.sheets.last_mut() {
            // Ensure we have enough rows
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

            // Ensure we have enough cells in the row
            while row_data.cells.len() <= col {
                row_data.cells.push(Cell {
                    text: String::new(),
                    value: CellValue::Empty,
                    formula: None,
                    annotation: None,
                    hyperlinks: Vec::new(),
                    rich_text: None,
                    range_source: None,
                    detective: None,
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

            // Set the cell value
            let text = match &value {
                CellValue::Number(n) => n.to_string(),
                CellValue::Text(t) => t.clone(),
                CellValue::Boolean(b) => b.to_string(),
                CellValue::Date(d) => d.clone(),
                CellValue::Currency(n, code) => format!("{} {}", n, code),
                CellValue::Percentage(n) => format!("{}%", n),
                CellValue::Time(t) => t.clone(),
                CellValue::Empty => String::new(),
            };

            let annotation = row_data.cells[col].annotation.take();
            let range_source = row_data.cells[col].range_source.take();
            let detective = row_data.cells[col].detective.take();
            let validation_name = row_data.cells[col].validation_name.take();
            let style_name = row_data.cells[col].style_name.take();
            let matrix_span = row_data.cells[col].matrix_span;
            let merge = row_data.cells[col].merge;
            let protect = row_data.cells[col].protect;
            let protected = row_data.cells[col].protected;
            row_data.cells[col] = Cell {
                text,
                value,
                formula: None,
                annotation,
                // Hyperlinks anchor into the replaced text, so they are dropped.
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source,
                detective,
                validation_name,
                style_name,
                matrix_span,
                merge,
                protect,
                protected,
                row,
                col,
            };
        }

        Ok(self)
    }

    /// Replace a cell's displayed value with one inert full-cell hyperlink.
    ///
    /// The target is serialized as ODF `text:a` metadata and is never
    /// dereferenced or fetched by this library.
    pub fn set_cell_hyperlink(
        &mut self,
        row: usize,
        col: usize,
        href: impl Into<String>,
        text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.set_cell_hyperlink_data(row, col, CellHyperlink::with_text(href, text)?)
    }

    /// Replace a cell's displayed value with a fully configured hyperlink.
    pub fn set_cell_hyperlink_data(
        &mut self,
        row: usize,
        col: usize,
        hyperlink: CellHyperlink,
    ) -> Result<&mut Self> {
        hyperlink.validate()?;
        if self
            .sheets
            .last()
            .and_then(|sheet| sheet.rows.get(row))
            .and_then(|row| row.cells.get(col))
            .is_some_and(|cell| cell.merge == crate::ods::CellMerge::Covered)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "cannot author a hyperlink in a covered cell".to_string(),
            ));
        }
        self.set_cell(row, col, CellValue::Text(hyperlink.text.clone()))?;
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .expect("set_cell creates the requested current-sheet cell")
            .set_hyperlink(hyperlink)?;
        Ok(self)
    }

    /// Add an inert hyperlink around a UTF-8 byte range of an existing cell.
    ///
    /// The link's visible text is derived from the selected cell range. The
    /// target is serialized as ODF `text:a` metadata and is never fetched or
    /// followed by this library.
    pub fn add_cell_hyperlink(
        &mut self,
        row: usize,
        col: usize,
        range: Range<usize>,
        href: impl Into<String>,
    ) -> Result<&mut Self> {
        let text = self
            .sheets
            .last()
            .and_then(|sheet| sheet.rows.get(row))
            .and_then(|row| row.cells.get(col))
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "cell at row {row}, column {col} does not exist"
                ))
            })?
            .text
            .get(range.clone())
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "cell hyperlink range is not on a UTF-8 character boundary".to_string(),
                )
            })?
            .to_string();
        self.add_cell_hyperlink_data(row, col, range, CellHyperlink::with_text(href, text)?)
    }

    /// Add a configured inert hyperlink around an existing cell-text range.
    ///
    /// The link text must exactly match the selected range, and link ranges
    /// must remain ordered and non-overlapping.
    pub fn add_cell_hyperlink_data(
        &mut self,
        row: usize,
        col: usize,
        range: Range<usize>,
        hyperlink: CellHyperlink,
    ) -> Result<&mut Self> {
        let cell = self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "cell at row {row}, column {col} does not exist"
                ))
            })?;
        cell.add_hyperlink(range, hyperlink)?;
        Ok(self)
    }

    /// Remove every hyperlink from a cell in the current sheet.
    pub fn remove_cell_hyperlinks(&mut self, row: usize, col: usize) -> Result<Vec<CellHyperlink>> {
        Ok(self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .map(Cell::clear_hyperlinks)
            .unwrap_or_default())
    }

    /// Remove one hyperlink by document-order index while preserving cell text.
    pub fn remove_cell_hyperlink(
        &mut self,
        row: usize,
        col: usize,
        hyperlink_index: usize,
    ) -> Result<Option<CellHyperlink>> {
        Ok(self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(|cell| cell.remove_hyperlink(hyperlink_index)))
    }

    /// Set a cell formula at a specific position in the current sheet
    ///
    /// # Arguments
    ///
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - The formula (e.g., "=SUM(A1:A10)")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.set_cell_formula(0, 0, "=SUM(A2:A10)")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_cell_formula(&mut self, row: usize, col: usize, formula: &str) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        if let Some(sheet) = self.sheets.last_mut() {
            // Ensure we have enough rows
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

            // Ensure we have enough cells in the row
            while row_data.cells.len() <= col {
                row_data.cells.push(Cell {
                    text: String::new(),
                    value: CellValue::Empty,
                    formula: None,
                    annotation: None,
                    hyperlinks: Vec::new(),
                    rich_text: None,
                    range_source: None,
                    detective: None,
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

            // Set the formula
            row_data.cells[col].formula = Some(formula.to_string());
        }

        Ok(self)
    }

    /// Attach or replace an annotation on a cell in the current sheet.
    pub fn set_cell_annotation(
        &mut self,
        row: usize,
        col: usize,
        annotation: CellAnnotation,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self
            .sheets
            .last_mut()
            .expect("a default sheet was added when the builder was empty");
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
                text: String::new(),
                value: CellValue::Empty,
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
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
        Ok(self)
    }

    /// Remove and return an annotation from a cell in the current sheet.
    pub fn remove_cell_annotation(
        &mut self,
        row: usize,
        col: usize,
    ) -> Result<Option<CellAnnotation>> {
        Ok(self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_annotation))
    }

    /// Attach or replace inert external-range metadata on a cell in the current sheet.
    pub fn set_cell_range_source(
        &mut self,
        row: usize,
        col: usize,
        source: CellRangeSource,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let exists = self
            .sheets
            .last()
            .and_then(|sheet| sheet.rows.get(row))
            .is_some_and(|row| row.cells.get(col).is_some());
        if !exists {
            self.set_cell(row, col, CellValue::Empty)?;
        }
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .expect("set_cell materialized the requested cell")
            .set_range_source(source);
        Ok(self)
    }

    /// Remove and return external-range metadata from a cell in the current sheet.
    pub fn remove_cell_range_source(&mut self, row: usize, col: usize) -> Option<CellRangeSource> {
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_range_source)
    }

    /// Attach or replace formula-auditing metadata on a cell in the current sheet.
    pub fn set_cell_detective(
        &mut self,
        row: usize,
        col: usize,
        detective: CellDetective,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let exists = self
            .sheets
            .last()
            .and_then(|sheet| sheet.rows.get(row))
            .is_some_and(|row| row.cells.get(col).is_some());
        if !exists {
            self.set_cell(row, col, CellValue::Empty)?;
        }
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .expect("set_cell materialized the requested cell")
            .set_detective(detective);
        Ok(self)
    }

    /// Remove and return formula-auditing metadata from a cell in the current sheet.
    pub fn remove_cell_detective(&mut self, row: usize, col: usize) -> Option<CellDetective> {
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(Cell::take_detective)
    }

    /// Apply a named content validation to a cell in the current sheet.
    pub fn set_cell_validation(
        &mut self,
        row: usize,
        col: usize,
        validation_name: &str,
    ) -> Result<&mut Self> {
        if !self
            .content_validations
            .iter()
            .any(|validation| validation.name == validation_name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "missing content validation '{validation_name}'"
            )));
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
                text: String::new(),
                value: CellValue::Empty,
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
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
        Ok(self)
    }

    /// Remove and return the content-validation name applied to a cell.
    pub fn clear_cell_validation(&mut self, row: usize, col: usize) -> Result<Option<String>> {
        Ok(self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(|cell| cell.validation_name.take()))
    }

    /// Set both ODF cell-protection attributes on a cell in the current sheet.
    pub fn set_cell_protection(
        &mut self,
        row: usize,
        col: usize,
        protect: Option<bool>,
        protected: Option<bool>,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
                text: String::new(),
                value: CellValue::Empty,
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
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
        Ok(self)
    }

    /// Apply a table-cell style name to a cell in the current sheet.
    pub fn set_cell_style_name(
        &mut self,
        row: usize,
        col: usize,
        style_name: impl Into<String>,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
                text: String::new(),
                value: CellValue::Empty,
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
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
        Ok(self)
    }

    /// Remove and return the table-cell style name applied to a cell.
    pub fn clear_cell_style_name(&mut self, row: usize, col: usize) -> Result<Option<String>> {
        Ok(self
            .sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .and_then(|cell| cell.style_name.take()))
    }

    /// Set matrix formula result dimensions on a cell in the current sheet.
    pub fn set_cell_matrix_span(
        &mut self,
        row: usize,
        col: usize,
        row_span: usize,
        column_span: usize,
    ) -> Result<&mut Self> {
        let matrix_span = crate::ods::CellMatrixSpan::new(row_span, column_span)?;
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
                text: String::new(),
                value: CellValue::Empty,
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
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
        Ok(self)
    }

    /// Remove matrix formula result dimensions from a cell.
    pub fn clear_cell_matrix_span(&mut self, row: usize, col: usize) -> bool {
        self.sheets
            .last_mut()
            .and_then(|sheet| sheet.rows.get_mut(row))
            .and_then(|row| row.cells.get_mut(col))
            .is_some_and(|cell| cell.matrix_span.take().is_some())
    }

    /// Set structural metadata for a row in the current sheet.
    pub fn set_row_metadata(
        &mut self,
        row: usize,
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        visibility: TableVisibility,
    ) -> Result<&mut Self> {
        if row >= MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "row index {row} exceeds the spreadsheet safety limit"
            )));
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        while sheet.rows.len() <= row {
            sheet.rows.push(Row {
                cells: Vec::new(),
                index: sheet.rows.len(),
                style_name: None,
                default_cell_style_name: None,
                visibility: TableVisibility::Visible,
            });
        }
        let row = &mut sheet.rows[row];
        row.style_name = style_name;
        row.default_cell_style_name = default_cell_style_name;
        row.visibility = visibility;
        Ok(self)
    }

    /// Set structural metadata for a logical column in the current sheet.
    pub fn set_column_metadata(
        &mut self,
        column: usize,
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        visibility: TableVisibility,
    ) -> Result<&mut Self> {
        if column >= MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "column index {column} exceeds the spreadsheet safety limit"
            )));
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
        Ok(self)
    }

    /// Replace the nested row grouping and header structure in the current sheet.
    pub fn set_row_structure(&mut self, structure: Vec<TableStructure>) -> Result<&mut Self> {
        let required = validate_table_structure(&structure, TableStructureAxis::Rows)?;
        if required > MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "row structure exceeds the {MAX_EXPANDED_ROWS_PER_SHEET} row safety limit"
            )));
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
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
        Ok(self)
    }

    /// Replace the nested column grouping and header structure in the current sheet.
    pub fn set_column_structure(&mut self, structure: Vec<TableStructure>) -> Result<&mut Self> {
        let required = validate_table_structure(&structure, TableStructureAxis::Columns)?;
        if required > MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "column structure exceeds the {MAX_EXPANDED_COLUMNS_PER_SHEET} column safety limit"
            )));
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        while sheet.columns.len() < required {
            sheet.columns.push(Column {
                index: sheet.columns.len(),
                ..Column::default()
            });
        }
        sheet.column_structure = structure;
        Ok(self)
    }

    /// Replace sheet-level table style and template settings in the current sheet.
    pub fn set_sheet_style(&mut self, style: SheetStyle) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .style = style;
        Ok(self)
    }

    /// Replace printing controls and ranges in the current sheet.
    pub fn set_sheet_print_settings(&mut self, settings: SheetPrintSettings) -> Result<&mut Self> {
        validate_sheet_print_settings(&settings)?;
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .print_settings = settings;
        Ok(self)
    }

    /// Set or clear the current sheet's human-readable title.
    pub fn set_sheet_title(&mut self, title: Option<String>) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .title = title;
        Ok(self)
    }

    /// Set or clear the current sheet's human-readable description.
    pub fn set_sheet_description(&mut self, description: Option<String>) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .description = description;
        Ok(self)
    }

    /// Set or clear the current sheet's inert external linked-table metadata.
    pub fn set_sheet_table_source(
        &mut self,
        table_source: Option<SheetTableSource>,
    ) -> Result<&mut Self> {
        if let Some(source) = &table_source {
            validate_table_source(source)?;
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .table_source = table_source;
        Ok(self)
    }

    /// Set or clear the current sheet's inert DDE source declaration.
    pub fn set_sheet_dde_source(
        &mut self,
        dde_source: Option<crate::DdeSource>,
    ) -> Result<&mut Self> {
        if let Some(source) = &dde_source {
            source.validate()?;
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .dde_source = dde_source;
        Ok(self)
    }

    /// Set or clear the current sheet's what-if scenario metadata.
    pub fn set_sheet_scenario(&mut self, scenario: Option<SheetScenario>) -> Result<&mut Self> {
        if let Some(scenario) = &scenario {
            validate_scenario(scenario)?;
        }
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .scenario = scenario;
        Ok(self)
    }

    /// Replace the current sheet's inert `calcext` conditional formats.
    ///
    /// Passing an empty collection removes all conditional formats from the
    /// sheet. Conditions are stored as typed data and are never evaluated.
    pub fn set_sheet_conditional_formats(
        &mut self,
        formats: Vec<ConditionalFormat>,
    ) -> Result<&mut Self> {
        validate_conditional_formats(&formats)?;
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        self.sheets
            .last_mut()
            .expect("default sheet was added")
            .conditional_formats = formats;
        Ok(self)
    }

    /// Append one inert conditional format to the current sheet.
    pub fn add_sheet_conditional_format(
        &mut self,
        format: ConditionalFormat,
    ) -> Result<&mut Self> {
        validate_conditional_format(&format)?;
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        if sheet.conditional_formats.len() >= MAX_CONDITIONAL_FORMATS_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
            )));
        }
        sheet.conditional_formats.push(format);
        Ok(self)
    }

    /// Add an inert image to the current sheet's `table:shapes` container.
    pub fn add_sheet_image(&mut self, image: crate::OdfImage) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        if sheet.images.len() >= MAX_IMAGES_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "sheet exceeds {MAX_IMAGES_PER_SHEET} images"
            )));
        }
        let image = normalize_sheet_image(image, &sheet.name)?;
        sheet.images.push(image);
        Ok(self)
    }

    /// Remove an inert image from the current sheet.
    pub fn remove_sheet_image(&mut self, index: usize) -> Option<crate::OdfImage> {
        let sheet = self.sheets.last_mut()?;
        (index < sheet.images.len()).then(|| sheet.images.remove(index))
    }

    /// Append an image alternative to a frame group on the current sheet.
    ///
    /// `primary_image_index` is the flat sheet-image index whose
    /// `alternative_index` is zero. The alternative must describe the same
    /// frame after its sheet ownership is normalized.
    pub fn append_sheet_image_alternative(
        &mut self,
        primary_image_index: usize,
        image: crate::OdfImage,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        append_sheet_image_alternative(&mut sheet.images, &sheet.name, primary_image_index, image)?;
        Ok(self)
    }

    /// Insert an image alternative at a group-local alternative index.
    ///
    /// Valid alternative indices start at one and include the position after
    /// the group's current final alternative.
    pub fn insert_sheet_image_alternative(
        &mut self,
        primary_image_index: usize,
        alternative_index: usize,
        image: crate::OdfImage,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        insert_sheet_image_alternative(
            &mut sheet.images,
            &sheet.name,
            primary_image_index,
            alternative_index,
            image,
        )?;
        Ok(self)
    }

    /// Add a general drawing shape to the current sheet's `table:shapes` container.
    ///
    /// The shape is inert authoring metadata: styles it references are not
    /// resolved and geometry is not rendered. Picture and embedded-object
    /// frames must use their dedicated APIs instead.
    pub fn add_sheet_shape(&mut self, shape: crate::ods::SheetShape) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        if sheet.shapes.len() >= super::shape::MAX_SHAPES_PER_SHEET {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "sheet exceeds {} drawing shapes",
                super::shape::MAX_SHAPES_PER_SHEET
            )));
        }
        super::shape::validate_sheet_shape(&shape)?;
        sheet.shapes.push(shape);
        Ok(self)
    }

    /// Remove a general drawing shape from the current sheet.
    pub fn remove_sheet_shape(&mut self, index: usize) -> Option<crate::ods::SheetShape> {
        let sheet = self.sheets.last_mut()?;
        (index < sheet.shapes.len()).then(|| sheet.shapes.remove(index))
    }

    /// Remove a non-primary image alternative from a current-sheet frame group.
    pub fn remove_sheet_image_alternative(
        &mut self,
        primary_image_index: usize,
        alternative_index: usize,
    ) -> Result<crate::OdfImage> {
        let sheet = self.sheets.last_mut().ok_or_else(|| {
            litchi_core::Error::InvalidFormat("spreadsheet has no current sheet".to_string())
        })?;
        remove_sheet_image_alternative(
            &mut sheet.images,
            &sheet.name,
            primary_image_index,
            alternative_index,
        )
    }

    /// Merge a rectangular range in the current sheet.
    pub fn merge_cells(
        &mut self,
        start_row: usize,
        start_col: usize,
        row_span: usize,
        column_span: usize,
    ) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }
        let sheet = self.sheets.last_mut().expect("default sheet was added");
        merge_cell_range(&mut sheet.rows, start_row, start_col, row_span, column_span)?;
        Ok(self)
    }

    /// Remove a merge anchored at a cell in the current sheet.
    pub fn unmerge_cells(&mut self, start_row: usize, start_col: usize) -> bool {
        self.sheets
            .last_mut()
            .is_some_and(|sheet| unmerge_cell_range(&mut sheet.rows, start_row, start_col))
    }

    /// Select a specific sheet by index for subsequent operations
    ///
    /// # Arguments
    ///
    /// * `index` - Sheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.add_sheet("Sheet2")?;
    /// builder.select_sheet(0)?; // Go back to Sheet1
    /// builder.add_row_with_values(&["Data for Sheet1"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn select_sheet(&mut self, index: usize) -> Result<&mut Self> {
        if index >= self.sheets.len() {
            return Err(litchi_core::Error::Other(format!(
                "Sheet index {} out of bounds (have {} sheets)",
                index,
                self.sheets.len()
            )));
        }

        // Move the selected sheet to the end (current working sheet)
        let sheet = self.sheets.remove(index);
        self.sheets.push(sheet);

        Ok(self)
    }

    /// Add a row with typed cell values
    ///
    /// # Arguments
    ///
    /// * `cells` - Cell values for the row
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::{CellValue, SCell, SpreadsheetBuilder};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    ///
    /// let cells = vec![SCell::new(CellValue::Number(100.0), "100", 0, 0)];
    /// builder.add_row(cells)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_row(&mut self, cells: Vec<Cell>) -> Result<&mut Self> {
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
        }

        let row_index = if let Some(sheet) = self.sheets.last() {
            sheet.rows.len()
        } else {
            0
        };

        let row = Row {
            cells,
            index: row_index,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };

        if let Some(sheet) = self.sheets.last_mut() {
            sheet.rows.push(row);
        }

        Ok(self)
    }

    /// Add a Sheet element directly
    ///
    /// # Arguments
    ///
    /// * `sheet` - A complete `Sheet` element to add
    pub fn add_sheet_element(&mut self, sheet: Sheet) -> Result<&mut Self> {
        self.sheets.push(sheet);
        Ok(self)
    }

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
            || self.dde_links.iter().any(DdeLink::has_formulas)
    }

    fn has_annotations(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .any(Cell::has_annotation)
            || self.dde_links.iter().any(DdeLink::has_annotations)
    }

    fn has_hyperlinks(&self) -> bool {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .any(Cell::has_hyperlinks)
            || self.dde_links.iter().any(DdeLink::has_hyperlinks)
    }

    fn push_table_start(out: &mut String, sheet: &Sheet) -> Result<()> {
        out.push_str("<table:table table:name=\"");
        out.push_str(&escape_xml(&sheet.name));
        out.push('"');
        write_sheet_formatting_attributes(out, &sheet.style, &sheet.print_settings)?;
        write_sheet_attributes(out, &sheet.protection);
        out.push('>');
        write_sheet_preamble(
            out,
            sheet.title.as_deref(),
            sheet.description.as_deref(),
            sheet.table_source.as_ref(),
            sheet.dde_source.as_ref(),
            sheet.scenario.as_ref(),
        )?;
        write_sheet_options(out, &sheet.protection.options);
        Ok(())
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

    /// Generate the content.xml body for spreadsheet
    fn generate_content_body(&self) -> Result<String> {
        let mut cell_count = 0usize;
        for sheet in &self.sheets {
            for row in &sheet.rows {
                cell_count += row.cells.len();
            }
        }

        let mut estimated = 256usize;
        estimated += self.sheets.len() * 96;
        estimated += cell_count * 96;
        estimated += self.sheets.iter().map(|s| s.name.len()).sum::<usize>();
        estimated += self.named_definitions.len() * 128;
        estimated += self
            .sheets
            .iter()
            .flat_map(|s| s.rows.iter())
            .flat_map(|r| r.cells.iter())
            .map(|c| c.text.len())
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        write_calculation_settings(&mut body, self.calculation_settings.as_ref())?;
        write_content_validations(&mut body, &self.content_validations);
        write_label_ranges(&mut body, &self.label_ranges)?;

        for sheet in &self.sheets {
            Self::push_table_start(&mut body, sheet)?;
            // ODF 1.3 orders `table:shapes` ahead of the column and row groups.
            super::shape::write_table_shapes(&mut body, &sheet.images, &sheet.shapes)?;
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

            write_conditional_formats(&mut body, &sheet.conditional_formats)?;

            body.push_str("</table:table>");
        }

        write_named_definitions(
            &mut body,
            self.named_definitions
                .iter()
                .filter(|definition| definition.scope() == &NamedDefinitionScope::Global),
        );

        write_database_ranges(&mut body, &self.database_ranges)?;
        write_data_pilot_tables(&mut body, &self.data_pilot_tables)?;
        write_consolidation(&mut body, self.consolidation.as_ref())?;
        write_dde_links(&mut body, &self.dde_links)?;

        Ok(body)
    }

    /// Generate the complete content.xml for spreadsheet
    fn generate_content_xml(&self) -> Result<String> {
        let body = self.generate_content_body()?;

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
        let has_hyperlinks = self.has_hyperlinks();
        let has_sheet_drawings = self
            .sheets
            .iter()
            .any(|sheet| !sheet.images.is_empty() || !sheet.shapes.is_empty());
        let has_3d_shapes = self
            .sheets
            .iter()
            .any(|sheet| super::shape::sheet_shapes_use_3d(&sheet.shapes));
        let has_shape_event_listeners = self
            .sheets
            .iter()
            .any(|sheet| super::shape::sheet_shapes_have_event_listeners(&sheet.shapes));
        let has_validation_event_listeners = self.has_validation_event_listeners();
        let has_table_sources = self.sheets.iter().any(|sheet| {
            sheet.table_source.is_some()
                || sheet
                    .rows
                    .iter()
                    .flat_map(|row| &row.cells)
                    .any(|cell| cell.range_source.is_some())
        }) || self.dde_links.iter().any(DdeLink::has_table_sources);
        let has_protection_extensions = has_protection_extensions(
            &self.protection,
            self.sheets.iter().map(|sheet| &sheet.protection).chain(
                self.dde_links
                    .iter()
                    .map(|link| &link.cached_table.protection),
            ),
        );
        if has_annotations {
            out.push_str(r#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#);
        }
        if has_sheet_drawings && !has_annotations {
            out.push_str(r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink""#);
        }
        if has_3d_shapes {
            out.push_str(r#" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0""#);
        }
        if has_validation_event_listeners || has_shape_event_listeners {
            out.push_str(r#" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0""#);
        }
        if (has_validation_event_listeners || has_table_sources || has_hyperlinks)
            && !has_annotations
            && !has_sheet_drawings
        {
            out.push_str(r#" xmlns:xlink="http://www.w3.org/1999/xlink""#);
        }
        if has_protection_extensions && !has_annotations {
            out.push_str(r#" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#);
        }
        let has_conditional_formats = self
            .sheets
            .iter()
            .any(|sheet| !sheet.conditional_formats.is_empty());
        if has_conditional_formats {
            out.push_str(CALCEXT_NAMESPACE_DECLARATION);
        }
        out.push_str(r#" office:version="1.3"><office:font-face-decls/>"#);
        out.push_str(
            &rewrite_managed_cell_styles(
                None,
                &self.conditional_cell_styles,
                &self.table_cell_protection_styles,
            )?
            .xml,
        );
        out.push_str("<office:body><office:spreadsheet");
        write_spreadsheet_attributes(&mut out, &self.protection);
        out.push('>');
        out.push_str(&body);
        out.push_str(r#"</office:spreadsheet></office:body></office:document-content>"#);
        Ok(out)
    }

    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Build the spreadsheet and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        self.validate_named_definitions()?;
        self.validate_annotations()?;
        self.validate_hyperlinks()?;
        self.validate_content_validations()?;
        self.validate_database_ranges()?;
        crate::ods::data_pilot::validate_data_pilot_tables(&self.data_pilot_tables)?;
        self.validate_label_ranges()?;
        if let Some(consolidation) = &self.consolidation {
            consolidation.validate()?;
        }
        self.dde_links.iter().try_for_each(DdeLink::validate)?;
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.spreadsheet")?;

        // Add content.xml
        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml
        let mut styles_xml = OdfStructure::default_styles_xml();
        if !self.common_table_cell_styles.is_empty() {
            let mut common = String::from("<office:styles>");
            for name in &self.common_table_cell_styles {
                common.push_str("<style:style style:name=\"");
                common.push_str(&escape_xml(name));
                common.push_str("\" style:family=\"table-cell\"/>");
            }
            common.push_str("</office:styles>");
            styles_xml = styles_xml.replace("<office:styles/>", &common);
        }
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the spreadsheet to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODS file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::SpreadsheetBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = SpreadsheetBuilder::new();
    /// builder.add_sheet("Sheet1")?;
    /// builder.save("output.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation, DetectiveOperationKind,
        Spreadsheet, TableGroup, TableRange,
    };
    use tempfile::tempdir;

    fn alternative_test_image(source: crate::OdfImageSource) -> crate::OdfImage {
        crate::OdfImage {
            part: crate::OdfImagePart::Content,
            source,
            frame: Some(crate::OdfImageFrame {
                name: Some("hero".to_string()),
                width: Some("3cm".to_string()),
                height: Some("2cm".to_string()),
                ..Default::default()
            }),
            xml_id: None,
            filter_name: None,
            declared_media_type: None,
            link_type: Some("simple".to_string()),
            show: Some("embed".to_string()),
            actuate: Some("onLoad".to_string()),
            alternative_index: 0,
        }
    }

    #[test]
    fn test_spreadsheet_builder_new() {
        let builder = SpreadsheetBuilder::new();
        assert_eq!(builder.sheets.len(), 0);
    }

    #[test]
    fn test_spreadsheet_builder_default() {
        let builder: SpreadsheetBuilder = Default::default();
        assert_eq!(builder.sheets.len(), 0);
    }

    #[test]
    fn named_definitions_round_trip_through_ods_package() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sales & Tax").unwrap();
        builder
            .add_named_range(
                NamedRange::new(
                    "PrintableSales",
                    "$'Sales & Tax'.$A$1:.$B$5",
                    NamedDefinitionScope::sheet("Sales & Tax"),
                )
                .unwrap()
                .with_base_cell("$'Sales & Tax'.$A$1")
                .unwrap()
                .with_usage(crate::ods::NamedRangeUsage::PrintRange)
                .with_usage(crate::ods::NamedRangeUsage::Filter),
            )
            .unwrap();
        builder
            .add_named_expression(
                NamedExpression::new("TaxRate", "of:=0.2", NamedDefinitionScope::Global).unwrap(),
            )
            .unwrap();

        let bytes = builder.build().unwrap();
        let spreadsheet = crate::ods::Spreadsheet::from_bytes(bytes).unwrap();
        let range = spreadsheet
            .named_range(
                "PrintableSales",
                &NamedDefinitionScope::sheet("Sales & Tax"),
            )
            .unwrap();
        assert_eq!(range.usable_as.len(), 2);
        assert_eq!(
            range.base_cell_address.as_deref(),
            Some("$'Sales & Tax'.$A$1")
        );
        assert_eq!(
            spreadsheet
                .named_expression("TaxRate", &NamedDefinitionScope::Global)
                .unwrap()
                .expression,
            "of:=0.2"
        );
    }

    #[test]
    fn named_definitions_require_existing_scope_and_unique_name() {
        let mut builder = SpreadsheetBuilder::new();
        let local = NamedRange::new(
            "LocalName",
            "$Missing.$A$1",
            NamedDefinitionScope::sheet("Missing"),
        )
        .unwrap();
        assert!(builder.add_named_range(local).is_err());

        builder.add_sheet("Sheet1").unwrap();
        let first =
            NamedRange::new("Duplicate", "$Sheet1.$A$1", NamedDefinitionScope::Global).unwrap();
        let second =
            NamedExpression::new("Duplicate", "of:=1", NamedDefinitionScope::Global).unwrap();
        builder.add_named_range(first).unwrap();
        assert!(builder.add_named_expression(second).is_err());
    }

    #[test]
    fn content_validations_round_trip_through_ods_package() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Input").unwrap();
        builder.set_cell(1, 2, CellValue::Number(7.0)).unwrap();

        let mut validation = ContentValidation::new("whole&number").unwrap();
        validation
            .set_condition("of:cell-content-is-whole-number()")
            .unwrap();
        validation.base_cell_address = Some("$Input.$C$2".to_string());
        validation.allow_empty_cell = Some(false);
        validation.display_list = Some(crate::ods::ValidationDisplayList::Unsorted);
        validation.help_message = Some(crate::ods::ValidationMessage {
            title: Some("Required input".to_string()),
            display: Some(true),
            paragraphs: vec!["Enter  a\twhole\nnumber".to_string()],
        });
        validation.error_message = Some(crate::ods::ValidationErrorMessage {
            title: Some("Invalid value".to_string()),
            display: Some(true),
            message_type: Some(crate::ods::ValidationMessageType::Stop),
            paragraphs: vec!["Try again".to_string()],
        });
        builder.add_content_validation(validation.clone()).unwrap();
        builder.set_cell_validation(1, 2, "whole&number").unwrap();

        let mut spreadsheet =
            crate::ods::Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            spreadsheet.content_validation("whole&number"),
            Some(&validation)
        );
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(
            sheets[0].rows[1].cells[2].validation_name(),
            Some("whole&number")
        );
    }

    #[test]
    fn content_validations_reject_duplicates_and_missing_references() {
        let validation = ContentValidation::new("known").unwrap();
        let mut builder = SpreadsheetBuilder::new();
        builder.add_content_validation(validation.clone()).unwrap();
        assert!(builder.add_content_validation(validation).is_err());
        assert!(builder.set_cell_validation(0, 0, "missing").is_err());

        builder.add_sheet("Sheet1").unwrap();
        builder.set_cell(0, 0, CellValue::Empty).unwrap();
        builder.sheets[0].rows[0].cells[0].set_validation_name("missing");
        assert!(builder.build().is_err());
    }

    #[test]
    fn protection_metadata_round_trips_through_ods_package() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Protected").unwrap();
        builder.protection_mut().structure_protected = Some(true);
        builder.protection_mut().key = crate::ods::ProtectionKey {
            value: Some("doc&key=".to_string()),
            digest_algorithm: Some("http://www.w3.org/2001/04/xmlenc#sha256".to_string()),
            secondary_digest_algorithm: Some("http://www.w3.org/2000/09/xmldsig#sha1".to_string()),
        };
        let sheet = builder.sheet_protection_mut(0).unwrap();
        sheet.protected = Some(true);
        sheet.key = crate::ods::ProtectionKey {
            value: Some("sheet-key".to_string()),
            digest_algorithm: Some("http://www.w3.org/2000/09/xmldsig#sha1".to_string()),
            secondary_digest_algorithm: None,
        };
        sheet.options.select_protected_cells = Some(true);
        sheet.options.select_unprotected_cells = Some(false);
        sheet.options.insert_rows = Some(true);
        builder
            .set_cell_protection(2, 3, Some(false), Some(true))
            .unwrap();

        let mut spreadsheet =
            crate::ods::Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(spreadsheet.protection().structure_protected, Some(true));
        assert_eq!(
            spreadsheet.protection().key.value.as_deref(),
            Some("doc&key=")
        );
        assert_eq!(
            spreadsheet
                .protection()
                .key
                .secondary_digest_algorithm
                .as_deref(),
            Some("http://www.w3.org/2000/09/xmldsig#sha1")
        );
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].protection.protected, Some(true));
        assert_eq!(
            sheets[0].protection.options.select_unprotected_cells,
            Some(false)
        );
        assert_eq!(sheets[0].rows[2].cells[3].protect(), Some(false));
        assert_eq!(sheets[0].rows[2].cells[3].protected(), Some(true));
    }

    #[test]
    fn test_add_sheet() {
        let mut builder = SpreadsheetBuilder::new();
        let result = builder.add_sheet("TestSheet");
        assert!(result.is_ok());
        assert_eq!(builder.sheets.len(), 1);
        assert_eq!(builder.sheets[0].name, "TestSheet");
    }

    #[test]
    fn test_add_multiple_sheets() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_sheet("Sheet2").unwrap();
        builder.add_sheet("Sheet3").unwrap();
        assert_eq!(builder.sheets.len(), 3);
    }

    #[test]
    fn test_add_row_with_values() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_values(&["A", "B", "C"]).unwrap();

        assert_eq!(builder.sheets[0].rows.len(), 1);
        assert_eq!(builder.sheets[0].rows[0].cells.len(), 3);
        assert_eq!(builder.sheets[0].rows[0].cells[0].text, "A");
        assert_eq!(builder.sheets[0].rows[0].cells[1].text, "B");
        assert_eq!(builder.sheets[0].rows[0].cells[2].text, "C");
    }

    #[test]
    fn test_add_row_with_values_auto_sheet() {
        let mut builder = SpreadsheetBuilder::new();
        // No sheet added explicitly - should auto-create Sheet1
        builder.add_row_with_values(&["A", "B"]).unwrap();

        assert_eq!(builder.sheets.len(), 1);
        assert_eq!(builder.sheets[0].name, "Sheet1");
        assert_eq!(builder.sheets[0].rows.len(), 1);
    }

    #[test]
    fn test_add_row_with_numbers() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder
            .add_row_with_numbers(&[1.0, 2.5, std::f64::consts::PI])
            .unwrap();

        assert_eq!(builder.sheets[0].rows[0].cells.len(), 3);
        match &builder.sheets[0].rows[0].cells[0].value {
            CellValue::Number(n) => assert!((n - 1.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
        match &builder.sheets[0].rows[0].cells[1].value {
            CellValue::Number(n) => assert!((n - 2.5).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_add_row_with_cell_values() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder
            .add_row_with_cell_values(&[
                CellValue::Text("Product".to_string()),
                CellValue::Number(99.99),
                CellValue::Boolean(true),
            ])
            .unwrap();

        assert_eq!(builder.sheets[0].rows[0].cells.len(), 3);
        match &builder.sheets[0].rows[0].cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "Product"),
            _ => panic!("Expected Text"),
        }
        match &builder.sheets[0].rows[0].cells[2].value {
            CellValue::Boolean(b) => assert!(*b),
            _ => panic!("Expected Boolean"),
        }
    }

    #[test]
    fn test_set_cell() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.set_cell(0, 0, CellValue::Number(42.0)).unwrap();
        builder
            .set_cell(0, 1, CellValue::Text("Hello".to_string()))
            .unwrap();
        builder.set_cell(5, 2, CellValue::Boolean(false)).unwrap();

        // Verify cells
        match &builder.sheets[0].rows[0].cells[0].value {
            CellValue::Number(n) => assert!((n - 42.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
        assert_eq!(builder.sheets[0].rows[0].cells[1].text, "Hello");

        // Row 5 should exist with row index 5
        assert_eq!(builder.sheets[0].rows.len(), 6);
        assert_eq!(builder.sheets[0].rows[5].index, 5);
    }

    #[test]
    fn builder_authors_full_cell_hyperlinks_and_round_trips_them() {
        let mut builder = SpreadsheetBuilder::new();
        builder
            .add_sheet("Links")
            .unwrap()
            .set_cell_hyperlink(0, 0, "https://example.test/a?x=1&y=2", "Example & link")
            .unwrap();

        let mut internal = CellHyperlink::with_text("#Sheet2.B10", "Jump").unwrap();
        internal.name = Some("bookmark-link".to_string());
        internal.title = Some("Jump to bookmark".to_string());
        internal.target_frame_name = Some("_self".to_string());
        internal.show = Some(crate::TextHyperlinkShow::Replace);
        internal.actuate = Some(crate::TextHyperlinkActuate::OnRequest);
        builder.set_cell_hyperlink_data(0, 1, internal).unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let cells = &spreadsheet.sheets().unwrap()[0].rows[0].cells;
        assert_eq!(cells[0].text, "Example & link");
        assert_eq!(
            cells[0].hyperlink().unwrap().href(),
            "https://example.test/a?x=1&y=2"
        );
        let internal = cells[1].hyperlink().unwrap();
        assert_eq!(internal.href(), "#Sheet2.B10");
        assert_eq!(internal.name.as_deref(), Some("bookmark-link"));
        assert_eq!(internal.show, Some(crate::TextHyperlinkShow::Replace));
        assert_eq!(
            internal.actuate,
            Some(crate::TextHyperlinkActuate::OnRequest)
        );
        assert!(spreadsheet.content_xml().contains("xmlns:xlink="));
        assert!(spreadsheet.content_xml().contains("xlink:type=\"simple\""));

        let mut invalid = SpreadsheetBuilder::new();
        assert!(
            invalid
                .set_cell_hyperlink(0, 0, "", "missing target")
                .is_err()
        );
        assert!(invalid.sheets.is_empty());
    }

    #[test]
    fn builder_authors_mixed_text_hyperlinks_and_round_trips_ranges() {
        let text = "Visit docs & issues";
        let docs = text.find("docs").unwrap();
        let issues = text.find("issues").unwrap();
        let mut builder = SpreadsheetBuilder::new();
        builder
            .add_sheet("Links")
            .unwrap()
            .set_cell(0, 0, CellValue::Text(text.to_string()))
            .unwrap()
            .add_cell_hyperlink(0, 0, docs..docs + "docs".len(), "https://docs.example/")
            .unwrap();

        let mut issue_link = CellHyperlink::with_text("#Issues.A1", "issues").unwrap();
        issue_link.title = Some("Open issues".to_string());
        builder
            .add_cell_hyperlink_data(0, 0, issues..issues + "issues".len(), issue_link)
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let cell = &spreadsheet.sheets().unwrap()[0].rows[0].cells[0];
        assert_eq!(cell.text, text);
        assert_eq!(cell.hyperlinks().len(), 2);
        assert_eq!(cell.hyperlinks()[0].range(), docs..docs + "docs".len());
        assert_eq!(
            cell.hyperlinks()[1].range(),
            issues..issues + "issues".len()
        );
        assert_eq!(cell.hyperlinks()[1].title.as_deref(), Some("Open issues"));
        assert!(spreadsheet.content_xml().contains("xmlns:xlink="));
        assert!(spreadsheet.content_xml().contains("Visit <text:a"));

        let mut invalid = SpreadsheetBuilder::new();
        invalid
            .add_sheet("Invalid")
            .unwrap()
            .set_cell(0, 0, CellValue::Text(text.to_string()))
            .unwrap();
        assert!(
            invalid
                .add_cell_hyperlink_data(
                    0,
                    0,
                    docs..docs + "docs".len(),
                    CellHyperlink::with_text("https://bad.example/", "mismatch").unwrap(),
                )
                .is_err()
        );
    }

    #[test]
    fn test_set_cell_auto_sheet() {
        let mut builder = SpreadsheetBuilder::new();
        // No sheet added - should auto-create
        builder
            .set_cell(0, 0, CellValue::Text("Auto".to_string()))
            .unwrap();
        assert_eq!(builder.sheets.len(), 1);
        assert_eq!(builder.sheets[0].name, "Sheet1");
    }

    #[test]
    fn test_set_cell_formula() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.set_cell_formula(0, 0, "=SUM(A2:A10)").unwrap();
        builder.set_cell_formula(1, 0, "=A1+B1").unwrap();

        assert_eq!(
            builder.sheets[0].rows[0].cells[0].formula,
            Some("=SUM(A2:A10)".to_string())
        );
        assert_eq!(
            builder.sheets[0].rows[1].cells[0].formula,
            Some("=A1+B1".to_string())
        );
    }

    #[test]
    fn test_select_sheet() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_sheet("Sheet2").unwrap();
        builder.add_row_with_values(&["Data for Sheet2"]).unwrap();

        // Select Sheet1 (index 0)
        builder.select_sheet(0).unwrap();
        builder.add_row_with_values(&["Data for Sheet1"]).unwrap();

        // Sheet1 should now be at index 1 (last position after move)
        assert_eq!(builder.sheets[1].name, "Sheet1");
        assert_eq!(builder.sheets[1].rows.len(), 1);
    }

    #[test]
    fn test_select_sheet_out_of_bounds() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        let result = builder.select_sheet(5);
        assert!(result.is_err());
    }

    #[test]
    fn test_add_row() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();

        let cells = vec![
            Cell {
                value: CellValue::Number(100.0),
                text: "100".to_string(),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row: 0,
                col: 0,
            },
            Cell {
                value: CellValue::Text("Test".to_string()),
                text: "Test".to_string(),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row: 0,
                col: 1,
            },
        ];
        builder.add_row(cells).unwrap();

        assert_eq!(builder.sheets[0].rows.len(), 1);
        assert_eq!(builder.sheets[0].rows[0].cells.len(), 2);
    }

    #[test]
    fn test_add_sheet_element() {
        let mut builder = SpreadsheetBuilder::new();
        let sheet = Sheet {
            name: "CustomSheet".to_string(),
            rows: vec![],
            columns: vec![],
            column_structure: vec![],
            row_structure: vec![],
            style: Default::default(),
            print_settings: Default::default(),
            title: None,
            description: None,
            table_source: None,
            dde_source: None,
            scenario: None,
            conditional_formats: Vec::new(),
            images: Vec::new(),
            shapes: Vec::new(),
            protection: crate::ods::SheetProtection::default(),
        };
        builder.add_sheet_element(sheet).unwrap();

        assert_eq!(builder.sheets.len(), 1);
        assert_eq!(builder.sheets[0].name, "CustomSheet");
    }

    #[test]
    fn test_set_metadata() {
        let mut builder = SpreadsheetBuilder::new();
        let metadata = Metadata {
            title: Some("Test Title".to_string()),
            author: Some("Test Author".to_string()),
            ..Default::default()
        };
        builder.set_metadata(metadata);

        assert_eq!(builder.metadata.title, Some("Test Title".to_string()));
        assert_eq!(builder.metadata.author, Some("Test Author".to_string()));
    }

    #[test]
    fn test_has_formulas() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        assert!(!builder.has_formulas());

        builder.set_cell_formula(0, 0, "=A1+B1").unwrap();
        assert!(builder.has_formulas());
    }

    #[test]
    fn test_sheet_max_cols() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_values(&["A", "B", "C", "D"]).unwrap();
        builder.add_row_with_values(&["X", "Y"]).unwrap();

        let max_cols = SpreadsheetBuilder::sheet_max_cols(&builder.sheets[0]);
        assert_eq!(max_cols, 4);
    }

    #[test]
    fn test_generate_content_body() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("TestSheet").unwrap();
        builder.add_row_with_values(&["A", "B"]).unwrap();

        let content = builder.generate_content_body().unwrap();
        assert!(content.contains("TestSheet"));
        assert!(content.contains("table:table"));
        assert!(content.contains("table:table-row"));
    }

    #[test]
    fn test_generate_content_xml() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_numbers(&[42.0]).unwrap();

        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.starts_with(r#"<?xml version="1.0" encoding="UTF-8"?>"#));
        assert!(xml.contains("office:document-content"));
        assert!(xml.contains("Sheet1"));
        assert!(xml.contains("42")); // Check number value
    }

    #[test]
    fn test_generate_content_xml_with_formula() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.set_cell_formula(0, 0, "=SUM(A1:A10)").unwrap();

        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains("xmlns:of="));
        assert!(xml.contains("table:formula"));
        assert!(xml.contains("=SUM(A1:A10)"));
    }

    #[test]
    fn test_generate_meta_xml() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.metadata.title = Some("Test Document".to_string());
        builder.metadata.author = Some("John Doe".to_string());

        let meta_xml = builder.generate_meta_xml();
        assert!(meta_xml.contains("office:document-meta"));
        assert!(meta_xml.contains("Litchi/"));
        assert!(meta_xml.contains("Test Document"));
        assert!(meta_xml.contains("John Doe"));
    }

    #[test]
    fn test_build() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_values(&["Test"]).unwrap();

        let result = builder.build();
        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(!bytes.is_empty());
        // Check it's a valid ZIP (starts with PK)
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn cell_annotations_round_trip_through_ods_package() {
        let mut annotation = CellAnnotation::new("plain & safe");
        annotation.set_creator(Some("Ada"));
        annotation.set_date(Some("2026-07-13T12:34:56Z"));
        annotation.set_display(Some(true));
        annotation
            .set_attribute("draw:style-name", "comment-style")
            .unwrap();
        annotation.set_attribute("svg:width", "4.5cm").unwrap();

        let mut span = crate::ods::AnnotationElement::new("text:span").unwrap();
        span.set_attribute("text:style-name", "Emphasis").unwrap();
        span.push_text("rich");
        let mut paragraph = crate::ods::AnnotationElement::new("text:p").unwrap();
        paragraph.push_element(span);
        annotation.push_element(paragraph);

        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Notes").unwrap();
        builder
            .set_cell(0, 0, CellValue::Text("value".to_string()))
            .unwrap()
            .set_cell_annotation(0, 0, annotation)
            .unwrap();
        builder
            .set_cell_annotation(1, 2, CellAnnotation::new("empty-cell note"))
            .unwrap();

        let mut spreadsheet =
            crate::ods::Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        let first = sheets[0].rows[0].cells[0].annotation().unwrap();
        assert_eq!(first.creator().as_deref(), Some("Ada"));
        assert_eq!(first.date().as_deref(), Some("2026-07-13T12:34:56Z"));
        assert_eq!(first.display(), Some(true));
        assert_eq!(first.attribute("draw:style-name"), Some("comment-style"));
        assert_eq!(first.attribute("svg:width"), Some("4.5cm"));
        assert_eq!(first.text(), "plain & safe\nrich");
        assert_eq!(
            sheets[0].rows[1].cells[2].annotation().unwrap().text(),
            "empty-cell note"
        );
    }

    #[test]
    fn test_save() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.ods");

        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_values(&["A", "B", "C"]).unwrap();

        let result = builder.save(&path);
        assert!(result.is_ok());
        assert!(path.exists());

        // Verify the file is a valid ZIP
        let bytes = std::fs::read(&path).unwrap();
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_chained_builder_api() {
        let mut builder = SpreadsheetBuilder::new();
        builder
            .add_sheet("Data")
            .unwrap()
            .add_row_with_values(&["Name", "Age"])
            .unwrap()
            .add_row_with_values(&["Alice", "30"])
            .unwrap()
            .add_row_with_numbers(&[25.0, 35.0])
            .unwrap();

        assert_eq!(builder.sheets[0].rows.len(), 3);
    }

    #[test]
    fn test_cell_value_types_in_content() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder
            .add_row_with_cell_values(&[
                CellValue::Text("Text".to_string()),
                CellValue::Number(123.45),
                CellValue::Currency(100.0, "USD".to_string()),
                CellValue::Percentage(0.5),
                CellValue::Boolean(true),
                CellValue::Date("2024-03-15".to_string()),
                CellValue::Time("PT12H30M00S".to_string()),
            ])
            .unwrap();

        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains(r#"office:value-type="string""#));
        assert!(xml.contains(r#"office:value-type="float""#));
        assert!(xml.contains(r#"office:value-type="currency""#));
        assert!(xml.contains(r#"office:value-type="percentage""#));
        assert!(xml.contains(r#"office:value-type="boolean""#));
        assert!(xml.contains(r#"office:value-type="date""#));
        assert!(xml.contains(r#"office:value-type="time""#));
    }

    #[test]
    fn test_empty_cell_with_formula() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.set_cell_formula(0, 0, "=IF(TRUE,1,0)").unwrap();

        let xml = builder.generate_content_xml().unwrap();
        // An unevaluated formula must not pretend that zero is its cached value.
        assert!(xml.contains(r#"table:formula="=IF(TRUE,1,0)""#));
        assert!(!xml.contains("office:value-type="));
        assert!(!xml.contains("<text:p>0</text:p>"));
    }

    #[test]
    fn builder_creates_and_removes_merged_ranges_safely() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Merged").unwrap();
        builder
            .set_cell(0, 0, CellValue::Text("anchor".to_string()))
            .unwrap();
        builder.merge_cells(0, 0, 2, 3).unwrap();
        assert_eq!(builder.sheets[0].rows[0].cells[0].span(), Some((2, 3)));
        assert_eq!(
            builder.sheets[0].rows[1].cells[2].merge(),
            crate::ods::CellMerge::Covered
        );
        assert!(builder.unmerge_cells(0, 0));
        assert!(!builder.unmerge_cells(0, 0));

        builder
            .set_cell(0, 1, CellValue::Text("occupied".to_string()))
            .unwrap();
        assert!(builder.merge_cells(0, 0, 1, 2).is_err());
    }

    #[test]
    fn builder_sets_and_clears_matrix_formula_spans() {
        let mut builder = SpreadsheetBuilder::new();
        assert!(builder.set_cell_matrix_span(0, 0, 0, 2).is_err());
        builder.set_cell_formula(0, 0, "of:=SEQUENCE(3;2)").unwrap();
        builder.set_cell_matrix_span(0, 0, 3, 2).unwrap();
        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains(r#"table:number-matrix-rows-spanned="3""#));
        assert!(xml.contains(r#"table:number-matrix-columns-spanned="2""#));
        assert!(builder.clear_cell_matrix_span(0, 0));
        assert!(!builder.clear_cell_matrix_span(0, 0));
    }

    #[test]
    fn builder_writes_row_and_column_metadata() {
        let mut builder = SpreadsheetBuilder::new();
        assert!(
            builder
                .set_row_metadata(
                    MAX_EXPANDED_ROWS_PER_SHEET,
                    None,
                    None,
                    TableVisibility::Visible,
                )
                .is_err()
        );
        assert!(
            builder
                .set_column_metadata(
                    MAX_EXPANDED_COLUMNS_PER_SHEET,
                    None,
                    None,
                    TableVisibility::Visible,
                )
                .is_err()
        );
        builder
            .set_row_metadata(
                2,
                Some("RowStyle".to_string()),
                Some("RowCell".to_string()),
                TableVisibility::Filter,
            )
            .unwrap()
            .set_column_metadata(
                1,
                Some("ColumnStyle".to_string()),
                Some("ColumnCell".to_string()),
                TableVisibility::Collapse,
            )
            .unwrap();
        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains(r#"table:style-name="RowStyle""#));
        assert!(xml.contains(r#"table:default-cell-style-name="ColumnCell""#));
        assert!(xml.contains(r#"table:visibility="filter""#));
        assert!(xml.contains(r#"table:visibility="collapse""#));

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].rows[2].style_name.as_deref(), Some("RowStyle"));
        assert_eq!(sheets[0].rows[2].visibility, TableVisibility::Filter);
        assert_eq!(
            sheets[0].columns[1].default_cell_style_name.as_deref(),
            Some("ColumnCell")
        );
        assert_eq!(sheets[0].columns[1].visibility, TableVisibility::Collapse);
    }

    #[test]
    fn builder_declares_default_columns_after_explicit_metadata() {
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_cell(0, 3, CellValue::Text("wide".to_string()))
            .unwrap()
            .set_column_metadata(
                0,
                Some("FirstColumn".to_string()),
                None,
                TableVisibility::Visible,
            )
            .unwrap();

        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains(r#"table:style-name="FirstColumn""#));
        assert!(xml.contains(r#"table:number-columns-repeated="3""#));
    }

    #[test]
    fn builder_round_trips_nested_groups_and_headers() {
        let structure = vec![TableStructure::Group(TableGroup {
            display: false,
            children: vec![
                TableStructure::Header(TableRange::new(0, 1).unwrap()),
                TableStructure::Group(TableGroup {
                    display: true,
                    children: vec![TableStructure::Range(TableRange::new(1, 3).unwrap())],
                }),
            ],
        })];
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_column_structure(structure.clone())
            .unwrap()
            .set_row_structure(structure.clone())
            .unwrap();
        let xml = builder.generate_content_xml().unwrap();
        assert!(xml.contains(r#"<table:table-column-group table:display="false">"#));
        assert!(xml.contains("<table:table-header-columns>"));
        assert!(xml.contains(r#"<table:table-row-group table:display="false">"#));
        assert!(xml.contains("<table:table-header-rows>"));

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].column_structure, structure);
        assert_eq!(sheets[0].row_structure, sheets[0].column_structure);
    }

    #[test]
    fn builder_rejects_invalid_table_structure() {
        let overlapping = vec![
            TableStructure::Range(TableRange::new(0, 2).unwrap()),
            TableStructure::Header(TableRange::new(1, 3).unwrap()),
        ];
        let discontinuous_group = vec![TableStructure::Group(TableGroup {
            display: true,
            children: vec![
                TableStructure::Range(TableRange::new(0, 1).unwrap()),
                TableStructure::Range(TableRange::new(2, 3).unwrap()),
            ],
        })];
        let mut builder = SpreadsheetBuilder::new();
        assert!(builder.set_row_structure(overlapping).is_err());
        assert!(builder.set_column_structure(discontinuous_group).is_err());

        let mut too_deep = TableStructure::Range(TableRange::new(0, 1).unwrap());
        for _ in 0..300 {
            too_deep = TableStructure::Group(TableGroup {
                display: true,
                children: vec![too_deep],
            });
        }
        assert!(builder.set_row_structure(vec![too_deep]).is_err());
    }

    #[test]
    fn builder_round_trips_sheet_style_and_print_settings() {
        let style = SheetStyle {
            style_name: Some("Sheet&Style".to_string()),
            template_name: Some("TemplateOne".to_string()),
            usage: crate::SheetStyleUsage {
                use_first_row_styles: Some(true),
                use_last_column_styles: Some(false),
                ..crate::SheetStyleUsage::default()
            },
        };
        let print = SheetPrintSettings::new(
            false,
            vec![
                "$Sheet1.$A$1:$B$2".to_string(),
                "'Q1 Sales'.$C$3:$D$4".to_string(),
            ],
        )
        .unwrap();
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_sheet_style(style.clone())
            .unwrap()
            .set_sheet_print_settings(print.clone())
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].style, style);
        assert_eq!(sheets[0].print_settings, print);
    }

    #[test]
    fn builder_rejects_invalid_individual_print_ranges() {
        let mut builder = SpreadsheetBuilder::new();
        assert!(
            builder
                .set_sheet_print_settings(SheetPrintSettings {
                    printable: true,
                    ranges: vec!["A1:B2 C3:D4".to_string()],
                })
                .is_err()
        );
        assert!(SheetPrintSettings::new(true, vec!["'Unclosed Sheet.$A$1".to_string()]).is_err());
    }

    #[test]
    fn builder_round_trips_sheet_text_and_scenario() {
        let mut scenario =
            SheetScenario::new(vec!["'Q1 Sales'.$A$1:$B$2".to_string()], true).unwrap();
        scenario.border_color = Some("#A1b2C3".to_string());
        scenario.copy_formulas = Some(true);
        scenario.comment = Some("Best & worst".to_string());
        let mut source = SheetTableSource::new("../Source & Data.ods");
        source.mode = Some(crate::ods::TableSourceMode::CopyResultsOnly);
        source.table_name = Some("Source <Sheet>".to_string());
        source.refresh_delay = Some("PT30M".to_string());
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_sheet_title(Some("Quarter & Forecast".to_string()))
            .unwrap()
            .set_sheet_description(Some("Best < worst".to_string()))
            .unwrap()
            .set_sheet_table_source(Some(source.clone()))
            .unwrap()
            .set_sheet_scenario(Some(scenario.clone()))
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].title.as_deref(), Some("Quarter & Forecast"));
        assert_eq!(sheets[0].description.as_deref(), Some("Best < worst"));
        assert_eq!(sheets[0].table_source.as_ref(), Some(&source));
        assert_eq!(sheets[0].scenario.as_ref(), Some(&scenario));
    }

    #[test]
    fn builder_rejects_invalid_scenario_metadata() {
        let mut scenario = SheetScenario::new(vec![".A1:.B2".to_string()], false).unwrap();
        scenario.border_color = Some("#12345Z".to_string());
        assert!(
            SpreadsheetBuilder::new()
                .set_sheet_scenario(Some(scenario))
                .is_err()
        );
        let mut source = SheetTableSource::new("source.ods");
        source.refresh_delay = Some("every hour".to_string());
        assert!(
            SpreadsheetBuilder::new()
                .set_sheet_table_source(Some(source))
                .is_err()
        );
    }

    #[test]
    fn builder_round_trips_inert_cell_range_sources() {
        let mut source = CellRangeSource::new("Prices & Tax", "../prices&tax.ods", 5, 2).unwrap();
        source.set_filter_options(Some("locale=en_US".to_string()));
        source.set_refresh_delay(Some("PT10M".to_string())).unwrap();

        let mut builder = SpreadsheetBuilder::new();
        builder
            .add_sheet("Imports")
            .unwrap()
            .set_cell(1, 2, CellValue::Text("cached".to_string()))
            .unwrap()
            .set_cell_range_source(1, 2, source.clone())
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        assert_eq!(sheets[0].rows[1].cells[2].range_source(), Some(&source));
        assert_eq!(
            sheets[0].rows[1].cells[2].value,
            CellValue::Text("cached".to_string())
        );
    }

    #[test]
    fn builder_round_trips_cell_detective_metadata() {
        let mut detective = CellDetective::new();
        detective
            .add_highlighted_range(
                DetectiveHighlightedRange::valid(
                    Some(".A1:.A4".to_string()),
                    DetectiveDirection::FromSameTable,
                    Some(false),
                )
                .unwrap(),
            )
            .add_operation(DetectiveOperation::new(
                DetectiveOperationKind::TraceDependents,
                2,
            ));
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_cell(0, 0, CellValue::Number(11.0))
            .unwrap()
            .set_cell_detective(0, 0, detective.clone())
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let cell = &spreadsheet.sheets().unwrap()[0].rows[0].cells[0];
        assert_eq!(cell.value, CellValue::Number(11.0));
        assert_eq!(cell.detective(), Some(&detective));
    }

    #[test]
    fn builder_round_trips_calculation_settings() {
        let settings = CalculationSettings {
            case_sensitive: Some(true),
            use_regular_expressions: Some(false),
            null_year: std::num::NonZeroUsize::new(1930),
            null_date: Some(crate::CalculationNullDate {
                value_type_date: true,
                date_value: Some("1899-12-30Z".to_string()),
            }),
            iteration: Some(crate::CalculationIteration {
                status: Some(crate::IterationStatus::Enable),
                steps: std::num::NonZeroUsize::new(50),
                maximum_difference: Some("INF".to_string()),
            }),
            ..CalculationSettings::default()
        };
        let mut builder = SpreadsheetBuilder::new();
        builder
            .set_calculation_settings(Some(settings.clone()))
            .unwrap();
        let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(spreadsheet.calculation_settings(), Some(&settings));
    }

    #[test]
    fn builder_sheet_image_alternatives_round_trip_and_reindex() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Pictures").unwrap();
        builder
            .add_sheet_image(alternative_test_image(crate::OdfImageSource::Linked {
                href: "https://example.test/primary.svg".to_string(),
            }))
            .unwrap();
        builder
            .append_sheet_image_alternative(
                0,
                alternative_test_image(crate::OdfImageSource::Inline {
                    bytes: vec![1, 2, 3, 4],
                    ignored_href: Some("ignored.png".to_string()),
                }),
            )
            .unwrap();
        builder
            .insert_sheet_image_alternative(
                0,
                1,
                alternative_test_image(crate::OdfImageSource::Linked {
                    href: "https://example.test/fallback.webp".to_string(),
                }),
            )
            .unwrap();
        let removed = builder.remove_sheet_image_alternative(0, 1).unwrap();
        assert!(matches!(
            removed.source,
            crate::OdfImageSource::Linked { .. }
        ));

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let sheets = spreadsheet.sheets().unwrap();
        let images = sheets[0].images();
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].alternative_index, 0);
        assert_eq!(images[1].alternative_index, 1);
        assert_eq!(images[0].frame, images[1].frame);
        assert!(matches!(
            images[0].source,
            crate::OdfImageSource::Linked { .. }
        ));
        assert_eq!(images[1].inline_bytes(), Some(&[1, 2, 3, 4][..]));
    }

    #[test]
    fn builder_sheet_image_alternative_validation_is_atomic() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Pictures").unwrap();
        builder
            .add_sheet_image(alternative_test_image(crate::OdfImageSource::Linked {
                href: "primary.svg".to_string(),
            }))
            .unwrap();
        builder
            .append_sheet_image_alternative(
                0,
                alternative_test_image(crate::OdfImageSource::Linked {
                    href: "fallback.png".to_string(),
                }),
            )
            .unwrap();
        let before = builder.sheets[0].images.clone();

        assert!(
            builder
                .append_sheet_image_alternative(
                    1,
                    alternative_test_image(crate::OdfImageSource::Linked {
                        href: "wrong-group.png".to_string(),
                    }),
                )
                .is_err()
        );
        assert!(
            builder
                .insert_sheet_image_alternative(
                    0,
                    0,
                    alternative_test_image(crate::OdfImageSource::Linked {
                        href: "wrong-index.png".to_string(),
                    }),
                )
                .is_err()
        );
        let mut wrong_frame = alternative_test_image(crate::OdfImageSource::Linked {
            href: "wrong-frame.png".to_string(),
        });
        wrong_frame.frame.as_mut().unwrap().width = Some("4cm".to_string());
        assert!(
            builder
                .append_sheet_image_alternative(0, wrong_frame)
                .is_err()
        );
        let mut no_frame = alternative_test_image(crate::OdfImageSource::Linked {
            href: "no-frame.png".to_string(),
        });
        no_frame.frame = None;
        assert!(builder.append_sheet_image_alternative(0, no_frame).is_err());
        assert_eq!(builder.sheets[0].images, before);
    }
}
