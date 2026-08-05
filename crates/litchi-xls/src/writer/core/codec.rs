use super::super::biff::AutoFilterConditionWrite;
use super::super::formatting::{CellStyle, ExtendedFormat, FormattingManager};
use super::comment;
use super::model::{
    a1_cell, prepare_data_validation, validate_list_object_style, validate_pivot_table_config,
};
use super::named_range;
use super::*;
use crate::encryption::{WriterEncryption, validate_writer_encryption};
use crate::error::{Error, Result};
use crate::{EncryptionProfile, ListObject};
use std::collections::{HashMap, HashSet};
use zeroize::Zeroizing;
impl Writer {
    /// Create a new XLS writer
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            shared_strings: Vec::new(),
            string_map: HashMap::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            sst_total: 0,
            fmt: FormattingManager::new(),
            workbook_protection: None,
            file_sharing: None,
            use_1904_dates: false,
            calculation_settings: CalculationSettings::default(),
            vba_metadata: None,
            environment_options: WorkbookEnvironmentOptions::default(),
            workbook_window_options: WorkbookWindowOptions::default(),
            function_group_options: FunctionGroupOptions::default(),
            external_workbooks: Vec::new(),
            external_names: Vec::new(),
            add_in_functions: Vec::new(),
            dde_or_ole_links: Vec::new(),
            custom_table_styles: None,
            book_ext: None,
            theme: None,
            mdx_metadata: None,
            real_time_data: Vec::new(),
            web_publications: Vec::new(),
            xf_extensions: Vec::new(),
            style_extensions: Vec::new(),
            encryption: None,
        }
    }

    /// Configure BIFF8 password-to-open encryption for subsequent writes.
    ///
    /// Validation is atomic: an invalid password or profile leaves the current
    /// encryption configuration unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<()> {
        let password = password.into();
        validate_writer_encryption(&password, profile)?;
        self.encryption = Some(WriterEncryption {
            password: Zeroizing::new(password),
            profile,
        });
        Ok(())
    }

    /// Remove password-to-open encryption from subsequent writes.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured password-to-open encryption profile.
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name (max 31 characters)
    ///
    /// # Returns
    ///
    /// * `Result<usize, Error>` - Worksheet index or error
    pub fn add_worksheet(&mut self, name: &str) -> Result<usize> {
        // Validate worksheet name
        if name.is_empty() || name.len() > 31 {
            return Err(Error::InvalidData(
                "Worksheet name must be 1-31 characters".to_string(),
            ));
        }

        // Check for duplicate names
        if self.worksheets.iter().any(|ws| ws.name == name) {
            return Err(Error::InvalidData(format!(
                "Worksheet '{}' already exists",
                name
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        self.synchronize_workbook_window_selection();
        Ok(index)
    }

    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> Result<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::String(value.to_string()), format_id)
    }

    /// Write a number value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Numeric value
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> Result<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    pub fn write_number_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: f64,
        format_id: u16,
    ) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::InvalidData(
                "cell number must be finite for BIFF8 serialization".to_string(),
            ));
        }
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Number(value), format_id)
    }

    /// Write a boolean value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Boolean value
    pub fn write_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> Result<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    pub fn write_boolean_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Boolean(value), format_id)
    }

    /// Write a formula to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - Formula string (without leading '=')
    ///
    /// The supported BIFF8 formula subset includes constants, cell/range
    /// references, arithmetic/comparison operators, and built-in functions
    /// recognized by [`FormulaTokenizer`](crate::writer::FormulaTokenizer).
    pub fn write_formula(&mut self, sheet: usize, row: u32, col: u16, formula: &str) -> Result<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(
            sheet,
            pos,
            CellValue::Formula(formula.to_string()),
            format_id,
        )
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This is a thin wrapper around the internal `FormattingManager`
    /// and mirrors Apache POI's `HSSFDataFormat.getFormat` API. The
    /// returned index can be stored in `ExtendedFormat.format_index`
    /// to apply number formats to cells.
    pub fn register_number_format(&mut self, pattern: &str) -> u16 {
        self.fmt.register_number_format(pattern)
    }

    /// Register a reusable cell style defined by `CellStyle`.
    ///
    /// The returned identifier can be passed to the `write_*_with_format`
    /// methods to apply this style to individual cells.
    pub fn add_cell_style(&mut self, style: CellStyle) -> u16 {
        self.fmt.register_cell_style(style)
    }

    pub fn add_cell_format(&mut self, format: ExtendedFormat) -> u16 {
        self.fmt.add_format(format)
    }

    /// Installs a complete custom table-style family.
    ///
    /// Validation happens before assignment, so an error leaves the current
    /// writer configuration unchanged.
    pub fn set_custom_table_styles(&mut self, styles: CustomTableStyles) -> Result<()> {
        styles.validate(&self.fmt)?;
        self.custom_table_styles = Some(styles);
        Ok(())
    }

    /// Removes caller-defined table styles and restores the default write path.
    pub fn clear_custom_table_styles(&mut self) {
        self.custom_table_styles = None;
    }

    /// Adds a legacy BIFF8 worksheet table and writes its header captions.
    pub fn add_list_object(&mut self, sheet: usize, table: ListObject) -> Result<()> {
        table.validate()?;
        let style = table.style().ok_or_else(|| {
            Error::InvalidData("validated table is missing its style".to_string())
        })?;
        validate_list_object_style(style.name(), self.custom_table_styles.as_ref())?;
        if self
            .worksheets
            .iter()
            .flat_map(|worksheet| &worksheet.list_objects)
            .any(|existing| {
                existing.id() == table.id() || existing.name().eq_ignore_ascii_case(table.name())
            })
        {
            return Err(Error::InvalidData(
                "table identifier or name collides within the workbook".to_string(),
            ));
        }
        if self
            .defined_names
            .iter()
            .any(|name| name.name.eq_ignore_ascii_case(table.name()))
            || self
                .defined_name_records
                .iter()
                .any(|(name, _)| name.name.eq_ignore_ascii_case(table.name()))
        {
            return Err(Error::InvalidData(
                "table name collides with a workbook defined name".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet
            .list_objects
            .iter()
            .any(|existing| existing.range().overlaps(table.range()))
        {
            return Err(Error::InvalidData(
                "table ranges overlap within the worksheet".to_string(),
            ));
        }
        if worksheet.auto_filter.is_some_and(|filter| {
            u32::from(table.range().first_row()) <= filter.last_row
                && filter.first_row <= u32::from(table.range().last_row())
                && table.range().first_column() <= filter.last_col
                && filter.first_col <= table.range().last_column()
        }) {
            return Err(Error::InvalidData(
                "table range overlaps the worksheet AutoFilter".to_string(),
            ));
        }
        let mut header_cells = Vec::new();
        for (offset, column) in table
            .columns()
            .iter()
            .enumerate()
            .filter(|_| table.has_header_row())
        {
            let key = (
                u32::from(table.range().first_row()),
                table.range().first_column() + offset as u16,
            );
            if let Some(cell) = worksheet.cells.get(&key)
                && !matches!(&cell.value, CellValue::String(value) if value == column.name())
            {
                return Err(Error::InvalidData(
                    "table header collides with a different cell value".to_string(),
                ));
            } else if !worksheet.cells.contains_key(&key) {
                header_cells.push(WritableCell::new(
                    CellPos::try_new(key.0, key.1)?,
                    CellValue::String(column.name().to_string()),
                    0,
                    None,
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.include_list_object_range(table.range());
        for cell in header_cells {
            worksheet.add_cell(cell);
        }
        worksheet.list_objects.push(table);
        Ok(())
    }

    pub fn clear_list_objects(&mut self, sheet: usize) -> Result<()> {
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .list_objects
            .clear();
        Ok(())
    }

    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 characters (Lbl.cch is a byte).
    fn validate_defined_name(name: &str) -> Result<()> {
        if name.is_empty() {
            return Err(Error::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        let char_count = name.chars().count();
        if char_count > u8::MAX as usize {
            return Err(Error::InvalidData(
                "Defined name must be at most 255 characters".to_string(),
            ));
        }

        Ok(())
    }

    fn hash_password(password: &str) -> u16 {
        let bytes = password.as_bytes();
        if bytes.is_empty() {
            return 0;
        }

        let mut hash: u16 = 0;
        for &b in bytes.iter().rev() {
            let high_bit = (hash >> 14) & 0x0001;
            hash = ((hash << 1) & 0x7FFF) | high_bit;
            hash ^= b as u16;
        }

        let high_bit = (hash >> 14) & 0x0001;
        hash = ((hash << 1) & 0x7FFF) | high_bit;
        hash ^= bytes.len() as u16;
        hash ^= 0xCE4B;
        hash
    }

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(Error::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        worksheet.add_hyperlink(Hyperlink {
            first_row: row,
            last_row: row,
            first_col: col,
            last_col: col,
            url: url.to_string(),
        });

        Ok(())
    }

    /// Add a canonical, macro-inert BIFF8 comment to a cell.
    pub fn add_comment(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
    ) -> Result<()> {
        self.add_comment_with_options(
            sheet,
            row,
            col,
            author,
            text,
            CommentWriteOptions::default(),
        )
    }

    /// Add a canonical BIFF8 comment with explicit visibility, anchor, rich runs, and GUID options.
    pub fn add_comment_with_options(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
        options: CommentWriteOptions,
    ) -> Result<()> {
        let (row, column) = comment::validate_comment(row, col, author, text, &options)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.comments.len() >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 canonical comment shapes".to_string(),
            ));
        }
        if worksheet
            .comments
            .iter()
            .any(|comment| comment.row == row && comment.column == column)
        {
            return Err(Error::InvalidData(
                "a cell cannot contain more than one comment".to_string(),
            ));
        }
        if let Some(guid) = options.guid
            && worksheet
                .comments
                .iter()
                .any(|comment| comment.options.guid == Some(guid))
        {
            return Err(Error::InvalidData(
                "comment GUID override is duplicated on the worksheet".to_string(),
            ));
        }
        let comment = comment::WritableComment::try_new(row, column, author, text, options)?;
        worksheet.add_comment(comment)
    }

    /// Add a validated, macro-inert primitive shape and return its worksheet OBJ identifier.
    pub fn add_shape(&mut self, sheet: usize, mut shape: ShapeWrite) -> Result<u16> {
        shape.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let reserved = collect_reserved_object_ids(worksheet, 0)?;
        let object_count =
            reserved
                .len()
                .checked_add(worksheet.comments.len())
                .ok_or(Error::Allocation(
                    "computing the worksheet drawing-object count",
                ))?;
        if object_count >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        let object_id = if let Some(requested) = shape.object_id {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
            requested
        } else {
            (1..u16::MAX)
                .find(|candidate| !reserved.contains(candidate))
                .ok_or_else(|| {
                    Error::InvalidData("worksheet object IDs are exhausted".to_string())
                })?
        };
        worksheet
            .shapes
            .try_reserve(1)
            .map_err(|_| Error::Allocation("reserving worksheet shape storage"))?;
        shape.object_id = Some(object_id);
        worksheet.shapes.push(shape);
        Ok(object_id)
    }

    /// Remove a primitive by its assigned OBJ identifier.
    pub fn remove_shape(&mut self, sheet: usize, object_id: u16) -> Result<ShapeWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shapes
            .iter()
            .position(|shape| shape.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shapes.remove(index))
    }

    /// Remove all writable primitive shapes from a worksheet.
    pub fn clear_shapes(&mut self, sheet: usize) -> Result<usize> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let count = worksheet.shapes.len();
        worksheet.shapes.clear();
        Ok(count)
    }

    /// Add a validated shape group and return the group's worksheet OBJ identifier.
    ///
    /// The group consumes one object ID for itself plus one per child; assigned
    /// child identifiers are stored back into the group before it is retained.
    pub fn add_shape_group(&mut self, sheet: usize, mut group: ShapeGroupWrite) -> Result<u16> {
        group.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let group_count = group.object_count()?;
        let mut reserved = collect_reserved_object_ids(worksheet, group_count)?;
        let object_count = reserved
            .len()
            .checked_add(worksheet.comments.len())
            .and_then(|count| count.checked_add(group_count))
            .ok_or(Error::Allocation(
                "computing the worksheet drawing-object count",
            ))?;
        if object_count > 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        for requested in group_object_ids(&group) {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
        }
        for requested in group_object_ids(&group) {
            reserved.insert(requested);
        }
        worksheet
            .shape_groups
            .try_reserve(1)
            .map_err(|_| Error::Allocation("reserving worksheet shape-group storage"))?;
        let group_id = assign_object_id(&mut reserved, group.object_id)?;
        group.object_id = Some(group_id);
        for child in &mut group.children {
            child.object_id = Some(assign_object_id(&mut reserved, child.object_id)?);
        }
        worksheet.shape_groups.push(group);
        Ok(group_id)
    }

    /// Remove a shape group by the group's assigned OBJ identifier.
    pub fn remove_shape_group(&mut self, sheet: usize, object_id: u16) -> Result<ShapeGroupWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shape_groups
            .iter()
            .position(|group| group.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shape_groups.remove(index))
    }

    pub fn set_auto_filter(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        MergedRange::try_new(first_row, last_row, first_col, last_col)?;
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.list_objects.iter().any(|table| {
            first_row <= u32::from(table.range().last_row())
                && u32::from(table.range().first_row()) <= last_row
                && first_col <= table.range().last_column()
                && table.range().first_column() <= last_col
        }) {
            return Err(Error::InvalidData(
                "set_auto_filter: range overlaps a worksheet table".to_string(),
            ));
        }
        let itab = sheet
            .checked_add(1)
            .and_then(|index| u16::try_from(index).ok())
            .ok_or_else(|| {
                Error::InvalidData(
                    "set_auto_filter: sheet index exceeds BIFF8 itab limit".to_string(),
                )
            })?;
        let target_sheet = u16::try_from(sheet).map_err(|_| {
            Error::InvalidData("set_auto_filter: sheet index exceeds BIFF8 limit".to_string())
        })?;

        let start_ref = a1_cell(first_row, first_col);
        let end_ref = a1_cell(last_row, last_col);
        let reference = format!("{start_ref}:{end_ref}");
        let defined_name = DefinedName {
            name: "_FilterDatabase".to_string(),
            reference,
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(target_sheet),
            hidden: true,
            is_function: false,
            is_built_in: true,
            built_in_code: Some(0x0D),
        };

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.auto_filter = Some(AutoFilterRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });
        self.defined_names.retain(|n| {
            !(n.is_built_in && n.built_in_code == Some(0x0D) && n.local_sheet == Some(itab))
        });
        self.defined_names.push(defined_name);

        Ok(())
    }

    /// Add a filter condition to a specific column within the AutoFilter range.
    ///
    /// The AutoFilter range must first be set via [`Self::set_auto_filter`]. The
    /// `column_index` is 0-based relative to the filter range start column.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `column_index` — column within the filter range (0-based relative)
    /// * `join_or` — `true` to join conditions with OR, `false` for AND
    /// * `cond1` — first filter condition
    /// * `cond2` — second filter condition (use `AutoFilterConditionWrite::None` if unused)
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xls::writer::biff::AutoFilterConditionWrite;
    ///
    /// // Filter column 2: value > 100
    /// writer.add_filter_condition(
    ///     sheet_idx, 2, false,
    ///     AutoFilterConditionWrite::Number { operator: 0x04, value: 100.0 },
    ///     AutoFilterConditionWrite::None,
    /// )?;
    /// ```
    pub fn add_filter_condition(
        &mut self,
        sheet: usize,
        column_index: u16,
        join_or: bool,
        cond1: AutoFilterConditionWrite,
        cond2: AutoFilterConditionWrite,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let Some(filter) = worksheet.auto_filter else {
            return Err(Error::InvalidData(
                "add_filter_condition: call set_auto_filter first".to_string(),
            ));
        };
        let width = filter.last_col - filter.first_col + 1;
        if column_index >= width {
            return Err(Error::InvalidCellReference(format!(
                "AutoFilter relative column {column_index} exceeds its {width}-column range"
            )));
        }
        if worksheet
            .auto_filter_columns
            .iter()
            .any(|entry| entry.column_index == column_index)
        {
            return Err(Error::InvalidData(
                "AutoFilter column already has a condition".to_string(),
            ));
        }

        worksheet.add_auto_filter_column(AutoFilterColumnDef {
            column_index,
            join_or,
            condition1: cond1,
            condition2: cond2,
        });

        Ok(())
    }

    /// Set the sort configuration for a worksheet.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `case_sensitive` — whether sorting is case-sensitive
    /// * `sort_by_columns` — `true` for left-to-right sort, `false` for top-to-bottom
    /// * `keys` — up to 3 sort keys as `(column_index, descending)` tuples
    pub fn set_sort(
        &mut self,
        sheet: usize,
        case_sensitive: bool,
        sort_by_columns: bool,
        keys: &[(u16, bool)],
    ) -> Result<()> {
        if keys.is_empty() || keys.len() > 3 {
            return Err(Error::InvalidData(
                "set_sort: must provide 1..3 sort keys".to_string(),
            ));
        }
        if keys.iter().any(|(col, _)| *col > u16::from(u8::MAX)) {
            return Err(Error::InvalidCellReference(
                "sort key column is outside the BIFF8 grid".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.set_sort_config(SortConfig {
            case_sensitive,
            sort_by_columns,
            keys: keys.to_vec(),
        });

        Ok(())
    }

    /// Replace the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Unlike [`set_sort`](Self::set_sort), this preserves the complete
    /// `SortData` model, including an explicit range, more than three keys,
    /// custom lists, differential-format colors, and icon sets. The previous
    /// owned configuration is returned.
    pub fn put_sort(
        &mut self,
        sheet: usize,
        sort: crate::writer::sort::Config,
    ) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if let crate::writer::sort::Parent::Table { id } = sort.parent()
            && !worksheet
                .list_objects
                .iter()
                .any(|table| table.id().value() == id)
        {
            return Err(Error::InvalidData(
                "table SortData references an unknown ListObject identifier".to_string(),
            ));
        }
        Ok(worksheet.put_sort(sort))
    }

    /// Remove and return the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Removing an absent configuration succeeds and returns `None`.
    pub fn remove_sort(&mut self, sheet: usize) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        Ok(worksheet.remove_sort())
    }

    /// Add a pivot table definition to a worksheet.
    ///
    /// This writes the SX* record family (SXVS, SXVIEW, SXVD, SXVI, SXDI,
    /// SXPI) to the worksheet stream. The pivot table must be fully
    /// configured before calling this method.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `config` — pivot table configuration (see [`PivotTableConfig`])
    pub fn add_pivot_table(&mut self, sheet: usize, config: PivotTableConfig) -> Result<()> {
        validate_pivot_table_config(&config)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Generate pivot output cells BEFORE consuming config.fields / config.data_items.
        // Excel validates that DIMENSIONS and cell content are consistent with the
        // pivot table definition; missing cells cause a "corrupt file" repair dialog.
        Self::generate_pivot_output_cells(worksheet, &config)?;
        self.fmt.enable_pivot_xfs();

        let fields: Vec<WritablePivotField> = config
            .fields
            .into_iter()
            .map(|f| {
                let mut items: Vec<WritablePivotItem> = f
                    .items
                    .into_iter()
                    .map(|i| WritablePivotItem {
                        item_type: i.item_type,
                        flags: i.flags,
                        cache_index: i.cache_index,
                        name: i.name,
                    })
                    .collect();

                // Sort data items (item_type=0x0000) alphabetically by their
                // cache label to match Excel's default SXVI ordering.  Non-data
                // items (subtotals etc.) stay at the end.
                let data_end = items
                    .iter()
                    .position(|i| i.item_type != 0x0000)
                    .unwrap_or(items.len());
                items[..data_end].sort_unstable_by(|a, b| {
                    let al = f
                        .cache_items
                        .get(a.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    let bl = f
                        .cache_items
                        .get(b.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    al.cmp(&bl)
                });

                WritablePivotField {
                    axis: f.axis,
                    subtotal_count: f.subtotal_count,
                    subtotal_flags: f.subtotal_flags,
                    items,
                    name: f.name,
                    cache_name: f.cache_name,
                    cache_items: f.cache_items,
                    is_numeric: f.is_numeric,
                    grouping: f.grouping,
                }
            })
            .collect();

        let data_items: Vec<WritablePivotDataItem> = config
            .data_items
            .into_iter()
            .map(|d| WritablePivotDataItem {
                source_field_index: d.source_field_index,
                function: d.function,
                display_format: d.display_format,
                base_field_index: d.base_field_index,
                base_item_index: d.base_item_index,
                num_format_index: d.num_format_index,
                name: d.name,
            })
            .collect();

        worksheet.add_pivot_table(WritablePivotTable {
            name: config.name,
            source_type: config.source_type,
            source_sheet_name: config.source_sheet_name,
            source_first_row: config.source_first_row,
            source_last_row: config.source_last_row,
            source_first_col: config.source_first_col,
            source_last_col: config.source_last_col,
            first_row: config.first_row,
            last_row: config.last_row,
            first_col: config.first_col,
            last_col: config.last_col,
            first_header_row: config.first_header_row,
            first_data_row: config.first_data_row,
            first_data_col: config.first_data_col,
            data_field_name: config.data_field_name,
            data_axis: config.data_axis,
            data_position: config.data_position,
            fields,
            data_items,
            page_entries: config.page_entries,
            source_data: config.source_data,
        });

        Ok(())
    }

    /// Generate the cell data that Excel expects in the SXVIEW output area.
    ///
    /// The layout (for a single row-field, single col-field, single page-field,
    /// single data-field configuration) is:
    ///
    /// ```text
    /// (first_row-2, 0)       : page field name    (first_row-2, 1)       : "(All)"
    /// (first_row,   0)       : data item name      (first_row, first_data_col): "Column Labels"
    /// (first_header_row, 0)  : "Row Labels"        (fhr, fdc+j)           : col item names …
    /// (first_data_row+i, 0)  : row item name       (fdr+i, fdc+j)         : aggregated value
    /// (last_row, 0)          : "Grand Total"        (lr, fdc+j)            : column totals
    /// ```
    fn generate_pivot_output_cells(
        ws: &mut WritableWorksheet,
        cfg: &PivotTableConfig,
    ) -> Result<()> {
        // Identify fields per axis.
        let row_field = cfg.fields.iter().find(|f| f.axis == 0x0001);
        let col_field = cfg.fields.iter().find(|f| f.axis == 0x0002);
        let page_field = cfg.fields.iter().find(|f| f.axis == 0x0004);

        let data_item = cfg.data_items.first();

        // Helper: find the field index for a given field by cache_name.
        let field_idx_of =
            |name: &str| -> Option<usize> { cfg.fields.iter().position(|f| f.cache_name == name) };

        // Collect row/col item labels from cache_items, sorted alphabetically
        // to match Excel's default SXVI ordering.  Also build a mapping from
        // cache_index → sorted position so the aggregation grid uses the same
        // order as the output rows/columns.
        let (row_items, row_cache_to_sorted) = Self::sorted_cache_items(row_field);
        let (col_items, col_cache_to_sorted) = Self::sorted_cache_items(col_field);

        let fr = cfg.first_row;
        let fhr = cfg.first_header_row;
        let fdr = cfg.first_data_row;
        let fdc = cfg.first_data_col;
        let lr = cfg.last_row;
        let lc = cfg.last_col;
        let fc = cfg.first_col;

        let offset = |base: u16, amount: usize| -> Result<u16> {
            let amount = u16::try_from(amount).map_err(|_| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })?;
            base.checked_add(amount).ok_or_else(|| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })
        };
        let mut staged = Vec::new();
        let mut add = |row: u16,
                       col: u16,
                       value: CellValue,
                       pivot_xf_role: Option<PivotCellXfRole>|
         -> Result<()> {
            staged.push(WritableCell::new(
                CellPos::try_new(u32::from(row), col)?,
                value,
                0,
                pivot_xf_role,
            ));
            Ok(())
        };

        // --- Page field area (above SXVIEW range) ---
        if let Some(pf) = page_field {
            let page_row = fr.saturating_sub(2);
            add(
                page_row,
                0,
                CellValue::String(pf.cache_name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
            add(
                page_row,
                1,
                CellValue::String("(All)".to_string()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }

        // --- Row at first_row: data item name + "Column Labels" ---
        if let Some(di) = data_item {
            add(
                fr,
                fc,
                CellValue::String(di.name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }
        if col_field.is_some() {
            add(
                fr,
                fdc,
                CellValue::String("Column Labels".to_string()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }

        // --- Row at first_header_row: "Row Labels" + column item names + "Grand Total" ---
        add(
            fhr,
            fc,
            CellValue::String("Row Labels".to_string()),
            Some(PivotCellXfRole::HeaderAccent),
        )?;
        for (j, ci) in col_items.iter().enumerate() {
            add(
                fhr,
                offset(fdc, j)?,
                CellValue::String(ci.clone()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }
        add(
            fhr,
            lc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::HeaderPlain),
        )?;

        // --- Compute aggregated values from source_data ---
        let row_fi = row_field.and_then(|f| field_idx_of(&f.cache_name));
        let col_fi = col_field.and_then(|f| field_idx_of(&f.cache_name));
        let data_fi = data_item.map(|di| di.source_field_index as usize);

        let nr = row_items.len();
        let nc = col_items.len();
        let mut grid = vec![vec![0.0f64; nc]; nr];
        let mut row_totals = vec![0.0f64; nr];
        let mut col_totals = vec![0.0f64; nc];
        let mut grand_total = 0.0f64;

        for row_data in &cfg.source_data {
            // Map cache indices through the sorted permutation so that
            // grid positions match the alphabetically-sorted output.
            let ri = row_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    row_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let ci = col_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    col_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let val = data_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::Number(v)) => Some(*v),
                _ => None,
            });

            if let (Some(ri), Some(ci), Some(val)) = (ri, ci, val)
                && ri < nr
                && ci < nc
            {
                grid[ri][ci] += val;
                row_totals[ri] += val;
                col_totals[ci] += val;
                grand_total += val;
            }
        }

        // --- Data rows ---
        for (i, (ri_name, row_total)) in row_items.iter().zip(row_totals.iter()).enumerate() {
            let r = offset(fdr, i)?;
            add(
                r,
                fc,
                CellValue::String(ri_name.clone()),
                Some(PivotCellXfRole::RowLabel),
            )?;
            for (j, cell_val) in grid[i].iter().enumerate() {
                add(
                    r,
                    offset(fdc, j)?,
                    CellValue::Number(*cell_val),
                    Some(PivotCellXfRole::Value),
                )?;
            }
            add(
                r,
                lc,
                CellValue::Number(*row_total),
                Some(PivotCellXfRole::Value),
            )?;
        }

        // --- Grand total row ---
        add(
            lr,
            fc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::RowLabel),
        )?;
        for (j, col_total) in col_totals.iter().enumerate() {
            add(
                lr,
                offset(fdc, j)?,
                CellValue::Number(*col_total),
                Some(PivotCellXfRole::Value),
            )?;
        }
        add(
            lr,
            lc,
            CellValue::Number(grand_total),
            Some(PivotCellXfRole::Value),
        )?;
        for cell in staged {
            ws.add_cell(cell);
        }
        Ok(())
    }

    /// Sort a field's cache items alphabetically and return the sorted labels
    /// plus a mapping from original cache index to sorted position.
    ///
    /// Returns `(sorted_labels, cache_to_sorted)` where `cache_to_sorted[i]`
    /// gives the position of original cache item `i` in the sorted output.
    fn sorted_cache_items(field: Option<&PivotFieldConfig>) -> (Vec<String>, Vec<usize>) {
        let Some(f) = field else {
            return (Vec::new(), Vec::new());
        };

        // Build (original_index, label) pairs and sort by label.
        let mut indexed: Vec<(usize, String)> = f
            .cache_items
            .iter()
            .enumerate()
            .map(|(i, item)| (i, item.display_text()))
            .collect();
        indexed.sort_unstable_by(|a, b| a.1.cmp(&b.1));

        let sorted_labels: Vec<String> = indexed.iter().map(|(_, value)| value.clone()).collect();

        // cache_to_sorted[original_cache_idx] = position in sorted output
        let mut cache_to_sorted = vec![0usize; f.cache_items.len()];
        for (sorted_pos, (orig_idx, _)) in indexed.iter().enumerate() {
            cache_to_sorted[*orig_idx] = sorted_pos;
        }

        (sorted_labels, cache_to_sorted)
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers. More complex formulas will be rejected
    /// at serialization time to avoid emitting invalid BIFF payloads.
    pub fn define_name(&mut self, name: &str, reference: &str) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name: workbook must have at least one worksheet".to_string(),
            ));
        }

        // For now, workbook-scoped names that refer to cell ranges are
        // anchored to the first worksheet. Users who need explicit
        // sheet scoping can use `define_name_local`.
        let target_sheet = 0u16;

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet: usize) -> Result<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            Error::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a workbook-scoped named range with a user-visible comment.
    pub fn define_name_with_comment(
        &mut self,
        name: &str,
        reference: &str,
        comment: &str,
    ) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name_with_comment: workbook must have at least one worksheet".to_string(),
            ));
        }

        let target_sheet = 0u16;

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Remove all defined names with the given name.
    ///
    /// Returns `true` if at least one name was removed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.defined_names.len();
        self.defined_names.retain(|n| n.name != name);
        self.defined_names.len() < initial_len
    }

    /// Get all defined names in this workbook.
    pub fn named_ranges(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Add complete inert BIFF8 defined-name metadata.
    pub fn add_defined_name_record(&mut self, options: DefinedNameRecordOptions) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records
            .push((options, Default::default()));
        Ok(index)
    }

    /// Add complete inert `Lbl` metadata and its ordered BIFF8 future records.
    pub fn add_defined_name_record_with_future_records(
        &mut self,
        options: DefinedNameRecordOptions,
        future: crate::DefinedNameFutureRecords,
    ) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        named_range::validate_future_records(&future, options.serialized_name())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records.push((options, future));
        Ok(index)
    }

    /// Set the width of a column in character units.
    ///
    /// The column index is 0-based (0 = column A), matching the rest of the
    /// XLS writer API. The width is specified in the same units as Excel's
    /// UI, i.e. the number of characters of the "0" glyph in the default
    /// font. Internally this is converted to BIFF8 units of 1/256 characters
    /// for the COLINFO record.
    pub fn set_column_width(&mut self, sheet: usize, col: u16, width_chars: f64) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "set_column_width: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        if !(width_chars.is_finite()) || width_chars <= 0.0 {
            return Err(Error::InvalidData(
                "set_column_width: width must be a positive finite value".to_string(),
            ));
        }

        let max_units = 255u32 * 256u32; // Excel maximum column width
        let width_units_f = (width_chars * 256.0).round();
        if width_units_f <= 0.0 || width_units_f > max_units as f64 {
            return Err(Error::InvalidData(
                "set_column_width: width exceeds Excel's maximum (255 characters)".to_string(),
            ));
        }

        let width_units = width_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_column_width(col, width_units);
        Ok(())
    }

    /// Hide a column.
    pub fn hide_column(&mut self, sheet: usize, col: u16) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "hide_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_column(col);
        Ok(())
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, sheet: usize, col: u16) -> Result<()> {
        if col >= 256 {
            return Err(Error::InvalidData(
                "show_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_column(col);
        Ok(())
    }

    pub fn merge_cells(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        let range = MergedRange::try_new(first_row, last_row, first_col, last_col)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if worksheet
            .merged_ranges
            .iter()
            .any(|existing| range.overlaps(*existing))
        {
            return Err(Error::InvalidData("merged-cell ranges overlap".to_string()));
        }

        worksheet.add_merged_range(range);

        Ok(())
    }

    /// Configure freeze panes for the specified worksheet.
    ///
    /// Row and column indices are 0-based and represent the number of
    /// rows/columns at the top/left that remain frozen.
    pub fn freeze_panes(&mut self, sheet: usize, freeze_rows: u32, freeze_cols: u16) -> Result<()> {
        if freeze_rows == 0 && freeze_cols == 0 {
            let worksheet = self
                .worksheets
                .get_mut(sheet)
                .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
            worksheet.clear_freeze_panes();
            return Ok(());
        }
        let freeze_rows = u16::try_from(freeze_rows).map_err(|_| {
            Error::InvalidCellReference("freeze-panes row is outside the BIFF8 grid".to_string())
        })?;
        let freeze_cols = u8::try_from(freeze_cols).map_err(|_| {
            Error::InvalidCellReference("freeze-panes column is outside the BIFF8 grid".to_string())
        })?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.set_freeze_panes(freeze_rows, freeze_cols)
    }

    /// Remove any freeze panes from the specified worksheet.
    pub fn unfreeze_panes(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.clear_freeze_panes();
        Ok(())
    }

    /// Replace the worksheet's checked BIFF8 zoom scale.
    pub fn put_scale(
        &mut self,
        sheet: usize,
        scale: Option<crate::writer::view::Scale>,
    ) -> Result<Option<crate::writer::view::Scale>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(worksheet.put_scale(scale))
    }

    /// Replace a worksheet view after validating the complete prospective state.
    pub fn put_view(
        &mut self,
        sheet: usize,
        view: crate::writer::view::View,
    ) -> Result<crate::writer::view::View> {
        view.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(std::mem::replace(&mut worksheet.view, view))
    }

    /// Replace a worksheet pane and its selections as one failure-atomic edit.
    pub fn put_pane(
        &mut self,
        sheet: usize,
        pane: crate::writer::view::Pane,
        selections: Vec<crate::writer::view::Selection>,
    ) -> Result<(
        Option<crate::writer::view::Pane>,
        Vec<crate::writer::view::Selection>,
    )> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.view.put_pane(pane, selections)
    }

    /// Set the height of a row in points.
    ///
    /// The row index is 0-based (0 = first row), and the height is specified
    /// in typographic points. Internally this is converted to twips
    /// (1/20th of a point) for the BIFF8 ROW record.
    pub fn set_row_height(&mut self, sheet: usize, row: u32, height_points: f64) -> Result<()> {
        if !(height_points.is_finite()) || height_points <= 0.0 {
            return Err(Error::InvalidData(
                "set_row_height: height must be a positive finite value".to_string(),
            ));
        }

        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "set_row_height: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let height_units_f = (height_points * 20.0).round();
        if height_units_f <= 0.0 || height_units_f > u16::MAX as f64 {
            return Err(Error::InvalidData(
                "set_row_height: height exceeds BIFF8 row height limit".to_string(),
            ));
        }

        let height_units = height_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_row_height(row, height_units);
        Ok(())
    }

    /// Hide a row.
    pub fn hide_row(&mut self, sheet: usize, row: u32) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "hide_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_row(row);
        Ok(())
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, sheet: usize, row: u32) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "show_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_row(row);
        Ok(())
    }

    /// Add a data validation rule to the specified worksheet.
    pub fn add_data_validation(&mut self, sheet: usize, validation: DataValidation) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let range = validation.range;
        worksheet.add_data_validation(
            validation,
            payload,
            vec![range],
            DataValidationOptions::default(),
        );

        Ok(())
    }

    /// Add a validation with typed flags and additional target ranges.
    pub fn add_data_validation_with_options(
        &mut self,
        sheet: usize,
        validation: DataValidation,
        additional_ranges: &[DataValidationRange],
        options: DataValidationOptions,
    ) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;
        let range_count = 1usize
            .checked_add(additional_ranges.len())
            .ok_or_else(|| Error::InvalidData("DV range count overflows".to_string()))?;
        if range_count > 432 {
            return Err(Error::InvalidData("DV range count exceeds 432".to_string()));
        }
        let mut ranges = Vec::with_capacity(range_count);
        ranges.push(validation.range);
        ranges.extend_from_slice(additional_ranges);
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.add_data_validation(validation, payload, ranges, options);
        Ok(())
    }

    /// Configure worksheet-level DVAL window/dropdown metadata.
    pub fn set_data_validation_table_options(
        &mut self,
        sheet: usize,
        options: DataValidationTableOptions,
    ) -> Result<()> {
        if options.x_left > 65_535
            || options.y_top > 65_535
            || matches!(options.dropdown_object_id, Some(0))
            || options.dropdown_object_id.is_some_and(|id| id > 32_767)
        {
            return Err(Error::InvalidData(
                "DVAL metadata is out of range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.data_validation_table_options = Some(options);
        Ok(())
    }

    pub fn add_conditional_format(&mut self, sheet: usize, cf: ConditionalFormat) -> Result<()> {
        if cf.first_row > cf.last_row
            || cf.first_col > cf.last_col
            || cf.last_row > 65_535
            || cf.last_col > 255
        {
            return Err(Error::InvalidData(
                "add_conditional_format: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_conditional_format(cf);

        Ok(())
    }

    /// Add one legacy `CONDFMT` collection with ordered ranges and one to three ordered rules.
    pub fn add_conditional_format_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormatGroup,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > 3 {
            return Err(Error::InvalidData(
                "legacy conditional-format rule count must be 1..=3".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        for rule in &group.rules {
            rule.format_type.to_biff_payload()?;
        }
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .add_conditional_format_group(group);
        Ok(())
    }

    /// Add one future `CondFmt12` collection. Formula tokens and visual
    /// payloads are serialized exactly and are never evaluated.
    pub fn add_conditional_format12_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormat12Group,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "future conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "future conditional-format rule count must be 1..=65535".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "future conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.conditional_formats.len() + worksheet.conditional_formats12.len() >= 32_768 {
            return Err(Error::InvalidData(
                "conditional-format group count exceeds the 15-bit BIFF identifier space"
                    .to_string(),
            ));
        }
        let mut priorities = worksheet
            .conditional_formats12
            .iter()
            .flat_map(|existing| existing.rules.iter().map(|rule| rule.priority))
            .collect::<HashSet<_>>();
        for rule in &group.rules {
            if rule.priority == 0 || !priorities.insert(rule.priority) {
                return Err(Error::InvalidData(
                    "future conditional-format priorities must be nonzero and unique per sheet"
                        .to_string(),
                ));
            }
            if !matches!(rule.template, 0..=5 | 7..=12 | 15..=27 | 29 | 30) {
                return Err(Error::InvalidData(
                    "future conditional-format template is invalid".to_string(),
                ));
            }
            let between = matches!(
                rule.format_type,
                ConditionalFormat12Type::CellValue {
                    operator: ConditionalFormatOperator::Between
                        | ConditionalFormatOperator::NotBetween,
                    ..
                }
            );
            if let ConditionalFormat12Type::CellValue { formula2, .. } = &rule.format_type
                && between != formula2.is_some()
            {
                return Err(Error::InvalidData(
                        "between/not-between CF12 rules require two formulas; other comparisons require one".to_string(),
                    ));
            }
            let visual = matches!(
                rule.format_type,
                ConditionalFormat12Type::ColorScale { .. }
                    | ConditionalFormat12Type::DataBar { .. }
                    | ConditionalFormat12Type::IconSet { .. }
            );
            if visual && (rule.stop_if_true || rule.differential_format != [0, 0, 0, 0, 0, 0]) {
                return Err(Error::InvalidData(
                    "visual CF12 rules require an empty DXFN12 and cannot stop-if-true".to_string(),
                ));
            }
            let (condition_type, comparison, formula1, formula2, active_formula, payload) =
                rule.format_type.biff_parts();
            let config = crate::writer::biff::Cf12Config {
                condition_type,
                comparison,
                differential_format: &rule.differential_format,
                formula1,
                formula2,
                active_formula,
                stop_if_true: rule.stop_if_true,
                priority: rule.priority,
                template: rule.template,
                template_parameters: rule.template_parameters,
                rule_payload: payload,
            };
            crate::writer::biff::write_cf12(&mut Vec::new(), &config)?;
        }
        worksheet.add_conditional_format12_group(group);
        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        pos: CellPos,
        value: CellValue,
        format_id: u16,
    ) -> Result<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(Error::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_cell(WritableCell::new(pos, value, format_id, None));

        Ok(())
    }

    /// Set the date system (1900 vs 1904)
    ///
    /// # Arguments
    ///
    /// * `use_1904` - True to use 1904 date system (Mac), false for 1900 (Windows, default)
    pub fn set_1904_dates(&mut self, use_1904: bool) {
        self.use_1904_dates = use_1904;
    }

    pub fn set_workbook_environment(&mut self, options: WorkbookEnvironmentOptions) -> Result<()> {
        if options.refresh_external_data_on_load && !options.template {
            return Err(Error::InvalidData(
                "RefreshAll requires a template workbook".to_string(),
            ));
        }
        if (options.envelope_visible || options.envelope_initialized) && !options.has_envelope {
            return Err(Error::InvalidData(
                "envelope state flags require has_envelope".to_string(),
            ));
        }
        if !(1..=981).contains(&options.default_country_code)
            || !(1..=981).contains(&options.current_country_code)
        {
            return Err(Error::InvalidData(
                "country codes must be 1..=981".to_string(),
            ));
        }
        self.environment_options = options;
        Ok(())
    }

    /// Set the workbook extension flags emitted as a `BookExt` record
    /// (MS-XLS 2.4.23); `None` emits no record.
    pub fn set_book_ext(&mut self, book_ext: Option<crate::BookExt>) {
        self.book_ext = book_ext;
    }

    /// Append a real-time data (RTD) topic emitted as a `RealTimeData`
    /// record (MS-XLS 2.4.214) in the workbook globals.
    ///
    /// When the topic shares a prefix with the previously added topic, set
    /// [`crate::RealTimeData::common_prefix_len`] and store only the
    /// trailing sub-strings in `topic_segments`, matching the on-disk prefix
    /// compression.
    pub fn add_real_time_data(&mut self, topic: crate::RealTimeData) -> Result<()> {
        if let Some(cell) = topic
            .cells
            .iter()
            .find(|cell| usize::from(cell.sheet_index) >= self.worksheets.len())
        {
            return Err(Error::WorksheetNotFound(format!(
                "Sheet {}",
                cell.sheet_index
            )));
        }
        self.real_time_data.push(topic);
        Ok(())
    }

    /// Append a Web page published from the workbook globals, emitted as a
    /// `WebPub` record (MS-XLS 2.4.344).
    pub fn add_web_publication(&mut self, publication: crate::WebPub) -> Result<()> {
        publication.validate_for_write()?;
        self.web_publications.push(publication);
        Ok(())
    }

    /// Append a Web page published from a worksheet, emitted as a `WebPub`
    /// record (MS-XLS 2.4.344) in that sheet's substream.
    pub fn add_sheet_web_publication(
        &mut self,
        sheet: usize,
        publication: crate::WebPub,
    ) -> Result<()> {
        publication.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.web_publications.push(publication);
        Ok(())
    }

    /// Set a worksheet's default phonetic format and visible phonetic ranges
    /// (PHONETICINFO, MS-XLS 2.4.192); `None` emits no record.
    pub fn set_phonetic_info(
        &mut self,
        sheet: usize,
        phonetic_info: Option<crate::PhoneticInfo>,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.phonetic_info = phonetic_info;
        Ok(())
    }

    /// Set the document theme emitted as a `Theme` record (MS-XLS 2.4.326);
    /// `None` emits no record. Large custom theme contents are chunked into
    /// ContinueFrt12 records automatically.
    pub fn set_theme(&mut self, theme: Option<crate::Theme>) {
        self.theme = theme;
    }

    /// Set the MDX (OLAP cube) metadata emitted as the workbook globals
    /// `METADATA` production (MS-XLS 2.1); `None` emits no records. Oversized
    /// record payloads are chunked into ContinueFrt12 records automatically.
    pub fn set_mdx_metadata(&mut self, metadata: Option<crate::MdxMetadata>) {
        self.mdx_metadata = metadata;
    }

    /// Set the `XFExt` formatting property extensions (MS-XLS 2.4.355)
    /// emitted after the XF table. Each extension's `xf_index` is validated
    /// against the written XF record count when the workbook is saved.
    pub fn set_xf_extensions(&mut self, xf_extensions: Vec<crate::XfExt>) {
        self.xf_extensions = xf_extensions;
    }

    /// Set the `StyleExt` cell-style extensions (MS-XLS 2.4.270) emitted
    /// after the built-in STYLE records.
    pub fn set_style_extensions(&mut self, style_extensions: Vec<crate::StyleExt>) {
        self.style_extensions = style_extensions;
    }

    pub fn set_workbook_window(&mut self, options: WorkbookWindowOptions) -> Result<()> {
        options.validate_intrinsic()?;
        self.workbook_window_options = options;
        self.synchronize_workbook_window_selection();
        Ok(())
    }

    fn synchronize_workbook_window_selection(&mut self) {
        let sheet_count = self.worksheets.len();
        let selected_count = usize::from(self.workbook_window_options.selected_sheet_count);
        let active = usize::from(self.workbook_window_options.active_sheet_index);
        if selected_count == 0 || selected_count > sheet_count || active >= sheet_count {
            return;
        }
        let first_selected = active.min(sheet_count - selected_count);
        let selected_range = first_selected..first_selected + selected_count;
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            worksheet.view.select(selected_range.contains(&index));
        }
    }

    pub fn set_function_groups(&mut self, options: FunctionGroupOptions) -> Result<()> {
        options.validate()?;
        self.function_group_options = options;
        Ok(())
    }

    pub fn add_external_workbook_link(
        &mut self,
        options: ExternalWorkbookOptions,
    ) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "external supporting-book count exceeds resource bound".to_string(),
            ));
        }
        let index = self.external_workbooks.len();
        self.external_workbooks.push(options);
        self.external_names.push(Vec::new());
        Ok(index)
    }

    fn external_name_count(&self) -> usize {
        self.external_names.iter().map(Vec::len).sum::<usize>()
            + self.add_in_functions.len()
            + self
                .dde_or_ole_links
                .iter()
                .map(|link| link.items.len())
                .sum::<usize>()
    }

    pub fn add_external_defined_name(
        &mut self,
        external_workbook: usize,
        options: ExternalDefinedNameOptions,
    ) -> Result<usize> {
        let book = self
            .external_workbooks
            .get(external_workbook)
            .ok_or_else(|| {
                Error::InvalidData("external workbook index is out of range".to_string())
            })?;
        options.validate(book.sheets.len())?;
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let names = &mut self.external_names[external_workbook];
        let index = names.len();
        names.push(options);
        Ok(index)
    }

    pub fn add_add_in_function(&mut self, options: AddInFunctionOptions) -> Result<usize> {
        options.validate()?;
        if self.add_in_functions.is_empty()
            && self.external_workbooks.len() + self.dde_or_ole_links.len() >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "add-in function count exceeds resource bound".to_string(),
            ));
        }
        let index = self.add_in_functions.len();
        self.add_in_functions.push(options);
        Ok(index)
    }

    pub fn add_dde_or_ole_link(&mut self, options: DdeOrOleLinkOptions) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self
            .external_name_count()
            .checked_add(options.items.len())
            .is_none_or(|count| count > 4096)
        {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let index = self.dde_or_ole_links.len();
        self.dde_or_ole_links.push(options);
        Ok(index)
    }

    pub fn set_calculation_settings(&mut self, settings: CalculationSettings) -> Result<()> {
        if !(1..=32_767).contains(&settings.maximum_iterations) {
            return Err(Error::InvalidData(
                "maximum calculation iterations must be 1..=32767".to_string(),
            ));
        }
        if !settings.iteration_delta.is_finite() || settings.iteration_delta < 0.0 {
            return Err(Error::InvalidData(
                "calculation iteration delta must be finite and non-negative".to_string(),
            ));
        }
        self.calculation_settings = settings;
        Ok(())
    }

    pub fn set_recalculation_pending(&mut self, sheet: usize, pending: bool) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.formulas_pending_recalculation = pending;
        Ok(())
    }

    pub fn set_scenario_manager(
        &mut self,
        sheet: usize,
        manager: crate::ScenarioManager,
    ) -> Result<()> {
        manager.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = Some(manager);
        Ok(())
    }

    pub fn clear_scenario_manager(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = None;
        Ok(())
    }

    /// Configure an inert BIFF8 data-consolidation directory for a worksheet.
    pub fn set_consolidation(
        &mut self,
        sheet: usize,
        consolidation: crate::Consolidation,
    ) -> Result<()> {
        consolidation.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = Some(consolidation);
        Ok(())
    }

    pub fn clear_consolidation(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = None;
        Ok(())
    }

    /// Configure a complete inert VBA project with safe default limits.
    pub fn set_vba(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
    ) -> Result<()> {
        self.set_vba_with(workbook_code_name, project, &litchi_vba::Limits::default())
    }

    /// Configure a complete inert VBA project using explicit resource limits.
    ///
    /// Module source is serialized but never compiled, interpreted, or run.
    /// Validation and serialization finish before the writer state is changed.
    pub fn set_vba_with(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        let payload = project.finish(limits)?;
        self.put_vba(workbook_code_name, payload)
    }

    /// Configure an already validated and serialized inert VBA project.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    pub fn put_vba(
        &mut self,
        workbook_code_name: &str,
        payload: litchi_vba::Payload,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        self.vba_metadata = Some(VbaWriteMetadata {
            workbook_code_name: workbook_code_name.to_string(),
            project: payload,
        });
        Ok(())
    }

    /// Remove the configured project and all worksheet VBA code names.
    pub fn clear_vba(&mut self) {
        self.vba_metadata = None;
        for worksheet in &mut self.worksheets {
            worksheet.vba_code_name = None;
        }
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_metadata.is_some()
    }

    /// Set a worksheet's tab color as a BIFF8 palette index (SHEETEXT
    /// `icvPlain`, MS-XLS 2.4.259). Valid indices are 0x08 through 0x3F;
    /// `None` clears an explicitly set color.
    pub fn set_worksheet_tab_color(&mut self, sheet: usize, tab_color: Option<u8>) -> Result<()> {
        if let Some(index) = tab_color
            && !(0x08..=0x3F).contains(&index)
        {
            return Err(Error::InvalidData(format!(
                "sheet tab color index {index:#04X} is outside the Icv palette"
            )));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.tab_color = tab_color;
        Ok(())
    }

    /// Author a what-if data table (MS-XLS 2.4.319) anchored at a formula
    /// cell. The anchor cell is written as a `PtgTbl` formula immediately
    /// followed by the `Table` record; it must lie outside the table range
    /// and must not already carry a value.
    pub fn add_data_table(
        &mut self,
        sheet: usize,
        anchor_row: u32,
        anchor_col: u16,
        table: crate::DataTable,
    ) -> Result<()> {
        let anchor_pos = CellPos::try_new(anchor_row, anchor_col)?;
        let range = table.range();
        let inside = (u32::from(range.first_row())..=u32::from(range.last_row()))
            .contains(&anchor_row)
            && (range.first_col()..=range.last_col()).contains(&anchor_pos.col());
        if inside {
            return Err(Error::InvalidData(
                "data-table anchor formula cell must lie outside the table range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet
            .data_tables
            .iter()
            .any(|(row, col, _)| (*row, *col) == (anchor_row, anchor_col))
        {
            return Err(Error::InvalidData(
                "duplicate data-table anchor cell".to_string(),
            ));
        }
        if let Some(cell) = worksheet.cells.get(&(anchor_row, anchor_col)) {
            if !matches!(cell.value, CellValue::Blank) {
                return Err(Error::InvalidData(
                    "data-table anchor cell already carries a value".to_string(),
                ));
            }
        } else {
            worksheet.add_cell(WritableCell::new(anchor_pos, CellValue::Blank, 0, None));
        }
        worksheet.data_tables.push((anchor_row, anchor_col, table));
        Ok(())
    }

    pub fn set_worksheet_vba_code_name(
        &mut self,
        sheet: usize,
        code_name: Option<&str>,
    ) -> Result<()> {
        if self.vba_metadata.is_none() && code_name.is_some() {
            return Err(Error::InvalidData(
                "worksheet VBA code names require an enabled VBA project".to_string(),
            ));
        }
        if let Some(value) = code_name {
            crate::vba::validate_code_name(value)?;
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.vba_code_name = code_name.map(str::to_string);
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_worksheet_layout(
        &mut self,
        sheet: usize,
        options: WorksheetLayoutOptions,
    ) -> Result<()> {
        options.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_layout = options;
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_page_setup(&mut self, sheet: usize, options: PageSetupOptions) -> Result<()> {
        let valid_margin = |value: f64| value.is_finite() && (0.0..49.0).contains(&value);
        if !valid_margin(options.left_margin_inches)
            || !valid_margin(options.right_margin_inches)
            || !valid_margin(options.top_margin_inches)
            || !valid_margin(options.bottom_margin_inches)
            || !valid_margin(options.header_margin_inches)
            || !valid_margin(options.footer_margin_inches)
        {
            return Err(Error::InvalidData(
                "page margins must be finite and between 0 and 49 inches".to_string(),
            ));
        }
        if options.header.encode_utf16().count() > 255
            || options.footer.encode_utf16().count() > 255
        {
            return Err(Error::InvalidData(
                "header and footer must not exceed 255 UTF-16 code units".to_string(),
            ));
        }
        if (118..=255).contains(&options.paper_size)
            || !(10..=400).contains(&options.scale_percent)
            || options.fit_width_pages > 32767
            || options.fit_height_pages > 32767
            || options.horizontal_resolution_dpi == 0
            || options.vertical_resolution_dpi == 0
            || options.copies == 0
            || options.copies > 32767
        {
            return Err(Error::InvalidData(
                "page setup contains an out-of-range dimension".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = Some(options);
        Ok(())
    }

    pub fn clear_page_setup(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = None;
        worksheet.horizontal_page_breaks.clear();
        worksheet.vertical_page_breaks.clear();
        Ok(())
    }

    /// Add a horizontal break at the first row below the break.
    pub fn add_horizontal_page_break(
        &mut self,
        sheet: usize,
        row: u32,
        col_start: u16,
        col_end: u16,
    ) -> Result<()> {
        let page_break = HorizontalPageBreak::try_new(row, col_start, col_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.horizontal_page_breaks.len() >= 1026 {
            return Err(Error::InvalidData(
                "horizontal page-break count exceeds 1026".to_string(),
            ));
        }
        if worksheet
            .horizontal_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "horizontal page-break ranges overlap".to_string(),
            ));
        }
        worksheet.horizontal_page_breaks.push(page_break);
        Ok(())
    }

    /// Add a vertical break at the first column right of the break.
    pub fn add_vertical_page_break(
        &mut self,
        sheet: usize,
        column: u16,
        row_start: u32,
        row_end: u32,
    ) -> Result<()> {
        let page_break = VerticalPageBreak::try_new(column, row_start, row_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.vertical_page_breaks.len() >= 255 {
            return Err(Error::InvalidData(
                "vertical page-break count exceeds 255".to_string(),
            ));
        }
        if worksheet
            .vertical_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "vertical page-break ranges overlap".to_string(),
            ));
        }
        worksheet.vertical_page_breaks.push(page_break);
        Ok(())
    }

    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        protect_structure: bool,
        protect_windows: bool,
    ) {
        if !protect_structure && !protect_windows && password.is_none() {
            self.workbook_protection = None;
            return;
        }

        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_structure = protect_structure;
        protection.protect_windows = protect_windows;
        protection.password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    pub fn unprotect_workbook(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_structure = false;
            protection.protect_windows = false;
            protection.password_hash = None;
            self.workbook_protection = protection.protect_revisions.then_some(protection);
        }
    }

    /// Configure legacy shared-workbook revision protection.
    pub fn protect_revisions(&mut self, password: Option<&str>) {
        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_revisions = true;
        protection.revision_password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    /// Remove shared-workbook revision protection.
    pub fn unprotect_revisions(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_revisions = false;
            protection.revision_password_hash = None;
            self.workbook_protection = (protection.protect_structure
                || protection.protect_windows
                || protection.password_hash.is_some())
            .then_some(protection);
        }
    }

    /// Configure read-only recommendation and an optional write-reservation password.
    pub fn set_file_sharing(
        &mut self,
        read_only_recommended: bool,
        password: Option<&str>,
        user_name: &str,
    ) -> Result<()> {
        if user_name.encode_utf16().count() > 54 {
            return Err(Error::InvalidData(
                "FILESHARING username exceeds 54 UTF-16 code units".to_string(),
            ));
        }
        self.file_sharing = Some(FileSharing {
            read_only_recommended,
            password_hash: password.map(Self::hash_password),
            user_name: user_name.to_string(),
        });
        Ok(())
    }

    pub fn clear_file_sharing(&mut self) {
        self.file_sharing = None;
    }

    pub fn protect_sheet(
        &mut self,
        sheet: usize,
        password: Option<&str>,
        protect_objects: bool,
        protect_scenarios: bool,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let password_hash = password.map(Self::hash_password);
        worksheet.sheet_protection = Some(SheetProtection {
            protect_objects,
            protect_scenarios,
            password_hash,
        });

        Ok(())
    }

    pub fn unprotect_sheet(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_protection = None;
        Ok(())
    }
}
/// Iterate every OBJ identifier requested or assigned inside a shape group.
fn group_object_ids(group: &ShapeGroupWrite) -> impl Iterator<Item = u16> + '_ {
    group
        .object_id
        .into_iter()
        .chain(group.children.iter().filter_map(|child| child.object_id))
}

/// Collect every existing worksheet drawing-object identifier with fallible
/// capacity reservation for any IDs the caller will add before retaining data.
fn collect_reserved_object_ids(
    worksheet: &WritableWorksheet,
    additional: usize,
) -> Result<HashSet<u16>> {
    let pivot_capacity = worksheet
        .pivot_tables
        .iter()
        .try_fold(0usize, |count, table| {
            count
                .checked_add(table.page_entries.len())
                .ok_or(Error::Allocation(
                    "computing worksheet pivot-object ID capacity",
                ))
        })?;
    let group_capacity = worksheet
        .shape_groups
        .iter()
        .try_fold(0usize, |count, group| {
            count
                .checked_add(group.object_count()?)
                .ok_or(Error::Allocation(
                    "computing worksheet shape-group ID capacity",
                ))
        })?;
    let capacity = pivot_capacity
        .checked_add(worksheet.shapes.len())
        .and_then(|count| count.checked_add(group_capacity))
        .and_then(|count| count.checked_add(additional))
        .ok_or(Error::Allocation(
            "computing worksheet drawing-object ID capacity",
        ))?;
    let mut reserved = HashSet::new();
    reserved
        .try_reserve(capacity)
        .map_err(|_| Error::Allocation("reserving worksheet drawing-object ID storage"))?;
    reserved.extend(
        worksheet
            .pivot_tables
            .iter()
            .flat_map(|table| table.page_entries.iter().map(|entry| entry.2))
            .filter(|id| *id != 0 && *id != u16::MAX),
    );
    reserved.extend(worksheet.shapes.iter().filter_map(|shape| shape.object_id));
    reserved.extend(worksheet.shape_groups.iter().flat_map(group_object_ids));
    Ok(reserved)
}

/// Reserve the requested OBJ identifier or the first free canonical one.
fn assign_object_id(reserved: &mut HashSet<u16>, requested: Option<u16>) -> Result<u16> {
    let object_id = match requested {
        Some(object_id) => object_id,
        None => (1..u16::MAX)
            .find(|candidate| !reserved.contains(candidate))
            .ok_or_else(|| Error::InvalidData("worksheet object IDs are exhausted".to_string()))?,
    };
    if requested.is_none() {
        reserved.insert(object_id);
    }
    Ok(object_id)
}
