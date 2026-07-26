//! Mutable XLSB worksheet for CRUD operations

use crate::xlsb::comments::Comment;
use crate::xlsb::conditional_formatting::{Cfvo, ConditionalFormatting, ConditionalFormattingRule};
use crate::xlsb::data_validation::{DataValidation, DataValidationSettings};
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaCompilationContext, FormulaCompiler, FormulaConverter, FormulaGroup,
    FormulaGroupKind, FormulaParser, FormulaRange,
};
use crate::xlsb::hyperlinks::Hyperlink;
use crate::xlsb::merged_cells::MergedCell;
use crate::xlsb::records::record_types;
use crate::xlsb::web_extension_bindings::XlsbWebExtensionBinding;
use crate::xlsb::writer::RecordWriter;
use litchi_core::sheet::CellValue;
use std::collections::{BTreeMap, HashSet};
use std::io::Write;

/// Cell data for storage
#[derive(Debug, Clone)]
pub struct CellData {
    pub value: CellValue,
    pub style: u32, // Style XF index
    /// Optional pre-encoded cell formula for lossless XLSB workflows.
    pub formula_binary: Option<CellParsedFormula>,
    /// `GrbitFmla` flags; only bit 1 (`fAlwaysCalc`) is defined.
    pub formula_flags: u16,
}

/// Column information for a single 0-based column.
///
/// This writer-side structure drives `BrtColInfo` emission and mirrors the
/// semantics of [MS-XLSB] 2.4.323 and SheetJS' `write_BrtColInfo` helper.
#[derive(Debug, Clone)]
pub struct ColumnInfo {
    /// Column width in character units. `None` uses the sheet default.
    pub width: Option<f64>,
    /// Whether the column is hidden.
    pub hidden: bool,
    /// Whether the column width was inferred via best-fit.
    pub best_fit: bool,
}

/// Row information for a single 0-based row.
#[derive(Debug, Clone)]
pub struct RowInfo {
    /// Row height in points. `None` uses the sheet default.
    pub height: Option<f64>,
    /// Whether the row is hidden.
    pub hidden: bool,
}

/// Auto-filter configuration for a rectangular range.
///
/// The indices are 0-based and inclusive.
#[derive(Debug, Clone)]
pub struct AutoFilter {
    /// First row (0-based).
    pub row_first: u32,
    /// Last row (0-based, inclusive).
    pub row_last: u32,
    /// First column (0-based).
    pub col_first: u32,
    /// Last column (0-based, inclusive).
    pub col_last: u32,
}

/// Sheet protection options for XLSB.
///
/// This is a minimal representation used to drive the `BrtSheetProtection`
/// record. Individual flags are optional; when `None` the default from the
/// [MS-XLSB] examples / SheetJS writer is used.
#[derive(Debug, Clone, Default)]
pub struct SheetProtection {
    /// Optional password hash (Method 1). When `None`, no password is enforced.
    pub password_hash: Option<u16>,
    pub objects: Option<bool>,
    pub scenarios: Option<bool>,
    pub format_cells: Option<bool>,
    pub format_columns: Option<bool>,
    pub format_rows: Option<bool>,
    pub insert_columns: Option<bool>,
    pub insert_rows: Option<bool>,
    pub insert_hyperlinks: Option<bool>,
    pub delete_columns: Option<bool>,
    pub delete_rows: Option<bool>,
    pub select_locked_cells: Option<bool>,
    pub sort: Option<bool>,
    pub auto_filter: Option<bool>,
    pub pivot_tables: Option<bool>,
    pub select_unlocked_cells: Option<bool>,
}

/// Mutable XLSB worksheet supporting CRUD operations
#[derive(Debug, Clone)]
pub struct MutableXlsbWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), CellData>,
    max_row: u32,
    max_col: u32,
    merged_cells: Vec<MergedCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    /// Column information (0-based column index).
    columns: BTreeMap<u32, ColumnInfo>,
    /// Row information (0-based row index).
    rows: BTreeMap<u32, RowInfo>,
    /// Optional auto-filter configuration.
    auto_filter: Option<AutoFilter>,
    /// Optional sheet protection configuration.
    sheet_protection: Option<SheetProtection>,
    /// Data validation rules.
    data_validations: Vec<DataValidation>,
    data_validation_settings: DataValidationSettings,
    data_validation14_settings: DataValidationSettings,
    /// Conditional formatting rules.
    conditional_formattings: Vec<ConditionalFormatting>,
    /// Inert Office Add-in range bindings.
    web_extension_bindings: Vec<XlsbWebExtensionBinding>,
    /// Array and shared formula definitions. Cell records contain only a
    /// `PtgExp` reference to one of these definitions.
    formula_groups: Vec<FormulaGroup>,
    /// Original text for formula groups created through the text API. Binary
    /// groups intentionally have no entry and are never recompiled.
    formula_group_sources: BTreeMap<(u32, u32), String>,
    /// Structured tables (ListObjects) hosted on this sheet.
    tables: Vec<crate::xlsb::table::XlsbTable>,
    /// Typed DrawingML charts anchored on this sheet.
    charts: Vec<crate::xlsx::WorksheetChart>,
    /// Relationship ID allocated for the sheet's Drawings part.
    drawing_rel_id: Option<String>,
    /// Relationship IDs allocated for `tables` by the workbook writer, in
    /// table order. Populated during `XlsbWorkbookWriter::save`.
    pub(crate) table_rel_ids: Vec<String>,
}

pub(crate) struct ContextualFormulaRestore {
    cell_positions: Vec<(u32, u32)>,
    group_formulas: Vec<(usize, CellParsedFormula)>,
    validation_formulas: Vec<(usize, bool, bool)>,
    conditional_rule_formulas: Vec<(usize, usize)>,
    conditional_value_formulas: Vec<(usize, usize, ConditionalValueLocation)>,
}

#[derive(Debug, Clone, Copy)]
enum ConditionalValueLocation {
    ColorScaleMin,
    ColorScaleMid,
    ColorScaleMax,
    DataBarMin,
    DataBarMax,
    IconSet(usize),
    ColorScale14Min,
    ColorScale14Mid,
    ColorScale14Max,
    DataBar14Min,
    DataBar14Max,
    IconSet14(usize),
}

fn conditional_value_mut(
    rule: &mut ConditionalFormattingRule,
    location: ConditionalValueLocation,
) -> Option<&mut Cfvo> {
    match location {
        ConditionalValueLocation::ColorScaleMin => {
            rule.color_scale.as_mut().map(|scale| &mut scale.min_cfvo)
        },
        ConditionalValueLocation::ColorScaleMid => rule
            .color_scale
            .as_mut()
            .and_then(|scale| scale.mid_cfvo.as_mut()),
        ConditionalValueLocation::ColorScaleMax => {
            rule.color_scale.as_mut().map(|scale| &mut scale.max_cfvo)
        },
        ConditionalValueLocation::DataBarMin => rule.data_bar.as_mut().map(|bar| &mut bar.min_cfvo),
        ConditionalValueLocation::DataBarMax => rule.data_bar.as_mut().map(|bar| &mut bar.max_cfvo),
        ConditionalValueLocation::IconSet(index) => rule
            .icon_set
            .as_mut()
            .and_then(|set| set.cfvos.get_mut(index)),
        ConditionalValueLocation::ColorScale14Min => {
            rule.color_scale14.as_mut().map(|scale| &mut scale.min_cfvo)
        },
        ConditionalValueLocation::ColorScale14Mid => rule
            .color_scale14
            .as_mut()
            .and_then(|scale| scale.mid_cfvo.as_mut()),
        ConditionalValueLocation::ColorScale14Max => {
            rule.color_scale14.as_mut().map(|scale| &mut scale.max_cfvo)
        },
        ConditionalValueLocation::DataBar14Min => {
            rule.data_bar14.as_mut().map(|bar| &mut bar.min_cfvo)
        },
        ConditionalValueLocation::DataBar14Max => {
            rule.data_bar14.as_mut().map(|bar| &mut bar.max_cfvo)
        },
        ConditionalValueLocation::IconSet14(index) => rule
            .icon_set14
            .as_mut()
            .and_then(|set| set.cfvos.get_mut(index)),
    }
}

fn formula_requires_workbook_context(error: &XlsbError) -> bool {
    matches!(
        error,
        XlsbError::UnsupportedFeature(message)
            if message.ends_with("requires workbook compilation context")
    )
}

impl MutableXlsbWorksheet {
    /// Create a new empty worksheet
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let sheet = MutableXlsbWorksheet::new("Sheet1");
    /// ```
    pub fn new<S: Into<String>>(name: S) -> Self {
        MutableXlsbWorksheet {
            name: name.into(),
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            columns: BTreeMap::new(),
            rows: BTreeMap::new(),
            auto_filter: None,
            sheet_protection: None,
            data_validations: Vec::new(),
            data_validation_settings: DataValidationSettings::default(),
            data_validation14_settings: DataValidationSettings::default(),
            conditional_formattings: Vec::new(),
            web_extension_bindings: Vec::new(),
            formula_groups: Vec::new(),
            formula_group_sources: BTreeMap::new(),
            tables: Vec::new(),
            charts: Vec::new(),
            drawing_rel_id: None,
            table_rel_ids: Vec::new(),
        }
    }

    pub(crate) fn compile_contextual_formulas(
        &mut self,
        context: &FormulaCompilationContext<'_>,
    ) -> XlsbResult<ContextualFormulaRestore> {
        let mut compiled_groups = Vec::new();
        for (index, group) in self.formula_groups.iter().enumerate() {
            let Some(source) = self.formula_group_sources.get(&group.range.top_left()) else {
                continue;
            };
            let formula = match group.kind {
                FormulaGroupKind::Array => FormulaCompiler::compile_with_context(source, context)?,
                FormulaGroupKind::Shared => FormulaCompiler::compile_shared_with_context(
                    source,
                    group.range.row_first,
                    group.range.col_first,
                    context,
                )?,
            };
            compiled_groups.push((index, formula));
        }

        let mut compiled = Vec::new();
        for (&position, cell) in &self.cells {
            let CellValue::Formula {
                formula,
                is_array,
                array_range,
                ..
            } = &cell.value
            else {
                continue;
            };
            let is_grouped = self
                .formula_groups
                .iter()
                .any(|group| group.range.contains(position.0, position.1));
            let is_array_anchor = *is_array
                && array_range
                    .as_deref()
                    .and_then(|range| FormulaRange::parse_a1(range).ok())
                    .is_some_and(|range| range.top_left() == position);
            if cell.formula_binary.is_none() && (!is_array || is_array_anchor) && !is_grouped {
                compiled.push((
                    position,
                    FormulaCompiler::compile_with_context(formula, context)?,
                ));
            }
        }
        let mut compiled_validations = Vec::new();
        for (index, validation) in self.data_validations.iter().enumerate() {
            let formula1 = if validation.formula1_binary.is_none() && !validation.string_list {
                validation
                    .formula1
                    .as_deref()
                    .filter(|formula| !formula.is_empty())
                    .map(|formula| FormulaCompiler::compile_with_context(formula, context))
                    .transpose()?
            } else {
                None
            };
            let formula2 = if validation.formula2_binary.is_none() {
                validation
                    .formula2
                    .as_deref()
                    .filter(|formula| !formula.is_empty())
                    .map(|formula| FormulaCompiler::compile_with_context(formula, context))
                    .transpose()?
            } else {
                None
            };
            if formula1.is_some() || formula2.is_some() {
                compiled_validations.push((index, formula1, formula2));
            }
        }
        let mut compiled_conditional_rules = Vec::new();
        let mut compiled_conditional_values = Vec::new();
        for (formatting_index, formatting) in self.conditional_formattings.iter().enumerate() {
            for (rule_index, rule) in formatting.rules.iter().enumerate() {
                if rule.formulas.is_empty() && !rule.formula_texts.is_empty() {
                    let formulas = rule
                        .formula_texts
                        .iter()
                        .map(|formula| FormulaCompiler::compile_with_context(formula, context))
                        .collect::<XlsbResult<Vec<_>>>()?;
                    compiled_conditional_rules.push((formatting_index, rule_index, formulas));
                }
                let mut values = Vec::new();
                if let Some(scale) = &rule.color_scale {
                    values.push((ConditionalValueLocation::ColorScaleMin, &scale.min_cfvo));
                    if let Some(midpoint) = &scale.mid_cfvo {
                        values.push((ConditionalValueLocation::ColorScaleMid, midpoint));
                    }
                    values.push((ConditionalValueLocation::ColorScaleMax, &scale.max_cfvo));
                }
                if let Some(bar) = &rule.data_bar {
                    values.push((ConditionalValueLocation::DataBarMin, &bar.min_cfvo));
                    values.push((ConditionalValueLocation::DataBarMax, &bar.max_cfvo));
                }
                if let Some(set) = &rule.icon_set {
                    values.extend(
                        set.cfvos.iter().enumerate().map(|(index, value)| {
                            (ConditionalValueLocation::IconSet(index), value)
                        }),
                    );
                }
                if let Some(scale) = &rule.color_scale14 {
                    values.push((ConditionalValueLocation::ColorScale14Min, &scale.min_cfvo));
                    if let Some(midpoint) = &scale.mid_cfvo {
                        values.push((ConditionalValueLocation::ColorScale14Mid, midpoint));
                    }
                    values.push((ConditionalValueLocation::ColorScale14Max, &scale.max_cfvo));
                }
                if let Some(bar) = &rule.data_bar14 {
                    values.push((ConditionalValueLocation::DataBar14Min, &bar.min_cfvo));
                    values.push((ConditionalValueLocation::DataBar14Max, &bar.max_cfvo));
                }
                if let Some(set) = &rule.icon_set14 {
                    values.extend(
                        set.cfvos.iter().enumerate().map(|(index, value)| {
                            (ConditionalValueLocation::IconSet14(index), value)
                        }),
                    );
                }
                for (location, value) in values {
                    let source = value.value.as_deref().filter(|source| {
                        value.formula_binary.is_none()
                            && (value.cfvo_type == 7
                                || (matches!(value.cfvo_type, 1 | 4 | 5)
                                    && source.parse::<f64>().is_err()))
                    });
                    if let Some(source) = source {
                        compiled_conditional_values.push((
                            formatting_index,
                            rule_index,
                            location,
                            FormulaCompiler::compile_with_context(source, context)?,
                        ));
                    }
                }
            }
        }
        let positions = compiled
            .iter()
            .map(|(position, _)| *position)
            .collect::<Vec<_>>();
        for (position, formula) in compiled {
            self.cells
                .get_mut(&position)
                .expect("formula cell collected from worksheet")
                .formula_binary = Some(formula);
        }
        let group_formulas = compiled_groups
            .into_iter()
            .map(|(index, formula)| {
                let old = std::mem::replace(&mut self.formula_groups[index].formula, formula);
                (index, old)
            })
            .collect();
        let validation_formulas = compiled_validations
            .into_iter()
            .map(|(index, formula1, formula2)| {
                let restore = (index, formula1.is_some(), formula2.is_some());
                if let Some(formula) = formula1 {
                    self.data_validations[index].formula1_binary = Some(formula);
                }
                if let Some(formula) = formula2 {
                    self.data_validations[index].formula2_binary = Some(formula);
                }
                restore
            })
            .collect();
        let conditional_rule_formulas = compiled_conditional_rules
            .into_iter()
            .map(|(formatting_index, rule_index, formulas)| {
                let rule = &mut self.conditional_formattings[formatting_index].rules[rule_index];
                rule.formulas = formulas
                    .iter()
                    .map(|formula| formula.rgce.clone())
                    .collect();
                rule.formula_extras = formulas.into_iter().map(|formula| formula.rgcb).collect();
                (formatting_index, rule_index)
            })
            .collect();
        let conditional_value_formulas = compiled_conditional_values
            .into_iter()
            .map(|(formatting_index, rule_index, location, formula)| {
                conditional_value_mut(
                    &mut self.conditional_formattings[formatting_index].rules[rule_index],
                    location,
                )
                .expect("conditional value collected from worksheet")
                .formula_binary = Some(formula);
                (formatting_index, rule_index, location)
            })
            .collect();
        Ok(ContextualFormulaRestore {
            cell_positions: positions,
            group_formulas,
            validation_formulas,
            conditional_rule_formulas,
            conditional_value_formulas,
        })
    }

    pub(crate) fn clear_compiled_formulas(&mut self, restore: ContextualFormulaRestore) {
        for position in restore.cell_positions {
            if let Some(cell) = self.cells.get_mut(&position) {
                cell.formula_binary = None;
            }
        }
        for (index, formula) in restore.group_formulas {
            self.formula_groups[index].formula = formula;
        }
        for (index, clear_first, clear_second) in restore.validation_formulas {
            if clear_first {
                self.data_validations[index].formula1_binary = None;
            }
            if clear_second {
                self.data_validations[index].formula2_binary = None;
            }
        }
        for (formatting_index, rule_index) in restore.conditional_rule_formulas {
            let rule = &mut self.conditional_formattings[formatting_index].rules[rule_index];
            rule.formulas.clear();
            rule.formula_extras.clear();
        }
        for (formatting_index, rule_index, location) in restore.conditional_value_formulas {
            if let Some(value) = conditional_value_mut(
                &mut self.conditional_formattings[formatting_index].rules[rule_index],
                location,
            ) {
                value.formula_binary = None;
            }
        }
    }

    /// Get the worksheet name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Rename the worksheet
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Set a cell value
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Hello");
    /// sheet.set_cell(0, 1, 42.0);
    /// sheet.set_cell(1, 0, true);
    /// ```
    pub fn set_cell<V: Into<CellValue>>(&mut self, row: u32, col: u32, value: V) {
        self.set_cell_with_style(row, col, value, 0);
    }

    /// Set a cell value with style
    pub fn set_cell_with_style<V: Into<CellValue>>(
        &mut self,
        row: u32,
        col: u32,
        value: V,
        style: u32,
    ) {
        self.remove_formula_groups_containing(row, col);
        let cell_data = CellData {
            value: value.into(),
            style,
            formula_binary: None,
            formula_flags: 0,
        };

        self.cells.insert((row, col), cell_data);
        self.max_row = self.max_row.max(row);
        self.max_col = self.max_col.max(col);
    }

    /// Set a formula using an already encoded `CellParsedFormula`.
    ///
    /// This is the lossless path for formulas containing tokens that the text
    /// compiler does not yet understand. `cached_value` determines which
    /// `BrtFmla*` record is emitted.
    pub fn set_cell_formula_binary(
        &mut self,
        row: u32,
        col: u32,
        cached_value: CellValue,
        formula: CellParsedFormula,
        always_calculate: bool,
        style: u32,
    ) {
        self.remove_formula_groups_containing(row, col);
        let cell_data = CellData {
            value: CellValue::Formula {
                formula: String::new(),
                cached_value: Some(Box::new(cached_value)),
                is_array: false,
                array_range: None,
            },
            style,
            formula_binary: Some(formula),
            formula_flags: if always_calculate { 0x0002 } else { 0 },
        };
        self.cells.insert((row, col), cell_data);
        self.max_row = self.max_row.max(row);
        self.max_col = self.max_col.max(col);
    }

    /// Set an XLSB array formula over an inclusive, 0-based cell range.
    ///
    /// Existing values in the range are retained as cached formula results.
    /// Cells without values are marked for recalculation.
    pub fn set_array_formula(
        &mut self,
        row_first: u32,
        col_first: u32,
        row_last: u32,
        col_last: u32,
        formula: &str,
    ) -> XlsbResult<()> {
        let range = FormulaRange::new(row_first, row_last, col_first, col_last)?;
        let definition = match FormulaCompiler::compile(formula) {
            Ok(definition) => definition,
            Err(error) if formula_requires_workbook_context(&error) => {
                FormulaCompiler::compile("0")?
            },
            Err(error) => return Err(error),
        };
        let group = FormulaGroup {
            kind: FormulaGroupKind::Array,
            range,
            formula: definition,
            always_calculate: true,
        };
        self.install_formula_group(group, Some(formula))?;
        self.formula_group_sources
            .insert(range.top_left(), formula.to_string());
        Ok(())
    }

    /// Set an XLSB shared formula over an inclusive, 0-based cell range.
    ///
    /// Relative references are encoded once relative to the top-left cell and
    /// are expanded for each cell when the workbook is read.
    pub fn set_shared_formula(
        &mut self,
        row_first: u32,
        col_first: u32,
        row_last: u32,
        col_last: u32,
        formula: &str,
    ) -> XlsbResult<()> {
        let range = FormulaRange::new(row_first, row_last, col_first, col_last)?;
        let definition = match FormulaCompiler::compile_shared(formula, row_first, col_first) {
            Ok(definition) => definition,
            Err(error) if formula_requires_workbook_context(&error) => {
                FormulaCompiler::compile_shared("0", row_first, col_first)?
            },
            Err(error) => return Err(error),
        };
        let group = FormulaGroup {
            kind: FormulaGroupKind::Shared,
            range,
            formula: definition,
            always_calculate: false,
        };
        self.install_formula_group(group, Some(formula))?;
        self.formula_group_sources
            .insert(range.top_left(), formula.to_string());
        Ok(())
    }

    /// Set an array or shared formula from an already encoded definition.
    ///
    /// This is the lossless path for grouped formulas containing tokens that
    /// the text compiler or converter does not understand. Existing values in
    /// the group range are retained as cached results.
    pub fn set_formula_group_binary(&mut self, group: FormulaGroup) -> XlsbResult<()> {
        // Validate both the range and parsed-formula framing before mutating
        // the worksheet.
        let _ = group.to_record_data()?;
        if group.formula.exp_cell()?.is_some() {
            return Err(XlsbError::InvalidFormula(
                "array/shared formula definition cannot contain PtgExp".to_string(),
            ));
        }
        self.install_formula_group(group, None)
    }

    fn install_formula_group(
        &mut self,
        group: FormulaGroup,
        anchor_formula: Option<&str>,
    ) -> XlsbResult<()> {
        if let Some(index) = self
            .formula_groups
            .iter()
            .position(|existing| existing.range.top_left() == group.range.top_left())
        {
            let replaced = self.formula_groups.remove(index);
            self.formula_group_sources
                .remove(&replaced.range.top_left());
            if replaced.kind == FormulaGroupKind::Array {
                self.normalize_array_formula_ranges(&[replaced.range.to_a1()]);
            }
        }

        let range_text = group.range.to_a1();
        for row in group.range.row_first..=group.range.row_last {
            for col in group.range.col_first..=group.range.col_last {
                let cached_value = self
                    .cells
                    .get(&(row, col))
                    .and_then(|cell| match &cell.value {
                        CellValue::Empty => None,
                        CellValue::Formula { cached_value, .. } => cached_value.clone(),
                        value => Some(Box::new(value.clone())),
                    });
                let style = self.cells.get(&(row, col)).map_or(0, |cell| cell.style);
                let decoded = || -> XlsbResult<String> {
                    let tokens = match group.kind {
                        FormulaGroupKind::Array => {
                            FormulaParser::with_extra(&group.formula.rgce, &group.formula.rgcb)
                                .parse()?
                        },
                        FormulaGroupKind::Shared => FormulaParser::with_base_cell_and_extra(
                            &group.formula.rgce,
                            &group.formula.rgcb,
                            row,
                            col,
                        )
                        .parse()?,
                    };
                    FormulaConverter::try_tokens_to_string(&tokens)
                };
                let formula = match (group.kind, anchor_formula) {
                    (FormulaGroupKind::Array, Some(formula)) => formula.to_string(),
                    (FormulaGroupKind::Shared, _) => decoded().or_else(|error| {
                        if anchor_formula.is_none() {
                            Ok(String::new())
                        } else {
                            Err(error)
                        }
                    })?,
                    (FormulaGroupKind::Array, None) => decoded().unwrap_or_default(),
                };
                self.cells.insert(
                    (row, col),
                    CellData {
                        value: CellValue::Formula {
                            formula,
                            cached_value,
                            is_array: group.kind == FormulaGroupKind::Array,
                            array_range: (group.kind == FormulaGroupKind::Array)
                                .then(|| range_text.clone()),
                        },
                        style,
                        formula_binary: None,
                        formula_flags: 0,
                    },
                );
            }
        }
        self.max_row = self.max_row.max(group.range.row_last);
        self.max_col = self.max_col.max(group.range.col_last);
        self.formula_groups.push(group);
        Ok(())
    }

    fn remove_formula_groups_containing(&mut self, row: u32, col: u32) {
        let removed_anchors = self
            .formula_groups
            .iter()
            .filter(|group| group.range.contains(row, col))
            .map(|group| group.range.top_left())
            .collect::<Vec<_>>();
        let removed_array_ranges: Vec<String> = self
            .formula_groups
            .iter()
            .filter(|group| group.kind == FormulaGroupKind::Array && group.range.contains(row, col))
            .map(|group| group.range.to_a1())
            .collect();
        self.formula_groups
            .retain(|group| !group.range.contains(row, col));
        for anchor in removed_anchors {
            self.formula_group_sources.remove(&anchor);
        }
        if removed_array_ranges.is_empty() {
            return;
        }
        self.normalize_array_formula_ranges(&removed_array_ranges);
    }

    fn normalize_array_formula_ranges(&mut self, ranges: &[String]) {
        for cell in self.cells.values_mut() {
            if let CellValue::Formula {
                is_array,
                array_range,
                ..
            } = &mut cell.value
            {
                if array_range
                    .as_ref()
                    .is_some_and(|range| ranges.contains(range))
                {
                    *is_array = false;
                    *array_range = None;
                }
            }
        }
    }

    fn dissolve_formula_groups_for_structure_change(&mut self) {
        self.formula_groups.clear();
        self.formula_group_sources.clear();
        for cell in self.cells.values_mut() {
            if let CellValue::Formula {
                is_array,
                array_range,
                ..
            } = &mut cell.value
            {
                *is_array = false;
                *array_range = None;
            }
        }
    }

    /// Get a cell value
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&CellValue> {
        self.cells.get(&(row, col)).map(|c| &c.value)
    }

    /// Delete a cell
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Hello");
    /// sheet.delete_cell(0, 0);
    /// assert!(sheet.get_cell(0, 0).is_none());
    /// ```
    pub fn delete_cell(&mut self, row: u32, col: u32) -> Option<CellValue> {
        self.remove_formula_groups_containing(row, col);
        self.cells.remove(&(row, col)).map(|c| c.value)
    }

    /// Clear all cells in the worksheet
    pub fn clear(&mut self) {
        self.cells.clear();
        self.max_row = 0;
        self.max_col = 0;
        self.merged_cells.clear();
        self.hyperlinks.clear();
        self.comments.clear();
        self.columns.clear();
        self.rows.clear();
        self.auto_filter = None;
        self.sheet_protection = None;
        self.data_validations.clear();
        self.data_validation_settings = DataValidationSettings::default();
        self.data_validation14_settings = DataValidationSettings::default();
        self.conditional_formattings.clear();
        self.charts.clear();
        self.drawing_rel_id = None;
        self.formula_groups.clear();
        self.formula_group_sources.clear();
    }

    /// Set a custom column width (in character units) for a 0-based column.
    ///
    /// This controls the `BrtColInfo` width field. The default width from the
    /// sheet format properties (`BrtSheetFormatPr`) is used when no explicit
    /// width is set.
    pub fn set_column_width(&mut self, col: u32, width: f64) {
        let entry = self.columns.entry(col).or_insert(ColumnInfo {
            width: None,
            hidden: false,
            best_fit: false,
        });
        entry.width = Some(width);
    }

    /// Set whether a zero-based column is hidden.
    pub fn set_column_hidden(&mut self, col: u32, hidden: bool) {
        let entry = self.columns.entry(col).or_insert(ColumnInfo {
            width: None,
            hidden: false,
            best_fit: false,
        });
        entry.hidden = hidden;
    }

    /// Set whether a zero-based column width was automatically best-fit.
    pub fn set_column_best_fit(&mut self, col: u32, best_fit: bool) {
        let entry = self.columns.entry(col).or_insert(ColumnInfo {
            width: None,
            hidden: false,
            best_fit: false,
        });
        entry.best_fit = best_fit;
    }

    /// Set a custom row height (in points) for a 0-based row.
    ///
    /// Heights are encoded in twips (1/20 of a point) in the `BrtRowHdr`
    /// records. When no explicit height is set, Excel's default of 15 points
    /// (300 twips) is used.
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        let entry = self.rows.entry(row).or_insert(RowInfo {
            height: None,
            hidden: false,
        });
        entry.height = Some(height);
    }

    /// Set whether a zero-based row is hidden.
    pub fn set_row_hidden(&mut self, row: u32, hidden: bool) {
        let entry = self.rows.entry(row).or_insert(RowInfo {
            height: None,
            hidden: false,
        });
        entry.hidden = hidden;
    }

    /// Configure a basic auto-filter range for the worksheet.
    ///
    /// The indices are 0-based and inclusive.
    pub fn set_auto_filter(
        &mut self,
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
    ) {
        self.auto_filter = Some(AutoFilter {
            row_first,
            row_last,
            col_first,
            col_last,
        });
    }

    /// Set sheet protection options. Passing `None` clears protection.
    pub fn set_sheet_protection(&mut self, protection: Option<SheetProtection>) {
        self.sheet_protection = protection;
    }

    /// Add a merged cell range
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi_ooxml::xlsb::advanced_features::MergedCell;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.add_merged_cell(MergedCell::new(0, 1, 0, 1)); // Merge A1:B2
    /// ```
    pub fn add_merged_cell(&mut self, merged: MergedCell) {
        self.merged_cells.push(merged);
    }

    /// Add a hyperlink
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi_ooxml::xlsb::advanced_features::Hyperlink;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string())
    ///     .with_tooltip("Visit website".to_string());
    /// sheet.add_hyperlink(link);
    /// ```
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    /// Add a comment
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi_ooxml::xlsb::advanced_features::Comment;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let comment = Comment::new(0, 0, "John".to_string(), "Important note".to_string());
    /// sheet.add_comment(comment);
    /// ```
    pub fn add_comment(&mut self, comment: Comment) {
        self.comments.push(comment);
    }

    /// Get all merged cells
    pub fn merged_cells(&self) -> &[MergedCell] {
        &self.merged_cells
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// Get mutable access to all hyperlinks.
    ///
    /// This is primarily used by the workbook writer to inject concrete
    /// relationship IDs (`rId`) after creating external OPC relationships
    /// but before serializing `BrtHLink` records.
    pub(crate) fn hyperlinks_mut(&mut self) -> &mut [Hyperlink] {
        &mut self.hyperlinks
    }

    /// Get all comments
    /// Add a structured table (ListObject) hosted on this sheet.
    ///
    /// The table part is serialized and related from this worksheet when the
    /// workbook is saved. The display name is required by Excel; the column
    /// list, when present, must match the range width.
    pub fn add_table(&mut self, table: crate::xlsb::table::XlsbTable) -> XlsbResult<()> {
        const MAX_TABLES_PER_SHEET: usize = 4_096;
        if self.tables.len() >= MAX_TABLES_PER_SHEET {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "worksheet table count exceeds the safety limit".to_string(),
            ));
        }
        if table.display_name.as_deref().is_none_or(str::is_empty) {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "structured table requires a display name".to_string(),
            ));
        }
        let range = &table.range;
        if range.first_row > range.last_row || range.first_column > range.last_column {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "structured table range is inverted".to_string(),
            ));
        }
        let width = u64::from(range.last_column) - u64::from(range.first_column) + 1;
        if !table.columns.is_empty() && table.columns.len() as u64 != width {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "structured table declares {} columns for a range {width} wide",
                table.columns.len()
            )));
        }
        if self.tables.iter().any(|existing| existing.id == table.id) {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "duplicate structured table id {}",
                table.id
            )));
        }
        self.tables.push(table);
        Ok(())
    }

    /// The structured tables hosted on this sheet.
    pub fn tables(&self) -> &[crate::xlsb::table::XlsbTable] {
        &self.tables
    }

    /// Add a typed DrawingML chart anchored on this worksheet.
    ///
    /// XLSB uses the same chart and SpreadsheetDrawing XML as XLSX. The
    /// workbook writer emits the chart and drawing parts and stores their
    /// relationship in the binary `BrtDrawing` record. Relationship-bearing
    /// external data, user shapes, extension fragments, and pivot charts are
    /// rejected until their complete package graphs can be authored.
    pub fn add_chart(&mut self, chart: crate::xlsx::WorksheetChart) -> XlsbResult<()> {
        if self.charts.len() >= crate::xlsb::drawing_write::MAX_CHARTS_PER_SHEET {
            return Err(XlsbError::InvalidFormula(
                "worksheet chart count exceeds the safety limit".to_string(),
            ));
        }
        let _ = crate::xlsb::drawing_write::serialize_chart(&chart)?;
        self.charts.push(chart);
        Ok(())
    }

    /// Typed DrawingML charts in drawing order.
    pub fn charts(&self) -> &[crate::xlsx::WorksheetChart] {
        &self.charts
    }

    /// Remove one chart by drawing order.
    pub fn remove_chart(&mut self, index: usize) -> XlsbResult<crate::xlsx::WorksheetChart> {
        if index >= self.charts.len() {
            return Err(XlsbError::InvalidFormula(format!(
                "chart index {index} is out of bounds for {} charts",
                self.charts.len()
            )));
        }
        Ok(self.charts.remove(index))
    }

    /// Remove every authored chart from this worksheet.
    pub fn clear_charts(&mut self) {
        self.charts.clear();
        self.drawing_rel_id = None;
    }

    pub(crate) fn set_drawing_rel_id(&mut self, rel_id: Option<String>) {
        self.drawing_rel_id = rel_id;
    }

    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Add a data validation rule.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi_ooxml::xlsb::data_validation::DataValidation;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let mut dv = DataValidation::new(3, "A1:A10".to_string()); // list
    /// dv.formula1 = Some("Yes,No".to_string());
    /// sheet.add_data_validation(dv);
    /// ```
    pub fn add_data_validation(&mut self, dv: DataValidation) {
        self.data_validations.push(dv);
    }

    /// Get all data validations.
    pub fn data_validations(&self) -> &[DataValidation] {
        &self.data_validations
    }

    /// Set UI prompt settings for classic `BrtDVal` rules.
    pub fn set_data_validation_settings(&mut self, settings: DataValidationSettings) {
        self.data_validation_settings = settings;
    }

    /// Set UI prompt settings for Office 2013 `BrtDVal14` rules.
    pub fn set_data_validation14_settings(&mut self, settings: DataValidationSettings) {
        self.data_validation14_settings = settings;
    }

    /// Add a conditional formatting block.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi_ooxml::xlsb::conditional_formatting::{
    ///     ConditionalFormatting, ConditionalFormattingRule, CfRuleType,
    /// };
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let mut cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
    /// let rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
    /// cf.add_rule(rule);
    /// sheet.add_conditional_formatting(cf);
    /// ```
    pub fn add_conditional_formatting(&mut self, cf: ConditionalFormatting) {
        self.conditional_formattings.push(cf);
    }

    /// Get all conditional formatting blocks.
    pub fn conditional_formattings(&self) -> &[ConditionalFormatting] {
        &self.conditional_formattings
    }

    /// Replace worksheet Office Add-in bindings after validating their payloads
    /// and unique application references.
    pub fn set_web_extension_bindings(
        &mut self,
        bindings: Vec<XlsbWebExtensionBinding>,
    ) -> XlsbResult<()> {
        let mut app_refs = HashSet::with_capacity(bindings.len());
        for binding in &bindings {
            binding.to_payload()?;
            if !app_refs.insert(binding.application_reference.as_str()) {
                return Err(XlsbError::Unrecognized {
                    typ: "WEBEXTENSIONS".to_string(),
                    val: "duplicate binding appRef".to_string(),
                });
            }
        }
        if bindings.len() > 65_536 {
            return Err(XlsbError::Unrecognized {
                typ: "WEBEXTENSIONS".to_string(),
                val: "binding count exceeds 65,536".to_string(),
            });
        }
        self.web_extension_bindings = bindings;
        Ok(())
    }

    /// Office Add-in bindings that will be written to this worksheet.
    pub fn web_extension_bindings(&self) -> &[XlsbWebExtensionBinding] {
        &self.web_extension_bindings
    }

    /// Get the number of non-empty cells
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get dimensions (min_row, min_col, max_row, max_col)
    pub fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            None
        } else {
            Some((0, 0, self.max_row, self.max_col))
        }
    }

    /// Delete a row (shifts remaining rows up)
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Row 0");
    /// sheet.set_cell(1, 0, "Row 1");
    /// sheet.set_cell(2, 0, "Row 2");
    ///
    /// sheet.delete_row(1);
    ///
    /// // Row 2 becomes row 1
    /// assert_eq!(sheet.get_cell(1, 0).and_then(|v| v.as_str()), Some("Row 2"));
    /// ```
    pub fn delete_row(&mut self, row: u32) {
        self.dissolve_formula_groups_for_structure_change();
        // Remove all cells in the row
        self.cells.retain(|(r, _), _| *r != row);

        // Shift rows after the deleted row up
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((r, _), _)| *r > row)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r - 1, c), cell);
        }

        // Recalculate max_row
        self.max_row = self.cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
    }

    /// Delete a column (shifts remaining columns left)
    pub fn delete_column(&mut self, col: u32) {
        self.dissolve_formula_groups_for_structure_change();
        // Remove all cells in the column
        self.cells.retain(|(_, c), _| *c != col);

        // Shift columns after the deleted column left
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((_, c), _)| *c > col)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r, c - 1), cell);
        }

        // Recalculate max_col
        self.max_col = self.cells.keys().map(|(_, c)| *c).max().unwrap_or(0);
    }

    /// Insert a row (shifts existing rows down)
    pub fn insert_row(&mut self, row: u32) {
        self.dissolve_formula_groups_for_structure_change();
        // Shift rows at and after the insert position down
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((r, _), _)| *r >= row)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r + 1, c), cell);
        }

        // Recalculate max_row
        self.max_row = self.cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
    }

    /// Insert a column (shifts existing columns right)
    pub fn insert_column(&mut self, col: u32) {
        self.dissolve_formula_groups_for_structure_change();
        // Shift columns at and after the insert position right
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((_, c), _)| *c >= col)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r, c + 1), cell);
        }

        // Recalculate max_col
        self.max_col = self.cells.keys().map(|(_, c)| *c).max().unwrap_or(0);
    }

    /// Write worksheet to binary format
    ///
    /// Following Excel's required structure
    pub(crate) fn write<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        shared_strings: &mut crate::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        // Write BrtBeginSheet
        writer.write_record(record_types::BEGIN_SHEET, &[])?;

        // Write worksheet properties and basic formatting information.
        self.write_ws_properties(writer)?;

        // Write worksheet dimensions
        self.write_dimensions(writer)?;

        // Write worksheet views (minimal SheetJS-style layout)
        self.write_ws_views(writer)?;

        // Write sheet formatting properties (BrtSheetFormatPr)
        self.write_sheet_format_pr(writer)?;

        // Column information (BrtBeginColInfos / BrtColInfo / BrtEndColInfos)
        self.write_col_infos(writer)?;

        // Write sheet data
        writer.write_record(record_types::BEGIN_SHEET_DATA, &[])?;
        self.write_cells(writer, shared_strings)?;
        writer.write_record(record_types::END_SHEET_DATA, &[])?;

        // Sheet protection (BrtSheetProtection) - minimal skeleton mirroring
        // SheetJS and [MS-XLSB] examples.
        self.write_sheet_protection(writer)?;

        // AutoFilter skeleton (BrtBeginAFilter / BrtEndAFilter).
        self.write_auto_filter(writer)?;

        // Write merged cells if present
        if !self.merged_cells.is_empty() {
            self.write_merged_cells(writer)?;
        }

        // Write hyperlinks if present
        if !self.hyperlinks.is_empty() {
            self.write_hyperlinks(writer)?;
        }

        // Write data validations if present
        if !self.data_validations.is_empty() {
            crate::xlsb::writer::data_validation::write_data_validations(
                writer,
                &self.data_validations,
                self.data_validation_settings,
                self.data_validation14_settings,
            )?;
        }

        // Write conditional formatting if present
        if !self.conditional_formattings.is_empty() {
            crate::xlsb::writer::conditional_formatting::write_conditional_formattings(
                writer,
                &self.conditional_formattings,
            )?;
        }

        if !self.web_extension_bindings.is_empty() {
            writer.write_record(record_types::BEGIN_WEB_EXTENSIONS, &[])?;
            for binding in &self.web_extension_bindings {
                writer.write_record(record_types::WEB_EXTENSION, &binding.to_payload()?)?;
            }
            writer.write_record(record_types::END_WEB_EXTENSIONS, &[])?;
        }

        if !self.charts.is_empty() {
            let rel_id = self.drawing_rel_id.as_deref().ok_or_else(|| {
                XlsbError::InvalidFormula(
                    "worksheet charts lack a Drawings relationship ID".to_string(),
                )
            })?;
            let mut payload = Vec::with_capacity(4 + rel_id.len() * 2);
            RecordWriter::new(&mut payload).write_wide_string(rel_id)?;
            writer.write_record(record_types::DRAWING, &payload)?;
        }

        // Write table references (BrtBeginListParts / BrtListPart /
        // BrtEndListParts) after all other sheet features.
        if !self.tables.is_empty() {
            if self.table_rel_ids.len() != self.tables.len() {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "worksheet tables lack relationship IDs from the workbook writer".to_string(),
                ));
            }
            crate::xlsb::table::write::write_list_parts(writer, &self.table_rel_ids)?;
        }

        // Write BrtEndSheet
        writer.write_record(record_types::END_SHEET, &[])?;

        Ok(())
    }

    /// Write worksheet properties (BrtWsProp) - REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.864 + spec example 3.7.21: 23 bytes total
    /// Structure: flags (3 bytes) + brtcolorTab (8 bytes) + rwSync (4) + colSync (4) + strName (4)
    fn write_ws_properties<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Flags (3 bytes per spec example 3.7.21):
        // Byte 0-1 (USHORT): flags A-O
        // Byte 2 (BYTE): flags P-Q + reserved
        //
        // From spec example: 0xC9, 0x04, 0x02
        // 0xC9 = fShowAutoBreaks(1) + fPublish(1) + fRowSumsBelow(1) + fColSumsRight(1) + fShowOutlineSymbols(1)
        // 0x04 = remaining bits
        // 0x02 = fCondFmtCalc(1) at bit 1
        temp_writer.write_u8(0xC9)?;
        temp_writer.write_u8(0x04)?;
        temp_writer.write_u8(0x02)?; // Third byte - fCondFmtCalc flag

        // brtcolorTab (8 bytes) - BrtColor structure
        // From spec example: xColorType=0x00 (auto), index=0x40
        temp_writer.write_u8(0x00)?; // fValidRGB(0) + xColorType(0x00)
        temp_writer.write_u8(0x40)?; // index
        temp_writer.write_u16(0)?; // nTintAndShade
        temp_writer.write_u8(0)?; // bRed
        temp_writer.write_u8(0)?; // bGreen
        temp_writer.write_u8(0)?; // bBlue
        temp_writer.write_u8(0)?; // bAlpha

        // rwSync (4 bytes) - RwNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // colSync (4 bytes) - ColNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // strName - CodeName (XLWideString): empty string
        temp_writer.write_u32(0)?;

        writer.write_record(record_types::WS_PROP, &data)?;
        Ok(())
    }

    /// Write worksheet views (REQUIRED by Excel)
    ///
    /// [MS-XLSB] 2.4.304: Specifies sheet view settings
    fn write_ws_views<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        writer.write_record(record_types::BEGIN_WS_VIEWS, &[])?;

        // BrtBeginWsView (30 bytes according to spec)
        let mut view_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut view_data);

        // Flags (2 bytes) - bits A-K + reserved
        // Default: fDspGrid=1, fDspRwCol=1, fDspZeros=1, fDefaultHdr=1
        // 0xDC = 11011100 = fDefaultHdr(1) + fDspGuts(1) + fSelected(1) + fDspZeros(1) + fDspRwCol(1) + fDspGrid(1)
        // 0x03 = 00000011 = reserved bits
        temp_writer.write_u8(0xDC)?;
        temp_writer.write_u8(0x03)?;

        // xlView (4 bytes) - XLView: 0 = normal view
        temp_writer.write_u32(0)?;

        // rwTop (4 bytes) - first row displayed
        temp_writer.write_u32(0)?;

        // colLeft (4 bytes) - first column displayed
        temp_writer.write_u32(0)?;

        // icvHdr (1 byte) - Icv: gridline color (0x40 = default)
        temp_writer.write_u8(0x40)?;

        // reserved2 (1 byte)
        temp_writer.write_u8(0)?;

        // reserved3 (2 bytes)
        temp_writer.write_u16(0)?;

        // wScale (2 bytes) - zoom level (100%)
        temp_writer.write_u16(100)?;

        // wScaleNormal (2 bytes) - per spec example: 0 means default 100
        temp_writer.write_u16(0)?;

        // wScaleSLV (2 bytes) - zoom for page break preview (0 = default 100%)
        temp_writer.write_u16(0)?;

        // wScalePLV (2 bytes) - zoom for page layout view (0 = default 100%)
        temp_writer.write_u16(0)?;

        // iWbkView (4 bytes) - workbook view index
        temp_writer.write_u32(0)?;

        // Minimal SheetJS-style view: BrtBeginWsViews / BrtBeginWsView / BrtEndWsView / BrtEndWsViews
        writer.write_record(record_types::BEGIN_WS_VIEW, &view_data)?;

        writer.write_record(record_types::END_WS_VIEW, &[])?;
        writer.write_record(record_types::END_WS_VIEWS, &[])?;

        Ok(())
    }

    /// Write SHEET_FORMAT_PR record (0x01E5) - sheet formatting properties
    /// REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.862 + spec example 3.7.28: 12 bytes total
    fn write_sheet_format_pr<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // dxGCol (4 bytes) - 0xFFFFFFFF = use cchDefColWidth instead
        temp_writer.write_u32(0xFFFFFFFF)?;

        // cchDefColWidth (2 bytes) - default column width in characters
        // Spec example 3.7.28: 0x0008 (8 characters)
        temp_writer.write_u16(8)?;

        // miyDefRwHeight (2 bytes) - default row height in twips
        // Spec example 3.7.28: 0x012C (300 twips = 15 points)
        temp_writer.write_u16(300)?;

        // Flags (4 bytes): all zeros per spec example
        // fUnsynced=0, fDyZero=0, fExAsc=0, fExDesc=0, reserved=0, iOutLevelRw=0, iOutLevelCol=0
        temp_writer.write_u32(0)?;

        writer.write_record(0x01E5, &data)?;
        Ok(())
    }

    /// Write worksheet dimensions record
    fn write_dimensions<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        if let Some((min_row, min_col, max_row, max_col)) = self.dimensions() {
            let (max_row, max_col) = self
                .formula_groups_for_write()?
                .iter()
                .fold((max_row, max_col), |(row, col), group| {
                    (row.max(group.range.row_last), col.max(group.range.col_last))
                });
            temp_writer.write_u32(min_row)?;
            temp_writer.write_u32(max_row)?;
            temp_writer.write_u32(min_col)?;
            temp_writer.write_u32(max_col)?;
        } else {
            // Empty worksheet
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
        }

        writer.write_record(record_types::WS_DIM, &data)?;
        Ok(())
    }

    /// Write all cells
    fn write_cells<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        shared_strings: &mut crate::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        let formula_groups = self.formula_groups_for_write()?;
        if formula_groups.is_empty() {
            return self.write_cells_from(writer, shared_strings, &self.cells, &formula_groups);
        }

        // Grouped formulas require a formula cell record at every position in
        // their range. Materialize only for this uncommon path, keeping the
        // ordinary worksheet writer allocation-free.
        let mut cells = self.cells.clone();
        for group in &formula_groups {
            for row in group.range.row_first..=group.range.row_last {
                for col in group.range.col_first..=group.range.col_last {
                    cells.entry((row, col)).or_insert_with(|| CellData {
                        value: CellValue::Empty,
                        style: 0,
                        formula_binary: None,
                        formula_flags: 0,
                    });
                }
            }
        }
        self.write_cells_from(writer, shared_strings, &cells, &formula_groups)
    }

    fn write_cells_from<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        shared_strings: &mut crate::xlsb::writer::MutableSharedStringsWriter,
        cells: &BTreeMap<(u32, u32), CellData>,
        formula_groups: &[FormulaGroup],
    ) -> XlsbResult<()> {
        let mut current_row: Option<u32> = None;

        for ((row, col), cell_data) in cells {
            // Write row header if row changed
            if current_row != Some(*row) {
                self.write_row_header(writer, *row, cells)?;
                current_row = Some(*row);
            }

            if let Some(group) = Self::formula_group_for_cell(formula_groups, *row, *col) {
                self.write_grouped_formula_cell(writer, *row, *col, cell_data, group)?;
            } else {
                self.write_cell(writer, *row, *col, cell_data, shared_strings)?;
            }
        }

        Ok(())
    }

    fn formula_groups_for_write(&self) -> XlsbResult<Vec<FormulaGroup>> {
        let mut groups = self.formula_groups.clone();
        for (&position, cell) in &self.cells {
            let CellValue::Formula {
                formula,
                cached_value,
                is_array: true,
                array_range,
            } = &cell.value
            else {
                continue;
            };
            let range_text = array_range.as_deref().ok_or_else(|| {
                XlsbError::InvalidFormula(format!(
                    "array formula at {} has no array range",
                    crate::xlsb::utils::cell_reference(position.0, position.1)
                ))
            })?;
            let range = FormulaRange::parse_a1(range_text)?;
            if range.top_left() != position {
                continue;
            }
            if groups
                .iter()
                .any(|group| group.kind == FormulaGroupKind::Array && group.range == range)
            {
                continue;
            }
            groups.push(FormulaGroup {
                kind: FormulaGroupKind::Array,
                range,
                formula: if let Some(formula) = &cell.formula_binary {
                    formula.clone()
                } else {
                    FormulaCompiler::compile(formula)?
                },
                always_calculate: cached_value.is_none(),
            });
        }

        for (index, group) in groups.iter().enumerate() {
            if group.formula.exp_cell()?.is_some() {
                return Err(XlsbError::InvalidFormula(
                    "array/shared formula definition cannot contain PtgExp".to_string(),
                ));
            }
            if groups[..index]
                .iter()
                .any(|existing| existing.range.top_left() == group.range.top_left())
            {
                return Err(XlsbError::InvalidFormula(format!(
                    "multiple formula definitions cannot share anchor {}",
                    crate::xlsb::utils::cell_reference(
                        group.range.row_first,
                        group.range.col_first
                    )
                )));
            }
        }
        Ok(groups)
    }

    fn formula_group_for_cell(
        groups: &[FormulaGroup],
        row: u32,
        col: u32,
    ) -> Option<&FormulaGroup> {
        groups
            .iter()
            .find(|group| group.range.top_left() == (row, col))
            .or_else(|| {
                groups
                    .iter()
                    .filter(|group| group.range.contains(row, col))
                    .min_by_key(|group| {
                        u64::from(group.range.row_last - group.range.row_first + 1)
                            * u64::from(group.range.col_last - group.range.col_first + 1)
                    })
            })
    }

    fn write_grouped_formula_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        row: u32,
        col: u32,
        cell_data: &CellData,
        group: &FormulaGroup,
    ) -> XlsbResult<()> {
        let cached_value = match &cell_data.value {
            CellValue::Formula { cached_value, .. } => cached_value.as_deref(),
            CellValue::Empty => None,
            value => Some(value),
        };
        let placeholder = CellParsedFormula::exp(group.range.row_first, group.range.col_first)?;
        self.write_formula_cell(
            writer,
            col,
            cell_data.style,
            "",
            cached_value,
            false,
            Some(&placeholder),
            cell_data.formula_flags,
        )?;

        if group.range.top_left() == (row, col) {
            let record_type = match group.kind {
                FormulaGroupKind::Array => record_types::ARR_FMLA,
                FormulaGroupKind::Shared => record_types::SHR_FMLA,
            };
            writer.write_record(record_type, &group.to_record_data()?)?;
        }
        Ok(())
    }

    /// Write row header record with BrtColSpan elements
    ///
    /// BrtRowHdr structure (2.4.761):
    /// - rw (4 bytes): Row index
    /// - ixfe (4 bytes): Style index
    /// - miyRw (2 bytes): Row height in twips (1/20 of a point)
    /// - flags1 (1 byte): fExtraAsc | fExtraDsc | reserved
    /// - flags2 (1 byte): outline/visibility flags
    /// - phonetic (1 byte): phonetic guide flags
    /// - ccolspan (4 bytes): number of BrtColSpan elements
    /// - rgBrtColspan (variable): array of BrtColSpan, each 8 bytes
    ///   (colFirst (u32) + colLast (u32))
    fn write_row_header<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        row: u32,
        cells: &BTreeMap<(u32, u32), CellData>,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Fixed part
        temp_writer.write_u32(row)?; // rw: Row index
        temp_writer.write_u32(0)?; // ixfe: Style index (0 = default)

        // Row height in twips (1/20 of a point). When no explicit height is
        // configured, use Excel's default of 15 points (300 twips).
        let miy_rw: u16 = if let Some(info) = self.rows.get(&row) {
            if let Some(height_pts) = info.height {
                (height_pts * 20.0).round() as u16
            } else {
                0x012C
            }
        } else {
            0x012C
        };
        temp_writer.write_u16(miy_rw)?;

        // flags1: extra ascender/descender padding (unused here).
        temp_writer.write_u8(0)?;

        // flags2: outline / visibility / custom height flags.
        // Bits 0-2: outline level, 0x10: hidden, 0x20: custom height.
        let mut flags2: u8 = 0;
        if let Some(info) = self.rows.get(&row) {
            if info.hidden {
                flags2 |= 0x10;
            }
            if info.height.is_some() {
                flags2 |= 0x20;
            }
        }
        temp_writer.write_u8(flags2)?;

        // phonetic guide: 0 = no phonetic information
        temp_writer.write_u8(0)?;

        // Collect all columns that have cells in this row (BTreeMap preserves sorted order)
        let cells_in_row: Vec<u32> = cells
            .keys()
            .filter(|(r, _)| *r == row)
            .map(|(_, c)| *c)
            .collect();

        if cells_in_row.is_empty() {
            // No cells in row - write 0 colspans
            temp_writer.write_u32(0)?;
        } else {
            // Group columns by 1024-wide segments, as in [MS-XLSB] BrtColSpan and SheetJS
            let mut spans: Vec<(u32, u32)> = Vec::new();
            let mut current_segment = cells_in_row[0] / 1024;
            let mut segment_first = cells_in_row[0];
            let mut segment_last = cells_in_row[0];

            for &col in &cells_in_row[1..] {
                let segment = col / 1024;
                if segment == current_segment {
                    segment_last = col;
                } else {
                    spans.push((segment_first, segment_last));
                    current_segment = segment;
                    segment_first = col;
                    segment_last = col;
                }
            }
            spans.push((segment_first, segment_last));

            // Number of spans
            temp_writer.write_u32(spans.len() as u32)?;

            // Each span is a BrtColSpan: colFirst (u32) + colLast (u32)
            for (first, last) in spans {
                temp_writer.write_u32(first)?; // colFirst
                temp_writer.write_u32(last)?; // colLast
            }
        }

        writer.write_record(record_types::ROW_HDR, &data)?;
        Ok(())
    }

    /// Write a single cell record
    fn write_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        _row: u32,
        col: u32,
        cell_data: &CellData,
        shared_strings: &mut crate::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        match &cell_data.value {
            CellValue::Empty => self.write_blank_cell(writer, col, cell_data.style)?,
            CellValue::String(s) => {
                self.write_shared_string_cell(writer, col, s, cell_data.style, shared_strings)?
            },
            CellValue::Int(i) => self.write_number_cell(writer, col, *i as f64, cell_data.style)?,
            CellValue::Float(f) => self.write_number_cell(writer, col, *f, cell_data.style)?,
            CellValue::Bool(b) => self.write_bool_cell(writer, col, *b, cell_data.style)?,
            CellValue::Error(e) => self.write_error_cell(writer, col, e, cell_data.style)?,
            CellValue::DateTime(dt) => {
                // Excel DateTime is already stored as serial number (days since epoch)
                // CellValue::DateTime stores the Excel serial number directly
                self.write_number_cell(writer, col, *dt, cell_data.style)?;
            },
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                ..
            } => self.write_formula_cell(
                writer,
                col,
                cell_data.style,
                formula,
                cached_value.as_deref(),
                *is_array,
                cell_data.formula_binary.as_ref(),
                cell_data.formula_flags,
            )?,
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn write_formula_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        style: u32,
        formula_text: &str,
        cached_value: Option<&CellValue>,
        is_array: bool,
        encoded: Option<&CellParsedFormula>,
        flags: u16,
    ) -> XlsbResult<()> {
        if is_array {
            return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(
                "XLSB array formula writing requires BrtArrFmla".to_string(),
            ));
        }
        if flags & !0x0002 != 0 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "invalid GrbitFmla flags 0x{flags:04X}"
            )));
        }
        let effective_flags = if cached_value.is_none() {
            flags | 0x0002
        } else {
            flags
        };

        let compiled;
        let parsed = if let Some(encoded) = encoded {
            encoded
        } else {
            compiled = FormulaCompiler::compile(formula_text)?;
            &compiled
        };
        let formula_bytes = parsed.to_bytes()?;
        let cached = cached_value.unwrap_or(&CellValue::Float(0.0));

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);
        Self::write_cell_structure(&mut temp_writer, col, style)?;

        let record_type = match cached {
            CellValue::String(value) => {
                temp_writer.write_wide_string(value)?;
                record_types::FMLA_STRING
            },
            CellValue::Bool(value) => {
                temp_writer.write_u8(u8::from(*value))?;
                record_types::FMLA_BOOL
            },
            CellValue::Error(error) => {
                temp_writer.write_u8(Self::error_code(error))?;
                record_types::FMLA_ERROR
            },
            CellValue::Int(value) => {
                temp_writer.write_f64(*value as f64)?;
                record_types::FMLA_NUM
            },
            CellValue::Float(value) | CellValue::DateTime(value) => {
                temp_writer.write_f64(*value)?;
                record_types::FMLA_NUM
            },
            CellValue::Empty => {
                temp_writer.write_f64(0.0)?;
                record_types::FMLA_NUM
            },
            CellValue::Formula { .. } => {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "formula cached value cannot itself be a formula".to_string(),
                ));
            },
        };
        temp_writer.write_u16(effective_flags)?;
        data.extend_from_slice(&formula_bytes);
        writer.write_record(record_type, &data)
    }

    /// Write the Cell structure (2.5.10) - 8 bytes
    ///
    /// Cell structure:
    /// - column (4 bytes): Column index
    /// - iStyleRef (3 bytes, 24-bit): Style XF index
    /// - fPhShow (1 bit): Phonetic info flag
    /// - reserved (7 bits): Reserved
    fn write_cell_structure<W: Write>(
        temp_writer: &mut RecordWriter<W>,
        col: u32,
        style: u32,
    ) -> XlsbResult<()> {
        // Column (4 bytes)
        temp_writer.write_u32(col)?;

        // iStyleRef (3 bytes) + flags (1 byte) = 4 bytes total
        temp_writer.write_u8((style & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 8) & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 16) & 0xFF) as u8)?;
        temp_writer.write_u8(0)?; // fPhShow=0, reserved=0

        Ok(())
    }

    /// Write a blank cell (BrtCellBlank - 8 bytes)
    fn write_blank_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        Self::write_cell_structure(&mut temp_writer, col, style)?;

        writer.write_record(record_types::CELL_BLANK, &data)?;
        Ok(())
    }

    /// Write a shared string cell (BrtCellIsst - Cell + u32 = 12 bytes)
    fn write_shared_string_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: &str,
        style: u32,
        shared_strings: &mut crate::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        // Add string to shared strings table and get index
        let string_index = shared_strings.add_string(value.to_string());

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + isst index (4 bytes) = 12 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u32(string_index)?;

        writer.write_record(record_types::CELL_ISST, &data)?;
        Ok(())
    }

    /// Write a number cell (BrtCellReal - Cell + f64 = 16 bytes)
    fn write_number_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: f64,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + Xnum value (8 bytes) = 16 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_f64(value)?;

        writer.write_record(record_types::CELL_REAL, &data)?;
        Ok(())
    }

    /// Write a boolean cell (BrtCellBool - Cell + u8 = 9 bytes)
    fn write_bool_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: bool,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + fBool (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(if value { 1 } else { 0 })?;

        writer.write_record(record_types::CELL_BOOL, &data)?;
        Ok(())
    }

    /// Write an error cell (BrtCellError - Cell + u8 = 9 bytes)
    fn write_error_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        error: &str,
        style: u32,
    ) -> XlsbResult<()> {
        let error_code = Self::error_code(error);

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + bError (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(error_code)?;

        writer.write_record(record_types::CELL_ERROR, &data)?;
        Ok(())
    }

    fn error_code(error: &str) -> u8 {
        match error {
            "#NULL!" => 0x00,
            "#DIV/0!" => 0x07,
            "#VALUE!" => 0x0F,
            "#REF!" => 0x17,
            "#NAME?" => 0x1D,
            "#NUM!" => 0x24,
            "#N/A" => 0x2A,
            "#GETTING_DATA" => 0x2B,
            _ => 0x2A, // Default to #N/A
        }
    }

    /// Write merged cells
    fn write_merged_cells<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // BrtBeginMergeCells (0x00B1) payload is a single DWORD count of BrtMergeCell
        // records that follow. SheetJS writes this as write_BrtBeginMergeCells(cnt).
        let mut header = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut header);
        temp_writer.write_u32(self.merged_cells.len() as u32)?;

        writer.write_record(record_types::BEGIN_MERGE_CELLS, &header)?;

        for merged in &self.merged_cells {
            let data = merged.serialize();
            writer.write_record(record_types::MERGE_CELL, &data)?;
        }

        writer.write_record(record_types::END_MERGE_CELLS, &[])?;
        Ok(())
    }

    /// Write column information records.
    fn write_col_infos<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        if self.columns.is_empty() {
            return Ok(());
        }

        writer.write_record(record_types::BEGIN_COL_INFOS, &[])?;

        for (col, info) in &self.columns {
            let mut data = Vec::new();
            let mut temp_writer = RecordWriter::new(&mut data);

            // firstCol / lastCol (both 0-based inclusive).
            temp_writer.write_u32(*col)?;
            temp_writer.write_u32(*col)?;

            // Width is stored as 256ths of a character, mirroring SheetJS
            // write_BrtColInfo and [MS-XLSB] 2.4.323.
            let width_chars = info.width.unwrap_or(10.0);
            let width_raw = (width_chars * 256.0).round() as u32;
            temp_writer.write_u32(width_raw)?;

            // Style XF index (we currently do not support per-column styles).
            temp_writer.write_u32(0)?;

            // Flags (2 bytes): 0x0001 = hidden, 0x0002 = custom width,
            // 0x0004 = best fit.
            let mut flags: u16 = 0;
            if info.hidden {
                flags |= 0x0001;
            }
            if info.width.is_some() {
                flags |= 0x0002;
            }
            if info.best_fit {
                flags |= 0x0004;
            }
            temp_writer.write_u16(flags)?;

            writer.write_record(record_types::COL_INFO, &data)?;
        }

        writer.write_record(record_types::END_COL_INFOS, &[])?;
        Ok(())
    }

    /// Write hyperlinks
    fn write_hyperlinks<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        for hyperlink in &self.hyperlinks {
            let data = hyperlink.serialize();
            writer.write_record(record_types::H_LINK, &data)?;
        }
        Ok(())
    }

    /// Write sheet protection if configured.
    fn write_sheet_protection<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let Some(ref prot) = self.sheet_protection else {
            return Ok(());
        };

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Password hash (Method 1). When absent, write 0.
        temp_writer.write_u16(prot.password_hash.unwrap_or(0))?;

        // Guard DWORD: this record should not be written if no protection.
        temp_writer.write_u32(1)?;

        fn flag(default_true: bool, value: Option<bool>) -> u32 {
            if default_true {
                if let Some(v) = value {
                    if !v { 1 } else { 0 }
                } else {
                    0
                }
            } else if let Some(v) = value {
                if v { 0 } else { 1 }
            } else {
                1
            }
        }

        temp_writer.write_u32(flag(false, prot.objects))?;
        temp_writer.write_u32(flag(false, prot.scenarios))?;
        temp_writer.write_u32(flag(true, prot.format_cells))?;
        temp_writer.write_u32(flag(true, prot.format_columns))?;
        temp_writer.write_u32(flag(true, prot.format_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_columns))?;
        temp_writer.write_u32(flag(true, prot.insert_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_hyperlinks))?;
        temp_writer.write_u32(flag(true, prot.delete_columns))?;
        temp_writer.write_u32(flag(true, prot.delete_rows))?;
        temp_writer.write_u32(flag(false, prot.select_locked_cells))?;
        temp_writer.write_u32(flag(true, prot.sort))?;
        temp_writer.write_u32(flag(true, prot.auto_filter))?;
        temp_writer.write_u32(flag(true, prot.pivot_tables))?;
        temp_writer.write_u32(flag(false, prot.select_unlocked_cells))?;

        writer.write_record(record_types::SHEET_PROTECTION, &data)?;
        Ok(())
    }

    /// Write basic auto-filter range if configured.
    fn write_auto_filter<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let Some(ref af) = self.auto_filter else {
            return Ok(());
        };
        if af.row_first > af.row_last
            || af.row_last >= 0x10_0000
            || af.col_first > af.col_last
            || af.col_last >= 0x4000
        {
            return Err(XlsbError::Encoding(format!(
                "invalid AutoFilter range: rows {}..={}, columns {}..={}",
                af.row_first, af.row_last, af.col_first, af.col_last
            )));
        }

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // UncheckedRfX: row_first, row_last, col_first, col_last
        temp_writer.write_u32(af.row_first)?;
        temp_writer.write_u32(af.row_last)?;
        temp_writer.write_u32(af.col_first)?;
        temp_writer.write_u32(af.col_last)?;

        writer.write_record(record_types::BEGIN_A_FILTER, &data)?;
        writer.write_record(record_types::END_A_FILTER, &[])?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::comments::Comment;
    use crate::xlsb::conditional_formatting::ConditionalFormatting;
    use crate::xlsb::data_validation::DataValidation;
    use crate::xlsb::hyperlinks::Hyperlink;
    use crate::xlsb::merged_cells::MergedCell;
    use crate::xlsb::records::XlsbRecordIter;
    use crate::xlsb::web_extension_bindings::XlsbWebExtensionBinding;
    use litchi_core::binary;
    use std::io::Cursor;

    #[test]
    fn test_set_and_get_cell() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Hello");
        sheet.set_cell(1, 1, 42.0);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Hello"));
        assert_eq!(sheet.get_cell(1, 1).and_then(|v| v.as_float()), Some(42.0));
    }

    #[test]
    fn writes_worksheet_web_extension_collection() {
        let formula = CellParsedFormula {
            rgce: vec![0x3B, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 1, 0],
            rgcb: Vec::new(),
        };
        let binding =
            XlsbWebExtensionBinding::new("sales-table", formula, |index| index == 0).unwrap();
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet
            .set_web_extension_bindings(vec![binding.clone()])
            .unwrap();
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet.write(&mut writer, &mut shared_strings).unwrap();

        let records = XlsbRecordIter::new(Cursor::new(buffer))
            .collect::<XlsbResult<Vec<_>>>()
            .unwrap();
        let begin = records
            .iter()
            .position(|record| record.header.record_type == record_types::BEGIN_WEB_EXTENSIONS)
            .unwrap();
        assert!(records[begin].data.is_empty());
        assert_eq!(
            records[begin + 1].header.record_type,
            record_types::WEB_EXTENSION
        );
        assert_eq!(
            XlsbWebExtensionBinding::parse_payload(&records[begin + 1].data, |index| index == 0)
                .unwrap(),
            binding
        );
        assert_eq!(
            records[begin + 2].header.record_type,
            record_types::END_WEB_EXTENSIONS
        );
        assert!(records[begin + 2].data.is_empty());
    }

    #[test]
    fn test_set_cell_with_style() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell_with_style(0, 0, "Styled", 5);

        assert_eq!(
            sheet.get_cell(0, 0).and_then(|v| v.as_str()),
            Some("Styled")
        );
        assert_eq!(sheet.cell_count(), 1);
    }

    #[test]
    fn test_delete_cell() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Hello");

        assert!(sheet.delete_cell(0, 0).is_some());
        assert!(sheet.get_cell(0, 0).is_none());
    }

    #[test]
    fn test_delete_row() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Row 0");
        sheet.set_cell(1, 0, "Row 1");
        sheet.set_cell(2, 0, "Row 2");

        sheet.delete_row(1);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Row 0"));
        assert_eq!(sheet.get_cell(1, 0).and_then(|v| v.as_str()), Some("Row 2"));
        assert!(sheet.get_cell(2, 0).is_none());
    }

    #[test]
    fn test_delete_column() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Col 0");
        sheet.set_cell(0, 1, "Col 1");
        sheet.set_cell(0, 2, "Col 2");

        sheet.delete_column(1);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Col 0"));
        assert_eq!(sheet.get_cell(0, 1).and_then(|v| v.as_str()), Some("Col 2"));
        assert!(sheet.get_cell(0, 2).is_none());
    }

    #[test]
    fn test_insert_row() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Row 0");
        sheet.set_cell(1, 0, "Row 1");

        sheet.insert_row(1);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Row 0"));
        assert!(sheet.get_cell(1, 0).is_none()); // Inserted row is empty
        assert_eq!(sheet.get_cell(2, 0).and_then(|v| v.as_str()), Some("Row 1"));
    }

    #[test]
    fn test_insert_column() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Col 0");
        sheet.set_cell(0, 1, "Col 1");

        sheet.insert_column(1);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Col 0"));
        assert!(sheet.get_cell(0, 1).is_none()); // Inserted column is empty
        assert_eq!(sheet.get_cell(0, 2).and_then(|v| v.as_str()), Some("Col 1"));
    }

    #[test]
    fn test_dimensions() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        assert!(sheet.dimensions().is_none());

        sheet.set_cell(5, 10, "Test");
        assert_eq!(sheet.dimensions(), Some((0, 0, 5, 10)));
    }

    #[test]
    fn test_cell_count() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        assert_eq!(sheet.cell_count(), 0);

        sheet.set_cell(0, 0, "A");
        sheet.set_cell(0, 1, "B");
        sheet.set_cell(1, 0, "C");

        assert_eq!(sheet.cell_count(), 3);
    }

    #[test]
    fn test_name() {
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        assert_eq!(sheet.name(), "Sheet1");
    }

    #[test]
    fn test_set_name() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_name("RenamedSheet");
        assert_eq!(sheet.name(), "RenamedSheet");
    }

    #[test]
    fn test_set_column_width() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_column_width(0, 15.5);
        sheet.set_column_width(2, 20.0);

        // Verify columns are set
        assert_eq!(sheet.columns.len(), 2);
    }

    #[test]
    fn test_set_row_height() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_row_height(0, 25.0);
        sheet.set_row_height(3, 30.5);

        // Verify rows are set
        assert_eq!(sheet.rows.len(), 2);
    }

    #[test]
    fn test_clear() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Test");
        sheet.add_merged_cell(MergedCell::new(0, 1, 0, 1));

        sheet.clear();

        assert_eq!(sheet.cell_count(), 0);
        assert!(sheet.merged_cells.is_empty());
        assert!(sheet.dimensions().is_none());
    }

    #[test]
    fn test_add_merged_cell() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let merged = MergedCell::new(0, 1, 0, 1);
        sheet.add_merged_cell(merged);

        assert_eq!(sheet.merged_cells().len(), 1);
        assert_eq!(sheet.merged_cells()[0].row_first, 0);
        assert_eq!(sheet.merged_cells()[0].row_last, 1);
    }

    #[test]
    fn test_add_hyperlink() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string());
        sheet.add_hyperlink(link);

        assert_eq!(sheet.hyperlinks().len(), 1);
    }

    #[test]
    fn test_hyperlinks_mut() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string());
        sheet.add_hyperlink(link);

        let links = sheet.hyperlinks_mut();
        assert_eq!(links.len(), 1);
    }

    #[test]
    fn test_add_comment() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let comment = Comment::new(0, 0, "Author".to_string(), "Comment text".to_string());
        sheet.add_comment(comment);

        assert_eq!(sheet.comments().len(), 1);
    }

    #[test]
    fn test_set_auto_filter() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_auto_filter(0, 10, 0, 5);

        assert!(sheet.auto_filter.is_some());
        let af = sheet.auto_filter.unwrap();
        assert_eq!(af.row_first, 0);
        assert_eq!(af.row_last, 10);
        assert_eq!(af.col_first, 0);
        assert_eq!(af.col_last, 5);
    }

    #[test]
    fn test_add_data_validation() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let dv = DataValidation::new(3, "A1:A10".to_string());
        sheet.add_data_validation(dv);

        assert_eq!(sheet.data_validations().len(), 1);
        assert_eq!(sheet.data_validations()[0].validation_type, 3);
    }

    #[test]
    fn test_add_conditional_formatting() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        let cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        sheet.add_conditional_formatting(cf);

        assert_eq!(sheet.conditional_formattings().len(), 1);
    }

    #[test]
    fn test_cell_data_types() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");

        // String
        sheet.set_cell(0, 0, "String");
        assert_eq!(
            sheet.get_cell(0, 0).and_then(|v| v.as_str()),
            Some("String")
        );

        // Integer - stored as CellValue::Int
        sheet.set_cell(0, 1, 42i32);
        match sheet.get_cell(0, 1) {
            Some(CellValue::Int(i)) => assert_eq!(*i, 42),
            _ => panic!("Expected Int(42)"),
        }

        // Float
        sheet.set_cell(0, 2, 1.5f64);
        assert_eq!(sheet.get_cell(0, 2).and_then(|v| v.as_float()), Some(1.5));

        // Bool - check by matching the enum variant directly
        sheet.set_cell(0, 3, true);
        match sheet.get_cell(0, 3) {
            Some(CellValue::Bool(b)) => assert!(*b),
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[test]
    fn test_worksheet_write_empty() {
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();

        let result = sheet.write(&mut writer, &mut shared_strings);
        assert!(result.is_ok());
        assert!(!buffer.is_empty());
    }

    #[test]
    fn test_worksheet_write_with_data() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Hello");
        sheet.set_cell(0, 1, 42.0);
        sheet.set_cell(1, 0, true);

        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();

        let result = sheet.write(&mut writer, &mut shared_strings);
        assert!(result.is_ok());
        assert!(!buffer.is_empty());

        // Verify shared strings were added
        assert_eq!(shared_strings.len(), 1); // "Hello"
    }

    #[test]
    fn test_column_info_struct() {
        let info = ColumnInfo {
            width: Some(15.0),
            hidden: false,
            best_fit: true,
        };
        assert_eq!(info.width, Some(15.0));
        assert!(!info.hidden);
        assert!(info.best_fit);
    }

    #[test]
    fn writes_best_fit_in_the_specified_column_flag() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.columns.insert(
            2,
            ColumnInfo {
                width: Some(15.0),
                hidden: false,
                best_fit: true,
            },
        );
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet.write(&mut writer, &mut shared_strings).unwrap();

        let record = crate::xlsb::records::XlsbRecordIter::new(buffer.as_slice())
            .find_map(|record| {
                let record = record.unwrap();
                (record.header.record_type == record_types::COL_INFO).then_some(record)
            })
            .unwrap();
        assert_eq!(binary::read_u16_le_at(&record.data, 16).unwrap(), 0x0006);
    }

    #[test]
    fn test_row_info_struct() {
        let info = RowInfo {
            height: Some(20.0),
            hidden: true,
        };
        assert_eq!(info.height, Some(20.0));
        assert!(info.hidden);
    }

    #[test]
    fn test_auto_filter_struct() {
        let af = AutoFilter {
            row_first: 0,
            row_last: 10,
            col_first: 0,
            col_last: 5,
        };
        assert_eq!(af.row_first, 0);
        assert_eq!(af.row_last, 10);
        assert_eq!(af.col_first, 0);
        assert_eq!(af.col_last, 5);
    }

    #[test]
    fn test_cell_data_struct() {
        let cell = CellData {
            value: CellValue::String("Test".to_string()),
            style: 5,
            formula_binary: None,
            formula_flags: 0,
        };
        assert_eq!(cell.style, 5);
        assert_eq!(cell.value.as_str(), Some("Test"));
    }

    #[test]
    fn writes_ms_xlsb_brt_fmla_num_layout_without_downgrading_formula() {
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        let cell = CellData {
            value: CellValue::Formula {
                formula: "C13*2".to_string(),
                cached_value: Some(Box::new(CellValue::Float(4.0))),
                is_array: false,
                array_range: None,
            },
            style: 0,
            formula_binary: None,
            formula_flags: 0,
        };
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet
            .write_cell(&mut writer, 12, 1, &cell, &mut shared_strings)
            .unwrap();

        let mut expected = vec![0x09, 0x25]; // BrtFmlaNum, 37-byte payload
        expected.extend_from_slice(&1_u32.to_le_bytes()); // Cell.column
        expected.extend_from_slice(&[0; 4]); // style and phonetic flags
        expected.extend_from_slice(&4_f64.to_le_bytes()); // cached xnum
        expected.extend_from_slice(&0_u16.to_le_bytes()); // GrbitFmla
        expected.extend_from_slice(&11_u32.to_le_bytes()); // cce
        expected.extend_from_slice(&[
            0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
        ]);
        expected.extend_from_slice(&0_u32.to_le_bytes()); // cb
        assert_eq!(buffer, expected);
    }

    #[test]
    fn unsupported_formula_is_an_error_instead_of_a_cached_constant() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(
            0,
            0,
            CellValue::Formula {
                formula: "UNSUPPORTED(A1)".to_string(),
                cached_value: Some(Box::new(CellValue::Float(42.0))),
                is_array: false,
                array_range: None,
            },
        );
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        let error = sheet.write(&mut writer, &mut shared_strings).unwrap_err();
        assert!(matches!(
            error,
            crate::xlsb::error::XlsbError::UnsupportedFeature(_)
        ));
    }

    #[test]
    fn formula_without_cached_result_is_marked_for_recalculation() {
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        let cell = CellData {
            value: CellValue::Formula {
                formula: "1+1".to_string(),
                cached_value: None,
                is_array: false,
                array_range: None,
            },
            style: 0,
            formula_binary: None,
            formula_flags: 0,
        };
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet
            .write_cell(&mut writer, 0, 0, &cell, &mut shared_strings)
            .unwrap();

        // Two-byte record header, then Cell (8) + cached xnum (8).
        assert_eq!(u16::from_le_bytes([buffer[18], buffer[19]]), 0x0002);
    }

    #[test]
    fn writes_shared_definition_immediately_after_anchor_and_exp_followers() {
        use crate::xlsb::records::RecordIter;
        use std::io::Cursor;

        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(2, 2, 10.0);
        sheet.set_cell(3, 2, 20.0);
        sheet.set_shared_formula(2, 2, 3, 2, "B3").unwrap();

        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet.write_cells(&mut writer, &mut shared_strings).unwrap();

        let mut iter = RecordIter::new(Cursor::new(buffer));
        let mut records = Vec::new();
        while let Ok(record_type) = iter.read_type() {
            let mut data = Vec::new();
            iter.fill_buffer(&mut data).unwrap();
            records.push((record_type, data));
        }
        assert_eq!(
            records.iter().map(|record| record.0).collect::<Vec<_>>(),
            vec![
                record_types::ROW_HDR,
                record_types::FMLA_NUM,
                record_types::SHR_FMLA,
                record_types::ROW_HDR,
                record_types::FMLA_NUM,
            ]
        );

        let group = FormulaGroup::parse_shared(&records[2].1).unwrap();
        assert_eq!(group.range.to_a1(), "C3:C4");
        for formula_record in [&records[1].1, &records[4].1] {
            let (placeholder, consumed) = CellParsedFormula::parse(&formula_record[18..]).unwrap();
            assert_eq!(18 + consumed, formula_record.len());
            assert_eq!(placeholder.exp_cell().unwrap(), Some((2, 2)));
        }
    }

    #[test]
    fn writes_unsupported_group_definition_losslessly() {
        use crate::xlsb::records::RecordIter;
        use std::io::Cursor;

        let group = FormulaGroup {
            kind: FormulaGroupKind::Array,
            range: FormulaRange::new(8, 8, 2, 2).unwrap(),
            formula: CellParsedFormula {
                rgce: vec![0x23, 0x02, 0x00, 0x00, 0x00, 0x42, 0x01, 0xFF, 0x00],
                rgcb: Vec::new(),
            },
            always_calculate: true,
        };
        let expected = group.to_record_data().unwrap();
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_formula_group_binary(group).unwrap();

        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let mut shared_strings = crate::xlsb::writer::MutableSharedStringsWriter::new();
        sheet.write_cells(&mut writer, &mut shared_strings).unwrap();

        let mut iter = RecordIter::new(Cursor::new(buffer));
        let mut definition = None;
        while let Ok(record_type) = iter.read_type() {
            let mut data = Vec::new();
            iter.fill_buffer(&mut data).unwrap();
            if record_type == record_types::ARR_FMLA {
                definition = Some(data);
            }
        }
        assert_eq!(definition.as_deref(), Some(expected.as_slice()));
    }
}
