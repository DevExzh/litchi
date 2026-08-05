//! Mutable XLSB worksheet state and CRUD model.

use crate::conditional_formatting::{Formatting, Rule, Value};
use crate::package::comments::Comment;
use crate::package::data_validation::{Settings, Validation};
use crate::package::error::{Error, Result};
use crate::package::formula::{
    CompilationContext, Compiler, Group, GroupKind, ParsedFormula, Parser, Range,
};
use crate::package::hyperlinks::Hyperlink;
use crate::package::merged_cells::MergedCell;
use crate::package::sheet_view::SheetView;
use crate::package::web_extension_bindings::Binding;
use litchi_core::sheet::CellValue;
use std::collections::{BTreeMap, HashSet};

/// Cell data for storage
#[derive(Debug, Clone)]
pub struct CellData {
    pub value: CellValue,
    pub style: u32, // Style XF index
    /// Optional pre-encoded cell formula for lossless XLSB workflows.
    pub formula_binary: Option<ParsedFormula>,
    /// `GrbitFmla` flags; only bit 1 (`fAlwaysCalc`) is defined.
    pub formula_flags: u16,
}

/// Column information for a single 0-based column.
///
/// This writer-side structure drives `BrtColInfo` emission and mirrors the
/// semantics of [MS-XLSB] 2.4.323 and SheetJS' `write_BrtColInfo` helper.
#[derive(Debug, Clone)]
pub(super) struct ColumnInfo {
    /// Column width in character units. `None` uses the sheet default.
    pub width: Option<f64>,
    /// Whether the column is hidden.
    pub hidden: bool,
    /// Whether the column width was inferred via best-fit.
    pub best_fit: bool,
}

/// Row information for a single 0-based row.
#[derive(Debug, Clone)]
pub(super) struct RowInfo {
    /// Row height in points. `None` uses the sheet default.
    pub height: Option<f64>,
    /// Whether the row is hidden.
    pub hidden: bool,
}

/// Freeze panes configuration.
///
/// Freezes rows and columns in place while scrolling.
#[derive(Debug, Clone)]
pub(super) struct FreezePanes {
    /// Number of rows to freeze from the top.
    pub(super) freeze_rows: u32,
    /// Number of columns to freeze from the left.
    pub(super) freeze_cols: u32,
}

/// Auto-filter configuration for a rectangular range.
///
/// The indices are 0-based and inclusive.
#[derive(Debug, Clone)]
pub(super) struct AutoFilter {
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
pub struct MutableWorksheet {
    pub(super) name: String,
    pub(super) cells: BTreeMap<(u32, u32), CellData>,
    pub(super) max_row: u32,
    pub(super) max_col: u32,
    pub(super) merged_cells: Vec<MergedCell>,
    pub(super) hyperlinks: Vec<Hyperlink>,
    pub(super) comments: Vec<Comment>,
    /// Column information (0-based column index).
    pub(super) columns: BTreeMap<u32, ColumnInfo>,
    /// Row information (0-based row index).
    pub(super) rows: BTreeMap<u32, RowInfo>,
    /// Optional auto-filter configuration.
    pub(super) auto_filter: Option<AutoFilter>,
    /// Optional sheet protection configuration.
    pub(super) sheet_protection: Option<SheetProtection>,
    /// Optional worksheet view configuration (zoom, pane, selections).
    pub(super) sheet_view: Option<SheetView>,
    /// Optional freeze panes configuration.
    pub(super) freeze_panes: Option<FreezePanes>,
    /// Data validation rules.
    pub(super) data_validations: Vec<Validation>,
    pub(super) data_validation_settings: Settings,
    pub(super) data_validation14_settings: Settings,
    /// Conditional formatting rules.
    pub(super) conditional_formattings: Vec<Formatting>,
    /// Inert Office Add-in range bindings.
    pub(super) web_extension_bindings: Vec<Binding>,
    /// Array and shared formula definitions. Cell records contain only a
    /// `PtgExp` reference to one of these definitions.
    pub(super) formula_groups: Vec<Group>,
    /// Original text for formula groups created through the text API. Binary
    /// groups intentionally have no entry and are never recompiled.
    pub(super) formula_group_sources: BTreeMap<(u32, u32), String>,
    /// Structured tables (ListObjects) hosted on this sheet.
    pub(super) tables: Vec<crate::package::table::Table>,
    /// Losslessly preserved PivotTable definition parts hosted on this sheet.
    pub(super) pivot_table_views: Vec<crate::pivot_view::Part>,
    /// Typed DrawingML charts anchored on this sheet.
    pub(super) charts: Vec<crate::package::xlsx::Chart>,
    /// Typed image parts anchored in the same Drawings part.
    pub(super) images: Vec<crate::writer::Image>,
    /// Cached sum of encoded image bytes for constant-time safety checks.
    pub(super) image_bytes: usize,
    /// Top-level DrawingML shapes and text boxes.
    pub(super) shapes: Vec<crate::package::xlsx::writer::ShapeSpec>,
    /// Top-level DrawingML shape groups.
    pub(super) groups: Vec<crate::package::xlsx::writer::GroupSpec>,
    /// Top-level DrawingML connection shapes.
    pub(super) connections: Vec<crate::package::xlsx::writer::ConnectionShapeSpec>,
    /// Relationship ID allocated for the sheet's Drawings part.
    pub(super) drawing_rel_id: Option<String>,
    /// Relationship IDs allocated for `tables` by the workbook writer, in
    /// table order. Populated during `WorkbookWriter::save`.
    pub(crate) table_rel_ids: Vec<String>,
}

pub(crate) struct ContextualFormulaRestore {
    cell_positions: Vec<(u32, u32)>,
    group_formulas: Vec<(usize, ParsedFormula)>,
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
    rule: &mut Rule,
    location: ConditionalValueLocation,
) -> Option<&mut Value> {
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

fn formula_requires_workbook_context(error: &Error) -> bool {
    matches!(
        error,
        Error::UnsupportedFeature(message)
            if message.ends_with("requires workbook compilation context")
    )
}

impl MutableWorksheet {
    /// Create a new empty worksheet
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xlsb::writer::MutableWorksheet;
    ///
    /// let sheet = MutableWorksheet::new("Sheet1");
    /// ```
    pub fn new<S: Into<String>>(name: S) -> Self {
        MutableWorksheet {
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
            sheet_view: None,
            freeze_panes: None,
            data_validations: Vec::new(),
            data_validation_settings: Settings::default(),
            data_validation14_settings: Settings::default(),
            conditional_formattings: Vec::new(),
            web_extension_bindings: Vec::new(),
            formula_groups: Vec::new(),
            formula_group_sources: BTreeMap::new(),
            tables: Vec::new(),
            pivot_table_views: Vec::new(),
            charts: Vec::new(),
            images: Vec::new(),
            image_bytes: 0,
            shapes: Vec::new(),
            groups: Vec::new(),
            connections: Vec::new(),
            drawing_rel_id: None,
            table_rel_ids: Vec::new(),
        }
    }

    pub(crate) fn compile_contextual_formulas(
        &mut self,
        context: &CompilationContext<'_>,
    ) -> Result<ContextualFormulaRestore> {
        let mut compiled_groups = Vec::new();
        for (index, group) in self.formula_groups.iter().enumerate() {
            let Some(source) = self.formula_group_sources.get(&group.range.top_left()) else {
                continue;
            };
            let formula = match group.kind {
                GroupKind::Array => {
                    crate::package::formula::text::Compiler::compile_with_context(source, context)?
                },
                GroupKind::Shared => {
                    crate::package::formula::text::Compiler::compile_shared_with_context(
                        source,
                        group.range.row_first,
                        group.range.col_first,
                        context,
                    )?
                },
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
                    .and_then(|range| Range::parse_a1(range).ok())
                    .is_some_and(|range| range.top_left() == position);
            if cell.formula_binary.is_none() && (!is_array || is_array_anchor) && !is_grouped {
                compiled.push((
                    position,
                    crate::package::formula::text::Compiler::compile_with_context(
                        formula, context,
                    )?,
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
                    .map(|formula| {
                        crate::package::formula::text::Compiler::compile_with_context(
                            formula, context,
                        )
                    })
                    .transpose()?
            } else {
                None
            };
            let formula2 = if validation.formula2_binary.is_none() {
                validation
                    .formula2
                    .as_deref()
                    .filter(|formula| !formula.is_empty())
                    .map(|formula| {
                        crate::package::formula::text::Compiler::compile_with_context(
                            formula, context,
                        )
                    })
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
                        .map(|formula| {
                            crate::package::formula::text::Compiler::compile_with_context(
                                formula, context,
                            )
                        })
                        .collect::<Result<Vec<_>>>()?;
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
                            crate::package::formula::text::Compiler::compile_with_context(
                                source, context,
                            )?,
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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

    /// Set a formula using an already encoded `ParsedFormula`.
    ///
    /// This is the lossless path for formulas containing tokens that the text
    /// compiler does not yet understand. `cached_value` determines which
    /// `BrtFmla*` record is emitted.
    pub fn set_cell_formula_binary(
        &mut self,
        row: u32,
        col: u32,
        cached_value: CellValue,
        formula: ParsedFormula,
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
    ) -> Result<()> {
        let range = Range::new(row_first, row_last, col_first, col_last)?;
        let definition = match crate::package::formula::text::Compiler::compile(formula) {
            Ok(definition) => definition,
            Err(error) if formula_requires_workbook_context(&error) => {
                crate::package::formula::text::Compiler::compile("0")?
            },
            Err(error) => return Err(error),
        };
        let group = Group {
            kind: GroupKind::Array,
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
    ) -> Result<()> {
        let range = Range::new(row_first, row_last, col_first, col_last)?;
        let definition = match crate::package::formula::text::Compiler::compile_shared(
            formula, row_first, col_first,
        ) {
            Ok(definition) => definition,
            Err(error) if formula_requires_workbook_context(&error) => {
                crate::package::formula::text::Compiler::compile_shared("0", row_first, col_first)?
            },
            Err(error) => return Err(error),
        };
        let group = Group {
            kind: GroupKind::Shared,
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
    pub fn set_formula_group_binary(&mut self, group: Group) -> Result<()> {
        // Validate both the range and parsed-formula framing before mutating
        // the worksheet.
        let _ = group.to_record_data()?;
        if group.formula.exp_cell()?.is_some() {
            return Err(Error::InvalidFormula(
                "array/shared formula definition cannot contain PtgExp".to_string(),
            ));
        }
        self.install_formula_group(group, None)
    }

    fn install_formula_group(&mut self, group: Group, anchor_formula: Option<&str>) -> Result<()> {
        if let Some(index) = self
            .formula_groups
            .iter()
            .position(|existing| existing.range.top_left() == group.range.top_left())
        {
            let replaced = self.formula_groups.remove(index);
            self.formula_group_sources
                .remove(&replaced.range.top_left());
            if replaced.kind == GroupKind::Array {
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
                let decoded = || -> Result<String> {
                    let tokens = match group.kind {
                        GroupKind::Array => {
                            Parser::with_extra(&group.formula.rgce, &group.formula.rgcb).parse()?
                        },
                        GroupKind::Shared => Parser::with_base_cell_and_extra(
                            &group.formula.rgce,
                            &group.formula.rgcb,
                            row,
                            col,
                        )
                        .parse()?,
                    };
                    Ok(Compiler::try_tokens_to_string(&tokens)?)
                };
                let formula = match (group.kind, anchor_formula) {
                    (GroupKind::Array, Some(formula)) => formula.to_string(),
                    (GroupKind::Shared, _) => decoded().or_else(|error| {
                        if anchor_formula.is_none() {
                            Ok(String::new())
                        } else {
                            Err(error)
                        }
                    })?,
                    (GroupKind::Array, None) => decoded().unwrap_or_default(),
                };
                self.cells.insert(
                    (row, col),
                    CellData {
                        value: CellValue::Formula {
                            formula,
                            cached_value,
                            is_array: group.kind == GroupKind::Array,
                            array_range: (group.kind == GroupKind::Array)
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
            .filter(|group| group.kind == GroupKind::Array && group.range.contains(row, col))
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
                && array_range
                    .as_ref()
                    .is_some_and(|range| ranges.contains(range))
            {
                *is_array = false;
                *array_range = None;
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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
        self.data_validation_settings = Settings::default();
        self.data_validation14_settings = Settings::default();
        self.conditional_formattings.clear();
        self.charts.clear();
        self.images.clear();
        self.image_bytes = 0;
        self.shapes.clear();
        self.groups.clear();
        self.connections.clear();
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

    /// Set the worksheet view (zoom scales, pane, selections, tab-selected flag).
    ///
    /// The view model is shared with XLSX worksheets; see
    /// [`crate::package::xlsx::views::SheetView`]. Pane and selection settings conflict
    /// with [`Self::freeze_panes`]; combining both fails at save time.
    pub fn set_sheet_view(&mut self, view: SheetView) {
        self.sheet_view = Some(view);
    }

    /// Worksheet view configured through [`Self::set_sheet_view`].
    pub fn sheet_view(&self) -> Option<&SheetView> {
        self.sheet_view.as_ref()
    }

    /// Freeze panes at the specified position.
    ///
    /// `freeze_rows` is the number of rows frozen from the top and
    /// `freeze_cols` the number of columns frozen from the left. The frozen
    /// pane uses the first scrolling cell as its selection anchor, mirroring
    /// the XLSX writer.
    pub fn freeze_panes(&mut self, freeze_rows: u32, freeze_cols: u32) {
        if freeze_rows > 0 || freeze_cols > 0 {
            self.freeze_panes = Some(FreezePanes {
                freeze_rows,
                freeze_cols,
            });
        }
    }

    /// Remove freeze panes.
    pub fn unfreeze_panes(&mut self) {
        self.freeze_panes = None;
    }

    /// Add a merged cell range
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xlsb::writer::MutableWorksheet;
    /// use litchi_xlsb::merged_cells::MergedCell;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    /// use litchi_xlsb::hyperlinks::Hyperlink;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    /// use litchi_xlsb::package::comments::Comment;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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
    pub fn add_table(&mut self, table: crate::package::table::Table) -> Result<()> {
        const MAX_TABLES_PER_SHEET: usize = 4_096;
        if self.tables.len() >= MAX_TABLES_PER_SHEET {
            return Err(Error::InvalidFormula(
                "worksheet table count exceeds the safety limit".to_string(),
            ));
        }
        if table.display_name.as_deref().is_none_or(str::is_empty) {
            return Err(Error::InvalidFormula(
                "structured table requires a display name".to_string(),
            ));
        }
        let range = &table.range;
        if range.first_row > range.last_row || range.first_column > range.last_column {
            return Err(Error::InvalidFormula(
                "structured table range is inverted".to_string(),
            ));
        }
        let width = u64::from(range.last_column) - u64::from(range.first_column) + 1;
        if !table.columns.is_empty() && table.columns.len() as u64 != width {
            return Err(Error::InvalidFormula(format!(
                "structured table declares {} columns for a range {width} wide",
                table.columns.len()
            )));
        }
        if self.tables.iter().any(|existing| existing.id == table.id) {
            return Err(Error::InvalidFormula(format!(
                "duplicate structured table id {}",
                table.id
            )));
        }
        self.tables.push(table);
        Ok(())
    }

    /// The structured tables hosted on this sheet.
    pub fn tables(&self) -> &[crate::package::table::Table] {
        &self.tables
    }

    /// Attach a losslessly preserved PivotTable definition part to this sheet.
    ///
    /// Its cache identifier is resolved against workbook PivotCaches at save
    /// time. Duplicate view names on one sheet are rejected immediately.
    pub fn add_pivot_table_view(&mut self, view: crate::pivot_view::Part) -> Result<()> {
        if self.pivot_table_views.len() >= 4_096 {
            return Err(Error::InvalidFormula(
                "worksheet PivotTable count exceeds the safety limit".to_string(),
            ));
        }
        if self
            .pivot_table_views
            .iter()
            .any(|existing| crate::package::formula::excel_name_eq(existing.name(), view.name()))
        {
            return Err(Error::InvalidFormula(format!(
                "duplicate PivotTable view name {:?} on worksheet {:?}",
                view.name(),
                self.name
            )));
        }
        self.pivot_table_views.push(view);
        Ok(())
    }

    /// Losslessly preserved PivotTable definition parts hosted on this sheet.
    pub fn pivot_table_views(&self) -> &[crate::pivot_view::Part] {
        &self.pivot_table_views
    }

    pub(crate) fn has_drawing_objects(&self) -> bool {
        !self.charts.is_empty()
            || !self.images.is_empty()
            || !self.shapes.is_empty()
            || !self.groups.is_empty()
            || !self.connections.is_empty()
    }

    fn clear_drawing_rel_id_if_empty(&mut self) {
        if !self.has_drawing_objects() {
            self.drawing_rel_id = None;
        }
    }

    /// Add a typed DrawingML chart anchored on this worksheet.
    ///
    /// XLSB uses the same chart and SpreadsheetDrawing XML as XLSX. The
    /// workbook writer emits the chart and drawing parts and stores their
    /// relationship in the binary `BrtDrawing` record. Pivot charts are
    /// resolved against PivotTable views attached to this workbook at save
    /// time.
    pub fn add_chart(&mut self, chart: crate::package::xlsx::Chart) -> Result<()> {
        if self.charts.len() >= crate::package::drawing_write::MAX_CHARTS_PER_SHEET {
            return Err(Error::InvalidFormula(
                "worksheet chart count exceeds the safety limit".to_string(),
            ));
        }
        crate::package::drawing_write::validate_chart(&chart)?;
        self.charts.push(chart);
        Ok(())
    }

    /// Typed DrawingML charts in drawing order.
    pub fn charts(&self) -> &[crate::package::xlsx::Chart] {
        &self.charts
    }

    /// Remove one chart by drawing order.
    pub fn remove_chart(&mut self, index: usize) -> Result<crate::package::xlsx::Chart> {
        if index >= self.charts.len() {
            return Err(Error::InvalidFormula(format!(
                "chart index {index} is out of bounds for {} charts",
                self.charts.len()
            )));
        }
        let removed = self.charts.remove(index);
        self.clear_drawing_rel_id_if_empty();
        Ok(removed)
    }

    /// Remove every authored chart from this worksheet.
    pub fn clear_charts(&mut self) {
        self.charts.clear();
        self.clear_drawing_rel_id_if_empty();
    }

    /// Add a typed embedded image to this worksheet's Drawings part.
    pub fn add_image(&mut self, image: crate::writer::Image) -> Result<()> {
        if self.images.len() >= crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGES {
            return Err(Error::InvalidFormula(
                "worksheet image count exceeds the safety limit".to_string(),
            ));
        }
        image.validate()?;
        let total_bytes =
            self.image_bytes
                .checked_add(image.data().len())
                .ok_or(Error::InvalidLength {
                    expected: crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                    found: usize::MAX,
                })?;
        if total_bytes > crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES {
            return Err(Error::InvalidLength {
                expected: crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                found: total_bytes,
            });
        }
        self.image_bytes = total_bytes;
        self.images.push(image);
        Ok(())
    }

    /// Typed embedded images in drawing order.
    pub fn images(&self) -> &[crate::writer::Image] {
        &self.images
    }

    /// Remove one embedded image by drawing order.
    pub fn remove_image(&mut self, index: usize) -> Result<crate::writer::Image> {
        if index >= self.images.len() {
            return Err(Error::InvalidFormula(format!(
                "image index {index} is out of bounds for {} images",
                self.images.len()
            )));
        }
        let removed = self.images.remove(index);
        self.image_bytes -= removed.data().len();
        self.clear_drawing_rel_id_if_empty();
        Ok(removed)
    }

    /// Remove every authored image from this worksheet.
    pub fn clear_images(&mut self) {
        self.images.clear();
        self.image_bytes = 0;
        self.clear_drawing_rel_id_if_empty();
    }

    fn drawing_shape_count(&self) -> usize {
        self.shapes.len() + self.groups.len() + self.connections.len()
    }

    /// Add a standard DrawingML shape or text box.
    pub fn add_shape(&mut self, shape: crate::package::xlsx::writer::ShapeSpec) -> Result<()> {
        shape
            .validate(self.drawing_shape_count())
            .map_err(Error::InvalidFormula)?;
        self.shapes.push(shape);
        Ok(())
    }

    /// Add a plain-text DrawingML text box.
    pub fn add_text_box(
        &mut self,
        name: impl Into<String>,
        anchor: crate::package::xlsx::ShapeAnchor,
        preset: crate::package::xlsx::Preset,
        text: &str,
    ) -> Result<()> {
        self.add_shape(crate::package::xlsx::writer::ShapeSpec::text_box(
            name, anchor, preset, text,
        ))
    }

    /// Authored top-level shapes in drawing order.
    pub fn shapes(&self) -> &[crate::package::xlsx::writer::ShapeSpec] {
        &self.shapes
    }

    /// Remove one top-level shape.
    pub fn remove_shape(
        &mut self,
        index: usize,
    ) -> Result<crate::package::xlsx::writer::ShapeSpec> {
        if index >= self.shapes.len() {
            return Err(Error::InvalidFormula(format!(
                "shape index {index} is out of bounds for {} shapes",
                self.shapes.len()
            )));
        }
        let removed = self.shapes.remove(index);
        self.clear_drawing_rel_id_if_empty();
        Ok(removed)
    }

    /// Add a nested DrawingML shape group.
    pub fn add_group(&mut self, group: crate::package::xlsx::writer::GroupSpec) -> Result<()> {
        group
            .validate(self.drawing_shape_count())
            .map_err(Error::InvalidFormula)?;
        self.groups.push(group);
        Ok(())
    }

    /// Authored top-level shape groups in drawing order.
    pub fn groups(&self) -> &[crate::package::xlsx::writer::GroupSpec] {
        &self.groups
    }

    /// Remove one top-level shape group.
    pub fn remove_group(
        &mut self,
        index: usize,
    ) -> Result<crate::package::xlsx::writer::GroupSpec> {
        if index >= self.groups.len() {
            return Err(Error::InvalidFormula(format!(
                "group index {index} is out of bounds for {} groups",
                self.groups.len()
            )));
        }
        let removed = self.groups.remove(index);
        self.clear_drawing_rel_id_if_empty();
        Ok(removed)
    }

    /// Add a DrawingML connection shape.
    pub fn add_connection(
        &mut self,
        connection: crate::package::xlsx::writer::ConnectionShapeSpec,
    ) -> Result<()> {
        connection
            .validate(self.drawing_shape_count())
            .map_err(Error::InvalidFormula)?;
        self.connections.push(connection);
        Ok(())
    }

    /// Authored top-level connection shapes in drawing order.
    pub fn connections(&self) -> &[crate::package::xlsx::writer::ConnectionShapeSpec] {
        &self.connections
    }

    /// Remove one top-level connection shape.
    pub fn remove_connection(
        &mut self,
        index: usize,
    ) -> Result<crate::package::xlsx::writer::ConnectionShapeSpec> {
        if index >= self.connections.len() {
            return Err(Error::InvalidFormula(format!(
                "connection index {index} is out of bounds for {} connection shapes",
                self.connections.len()
            )));
        }
        let removed = self.connections.remove(index);
        self.clear_drawing_rel_id_if_empty();
        Ok(removed)
    }

    /// Remove every authored shape, group, and connection shape.
    pub fn clear_drawing_shapes(&mut self) {
        self.shapes.clear();
        self.groups.clear();
        self.connections.clear();
        self.clear_drawing_rel_id_if_empty();
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    /// use litchi_xlsb::package::data_validation::Validation;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
    /// let mut dv = Validation::new(3, "A1:A10".to_string()); // list
    /// dv.formula1 = Some("Yes,No".to_string());
    /// sheet.add_data_validation(dv);
    /// ```
    pub fn add_data_validation(&mut self, dv: Validation) {
        self.data_validations.push(dv);
    }

    /// Get all data validations.
    pub fn data_validations(&self) -> &[Validation] {
        &self.data_validations
    }

    /// Set UI prompt settings for classic `BrtDVal` rules.
    pub fn set_data_validation_settings(&mut self, settings: Settings) {
        self.data_validation_settings = settings;
    }

    /// Set UI prompt settings for Office 2013 `BrtDVal14` rules.
    pub fn set_data_validation14_settings(&mut self, settings: Settings) {
        self.data_validation14_settings = settings;
    }

    /// Add a conditional formatting block.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xlsb::writer::MutableWorksheet;
    /// use crate::conditional_formatting::{
    ///     Formatting, Rule, RuleType,
    /// };
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
    /// let mut cf = Formatting::new(vec!["A1:A10".to_string()]);
    /// let rule = Rule::new(RuleType::CellIs, 1);
    /// cf.add_rule(rule);
    /// sheet.add_conditional_formatting(cf);
    /// ```
    pub fn add_conditional_formatting(&mut self, cf: Formatting) {
        self.conditional_formattings.push(cf);
    }

    /// Get all conditional formatting blocks.
    pub fn conditional_formattings(&self) -> &[Formatting] {
        &self.conditional_formattings
    }

    /// Replace worksheet Office Add-in bindings after validating their payloads
    /// and unique application references.
    pub fn set_web_extension_bindings(&mut self, bindings: Vec<Binding>) -> Result<()> {
        let mut app_refs = HashSet::with_capacity(bindings.len());
        for binding in &bindings {
            binding.to_payload()?;
            if !app_refs.insert(binding.application_reference.as_str()) {
                return Err(Error::Unrecognized {
                    typ: "WEBEXTENSIONS".to_string(),
                    val: "duplicate binding appRef".to_string(),
                });
            }
        }
        if bindings.len() > 65_536 {
            return Err(Error::Unrecognized {
                typ: "WEBEXTENSIONS".to_string(),
                val: "binding count exceeds 65,536".to_string(),
            });
        }
        self.web_extension_bindings = bindings;
        Ok(())
    }

    /// Office Add-in bindings that will be written to this worksheet.
    pub fn web_extension_bindings(&self) -> &[Binding] {
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
    /// use litchi_xlsb::writer::MutableWorksheet;
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
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
}
