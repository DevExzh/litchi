#![allow(
    clippy::expect_used,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! XLSB worksheet wire state and contextual formula restoration.

use crate::conditional_formatting::{Rule, Value};
use crate::package::error::{Error, Result};
use crate::package::formula::{CompilationContext, GroupKind, ParsedFormula, Range};
use litchi_core::sheet::CellValue;

use super::semantic::MutableWorksheet;

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
pub(crate) struct ColumnInfo {
    /// Column width in character units. `None` uses the sheet default.
    pub width: Option<f64>,
    /// Whether the column is hidden.
    pub hidden: bool,
    /// Whether the column width was inferred via best-fit.
    pub best_fit: bool,
}

/// Row information for a single 0-based row.
#[derive(Debug, Clone)]
pub(crate) struct RowInfo {
    /// Row height in points. `None` uses the sheet default.
    pub height: Option<f64>,
    /// Whether the row is hidden.
    pub hidden: bool,
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

pub(crate) fn formula_requires_workbook_context(error: &Error) -> bool {
    matches!(
        error,
        Error::UnsupportedFeature(message)
            if message.ends_with("requires workbook compilation context")
    )
}

impl MutableWorksheet {
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
}
