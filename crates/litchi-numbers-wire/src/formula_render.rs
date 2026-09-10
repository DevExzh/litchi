//! Generated-free Numbers formula event rendering shared by format adapters.
//!
//! The wire codec owns event decoding while this module owns the bounded
//! semantic projection of those events into formula text. Format adapters
//! supply only their error budget and borrowed table/category/function names.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The renderer keeps its generic contract next to its event and formatting helpers."
)]

use std::fmt::Write as _;
use std::result::Result;

use litchi_iwa_common::formula::render::{FormulaExpr, FormulaRenderBudget, FormulaRenderer};
use litchi_iwa_protos::numbers_formula_codec;

type RenderError<B> = <B as FormulaRenderBudget>::Error;
type RenderResult<T, B> = Result<T, RenderError<B>>;

/// Borrowed semantic table prefix used for a cross-table formula reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaTablePrefix<'a> {
    /// Sheet name selected by the format adapter.
    pub sheet: &'a str,
    /// Table name selected by the format adapter.
    pub table: &'a str,
}

/// Category identity used by a category-reference formula node.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaCategoryId {
    /// Lower 64 bits of the native category UUID.
    pub lower: u64,
    /// Upper 64 bits of the native category UUID.
    pub upper: u64,
}

/// Resolves native formula references to borrowed semantic names.
///
/// The resolver remains an adapter concern: this low-level wire module never
/// stores or exposes a native object map. Missing or partial native identities
/// intentionally resolve to the renderer's existing fallback text.
pub trait ReferenceResolver {
    /// Resolve a complete native table UUID to a borrowed sheet/table prefix.
    fn table_prefix(
        &self,
        id: &numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<FormulaTablePrefix<'_>>;

    /// Resolve a complete native table UUID to a table-only display name.
    ///
    /// Pages and Keynote can address a table without a Numbers sheet name.
    /// Adapters that do not have this form of reference keep the historical
    /// behavior by inheriting the empty default.
    fn table_only_name(&self, _id: &numbers_formula_codec::FormulaRenderCfuuid) -> Option<&str> {
        None
    }

    /// Resolve a category UUID to its borrowed display name.
    fn category_name(&self, id: FormulaCategoryId) -> Option<&str>;

    /// Resolve a built-in function identifier to its borrowed name.
    fn function_name(&self, index: u32) -> Option<&str>;
}

/// Error and resource adapter for shared formula event rendering.
pub trait FormulaEventRenderBudget: FormulaRenderBudget {
    /// Admit one semantic formula-array depth.
    fn check_render_depth(&self, depth: usize) -> Result<(), Self::Error>;

    /// Construct a format-owned parse error with a dynamic message.
    fn parse_error(&self, message: String) -> Self::Error;

    /// Construct a format-owned invalid-format error.
    fn invalid_format(&self, message: String) -> Self::Error;
}

/// Render one scalar generated-free formula node sequence.
///
/// The scalar decoder remains in each format adapter because it owns archive
/// admission and aggregate report charging. This shared function owns all
/// scalar-to-text semantics after that decoder has produced its borrowed node
/// stream.
pub fn render_scalar_formula_nodes<R, B>(
    nodes: &[numbers_formula_codec::FormulaNode],
    resolver: &R,
    budget: &mut B,
) -> Result<String, B::Error>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    budget.check_structure(nodes.len(), nodes.len())?;
    if nodes.is_empty() {
        return retain_text("=", budget);
    }
    let mut renderer = FormulaRenderer::default();
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(nodes.len())
        .map_err(|_| budget.allocation("Numbers scalar formula expression stack", nodes.len()))?;
    for node in nodes {
        use numbers_formula_codec::{BinaryOperator, FormulaNode};
        let expression = match *node {
            FormulaNode::Binary(operator) => {
                let (symbol, operation) = match operator {
                    BinaryOperator::Add => ("+", "addition"),
                    BinaryOperator::Subtract => ("-", "subtraction"),
                    BinaryOperator::Multiply => ("*", "multiplication"),
                    BinaryOperator::Divide => ("/", "division"),
                    BinaryOperator::Power => ("^", "power"),
                    BinaryOperator::Concatenate => ("&", "concatenation"),
                    BinaryOperator::GreaterThan => (">", "greater than"),
                    BinaryOperator::GreaterThanOrEqual => (">=", "greater than or equal"),
                    BinaryOperator::LessThan => ("<", "less than"),
                    BinaryOperator::LessThanOrEqual => ("<=", "less than or equal"),
                    BinaryOperator::Equal => ("=", "equality"),
                    BinaryOperator::NotEqual => ("<>", "inequality"),
                };
                Some(render_binary(
                    &mut stack,
                    &mut renderer,
                    symbol,
                    operation,
                    true,
                    budget,
                )?)
            },
            FormulaNode::Negation => stack
                .pop()
                .map(|operand| renderer.unary("-(", operand, ")", budget))
                .transpose()?,
            FormulaNode::Percent => {
                let operand = stack.pop().ok_or_else(|| {
                    budget.parse_error(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(renderer.unary("(", operand, ")%", budget)?)
            },
            FormulaNode::Function {
                identifier,
                argument_count,
            } => {
                let arguments =
                    pop_formula_arguments(&mut stack, argument_count, "function", budget)?;
                Some(renderer.comma_joined(
                    Some(fallible_function_name(
                        identifier, resolver, &renderer, budget,
                    )?),
                    arguments,
                    "(",
                    ")",
                    budget,
                )?)
            },
            FormulaNode::Number { bits } => {
                let value = fallible_formula_display(f64::from_bits(bits), &renderer, budget)?;
                Some(renderer.owned_expr(value, budget)?)
            },
            FormulaNode::Boolean(value) | FormulaNode::Token(value) => {
                Some(renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget)?)
            },
            FormulaNode::Empty => Some(renderer.static_expr("", budget)?),
            FormulaNode::LocalCell {
                coordinate,
                row_is_sticky,
                column_is_sticky,
            } => {
                let column = FormulaColumn(coordinate.column());
                let row = checked_formula_row_number(coordinate.row(), budget)?;
                let value = fallible_formula_format(&renderer, budget, |output| {
                    write!(
                        output,
                        "{}{column}{}{row}",
                        if column_is_sticky != 0 { "$" } else { "" },
                        if row_is_sticky != 0 { "$" } else { "" },
                    )
                })?;
                Some(renderer.owned_expr(value, budget)?)
            },
            FormulaNode::Colon | FormulaNode::ColonWithUids => Some(render_binary(
                &mut stack,
                &mut renderer,
                ":",
                "range",
                false,
                budget,
            )?),
            FormulaNode::PlusSign
            | FormulaNode::AppendWhitespace
            | FormulaNode::PrependWhitespace => None,
            FormulaNode::LocalCellReference { .. }
            | FormulaNode::LocalRange { .. }
            | FormulaNode::CellReference { .. }
            | FormulaNode::ResolvedCellReference { .. }
            | FormulaNode::ResolvedRange { .. } => {
                return Err(budget.invalid_format(
                    "Numbers scalar formula visitor received an owner-bearing node".to_owned(),
                ));
            },
        };
        if let Some(expression) = expression {
            stack.push(expression);
        }
    }
    let Some(root) = stack.pop() else {
        // Keep parity with the historical compatibility renderer: an archive
        // containing only ignored postfix markers (or a missing negation
        // operand) falls through to its FORMULA() placeholder.
        return retain_text("=FORMULA()", budget);
    };
    renderer.render(root, budget)
}

struct CompatibilityFormulaArray {
    expressions: Vec<FormulaExpr>,
    had_node: bool,
    is_thunk: bool,
}

/// Generated-free semantic visitor for already-decoded formula events.
///
/// The caller supplies a borrowed [`ReferenceResolver`] and its format-owned
/// [`FormulaEventRenderBudget`]. Feed events in codec order through
/// [`Self::visit_event`]. The stream must be one codec-order sequence with a
/// single root expression; this visitor projects that sequence and does not
/// validate the complete event grammar. Stop and discard the visitor after
/// the first error, and call [`Self::finish`] only after the stream has closed
/// every array.
pub struct CompatibilityFormulaVisitor<'references, 'budget, R, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    host_row: u32,
    host_column: u32,
    resolver: &'references R,
    budget: &'budget mut B,
    renderer: FormulaRenderer,
    arrays: Vec<CompatibilityFormulaArray>,
    root_expression: Option<FormulaExpr>,
    pending_thunk: bool,
}

impl<'references, 'budget, R, B> CompatibilityFormulaVisitor<'references, 'budget, R, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    /// Create a visitor for one decoded compatibility formula.
    pub fn new(
        host_row: u32,
        host_column: u32,
        resolver: &'references R,
        budget: &'budget mut B,
    ) -> Self {
        Self {
            host_row,
            host_column,
            resolver,
            budget,
            renderer: FormulaRenderer::default(),
            arrays: Vec::new(),
            root_expression: None,
            pending_thunk: false,
        }
    }

    fn current_array_mut(&mut self) -> RenderResult<&mut CompatibilityFormulaArray, B> {
        self.arrays.last_mut().ok_or_else(|| {
            self.budget
                .parse_error("Numbers formula event stream has no active array".to_owned())
        })
    }

    fn mark_node(&mut self) -> RenderResult<(), B> {
        self.current_array_mut()?.had_node = true;
        Ok(())
    }

    fn push_expression(&mut self, expression: FormulaExpr) -> RenderResult<(), B> {
        let budget = &*self.budget;
        let Some(array) = self.arrays.last_mut() else {
            return Err(
                budget.parse_error("Numbers formula event stream has no active array".to_owned())
            );
        };
        let current_len = array.expressions.len();
        if array.expressions.try_reserve(1).is_err() {
            return Err(budget.allocation(
                "Numbers formula compatibility expression stack",
                current_len.saturating_add(1),
            ));
        }
        array.expressions.push(expression);
        Ok(())
    }

    fn pop_binary(&mut self, operation: &str) -> RenderResult<(FormulaExpr, FormulaExpr), B> {
        let budget = &*self.budget;
        let Some(array) = self.arrays.last_mut() else {
            return Err(
                budget.parse_error("Numbers formula event stream has no active array".to_owned())
            );
        };
        pop_binary_operands(&mut array.expressions, operation, budget)
    }

    fn pop_arguments(&mut self, count: u32, node_kind: &str) -> RenderResult<Vec<FormulaExpr>, B> {
        let budget = &*self.budget;
        let Some(array) = self.arrays.last_mut() else {
            return Err(
                budget.parse_error("Numbers formula event stream has no active array".to_owned())
            );
        };
        pop_formula_arguments(&mut array.expressions, count, node_kind, budget)
    }

    fn begin_array(&mut self, depth: u32) -> RenderResult<(), B> {
        // The codec reports logical AST-array depth (root = 1, each thunk
        // adds one). Check that depth directly so the package limit remains
        // identical to the legacy renderer's semantic recursion bound.
        let semantic_depth = usize::try_from(depth).unwrap_or(usize::MAX);
        let expected_depth = self.arrays.len().saturating_add(1);
        if semantic_depth != expected_depth {
            return Err(self.budget.parse_error(format!(
                "Numbers formula event stream reported depth {semantic_depth}, expected {expected_depth}"
            )));
        }
        self.budget.check_render_depth(semantic_depth)?;
        let is_thunk = self.pending_thunk;
        self.pending_thunk = false;
        self.arrays.try_reserve(1).map_err(|_error| {
            self.budget.allocation(
                "Numbers formula compatibility arrays",
                self.arrays.len().saturating_add(1),
            )
        })?;
        self.arrays.push(CompatibilityFormulaArray {
            expressions: Vec::new(),
            had_node: false,
            is_thunk,
        });
        Ok(())
    }

    fn end_array(&mut self) -> RenderResult<(), B> {
        let array = self.arrays.pop().ok_or_else(|| {
            self.budget
                .parse_error("Numbers formula event stream ended an inactive array".to_owned())
        })?;
        let expression = if let Some(expression) = array.expressions.last().copied() {
            expression
        } else if array.is_thunk {
            self.renderer
                .static_expr(if array.had_node { "FORMULA()" } else { "" }, self.budget)?
        } else if array.had_node {
            self.renderer.static_expr("FORMULA()", self.budget)?
        } else {
            // The root empty archive is handled by `finish`; retaining no
            // expression here preserves the legacy decoder's `=` result.
            return Ok(());
        };
        if self.arrays.is_empty() {
            self.root_expression = Some(expression);
        } else if array.is_thunk {
            self.push_expression(expression)?;
        } else {
            return Err(self.budget.parse_error(
                "Numbers formula event stream nested an unexpected array".to_owned(),
            ));
        }
        Ok(())
    }

    fn render_event(
        &mut self,
        event: numbers_formula_codec::FormulaRenderEvent<'_>,
    ) -> RenderResult<Option<FormulaExpr>, B> {
        use numbers_formula_codec::{BinaryOperator, FormulaRenderEvent};
        let expression = match event {
            FormulaRenderEvent::Binary(operator) => {
                let (symbol, operation) = match operator {
                    BinaryOperator::Add => ("+", "addition"),
                    BinaryOperator::Subtract => ("-", "subtraction"),
                    BinaryOperator::Multiply => ("*", "multiplication"),
                    BinaryOperator::Divide => ("/", "division"),
                    BinaryOperator::Power => ("^", "power"),
                    BinaryOperator::Concatenate => ("&", "concatenation"),
                    BinaryOperator::GreaterThan => (">", "greater than"),
                    BinaryOperator::GreaterThanOrEqual => (">=", "greater than or equal"),
                    BinaryOperator::LessThan => ("<", "less than"),
                    BinaryOperator::LessThanOrEqual => ("<=", "less than or equal"),
                    BinaryOperator::Equal => ("=", "equality"),
                    BinaryOperator::NotEqual => ("<>", "inequality"),
                };
                let (left, right) = self.pop_binary(operation)?;
                Some(
                    self.renderer
                        .binary(left, symbol, right, true, self.budget)?,
                )
            },
            FormulaRenderEvent::Negation => self
                .current_array_mut()?
                .expressions
                .pop()
                .map(|operand| self.renderer.unary("-(", operand, ")", self.budget))
                .transpose()?,
            FormulaRenderEvent::Percent => {
                let operand = self.current_array_mut()?.expressions.pop().ok_or_else(|| {
                    self.budget.parse_error(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(self.renderer.unary("(", operand, ")%", self.budget)?)
            },
            FormulaRenderEvent::Number { value } => Some(self.renderer.owned_expr(
                fallible_formula_display(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::String(value) => Some(self.renderer.owned_expr(
                formula_string_literal(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::Boolean(value) | FormulaRenderEvent::Token(value) => Some(
                self.renderer
                    .static_expr(if value { "TRUE" } else { "FALSE" }, self.budget)?,
            ),
            FormulaRenderEvent::Date { value } => {
                let days = value / 86_400.0;
                Some(self.renderer.owned_expr(
                    fallible_formula_format(&self.renderer, self.budget, |output| {
                        write!(output, "(DATE(2001,1,1)+{days})")
                    })?,
                    self.budget,
                )?)
            },
            FormulaRenderEvent::Duration { value } => Some(self.renderer.owned_expr(
                fallible_formula_display(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::EmptyArgument => Some(self.renderer.static_expr("", self.budget)?),
            FormulaRenderEvent::Function {
                identifier,
                argument_count,
            } => {
                let arguments = self.pop_arguments(argument_count, "function")?;
                Some(self.renderer.comma_joined(
                    Some(fallible_function_name(
                        identifier,
                        self.resolver,
                        &self.renderer,
                        self.budget,
                    )?),
                    arguments,
                    "(",
                    ")",
                    self.budget,
                )?)
            },
            FormulaRenderEvent::List { argument_count } => {
                let arguments = self.pop_arguments(argument_count, "list")?;
                Some(
                    self.renderer
                        .comma_joined(None, arguments, "", "", self.budget)?,
                )
            },
            FormulaRenderEvent::Array { columns, rows } => {
                let count = columns.checked_mul(rows).ok_or_else(|| {
                    self.budget
                        .parse_error("Numbers formula array size overflow".to_owned())
                })?;
                let values = self.pop_arguments(count, "array")?;
                let columns = usize::try_from(columns).map_err(|_| {
                    self.budget
                        .parse_error("Numbers formula array width exceeds usize".to_owned())
                })?;
                Some(self.renderer.array(values, columns, self.budget)?)
            },
            FormulaRenderEvent::UnknownFunction {
                name,
                argument_count,
            } => {
                let arguments = self.pop_arguments(argument_count, "unknown function")?;
                Some(self.renderer.comma_joined(
                    Some(fallible_formula_owned(
                        name.unwrap_or("UNKNOWN"),
                        &self.renderer,
                        self.budget,
                    )?),
                    arguments,
                    "(",
                    ")",
                    self.budget,
                )?)
            },
            FormulaRenderEvent::CellReference(reference) => {
                Some(self.render_cell_reference(&reference)?)
            },
            FormulaRenderEvent::LocalCellReference(reference) => {
                Some(self.render_standalone_local_cell_reference(reference)?)
            },
            FormulaRenderEvent::CrossTableCellReference(reference) => {
                Some(self.render_cross_table_cell_reference(reference)?)
            },
            FormulaRenderEvent::Colon => Some(self.render_binary("colon", ":", false)?),
            FormulaRenderEvent::ColonWithUids => Some(self.render_binary("range", ":", false)?),
            FormulaRenderEvent::ColonTract(tract) => Some(self.render_colon_tract(&tract)?),
            FormulaRenderEvent::CategoryReference(category) => {
                Some(self.render_category_reference(category)?)
            },
            FormulaRenderEvent::ReferenceError => {
                Some(self.renderer.static_expr("#REF!", self.budget)?)
            },
            FormulaRenderEvent::Ignored { .. }
            | FormulaRenderEvent::PlusSign
            | FormulaRenderEvent::AppendWhitespace
            | FormulaRenderEvent::PrependWhitespace
            | FormulaRenderEvent::BeginArray { .. }
            | FormulaRenderEvent::EndArray
            | FormulaRenderEvent::ThunkBegin
            | FormulaRenderEvent::ThunkEnd => None,
        };
        Ok(expression)
    }

    fn render_binary(
        &mut self,
        operation: &str,
        operator: &'static str,
        wrapped: bool,
    ) -> RenderResult<FormulaExpr, B> {
        let (left, right) = self.pop_binary(operation)?;
        self.renderer
            .binary(left, operator, right, wrapped, self.budget)
    }

    fn render_cell_reference(
        &mut self,
        reference: &numbers_formula_codec::FormulaRenderCellReference,
    ) -> RenderResult<FormulaExpr, B> {
        if let Some(coordinates) = reference.coordinates {
            let column_absolute = coordinates.column.absolute;
            let row_absolute = coordinates.row.absolute;
            let column = FormulaColumn(resolve_formula_coordinate(
                self.host_column as usize,
                coordinates.column.coordinate,
                column_absolute,
                "column",
                self.budget,
            )?);
            let row = checked_formula_row_number(
                resolve_formula_coordinate(
                    self.host_row as usize,
                    coordinates.row.coordinate,
                    row_absolute,
                    "row",
                    self.budget,
                )?,
                self.budget,
            )?;
            let prefix = reference
                .cross_table_extra
                .as_ref()
                .and_then(|extra| formula_render_prefix_parts(&extra.table_id, self.resolver));
            return self.renderer.owned_expr(
                fallible_formula_format(&self.renderer, self.budget, |output| {
                    if reference.cross_table_extra.is_some() {
                        write_formula_reference_prefix(output, prefix)?;
                    }
                    write!(
                        output,
                        "{}{column}{}{row}",
                        if column_absolute { "$" } else { "" },
                        if row_absolute { "$" } else { "" },
                    )
                })?,
                self.budget,
            );
        }
        if let Some(local) = reference.local {
            return self.render_local_cell_reference(Some(local));
        }
        if let Some(cross) = reference.cross_table {
            return self.render_cross_table_cell_reference(Some(cross));
        }
        self.renderer.owned_expr(
            fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
            self.budget,
        )
    }

    fn render_local_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderLocalCellReference>,
    ) -> RenderResult<FormulaExpr, B> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle, self.budget)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write!(
                    output,
                    "{}{column}{}{row}",
                    if reference.column_is_sticky != 0 {
                        "$"
                    } else {
                        ""
                    },
                    if reference.row_is_sticky != 0 {
                        "$"
                    } else {
                        ""
                    },
                )
            })?,
            self.budget,
        )
    }

    /// A standalone LocalCellReferenceNode is rendered by the legacy
    /// generated path without consulting its sticky flags.  Keep that quirk
    /// distinct from CellReferenceNode's nested-local fallback, which does
    /// preserve the flags.
    fn render_standalone_local_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderLocalCellReference>,
    ) -> RenderResult<FormulaExpr, B> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle, self.budget)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write!(output, "{column}{row}")
            })?,
            self.budget,
        )
    }

    fn render_cross_table_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderCrossTableCellReference>,
    ) -> RenderResult<FormulaExpr, B> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let prefix = formula_render_prefix_parts(&reference.table_id, self.resolver);
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle, self.budget)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write_formula_reference_prefix(output, prefix)?;
                write!(output, "{column}{row}")
            })?,
            self.budget,
        )
    }

    fn render_colon_tract(
        &mut self,
        tract: &numbers_formula_codec::FormulaRenderColonTract,
    ) -> RenderResult<FormulaExpr, B> {
        let whole_rows = tract.relative_column.count == 0
            && tract.absolute_column.count == 1
            && tract.absolute_column.first_begin == Some(i16::MAX as i64)
            && tract.absolute_column.first_end.is_none();
        let whole_columns = tract.relative_row.count == 0
            && tract.absolute_row.count == 1
            && tract.absolute_row.first_begin == Some(i32::MAX as i64)
            && tract.absolute_row.first_end.is_none();
        let has_columns =
            !whole_rows && (tract.relative_column.count != 0 || tract.absolute_column.count != 0);
        let has_rows =
            !whole_columns && (tract.relative_row.count != 0 || tract.absolute_row.count != 0);
        if !has_columns && !has_rows {
            return Err(self.budget.parse_error(
                "Numbers formula colon tract has no row or column coordinates".to_owned(),
            ));
        }
        let columns = if has_columns {
            Some((
                resolve_render_colon_axis(
                    &tract.relative_column,
                    &tract.absolute_column,
                    tract.sticky.begin_column_is_absolute,
                    false,
                    self.host_column as usize,
                    "column",
                    self.budget,
                )?,
                resolve_render_colon_axis(
                    &tract.relative_column,
                    &tract.absolute_column,
                    tract.sticky.end_column_is_absolute,
                    true,
                    self.host_column as usize,
                    "column",
                    self.budget,
                )?,
            ))
        } else {
            None
        };
        let rows = if has_rows {
            Some((
                resolve_render_colon_axis(
                    &tract.relative_row,
                    &tract.absolute_row,
                    tract.sticky.begin_row_is_absolute,
                    false,
                    self.host_row as usize,
                    "row",
                    self.budget,
                )?,
                resolve_render_colon_axis(
                    &tract.relative_row,
                    &tract.absolute_row,
                    tract.sticky.end_row_is_absolute,
                    true,
                    self.host_row as usize,
                    "row",
                    self.budget,
                )?,
            ))
        } else {
            None
        };
        let rows = rows
            .map(|(begin, end)| {
                Ok::<(u64, u64), RenderError<B>>((
                    checked_formula_row_number(begin, self.budget)?,
                    checked_formula_row_number(end, self.budget)?,
                ))
            })
            .transpose()?;
        let prefix = tract
            .cross_table_extra
            .as_ref()
            .and_then(|extra| formula_render_prefix_parts(&extra.table_id, self.resolver));
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                if tract.cross_table_extra.is_some() {
                    write_formula_reference_prefix(output, prefix)?;
                }
                if let Some((begin, end)) = columns {
                    if tract.sticky.begin_column_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{}", FormulaColumn(begin))?;
                    if let Some((row_begin, _)) = rows {
                        if tract.sticky.begin_row_is_absolute {
                            output.write_char('$')?;
                        }
                        write!(output, "{row_begin}")?;
                    }
                    output.write_char(':')?;
                    if tract.sticky.end_column_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{}", FormulaColumn(end))?;
                    if let Some((_, row_end)) = rows {
                        if tract.sticky.end_row_is_absolute {
                            output.write_char('$')?;
                        }
                        write!(output, "{row_end}")?;
                    }
                } else if let Some((begin, end)) = rows {
                    if tract.sticky.begin_row_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{begin}")?;
                    output.write_char(':')?;
                    if tract.sticky.end_row_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{end}")?;
                }
                Ok(())
            })?,
            self.budget,
        )
    }

    fn render_category_reference(
        &mut self,
        category: Option<numbers_formula_codec::FormulaRenderCategoryReference>,
    ) -> RenderResult<FormulaExpr, B> {
        let category_uid = category.and_then(|category| {
            category
                .absolute_group_uid
                .or(category.relative_group_uid)
                .or(category.last_group_uid)
        });
        let Some(category_uid) = category_uid else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#CATEGORY!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let Some(label) = self.resolver.category_name(FormulaCategoryId {
            lower: category_uid.lower,
            upper: category_uid.upper,
        }) else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#CATEGORY!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        self.renderer.owned_expr(
            render_category_label_checked(label, &self.renderer, self.budget)?,
            self.budget,
        )
    }

    /// Finish the semantic projection after the codec has emitted all events.
    pub fn finish(self) -> RenderResult<String, B> {
        let this = self;
        if !this.arrays.is_empty() {
            return Err(this
                .budget
                .parse_error("Numbers formula event stream left an active array".to_owned()));
        }
        match this.root_expression {
            Some(expression) => this.renderer.render(expression, this.budget),
            None => retain_text("=", this.budget),
        }
    }

    /// Feed one already-decoded formula event to the semantic renderer.
    pub fn visit_event(
        &mut self,
        event: numbers_formula_codec::FormulaRenderEvent<'_>,
    ) -> RenderResult<(), B> {
        use numbers_formula_codec::FormulaRenderEvent;
        let result = match event {
            FormulaRenderEvent::BeginArray { depth } => self.begin_array(depth).map(|_| None),
            FormulaRenderEvent::EndArray => self.end_array().map(|_| None),
            FormulaRenderEvent::ThunkBegin => self.mark_node().map(|()| {
                self.pending_thunk = true;
                None
            }),
            FormulaRenderEvent::ThunkEnd => Ok(None),
            event => self.mark_node().and_then(|()| self.render_event(event)),
        };
        match result? {
            Some(expression) => self.push_expression(expression),
            None => Ok(()),
        }
    }

    /// Access the adapter budget for aggregate report charging.
    pub fn budget_mut(&mut self) -> &mut B {
        self.budget
    }
}

/// Thin Buffa visitor adapter that translates shared semantic failures back
/// into a decoder stop while retaining the typed error for the format adapter.
pub struct FormulaRenderCodecVisitor<'references, 'budget, R, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    visitor: CompatibilityFormulaVisitor<'references, 'budget, R, B>,
    error: Option<RenderError<B>>,
}

impl<'references, 'budget, R, B> FormulaRenderCodecVisitor<'references, 'budget, R, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    /// Create the codec-facing visitor for one formula archive.
    pub fn new(
        host_row: u32,
        host_column: u32,
        resolver: &'references R,
        budget: &'budget mut B,
    ) -> Self {
        Self {
            visitor: CompatibilityFormulaVisitor::new(host_row, host_column, resolver, budget),
            error: None,
        }
    }

    /// Access the adapter budget for report charging after decoding.
    pub fn budget_mut(&mut self) -> &mut B {
        self.visitor.budget_mut()
    }

    /// Take the first semantic error captured while stopping the codec.
    pub fn take_error(&mut self) -> Option<RenderError<B>> {
        self.error.take()
    }

    /// Finish the semantic projection.
    pub fn finish(self) -> RenderResult<String, B> {
        let Self { visitor, error } = self;
        if let Some(error) = error {
            Err(error)
        } else {
            visitor.finish()
        }
    }
}

impl<R, B> numbers_formula_codec::FormulaRenderVisitor for FormulaRenderCodecVisitor<'_, '_, R, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    fn visit(
        &mut self,
        event: numbers_formula_codec::FormulaRenderEvent<'_>,
    ) -> std::result::Result<(), numbers_formula_codec::DecodeError> {
        if self.error.is_some() {
            return Err(numbers_formula_codec::DecodeError::allocation(0));
        }
        self.visitor.visit_event(event).map_err(|error| {
            self.error = Some(error);
            numbers_formula_codec::DecodeError::allocation(0)
        })
    }
}

fn retain_text<B>(value: &str, budget: &mut B) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    budget.charge(value.len())?;
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_| budget.allocation("Numbers rendered formula", value.len()))?;
    retained.push_str(value);
    Ok(retained)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ResolvedTablePrefix<'a> {
    SheetTable(FormulaTablePrefix<'a>),
    TableOnly(&'a str),
}

fn formula_render_prefix_parts<'a, R>(
    owner: &numbers_formula_codec::FormulaRenderCfuuid,
    resolver: &'a R,
) -> Option<ResolvedTablePrefix<'a>>
where
    R: ReferenceResolver,
{
    resolver
        .table_only_name(owner)
        .map(ResolvedTablePrefix::TableOnly)
        .or_else(|| {
            resolver
                .table_prefix(owner)
                .map(ResolvedTablePrefix::SheetTable)
        })
}

fn render_category_label_checked<B>(
    label: &str,
    renderer: &FormulaRenderer,
    budget: &B,
) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    let escaped_extra = label
        .bytes()
        .filter(|byte| *byte == b'\\' || *byte == b']')
        .count();
    let required = "#CATEGORY!["
        .len()
        .checked_add(label.len())
        .and_then(|length| length.checked_add(escaped_extra))
        .and_then(|length| length.checked_add(1))
        .ok_or_else(|| budget.output_limit(usize::MAX))?;
    renderer.check_additional_owned(required, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(required)
        .map_err(|_| budget.allocation("Numbers formula category text", required))?;
    output.push_str("#CATEGORY![");
    for character in label.chars() {
        if character == '\\' || character == ']' {
            output.push('\\');
        }
        output.push(character);
    }
    output.push(']');
    Ok(output)
}

fn resolve_render_colon_axis<B>(
    relative: &numbers_formula_codec::FormulaRenderRangeSummary,
    absolute: &numbers_formula_codec::FormulaRenderRangeSummary,
    is_absolute: bool,
    is_end: bool,
    host: usize,
    axis: &str,
    budget: &B,
) -> RenderResult<u32, B>
where
    B: FormulaEventRenderBudget,
{
    let summary = if is_absolute { absolute } else { relative };
    let stored = if is_end {
        summary.first_end.or(summary.first_begin)
    } else {
        summary.first_begin
    }
    .ok_or_else(|| {
        budget.parse_error(format!(
            "Numbers formula colon tract has no {} {} coordinate",
            if is_absolute { "absolute" } else { "relative" },
            axis
        ))
    })?;
    if is_absolute {
        u32::try_from(stored).map_err(|_| {
            budget.parse_error(format!(
                "Numbers formula colon tract absolute {axis} coordinate is out of range"
            ))
        })
    } else {
        let stored = i32::try_from(stored).map_err(|_| {
            budget.parse_error(format!(
                "Numbers formula colon tract relative {axis} coordinate is out of range"
            ))
        })?;
        resolve_formula_coordinate(host, stored, false, axis, budget)
    }
}

fn fallible_formula_owned<B>(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &B,
) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    renderer.check_additional_owned(value.len(), budget)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_| budget.allocation("Numbers formula owned text", value.len()))?;
    owned.push_str(value);
    Ok(owned)
}

fn fallible_formula_display<B>(
    value: impl std::fmt::Display,
    renderer: &FormulaRenderer,
    budget: &B,
) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    #[derive(Default)]
    struct Counter {
        bytes: usize,
    }
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.bytes = self.bytes.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write!(&mut counter, "{value}").map_err(|_error| budget.output_limit(usize::MAX))?;
    renderer.check_additional_owned(counter.bytes, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.bytes)
        .map_err(|_| budget.allocation("Numbers formula owned text", counter.bytes))?;
    write!(&mut output, "{value}")
        .map_err(|_error| budget.invalid_format("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

fn fallible_formula_format<B>(
    renderer: &FormulaRenderer,
    budget: &B,
    write_value: impl Fn(&mut dyn std::fmt::Write) -> std::fmt::Result,
) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    #[derive(Default)]
    struct Counter(usize);
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.0 = self.0.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write_value(&mut counter).map_err(|_error| budget.output_limit(usize::MAX))?;
    renderer.check_additional_owned(counter.0, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.0)
        .map_err(|_| budget.allocation("Numbers formula owned text", counter.0))?;
    write_value(&mut output)
        .map_err(|_error| budget.invalid_format("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

#[derive(Clone, Copy)]
struct FormulaColumn(u32);

impl std::fmt::Display for FormulaColumn {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut bytes = [0_u8; 7];
        let mut cursor = bytes.len();
        let mut value = self.0;
        loop {
            cursor -= 1;
            bytes[cursor] = b'A' + u8::try_from(value % 26).map_err(|_error| std::fmt::Error)?;
            if value < 26 {
                break;
            }
            value = value / 26 - 1;
        }
        formatter
            .write_str(std::str::from_utf8(&bytes[cursor..]).map_err(|_error| std::fmt::Error)?)
    }
}

fn fallible_function_name<R, B>(
    index: u32,
    resolver: &R,
    renderer: &FormulaRenderer,
    budget: &B,
) -> RenderResult<String, B>
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    if let Some(name) = resolver.function_name(index) {
        fallible_formula_owned(name, renderer, budget)
    } else {
        fallible_formula_format(renderer, budget, |output| write!(output, "FUNC{index}"))
    }
}

fn write_formula_reference_prefix(
    output: &mut dyn std::fmt::Write,
    prefix: Option<ResolvedTablePrefix<'_>>,
) -> std::fmt::Result {
    match prefix {
        Some(ResolvedTablePrefix::SheetTable(name)) => {
            write!(output, "{}::{}::", name.sheet, name.table)
        },
        Some(ResolvedTablePrefix::TableOnly(name)) => write!(output, "{name}::"),
        None => output.write_str("Table::"),
    }
}

fn render_binary<B>(
    stack: &mut Vec<FormulaExpr>,
    renderer: &mut FormulaRenderer,
    operator: &'static str,
    operation: &str,
    wrapped: bool,
    budget: &B,
) -> RenderResult<FormulaExpr, B>
where
    B: FormulaEventRenderBudget,
{
    let (left, right) = pop_binary_operands(stack, operation, budget)?;
    renderer.binary(left, operator, right, wrapped, budget)
}

fn formula_string_literal<B>(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &B,
) -> RenderResult<String, B>
where
    B: FormulaEventRenderBudget,
{
    let quote_count = value.bytes().filter(|byte| *byte == b'"').count();
    let length = value
        .len()
        .checked_add(quote_count)
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| budget.allocation("Numbers formula string literal", usize::MAX))?;
    renderer.check_additional_owned(length, budget)?;
    let mut literal = String::new();
    literal
        .try_reserve_exact(length)
        .map_err(|_| budget.allocation("Numbers formula string literal", length))?;
    literal.push('"');
    for character in value.chars() {
        if character == '"' {
            literal.push('"');
        }
        literal.push(character);
    }
    literal.push('"');
    Ok(literal)
}

fn resolve_formula_coordinate<B>(
    host: usize,
    stored: i32,
    absolute: bool,
    axis: &str,
    budget: &B,
) -> RenderResult<u32, B>
where
    B: FormulaEventRenderBudget,
{
    let coordinate = if absolute {
        i64::from(stored)
    } else {
        i64::try_from(host)
            .map_err(|_error| {
                budget.parse_error(format!("Numbers formula host {axis} exceeds i64"))
            })?
            .checked_add(i64::from(stored))
            .ok_or_else(|| budget.parse_error(format!("Numbers formula {axis} overflow")))?
    };
    u32::try_from(coordinate).map_err(|_error| {
        budget.parse_error(format!(
            "Numbers formula {axis} coordinate {coordinate} is out of range"
        ))
    })
}

fn checked_formula_row_number<B>(row: u32, budget: &B) -> RenderResult<u64, B>
where
    B: FormulaEventRenderBudget,
{
    u64::from(row)
        .checked_add(1)
        .ok_or_else(|| budget.parse_error("Numbers formula row coordinate overflow".to_owned()))
}

fn pop_binary_operands<T, B>(
    stack: &mut Vec<T>,
    operation: &str,
    budget: &B,
) -> RenderResult<(T, T), B>
where
    B: FormulaEventRenderBudget,
{
    let right = stack.pop().ok_or_else(|| {
        budget.parse_error(format!(
            "Malformed Numbers formula: {operation} is missing its right operand"
        ))
    })?;
    let left = stack.pop().ok_or_else(|| {
        budget.parse_error(format!(
            "Malformed Numbers formula: {operation} is missing its left operand"
        ))
    })?;
    Ok((left, right))
}

fn pop_formula_arguments<T, B>(
    stack: &mut Vec<T>,
    count: u32,
    node_kind: &str,
    budget: &B,
) -> RenderResult<Vec<T>, B>
where
    B: FormulaEventRenderBudget,
{
    let count = usize::try_from(count).map_err(|_| {
        budget.parse_error(format!(
            "Numbers formula {node_kind} argument count exceeds usize"
        ))
    })?;
    let start = stack.len().checked_sub(count).ok_or_else(|| {
        budget.parse_error(format!(
            "Malformed Numbers formula: {node_kind} requires {count} arguments but only {} are available",
            stack.len()
        ))
    })?;
    let mut arguments = Vec::new();
    arguments
        .try_reserve_exact(count)
        .map_err(|_| budget.allocation("Numbers formula arguments", count))?;
    arguments.extend(stack.drain(start..));
    Ok(arguments)
}
