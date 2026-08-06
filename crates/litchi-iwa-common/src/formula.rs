//! Dependency-free iWork formula vocabulary.
//!
//! Archive decoding, protobuf compilation, and calculation-engine mutation
//! remain owned by the concrete format adapters. This module owns the
//! semantic values that callers use to describe a formula, so Pages, Keynote,
//! and Numbers can share one allocation-conscious model without a concrete
//! format dependency.

#![allow(
    clippy::module_name_repetitions,
    reason = "Formula-prefixed names keep the shared public vocabulary explicit at call sites"
)]

/// A typed display cache stored alongside a native formula reference.
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaCachedValue {
    /// A finite numeric result.
    Number(f64),
    /// A textual result.
    Text(String),
    /// A Boolean result.
    Boolean(bool),
    /// An iWork date represented in Apple-epoch seconds.
    Date(f64),
    /// An iWork duration represented in seconds.
    Duration(f64),
}

impl FormulaCachedValue {
    /// Convert the formula cache into a value owned by a concrete cell model.
    ///
    /// The target type is deliberately generic: the neutral formula leaf does
    /// not depend on a concrete format's cell vocabulary. Format crates provide
    /// the appropriate `From` implementation at their archive boundary.
    #[must_use]
    pub fn into_value<T>(self) -> T
    where
        T: From<Self>,
    {
        T::from(self)
    }
}

/// A cell address used by an iWork table formula.
///
/// The row and column are zero-based logical table coordinates. Absolute flags
/// control how the address behaves when an iWork app fills or moves the formula;
/// they do not change which cell is referenced at the formula's current host.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaCellReference {
    pub row: usize,
    pub column: usize,
    pub absolute_row: bool,
    pub absolute_column: bool,
}

impl FormulaCellReference {
    /// Construct a reference whose row and column both move with the formula.
    #[must_use]
    pub const fn relative(row: usize, column: usize) -> Self {
        Self {
            row,
            column,
            absolute_row: false,
            absolute_column: false,
        }
    }

    /// Construct a reference fixed to an absolute row and column.
    #[must_use]
    pub const fn absolute(row: usize, column: usize) -> Self {
        Self {
            row,
            column,
            absolute_row: true,
            absolute_column: true,
        }
    }

    /// Construct a reference with independently controlled row/column modes.
    #[must_use]
    pub const fn mixed(
        row: usize,
        column: usize,
        absolute_row: bool,
        absolute_column: bool,
    ) -> Self {
        Self {
            row,
            column,
            absolute_row,
            absolute_column,
        }
    }
}

/// A row or column endpoint used by a whole-axis formula reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaAxisReference {
    pub index: usize,
    pub absolute: bool,
}

impl FormulaAxisReference {
    /// Construct an axis endpoint that moves with the formula.
    #[must_use]
    pub const fn relative(index: usize) -> Self {
        Self {
            index,
            absolute: false,
        }
    }

    /// Construct an axis endpoint fixed to its absolute index.
    #[must_use]
    pub const fn absolute(index: usize) -> Self {
        Self {
            index,
            absolute: true,
        }
    }
}

/// A 128-bit identifier used by an iWork pivot formula model.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FormulaUuid {
    pub lower: u64,
    pub upper: u64,
}

impl FormulaUuid {
    /// Construct a formula UUID from its two native 64-bit words.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }
}

/// An absolute category aggregate in an iWork pivot table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaPivotCategoryReference {
    pub group_by_uid: FormulaUuid,
    pub column_uid: FormulaUuid,
    pub group_uid: FormulaUuid,
    pub aggregate_type: u32,
    pub group_level: i32,
}

impl FormulaPivotCategoryReference {
    /// Construct a pivot category reference from its native identifiers.
    #[must_use]
    pub const fn new(
        group_by_uid: FormulaUuid,
        column_uid: FormulaUuid,
        group_uid: FormulaUuid,
        aggregate_type: u32,
        group_level: i32,
    ) -> Self {
        Self {
            group_by_uid,
            column_uid,
            group_uid,
            aggregate_type,
            group_level,
        }
    }
}

/// A binary operator in an iWork table formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaBinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concatenate,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
}

/// A typed formula expression compiled to iWork's postfix protobuf AST by the
/// owning archive crate.
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaExpression {
    Number(f64),
    Text(String),
    Boolean(bool),
    PivotCategory(FormulaPivotCategoryReference),
    Cell(FormulaCellReference),
    TableCell {
        table_id: u64,
        reference: FormulaCellReference,
    },
    TableRange {
        table_id: u64,
        start: FormulaCellReference,
        end: FormulaCellReference,
    },
    Rows {
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    },
    Columns {
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    },
    TableRows {
        table_id: u64,
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    },
    TableColumns {
        table_id: u64,
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    },
    Range {
        start: FormulaCellReference,
        end: FormulaCellReference,
    },
    Function {
        name: String,
        arguments: Vec<FormulaExpression>,
    },
    Binary {
        operator: FormulaBinaryOperator,
        left: Box<FormulaExpression>,
        right: Box<FormulaExpression>,
    },
    Negate(Box<FormulaExpression>),
    Percent(Box<FormulaExpression>),
}

impl FormulaExpression {
    /// Construct a function call from a name and its arguments.
    #[must_use]
    pub fn function(
        name: impl Into<String>,
        arguments: impl IntoIterator<Item = FormulaExpression>,
    ) -> Self {
        Self::Function {
            name: name.into(),
            arguments: arguments.into_iter().collect(),
        }
    }

    /// Construct a binary expression.
    #[must_use]
    pub fn binary(
        operator: FormulaBinaryOperator,
        left: FormulaExpression,
        right: FormulaExpression,
    ) -> Self {
        Self::Binary {
            operator,
            left: Box::new(left),
            right: Box::new(right),
        }
    }

    /// Construct a unary negation without exposing the backing allocation.
    #[must_use]
    pub fn negate(value: FormulaExpression) -> Self {
        Self::Negate(Box::new(value))
    }

    /// Construct a percentage expression without exposing the backing
    /// allocation.
    #[must_use]
    pub fn percent(value: FormulaExpression) -> Self {
        Self::Percent(Box::new(value))
    }

    /// Construct a local cell reference expression.
    #[must_use]
    pub const fn cell(reference: FormulaCellReference) -> Self {
        Self::Cell(reference)
    }

    /// Construct a pivot category expression.
    #[must_use]
    pub const fn pivot_category(reference: FormulaPivotCategoryReference) -> Self {
        Self::PivotCategory(reference)
    }

    /// Construct a relative local cell reference expression.
    #[must_use]
    pub const fn relative_cell(row: usize, column: usize) -> Self {
        Self::Cell(FormulaCellReference::relative(row, column))
    }

    /// Construct a cross-table cell reference expression.
    #[must_use]
    pub const fn table_cell(table_id: u64, reference: FormulaCellReference) -> Self {
        Self::TableCell {
            table_id,
            reference,
        }
    }

    /// Construct a cross-table rectangular range expression.
    #[must_use]
    pub const fn table_range(
        table_id: u64,
        start: FormulaCellReference,
        end: FormulaCellReference,
    ) -> Self {
        Self::TableRange {
            table_id,
            start,
            end,
        }
    }

    /// Construct a local whole-row range expression.
    #[must_use]
    pub const fn rows(start: FormulaAxisReference, end: FormulaAxisReference) -> Self {
        Self::Rows { start, end }
    }

    /// Construct a local whole-column range expression.
    #[must_use]
    pub const fn columns(start: FormulaAxisReference, end: FormulaAxisReference) -> Self {
        Self::Columns { start, end }
    }

    /// Construct a cross-table whole-row range expression.
    #[must_use]
    pub const fn table_rows(
        table_id: u64,
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    ) -> Self {
        Self::TableRows {
            table_id,
            start,
            end,
        }
    }

    /// Construct a cross-table whole-column range expression.
    #[must_use]
    pub const fn table_columns(
        table_id: u64,
        start: FormulaAxisReference,
        end: FormulaAxisReference,
    ) -> Self {
        Self::TableColumns {
            table_id,
            start,
            end,
        }
    }

    /// Construct a local rectangular range expression.
    #[must_use]
    pub const fn range(start: FormulaCellReference, end: FormulaCellReference) -> Self {
        Self::Range { start, end }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Debug, PartialEq)]
    enum TestCellValue {
        Number(f64),
        Text(String),
        Boolean(bool),
        Date(f64),
        Duration(f64),
    }

    impl From<FormulaCachedValue> for TestCellValue {
        fn from(value: FormulaCachedValue) -> Self {
            match value {
                FormulaCachedValue::Number(value) => Self::Number(value),
                FormulaCachedValue::Text(value) => Self::Text(value),
                FormulaCachedValue::Boolean(value) => Self::Boolean(value),
                FormulaCachedValue::Date(value) => Self::Date(value),
                FormulaCachedValue::Duration(value) => Self::Duration(value),
            }
        }
    }

    #[test]
    fn cached_values_convert_through_the_concrete_cell_boundary() {
        assert_eq!(
            FormulaCachedValue::Number(3.5).into_value::<TestCellValue>(),
            TestCellValue::Number(3.5)
        );
        assert_eq!(
            FormulaCachedValue::Text("ok".to_owned()).into_value::<TestCellValue>(),
            TestCellValue::Text("ok".to_owned())
        );
        assert_eq!(
            FormulaCachedValue::Boolean(true).into_value::<TestCellValue>(),
            TestCellValue::Boolean(true)
        );
        assert_eq!(
            FormulaCachedValue::Date(12.0).into_value::<TestCellValue>(),
            TestCellValue::Date(12.0)
        );
        assert_eq!(
            FormulaCachedValue::Duration(4.5).into_value::<TestCellValue>(),
            TestCellValue::Duration(4.5)
        );
    }

    #[test]
    fn constructors_preserve_reference_modes_and_nesting() {
        let reference = FormulaCellReference::mixed(4, 2, true, false);
        let expression = FormulaExpression::binary(
            FormulaBinaryOperator::Add,
            FormulaExpression::cell(reference),
            FormulaExpression::relative_cell(0, 1),
        );
        assert_eq!(
            expression,
            FormulaExpression::Binary {
                operator: FormulaBinaryOperator::Add,
                left: Box::new(FormulaExpression::Cell(reference)),
                right: Box::new(FormulaExpression::Cell(FormulaCellReference::relative(
                    0, 1
                ))),
            }
        );
        assert_eq!(
            FormulaExpression::negate(FormulaExpression::Number(1.0)),
            FormulaExpression::Negate(Box::new(FormulaExpression::Number(1.0)))
        );
        assert_eq!(
            FormulaExpression::percent(FormulaExpression::Number(50.0)),
            FormulaExpression::Percent(Box::new(FormulaExpression::Number(50.0)))
        );
    }

    #[test]
    fn constructors_cover_axis_and_pivot_references() {
        let group_by = FormulaUuid::new(1, 2);
        let column = FormulaUuid::new(3, 4);
        let group = FormulaUuid::new(5, 6);
        let pivot = FormulaPivotCategoryReference::new(group_by, column, group, 2, 1);
        assert_eq!(
            FormulaExpression::pivot_category(pivot),
            FormulaExpression::PivotCategory(pivot)
        );
        assert_eq!(
            FormulaExpression::table_rows(
                7,
                FormulaAxisReference::relative(1),
                FormulaAxisReference::absolute(3)
            ),
            FormulaExpression::TableRows {
                table_id: 7,
                start: FormulaAxisReference::relative(1),
                end: FormulaAxisReference::absolute(3),
            }
        );
    }
}
