//! Bounded database-function evaluation for the value VM.
//!
//! This module is deliberately separate from the ordinary scalar and matrix
//! dispatch.  A database function consumes two rectangular values (the
//! database and criteria) and one scalar field selector.  Treating those
//! inputs as ordinary broadcast arguments would implicitly intersect a range
//! and silently change the query.  The parent VM therefore evaluates the
//! database and criteria arguments in matrix context and calls [`apply`]
//! after the field argument has been evaluated.
//!
//! The implementation keeps only database/criteria headers, compiled
//! criterion clauses, and the columns needed by those clauses plus the
//! selected field.  Records are read one at a time.  A reference database is
//! currently one contiguous area on one sheet; a reference list or a 3-D
//! result is a different pseudotype and is refused before any cell read.
//!
//! The selected profile uses OR between criteria rows and AND within one row.
//! Criteria text is exact and case-sensitive; regular expressions, wildcards,
//! and substring matching are disabled.  Database and criteria field headers
//! are matched case-insensitively, while duplicate matches are an explicit
//! formula error.  Formula values encountered through the resolver remain a
//! typed provider refusal; this layer never uses an inert cached formula.

use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeArea, RuntimeArrayValue, RuntimeElement,
    RuntimeValue, ScalarError, ValueEvaluator, WorkingValue, ensure_capacity,
};
use litchi_core::{Reservation, Resource};
use std::cmp::Ordering;

mod numerics;
use numerics::{NumericAggregate, NumericOperation};

/// Return whether `name` is one of the twelve functions in OpenFormula §6.9.
pub(super) fn is_database_function(name: &str) -> bool {
    database_function(name).is_some()
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DatabaseFunction {
    Average,
    Count,
    CountA,
    Get,
    Max,
    Min,
    Product,
    StdSample,
    StdPopulation,
    Sum,
    VarSample,
    VarPopulation,
}

fn database_function(name: &str) -> Option<DatabaseFunction> {
    Some(if name.eq_ignore_ascii_case("DAVERAGE") {
        DatabaseFunction::Average
    } else if name.eq_ignore_ascii_case("DCOUNT") {
        DatabaseFunction::Count
    } else if name.eq_ignore_ascii_case("DCOUNTA") {
        DatabaseFunction::CountA
    } else if name.eq_ignore_ascii_case("DGET") {
        DatabaseFunction::Get
    } else if name.eq_ignore_ascii_case("DMAX") {
        DatabaseFunction::Max
    } else if name.eq_ignore_ascii_case("DMIN") {
        DatabaseFunction::Min
    } else if name.eq_ignore_ascii_case("DPRODUCT") {
        DatabaseFunction::Product
    } else if name.eq_ignore_ascii_case("DSTDEV") {
        DatabaseFunction::StdSample
    } else if name.eq_ignore_ascii_case("DSTDEVP") {
        DatabaseFunction::StdPopulation
    } else if name.eq_ignore_ascii_case("DSUM") {
        DatabaseFunction::Sum
    } else if name.eq_ignore_ascii_case("DVAR") {
        DatabaseFunction::VarSample
    } else if name.eq_ignore_ascii_case("DVARP") {
        DatabaseFunction::VarPopulation
    } else {
        return None;
    })
}

/// Apply one database function after its arguments have been evaluated.
///
/// The parent VM owns argument scheduling and invokes this function before
/// generic scalar/matrix dispatch.  `DCOUNT` and `DCOUNTA` accept either the
/// explicit three-argument form with a missing middle argument or the
/// two-argument omitted-field form.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = database_function(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };

    // A direct formula error is an argument value and must retain its source
    // order before shape or header preparation starts. Errors in records are
    // considered only when the required record column is actually read.
    for argument in &arguments {
        if let RuntimeValue::Scalar(WorkingValue::Error(error)) = argument {
            return Ok(formula_error(*error));
        }
    }

    let (database_index, field_index, criteria_index) = match (function, arguments.len()) {
        (DatabaseFunction::Count | DatabaseFunction::CountA, 2) => (0, None, 1),
        (DatabaseFunction::Count | DatabaseFunction::CountA, 3) => (0, Some(1), 2),
        (_, 3) => (0, Some(1), 2),
        _ => return Ok(formula_error(ScalarError::Value)),
    };

    // The parent schedules Field through its scalar-parameter frame, but the
    // call remains defensive because a computed one-cell array can still
    // reach this boundary.  This also handles an explicit missing field for
    // DCOUNT/DCOUNTA without turning it into an empty field name.
    let field_value = if let Some(index) = field_index {
        let value = std::mem::replace(&mut arguments[index], RuntimeValue::Missing);
        evaluator.matrix_scalar_parameter(value)?
    } else {
        RuntimeValue::Missing
    };
    if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &field_value {
        return Ok(formula_error(*error));
    }

    let database_source = match RectSource::from_runtime(&arguments[database_index]) {
        Ok(Some(source)) => source,
        Ok(None) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Err(error),
    };
    let criteria_source = match RectSource::from_runtime(&arguments[criteria_index]) {
        Ok(Some(source)) => source,
        Ok(None) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Err(error),
    };

    let database = match DatabasePlan::build(evaluator, database_source)? {
        Ok(database) => database,
        Err(error) => return Ok(formula_error(error)),
    };
    let has_selected_field = field_index.is_some()
        && !(matches!(function, DatabaseFunction::Count | DatabaseFunction::CountA)
            && matches!(&field_value, RuntimeValue::Missing));
    let selected_field = if has_selected_field {
        match select_field(evaluator, &database.headers, &field_value)? {
            Ok(index) => Some(index),
            Err(error) => return Ok(formula_error(error)),
        }
    } else {
        None
    };
    let criteria = match CriteriaPlan::build(evaluator, criteria_source, &database.headers)? {
        Ok(criteria) => criteria,
        Err(error) => return Ok(formula_error(error)),
    };

    run_query(evaluator, function, &database, &criteria, selected_field)
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

/// One rectangular source backed by an inline array, one reference area, or
/// a single scalar database cell.  The source borrow is tied to the argument
/// vector, not to this wrapper, so header/criterion text can safely be kept
/// in the plans without making either plan self-referential.
#[derive(Clone, Copy)]
enum RectSource<'source, 'expr> {
    Array(&'source RuntimeArrayValue<'expr>),
    Reference(RuntimeArea<'expr>),
    Scalar(&'source RuntimeValue<'expr>),
}

impl<'source, 'expr> RectSource<'source, 'expr> {
    fn from_runtime(
        value: &'source RuntimeValue<'expr>,
    ) -> Result<Option<Self>, EvaluationFailure> {
        match value {
            RuntimeValue::Array(array) => {
                let Some(cells) = array.shape.cell_count() else {
                    return Ok(None);
                };
                if cells != array.cells.len() {
                    return Ok(None);
                }
                Ok(Some(Self::Array(array)))
            },
            RuntimeValue::Areas(areas) => {
                if areas.is_list || areas.areas.len() != 1 {
                    return Err(EvaluationFailure::Unsupported(
                        super::super::UnsupportedKind::Reference,
                    ));
                }
                let Some(area) = areas.areas.first().copied() else {
                    return Ok(None);
                };
                Ok(Some(Self::Reference(area)))
            },
            RuntimeValue::Scalar(WorkingValue::Error(_)) => Ok(Some(Self::Scalar(value))),
            RuntimeValue::Scalar(_) => Ok(Some(Self::Scalar(value))),
            RuntimeValue::Empty | RuntimeValue::Missing => Ok(None),
            RuntimeValue::ScalarCell(_) => Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference,
            )),
        }
    }

    fn rows(self) -> usize {
        match self {
            Self::Array(array) => array.shape.rows(),
            Self::Reference(area) => area.rect.rows(),
            Self::Scalar(_) => 1,
        }
    }

    fn columns(self) -> usize {
        match self {
            Self::Array(array) => array.shape.columns(),
            Self::Reference(area) => area.rect.columns(),
            Self::Scalar(_) => 1,
        }
    }

    fn cell<R>(
        self,
        evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
        row: usize,
        column: usize,
    ) -> EvaluationResult<DbCell<'source>>
    where
        'expr: 'source,
        R: Resolver + ?Sized,
    {
        if row >= self.rows() || column >= self.columns() {
            return Err(EvaluationFailure::InvalidExpression(
                "database source coordinate is out of bounds",
            ));
        }
        evaluator.scalar.charge_work(1)?;
        match self {
            Self::Array(array) => {
                let index = row
                    .checked_mul(array.shape.columns())
                    .and_then(|index| index.checked_add(column))
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "database array index overflows",
                    ))?;
                let element: &'source RuntimeElement<'expr> =
                    array
                        .cells
                        .get(index)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "database array cell is missing",
                        ))?;
                Ok(DbCell::from_element(element))
            },
            Self::Reference(area) => {
                let read = evaluator.read_reference_cell(
                    area.sheet,
                    area.rect.row_start.checked_add(row).ok_or(
                        EvaluationFailure::InvalidExpression("database row overflows"),
                    )?,
                    area.rect.column_start.checked_add(column).ok_or(
                        EvaluationFailure::InvalidExpression("database column overflows"),
                    )?,
                )?;
                if let super::CellRead::Text(text) = read {
                    if text.len() > evaluator.limits.scalar.max_text_bytes {
                        return Err(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                            Resource::Memory,
                            u64::try_from(text.len()).unwrap_or(u64::MAX),
                            evaluator.limits.scalar.max_text_bytes,
                        )));
                    }
                }
                DbCell::from_cell_read(read)
            },
            Self::Scalar(value) => Ok(DbCell::from_runtime_value(value)),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum DbCell<'a> {
    Empty,
    Number(f64),
    Logical(bool),
    Text(&'a str),
    Error(ScalarError),
    Complex,
}

impl<'a> DbCell<'a> {
    fn from_element(element: &'a RuntimeElement<'_>) -> Self {
        match element {
            RuntimeElement::Empty => Self::Empty,
            RuntimeElement::Missing => Self::Error(ScalarError::NotAvailable),
            RuntimeElement::Present(WorkingValue::Number(value)) => Self::Number(*value),
            RuntimeElement::Present(WorkingValue::Logical(value)) => Self::Logical(*value),
            RuntimeElement::Present(WorkingValue::Text(value)) => Self::Text(value.text.as_ref()),
            RuntimeElement::Present(WorkingValue::Error(error)) => Self::Error(*error),
            RuntimeElement::Present(WorkingValue::Complex(_)) => Self::Complex,
        }
    }

    fn from_runtime_value(value: &'a RuntimeValue<'_>) -> Self {
        match value {
            RuntimeValue::Empty | RuntimeValue::Missing => Self::Empty,
            RuntimeValue::Scalar(WorkingValue::Number(value)) => Self::Number(*value),
            RuntimeValue::Scalar(WorkingValue::Logical(value)) => Self::Logical(*value),
            RuntimeValue::Scalar(WorkingValue::Text(value)) => Self::Text(value.text.as_ref()),
            RuntimeValue::Scalar(WorkingValue::Error(error)) => Self::Error(*error),
            RuntimeValue::Scalar(WorkingValue::Complex(_)) => Self::Complex,
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
                Self::Error(ScalarError::Value)
            },
        }
    }

    fn from_cell_read(read: super::CellRead<'a>) -> EvaluationResult<Self> {
        Ok(match read {
            super::CellRead::Empty => Self::Empty,
            super::CellRead::Number(value) if value.is_finite() => Self::Number(value),
            super::CellRead::Number(_) => Self::Error(ScalarError::Number),
            super::CellRead::Logical(value) => Self::Logical(value),
            super::CellRead::Text(value) => Self::Text(value),
            super::CellRead::Error(error) => Self::Error(error),
            super::CellRead::Unsupported => {
                return Err(EvaluationFailure::Unsupported(
                    super::super::UnsupportedKind::CellValue,
                ));
            },
        })
    }

    fn is_empty(self) -> bool {
        matches!(self, Self::Empty)
    }
}

/// Header values and source geometry retained during one query.  `headers`
/// contains only the first row, never the complete database table.
struct DatabasePlan<'source, 'expr> {
    source: RectSource<'source, 'expr>,
    headers: Vec<DbCell<'source>>,
    _header_reservation: Option<Reservation>,
}

impl<'source, 'expr> DatabasePlan<'source, 'expr> {
    fn build<'scalar, 'exec, 'position, R>(
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
        source: RectSource<'source, 'expr>,
    ) -> EvaluationResult<Result<Self, ScalarError>>
    where
        R: Resolver + ?Sized,
    {
        let rows = source.rows();
        let columns = source.columns();
        if rows == 0 || columns == 0 {
            return Ok(Err(ScalarError::Value));
        }
        let mut headers = Vec::new();
        let mut header_reservation = None;
        ensure_capacity(
            &mut headers,
            &mut header_reservation,
            columns,
            evaluator.limits.max_array_cells,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula database headers",
        )?;
        for column in 0..columns {
            let header = source.cell(evaluator, 0, column)?;
            headers.push(header);
        }
        Ok(Ok(Self {
            source,
            headers,
            _header_reservation: header_reservation,
        }))
    }

    fn rows(&self) -> usize {
        self.source.rows()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CriterionOperator {
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
}

#[derive(Clone, Copy, Debug)]
enum CriterionMatcher<'a> {
    Empty(CriterionOperator),
    Number(CriterionOperator, f64),
    Logical(CriterionOperator, bool),
    Text(CriterionOperator, &'a str),
}

#[derive(Clone, Copy, Debug)]
struct CriterionClause<'a> {
    field: usize,
    matcher: CriterionMatcher<'a>,
}

/// Compiled criteria cells.  The text in a matcher remains borrowed from the
/// evaluated criteria argument; no per-record parse or string copy occurs.
struct CriteriaPlan<'source> {
    rows: usize,
    columns: usize,
    clauses: Vec<CriterionClause<'source>>,
    _clause_reservation: Option<Reservation>,
}

impl<'source> CriteriaPlan<'source> {
    fn build<'expr, 'scalar, 'exec, 'position, R>(
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
        source: RectSource<'source, 'expr>,
        headers: &[DbCell<'source>],
    ) -> EvaluationResult<Result<Self, ScalarError>>
    where
        R: Resolver + ?Sized,
    {
        let rows = source.rows();
        let columns = source.columns();
        if rows < 2 || columns == 0 {
            return Ok(Err(ScalarError::Value));
        }
        let mut fields = Vec::new();
        let mut field_reservation = None;
        ensure_capacity(
            &mut fields,
            &mut field_reservation,
            columns,
            evaluator.limits.max_array_cells,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula database criteria fields",
        )?;
        for column in 0..columns {
            let header = source.cell(evaluator, 0, column)?;
            let field = match criteria_field(evaluator, headers, header)? {
                Ok(field) => field,
                Err(error) => return Ok(Err(error)),
            };
            fields.push(field);
        }

        let body_rows = rows - 1;
        let clause_count = body_rows.checked_mul(columns).ok_or_else(|| {
            EvaluationFailure::ResourceLimit(evaluator.local_limit(
                Resource::Objects,
                u64::MAX,
                evaluator.limits.max_array_cells,
            ))
        })?;
        let mut clauses = Vec::new();
        let mut clause_reservation = None;
        ensure_capacity(
            &mut clauses,
            &mut clause_reservation,
            clause_count,
            evaluator.limits.max_array_cells,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula database criteria clauses",
        )?;
        for row in 1..rows {
            for (column, field) in fields.iter().copied().enumerate() {
                let cell = source.cell(evaluator, row, column)?;
                let matcher = match compile_criterion(evaluator, cell)? {
                    Ok(matcher) => matcher,
                    Err(error) => return Ok(Err(error)),
                };
                clauses.push(CriterionClause { field, matcher });
            }
        }
        drop(fields);
        drop(field_reservation);
        Ok(Ok(Self {
            rows: body_rows,
            columns,
            clauses,
            _clause_reservation: clause_reservation,
        }))
    }

    fn required_columns<'scalar, 'exec, 'position, R>(
        &self,
        evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    ) -> EvaluationResult<ColumnSet>
    where
        R: Resolver + ?Sized,
    {
        let mut columns = Vec::new();
        let mut reservation = None;
        let maximum = evaluator.limits.max_array_cells;
        for clause in &self.clauses {
            let field = clause.field;
            add_required_column(evaluator, &mut columns, &mut reservation, field, maximum)?;
        }
        // Sorting permits binary lookup of the slots; actual provider reads
        // remain lazy in criterion clause order. The selected field is
        // deliberately admitted/read only after a record matches; this keeps
        // errors and unsupported values in nonmatching records uninspected.
        evaluator
            .scalar
            .charge_work(u64::try_from(columns.len()).unwrap_or(u64::MAX))?;
        columns.sort_unstable();
        Ok(ColumnSet {
            columns,
            _reservation: reservation,
        })
    }

    fn matches_row<'expr, 'scalar, 'exec, 'position, R>(
        &self,
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
        required: &[usize],
        values: &mut [Option<DbCell<'source>>],
        source: RectSource<'source, 'expr>,
        database_row: usize,
        row: usize,
    ) -> EvaluationResult<Result<bool, ScalarError>>
    where
        R: Resolver + ?Sized,
    {
        // Each attempted criteria alternative is a bounded query step.
        evaluator.scalar.charge_work(1)?;
        let start = row
            .checked_mul(self.columns)
            .ok_or(EvaluationFailure::InvalidExpression(
                "criteria row offset overflows",
            ))?;
        let end = start
            .checked_add(self.columns)
            .ok_or(EvaluationFailure::InvalidExpression(
                "criteria row end overflows",
            ))?;
        let clauses = self
            .clauses
            .get(start..end)
            .ok_or(EvaluationFailure::InvalidExpression(
                "criteria clauses do not match their dimensions",
            ))?;
        for clause in clauses {
            let field = clause.field;
            let slot = required.binary_search(&field).map_err(|_| {
                EvaluationFailure::InvalidExpression("criteria field was not admitted")
            })?;
            let cached = values
                .get_mut(slot)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "criteria row value is missing",
                ))?;
            let candidate = match *cached {
                Some(value) => value,
                None => {
                    let value = source.cell(evaluator, database_row, field)?;
                    *cached = Some(value);
                    value
                },
            };
            match criterion_matches(evaluator, clause.matcher, candidate)? {
                Ok(true) => {},
                Ok(false) => return Ok(Ok(false)),
                Err(error) => return Ok(Err(error)),
            }
        }
        Ok(Ok(true))
    }
}

struct ColumnSet {
    columns: Vec<usize>,
    _reservation: Option<Reservation>,
}

struct RowScratch<'a> {
    values: Vec<Option<DbCell<'a>>>,
    _reservation: Option<Reservation>,
}

impl<'a> RowScratch<'a> {
    fn new<'scalar, 'exec, 'position, R>(
        evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
        length: usize,
    ) -> EvaluationResult<Self>
    where
        R: Resolver + ?Sized,
    {
        let mut values = Vec::new();
        let mut reservation = None;
        ensure_capacity(
            &mut values,
            &mut reservation,
            length,
            evaluator.limits.max_array_cells,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula database row scratch",
        )?;
        Ok(Self {
            values,
            _reservation: reservation,
        })
    }
}

fn add_required_column<'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    columns: &mut Vec<usize>,
    reservation: &mut Option<Reservation>,
    field: usize,
    maximum: usize,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for existing in columns.iter().copied() {
        evaluator.scalar.charge_work(1)?;
        if existing == field {
            return Ok(());
        }
    }
    ensure_capacity(
        columns,
        reservation,
        1,
        maximum,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula database required columns",
    )?;
    columns.push(field);
    Ok(())
}

fn criteria_field<'source, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    headers: &[DbCell<'source>],
    value: DbCell<'source>,
) -> EvaluationResult<Result<usize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    match value {
        DbCell::Empty => Ok(Err(ScalarError::Value)),
        DbCell::Error(error) => Ok(Err(error)),
        DbCell::Number(value) => Ok(number_field_index(value, headers.len())),
        DbCell::Text(name) => {
            let mut found = None;
            for (index, header) in headers.iter().copied().enumerate() {
                let DbCell::Text(header) = header else {
                    continue;
                };
                if header_equal(evaluator, name, header)? {
                    if found.is_some() {
                        return Ok(Err(ScalarError::Value));
                    }
                    found = Some(index);
                }
            }
            Ok(found.ok_or(ScalarError::Value))
        },
        DbCell::Logical(_) | DbCell::Complex => Ok(Err(ScalarError::Value)),
    }
}

fn select_field<'source, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    headers: &[DbCell<'source>],
    value: &RuntimeValue<'_>,
) -> EvaluationResult<Result<usize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = match value {
        RuntimeValue::Empty | RuntimeValue::Missing => return Ok(Err(ScalarError::Value)),
        RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(Err(*error)),
        RuntimeValue::Scalar(WorkingValue::Number(value)) => DbCell::Number(*value),
        RuntimeValue::Scalar(WorkingValue::Logical(value)) => DbCell::Logical(*value),
        RuntimeValue::Scalar(WorkingValue::Text(value)) => DbCell::Text(value.text.as_ref()),
        RuntimeValue::Scalar(WorkingValue::Complex(_)) => DbCell::Complex,
        RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
            return Ok(Err(ScalarError::Value));
        },
    };
    match value {
        DbCell::Number(value) => Ok(number_field_index(value, headers.len())),
        DbCell::Text(name) => {
            let mut found = None;
            for (index, header) in headers.iter().copied().enumerate() {
                let DbCell::Text(header) = header else {
                    continue;
                };
                if header_equal(evaluator, name, header)? {
                    if found.is_some() {
                        return Ok(Err(ScalarError::Value));
                    }
                    found = Some(index);
                }
            }
            Ok(found.ok_or(ScalarError::Value))
        },
        DbCell::Empty | DbCell::Logical(_) | DbCell::Error(_) | DbCell::Complex => {
            Ok(Err(ScalarError::Value))
        },
    }
}

fn number_field_index(value: f64, columns: usize) -> Result<usize, ScalarError> {
    if !value.is_finite() || value < 1.0 || value.fract() != 0.0 || value > columns as f64 {
        return Err(ScalarError::Value);
    }
    Ok(value as usize - 1)
}

fn header_equal<'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    left: &str,
    right: &str,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    // `to_lowercase` is allocation-free and gives a deterministic Unicode
    // scalar case-insensitive header policy without changing criterion text.
    let mut left = left.chars().flat_map(|character| character.to_lowercase());
    let mut right = right.chars().flat_map(|character| character.to_lowercase());
    loop {
        evaluator.scalar.charge_work(1)?;
        match (left.next(), right.next()) {
            (None, None) => return Ok(true),
            (Some(left), Some(right)) if left == right => {},
            _ => return Ok(false),
        }
    }
}

fn compile_criterion<'source, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    value: DbCell<'source>,
) -> EvaluationResult<Result<CriterionMatcher<'source>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    match value {
        DbCell::Empty => Ok(Ok(CriterionMatcher::Number(CriterionOperator::Equal, 0.0))),
        DbCell::Number(value) => Ok(Ok(CriterionMatcher::Number(
            CriterionOperator::Equal,
            value,
        ))),
        DbCell::Logical(value) => Ok(Ok(CriterionMatcher::Logical(
            CriterionOperator::Equal,
            value,
        ))),
        DbCell::Error(error) => Ok(Err(error)),
        DbCell::Complex => Ok(Err(ScalarError::Value)),
        DbCell::Text(text) => {
            evaluator.scalar.charge_bytes(text.len())?;
            let (operator, rhs) = criterion_parts(text);
            if rhs.is_empty() {
                return Ok(Ok(CriterionMatcher::Empty(operator)));
            }
            if let Ok(number) = fast_float2::parse::<f64, _>(rhs)
                && number.is_finite()
            {
                return Ok(Ok(CriterionMatcher::Number(operator, number)));
            }
            Ok(Ok(CriterionMatcher::Text(operator, rhs)))
        },
    }
}

fn criterion_parts(value: &str) -> (CriterionOperator, &str) {
    if let Some(rest) = value.strip_prefix(">=") {
        (CriterionOperator::GreaterEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<=") {
        (CriterionOperator::LessEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<>") {
        (CriterionOperator::NotEqual, rest)
    } else if let Some(rest) = value.strip_prefix('>') {
        (CriterionOperator::Greater, rest)
    } else if let Some(rest) = value.strip_prefix('<') {
        (CriterionOperator::Less, rest)
    } else if let Some(rest) = value.strip_prefix('=') {
        (CriterionOperator::Equal, rest)
    } else {
        (CriterionOperator::Equal, value)
    }
}

fn criterion_matches<'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'_, 'scalar, 'exec, 'position, R>,
    matcher: CriterionMatcher<'_>,
    candidate: DbCell<'_>,
) -> EvaluationResult<Result<bool, ScalarError>>
where
    R: Resolver + ?Sized,
{
    if let DbCell::Error(error) = candidate {
        return Ok(Err(error));
    }
    let result = match matcher {
        CriterionMatcher::Empty(operator) => match operator {
            CriterionOperator::Equal => candidate.is_empty(),
            CriterionOperator::NotEqual => !candidate.is_empty(),
            CriterionOperator::Less
            | CriterionOperator::LessEqual
            | CriterionOperator::Greater
            | CriterionOperator::GreaterEqual => false,
        },
        CriterionMatcher::Number(operator, expected) => match candidate {
            DbCell::Number(actual) => compare_numbers(operator, actual, expected),
            DbCell::Empty | DbCell::Logical(_) | DbCell::Text(_) | DbCell::Complex => {
                matches!(operator, CriterionOperator::NotEqual)
            },
            DbCell::Error(_) => unreachable!("formula errors are handled above"),
        },
        CriterionMatcher::Logical(operator, expected) => match candidate {
            DbCell::Logical(actual) => compare_booleans(operator, actual, expected),
            DbCell::Empty | DbCell::Number(_) | DbCell::Text(_) | DbCell::Complex => {
                matches!(operator, CriterionOperator::NotEqual)
            },
            DbCell::Error(_) => unreachable!("formula errors are handled above"),
        },
        CriterionMatcher::Text(operator, expected) => match candidate {
            DbCell::Text(actual) => {
                evaluator.scalar.charge_bytes(
                    actual
                        .len()
                        .checked_add(expected.len())
                        .unwrap_or(usize::MAX),
                )?;
                compare_text(operator, actual, expected)
            },
            DbCell::Empty | DbCell::Number(_) | DbCell::Logical(_) | DbCell::Complex => {
                matches!(operator, CriterionOperator::NotEqual)
            },
            DbCell::Error(_) => unreachable!("formula errors are handled above"),
        },
    };
    Ok(Ok(result))
}

fn compare_numbers(operator: CriterionOperator, left: f64, right: f64) -> bool {
    match operator {
        CriterionOperator::Equal => left == right,
        CriterionOperator::NotEqual => left != right,
        CriterionOperator::Less => left < right,
        CriterionOperator::LessEqual => left <= right,
        CriterionOperator::Greater => left > right,
        CriterionOperator::GreaterEqual => left >= right,
    }
}

fn compare_booleans(operator: CriterionOperator, left: bool, right: bool) -> bool {
    compare_numbers(operator, u8::from(left) as f64, u8::from(right) as f64)
}

fn compare_text(operator: CriterionOperator, left: &str, right: &str) -> bool {
    let ordering = left.cmp(right);
    match operator {
        CriterionOperator::Equal => ordering == Ordering::Equal,
        CriterionOperator::NotEqual => ordering != Ordering::Equal,
        CriterionOperator::Less => ordering == Ordering::Less,
        CriterionOperator::LessEqual => ordering != Ordering::Greater,
        CriterionOperator::Greater => ordering == Ordering::Greater,
        CriterionOperator::GreaterEqual => ordering != Ordering::Less,
    }
}

struct AggregateState<'a> {
    function: DatabaseFunction,
    matched_records: usize,
    first_selected: Option<DbCell<'a>>,
    first_formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
    nonblank_count: usize,
    numeric: NumericAggregate,
}

impl<'a> AggregateState<'a> {
    fn new(function: DatabaseFunction) -> Self {
        let operation = match function {
            DatabaseFunction::Average => NumericOperation::Average,
            DatabaseFunction::Sum => NumericOperation::Sum,
            DatabaseFunction::Product => NumericOperation::Product,
            DatabaseFunction::Min => NumericOperation::Minimum,
            DatabaseFunction::Max => NumericOperation::Maximum,
            DatabaseFunction::StdSample => NumericOperation::SampleStandardDeviation,
            DatabaseFunction::StdPopulation => NumericOperation::PopulationStandardDeviation,
            DatabaseFunction::VarSample => NumericOperation::SampleVariance,
            DatabaseFunction::VarPopulation => NumericOperation::PopulationVariance,
            DatabaseFunction::Count | DatabaseFunction::CountA | DatabaseFunction::Get => {
                NumericOperation::Count
            },
        };
        Self {
            function,
            matched_records: 0,
            first_selected: None,
            first_formula_error: None,
            generated_error: None,
            nonblank_count: 0,
            numeric: NumericAggregate::new(operation),
        }
    }

    fn matched(&mut self) -> EvaluationResult<()> {
        self.matched_records =
            self.matched_records
                .checked_add(1)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "database match count overflows",
                ))?;
        Ok(())
    }

    fn observe(&mut self, value: DbCell<'a>) {
        if self.function == DatabaseFunction::Get {
            if self.first_selected.is_none() {
                self.first_selected = Some(value);
            }
            return;
        }
        if self.function == DatabaseFunction::CountA {
            if !value.is_empty() {
                self.nonblank_count += 1;
            }
            return;
        }
        if let DbCell::Error(error) = value {
            if self.function != DatabaseFunction::Count && self.first_formula_error.is_none() {
                self.first_formula_error = Some(error);
            }
            return;
        }
        if let DbCell::Number(value) = value {
            if let Err(error) = self.numeric.push_number(value) {
                self.generated_error.get_or_insert(error);
            }
        }
    }
}

fn run_query<'source, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: DatabaseFunction,
    database: &DatabasePlan<'source, 'expr>,
    criteria: &CriteriaPlan<'source>,
    selected_field: Option<usize>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let required = criteria.required_columns(evaluator)?;
    let mut row = RowScratch::new(evaluator, required.columns.len())?;
    let mut state = AggregateState::new(function);

    for database_row in 1..database.rows() {
        // Charge each record even before the first lazy criterion read.
        evaluator.scalar.charge_work(1)?;
        row.values.clear();
        row.values.resize(required.columns.len(), None);
        // A criteria body row is an alternative: all expressions in one row
        // are AND-ed, while successive rows are OR-ed.  Keep the row walk
        // here rather than folding it into `matches_row`, so a false first
        // row does not make the database row index select a different
        // criteria row (and so criteria with fewer rows than the database do
        // not run past their clause arena).
        let mut matches = false;
        for criteria_row in 0..criteria.rows {
            match criteria.matches_row(
                evaluator,
                &required.columns,
                &mut row.values,
                database.source,
                database_row,
                criteria_row,
            )? {
                Ok(true) => {
                    matches = true;
                    break;
                },
                Ok(false) => {},
                Err(error) => return Ok(formula_error(error)),
            }
        }
        if !matches {
            continue;
        }
        state.matched()?;
        if let Some(field) = selected_field {
            if let Ok(slot) = required.columns.binary_search(&field) {
                let value =
                    row.values
                        .get(slot)
                        .copied()
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "selected database field is missing",
                        ))?;
                let value = match value {
                    Some(value) => value,
                    None => database.source.cell(evaluator, database_row, field)?,
                };
                state.observe(value);
            } else {
                // A selected value is part of the query only for a matching
                // record.  In particular, an error or unsupported provider
                // cell in a nonmatching record must not abort the operation.
                let value = database.source.cell(evaluator, database_row, field)?;
                state.observe(value);
            }
        }
    }

    finish_query(evaluator, state, selected_field.is_some())
}

fn finish_query<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: AggregateState<'_>,
    has_selected_field: bool,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if state.function != DatabaseFunction::Count && state.function != DatabaseFunction::CountA {
        if let Some(error) = state.first_formula_error {
            return Ok(formula_error(error));
        }
        if let Some(error) = state.generated_error {
            return Ok(formula_error(error));
        }
    }

    match state.function {
        DatabaseFunction::Get => {
            if state.matched_records != 1 {
                return Ok(formula_error(ScalarError::Value));
            }
            let value = state
                .first_selected
                .ok_or(EvaluationFailure::InvalidExpression(
                    "DGET matched record has no selected value",
                ))?;
            dget_value(evaluator, value)
        },
        DatabaseFunction::Count => {
            let value = if has_selected_field {
                state.numeric.count() as usize
            } else {
                state.matched_records
            };
            Ok(RuntimeValue::Scalar(WorkingValue::Number(value as f64)))
        },
        DatabaseFunction::CountA => {
            let value = if has_selected_field {
                state.nonblank_count
            } else {
                state.matched_records
            };
            Ok(RuntimeValue::Scalar(WorkingValue::Number(value as f64)))
        },
        _ => {
            let count = state.numeric.count();
            let minimum_count = match state.function {
                DatabaseFunction::StdSample | DatabaseFunction::VarSample => 2,
                DatabaseFunction::Average
                | DatabaseFunction::StdPopulation
                | DatabaseFunction::VarPopulation => 1,
                _ => 0,
            };
            if count < minimum_count {
                return Ok(formula_error(ScalarError::Value));
            }
            match state.numeric.result() {
                Ok(Some(value)) => Ok(number(value)),
                Ok(None) => Ok(number(if state.function == DatabaseFunction::Product {
                    1.0
                } else {
                    0.0
                })),
                Err(error) => Ok(formula_error(error)),
            }
        },
    }
}

fn number<'a>(value: f64) -> RuntimeValue<'a> {
    if value.is_finite() {
        RuntimeValue::Scalar(WorkingValue::Number(value))
    } else {
        formula_error(ScalarError::Number)
    }
}

fn dget_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: DbCell<'_>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    match value {
        DbCell::Empty => Ok(RuntimeValue::Scalar(WorkingValue::Number(0.0))),
        DbCell::Number(value) => Ok(number(value)),
        DbCell::Logical(value) => Ok(RuntimeValue::Scalar(WorkingValue::Number(if value {
            1.0
        } else {
            0.0
        }))),
        DbCell::Text(value) => {
            evaluator.scalar.charge_bytes(value.len())?;
            match fast_float2::parse::<f64, _>(value)
                .ok()
                .filter(|value| value.is_finite())
            {
                Some(value) => Ok(number(value)),
                None => Ok(formula_error(ScalarError::Value)),
            }
        },
        DbCell::Error(error) => Ok(formula_error(error)),
        DbCell::Complex => Ok(formula_error(ScalarError::Value)),
    }
}
