//! Semantic coverage for the UTF-8 byte-position text profile in ODF 1.4 §6.7.
//!
//! The selected byte unit is UTF-8 octets of semantic Text.  Returned values
//! remain valid Unicode text: interior starts snap backward to a scalar start
//! and byte-length clipping never emits a partial scalar.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

#[derive(Debug, Clone)]
enum ByteCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct ByteResolver {
    rows: usize,
    columns: usize,
    cells: Vec<ByteCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl ByteResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 8,
            columns: 8,
            cells: (0..64).map(|_| ByteCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };
        resolver.set(0, 0, ByteCell::Text("Aé界🙂".to_owned()));
        resolver.set(1, 0, ByteCell::Text("aß".to_owned()));
        resolver.set(2, 0, ByteCell::Text("needle".to_owned()));
        resolver.set(3, 0, ByteCell::Error(ScalarError::NotAvailable));
        resolver.set(4, 0, ByteCell::Unsupported);
        resolver.set(0, 1, ByteCell::Text("é".to_owned()));
        resolver.set(1, 1, ByteCell::Number(12.5));
        resolver.set(2, 1, ByteCell::Logical(true));
        resolver.set(3, 1, ByteCell::Empty);
        resolver.set(0, 2, ByteCell::Text("Needle".to_owned()));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: ByteCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|value| value.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear_reads(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
    }
}

impl Resolver for ByteResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((matches!(sheet, "Main" | "Data" | "Archive"))
            .then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        self.read_order
            .borrow_mut()
            .push((sheet.to_owned(), row, column));
        if !matches!(sheet, "Main" | "Data" | "Archive") {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if row == 0 && column == 0 {
            return Ok(match sheet {
                "Main" => CellRead::Text("Aé界🙂"),
                "Data" => CellRead::Text("data"),
                "Archive" => CellRead::Text("archive"),
                _ => unreachable!(),
            });
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            ByteCell::Empty => CellRead::Empty,
            ByteCell::Number(value) => CellRead::Number(*value),
            ByteCell::Logical(value) => CellRead::Logical(*value),
            ByteCell::Text(value) => CellRead::Text(value.as_str()),
            ByteCell::Error(error) => CellRead::Error(*error),
            ByteCell::Unsupported => CellRead::Unsupported,
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(match sheet {
            "Main" => Some(0),
            "Data" => Some(1),
            "Archive" => Some(2),
            _ => None,
        })
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(3)
    }

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions else {
            return Ok(None);
        };
        let call = self.source_version_calls.get();
        self.source_version_calls.set(call.saturating_add(1));
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn parse(source: &str) -> Expression {
    Expression::parse(source).unwrap_or_else(|error| panic!("{source:?} should parse: {error}"))
}

#[derive(Debug, PartialEq)]
enum ScalarObserved {
    Text(String),
    Number(f64),
    Logical(bool),
    Error(ScalarError),
    Other,
}

fn scalar(source: &str, execution: &ExecutionContext) -> Result<ScalarObserved, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Text(value) => ScalarObserved::Text(value.to_string()),
        ScalarValue::Number(value) => ScalarObserved::Number(*value),
        ScalarValue::Logical(value) => ScalarObserved::Logical(*value),
        ScalarValue::Error(error) => ScalarObserved::Error(*error),
        ScalarValue::Complex(_) => ScalarObserved::Other,
        _ => ScalarObserved::Other,
    })
}

#[derive(Debug, PartialEq)]
enum CellObserved {
    Empty,
    Text(String),
    Number(f64),
    Logical(bool),
    Error(ScalarError),
    Other,
}

#[derive(Debug, PartialEq)]
enum Observed {
    Empty,
    Text(String),
    Number(f64),
    Logical(bool),
    Error(ScalarError),
    Array {
        rows: usize,
        columns: usize,
        cells: Vec<CellObserved>,
    },
    Other,
}

fn observe_cell(value: Option<Value<'_>>) -> CellObserved {
    match value {
        Some(Value::Empty) => CellObserved::Empty,
        Some(Value::Text(value)) => CellObserved::Text(value.to_owned()),
        Some(Value::Number(value)) => CellObserved::Number(value),
        Some(Value::Logical(value)) => CellObserved::Logical(value),
        Some(Value::Error(error)) => CellObserved::Error(error),
        Some(_) | None => CellObserved::Other,
    }
}

fn observe(result: &value::Evaluated<'_>) -> Observed {
    match result.value() {
        Value::Empty => Observed::Empty,
        Value::Text(value) => Observed::Text((*value).to_owned()),
        Value::Number(value) => Observed::Number(value),
        Value::Logical(value) => Observed::Logical(value),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) => Observed::Array {
            rows: array.shape().rows(),
            columns: array.shape().columns(),
            cells: (0..array.len())
                .map(|index| observe_cell(array.get(index)))
                .collect(),
        },
        Value::Complex(_) | Value::Reference(_) | Value::ReferenceList(_) => Observed::Other,
        _ => Observed::Other,
    }
}

fn value(
    source: &str,
    resolver: &ByteResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(observe(&result))
}

fn assert_scalar_text(source: &str, expected: &str, execution: &ExecutionContext) {
    assert_eq!(
        scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}")),
        ScalarObserved::Text(expected.to_owned()),
        "{source}"
    );
}

fn assert_scalar_number(source: &str, expected: f64, execution: &ExecutionContext) {
    assert_eq!(
        scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}")),
        ScalarObserved::Number(expected),
        "{source}"
    );
}

fn assert_scalar_error(source: &str, expected: ScalarError, execution: &ExecutionContext) {
    assert_eq!(
        scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}")),
        ScalarObserved::Error(expected),
        "{source}"
    );
}

fn assert_value(
    source: &str,
    expected: Observed,
    resolver: &ByteResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    assert_eq!(
        value(source, resolver, execution, mode, &Limits::default())
            .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
        expected,
        "{source} in {mode:?}"
    );
}

#[test]
fn scalar_byte_functions_follow_the_utf8_profile() {
    let (_budget, _cancellation, execution) = execution("ods-formula-byte-text-scalar");
    for (source, expected) in [
        (r#"=LENB("Aé界🙂")"#, 10.0),
        (r#"=FINDB("界";"Aé界🙂")"#, 4.0),
        (r#"=FINDB("";"é";3)"#, 3.0),
        (r#"=SEARCHB("ss";"aß")"#, 2.0),
        (r#"=SEARCHB("NEEDLE";"aNeedle")"#, 2.0),
    ] {
        assert_scalar_number(source, expected, &execution);
    }
    for (source, expected) in [
        (r#"=LEFTB("Aé界🙂";3)"#, "Aé"),
        (r#"=RIGHTB("Aé界🙂";1)"#, ""),
        (r#"=MIDB("Aé界";3;2)"#, "é"),
        (r#"=REPLACEB("Aé界";3;2;"X")"#, "AX界"),
        (r#"=LEFTB("é")"#, ""),
        (r#"=RIGHTB("🙂")"#, ""),
        (r#"=REPLACEB("Aé";99;0;"X")"#, "AéX"),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
}

#[test]
fn scalar_byte_arguments_conversion_and_errors_are_typed() {
    let (_budget, _cancellation, execution) = execution("ods-formula-byte-text-arguments");
    assert_scalar_text(r#"=LEFTB("Aé";TRUE())"#, "A", &execution);
    assert_scalar_text(r#"=LEFTB("Aé";FALSE())"#, "", &execution);
    assert_scalar_text(r#"=LEFTB("Aé";"2.9")"#, "A", &execution);
    assert_scalar_text(r#"=LEFTB("x";)"#, "x", &execution);
    assert_scalar_text(r#"=RIGHTB("x";)"#, "x", &execution);
    assert_scalar_number(r#"=FINDB("é";"Aé";"2.9")"#, 2.0, &execution);
    assert_scalar_number(r#"=FINDB("é";"Aé";TRUE())"#, 2.0, &execution);
    assert_scalar_number(r#"=FINDB("x";"x";)"#, 1.0, &execution);
    assert_scalar_number(r#"=SEARCHB("x";"x";)"#, 1.0, &execution);
    assert_scalar_text(r#"=MIDB("Aé";TRUE();TRUE())"#, "A", &execution);
    assert_scalar_number(r#"=LENB(12.5)"#, 4.0, &execution);
    for (source, expected) in [
        (r#"=LEFTB("x";0)"#, ""),
        (r#"=LEFTB("Aé";99)"#, "Aé"),
        (r#"=RIGHTB("Aé";99)"#, "Aé"),
        (r#"=MIDB("Aé";99;1)"#, ""),
        (r#"=REPLACEB("Aé";3;0;"X")"#, "AXé"),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
    for (source, expected) in [
        (r#"=LEFTB("Aé";1e308)"#, "Aé"),
        (r#"=RIGHTB("Aé";1e308)"#, "Aé"),
        (r#"=MIDB("Aé";1e308;1e308)"#, ""),
        (r#"=REPLACEB("Aé";1e308;1e308;"X")"#, "AéX"),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
    for source in [r#"=FINDB("é";"Aé";1e308)"#, r#"=SEARCHB("é";"Aé";1e308)"#] {
        assert_scalar_error(source, ScalarError::Value, &execution);
    }
    for (source, expected) in [
        (r#"=LEFTB("x";-1)"#, ScalarError::Value),
        (r#"=RIGHTB("x";"bad")"#, ScalarError::Value),
        (r#"=MIDB("x";0;1)"#, ScalarError::Value),
        (r#"=FINDB("x";"x";0)"#, ScalarError::Value),
        (r#"=SEARCHB("x";"y")"#, ScalarError::Value),
        (r#"=LEFTB("x";1e999)"#, ScalarError::Number),
        (r#"=MIDB("x";)"#, ScalarError::Value),
        (r#"=LENB()"#, ScalarError::Value),
        (r#"=FINDB(;"x")"#, ScalarError::Value),
        (r#"=FINDB("x";)"#, ScalarError::Value),
        (r#"=SEARCHB(;"x")"#, ScalarError::Value),
        (r#"=SEARCHB("x";)"#, ScalarError::Value),
        (r#"=MIDB("x";;1)"#, ScalarError::Value),
        (r#"=MIDB("x";1;)"#, ScalarError::Value),
        (r#"=REPLACEB("x";;1;"y")"#, ScalarError::Value),
        (r#"=REPLACEB("x";1;;"y")"#, ScalarError::Value),
        (r#"=REPLACEB("x";1;1;)"#, ScalarError::Value),
        (r#"=FINDB(COMPLEX(1;2);"x")"#, ScalarError::Value),
        (r#"=FINDB("x";COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=FINDB("x";"x";COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=LEFTB(COMPLEX(1;2);1)"#, ScalarError::Value),
        (r#"=LEFTB("x";COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=LENB(COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=MIDB(COMPLEX(1;2);1;1)"#, ScalarError::Value),
        (r#"=MIDB("x";COMPLEX(1;2);1)"#, ScalarError::Value),
        (r#"=MIDB("x";1;COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=REPLACEB(COMPLEX(1;2);1;1;"x")"#, ScalarError::Value),
        (r#"=REPLACEB("x";COMPLEX(1;2);1;"x")"#, ScalarError::Value),
        (r#"=REPLACEB("x";1;COMPLEX(1;2);"x")"#, ScalarError::Value),
        (r#"=REPLACEB("x";1;1;COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=RIGHTB("x";COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=SEARCHB(COMPLEX(1;2);"x")"#, ScalarError::Value),
        (r#"=SEARCHB("x";COMPLEX(1;2))"#, ScalarError::Value),
        (r#"=SEARCHB("x";"x";COMPLEX(1;2))"#, ScalarError::Value),
    ] {
        assert_scalar_error(source, expected, &execution);
    }
}

#[test]
fn numeric_text_overflow_stays_num_and_malformed_text_stays_value() {
    let resolver = ByteResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-byte-text-numeric-text");
    let cases = [
        (r#"=FINDB("x";"x";"1e999")"#, ScalarError::Number),
        (r#"=FINDB("x";"x";"-1e999")"#, ScalarError::Number),
        (r#"=SEARCHB("x";"x";"1e999")"#, ScalarError::Number),
        (r#"=SEARCHB("x";"x";"-1e999")"#, ScalarError::Number),
        (r#"=LEFTB("x";"1e999")"#, ScalarError::Number),
        (r#"=LEFTB("x";"-1e999")"#, ScalarError::Number),
        (r#"=RIGHTB("x";"1e999")"#, ScalarError::Number),
        (r#"=RIGHTB("x";"-1e999")"#, ScalarError::Number),
        (r#"=MIDB("x";"1e999";1)"#, ScalarError::Number),
        (r#"=MIDB("x";"-1e999";1)"#, ScalarError::Number),
        (r#"=MIDB("x";1;"1e999")"#, ScalarError::Number),
        (r#"=MIDB("x";1;"-1e999")"#, ScalarError::Number),
        (r#"=REPLACEB("x";"1e999";1;"y")"#, ScalarError::Number),
        (r#"=REPLACEB("x";"-1e999";1;"y")"#, ScalarError::Number),
        (r#"=REPLACEB("x";1;"1e999";"y")"#, ScalarError::Number),
        (r#"=REPLACEB("x";1;"-1e999";"y")"#, ScalarError::Number),
        (r#"=FINDB("x";"x";"bad")"#, ScalarError::Value),
        (r#"=SEARCHB("x";"x";"bad")"#, ScalarError::Value),
        (r#"=LEFTB("x";"bad")"#, ScalarError::Value),
        (r#"=RIGHTB("x";"bad")"#, ScalarError::Value),
        (r#"=MIDB("x";"bad";1)"#, ScalarError::Value),
        (r#"=MIDB("x";1;"bad")"#, ScalarError::Value),
        (r#"=REPLACEB("x";"bad";1;"y")"#, ScalarError::Value),
        (r#"=REPLACEB("x";1;"bad";"y")"#, ScalarError::Value),
    ];
    for (source, expected) in cases {
        assert_scalar_error(source, expected, &execution);
        assert_value(
            source,
            Observed::Error(expected),
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }
}

#[test]
fn value_byte_functions_lift_scalar_results_and_preserve_utf8_boundaries() {
    let resolver = ByteResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-byte-text-value");
    for (source, expected) in [
        ("=LENB([.A1])", Observed::Number(10.0)),
        ("=FINDB(\"界\";[.A1])", Observed::Number(4.0)),
        ("=SEARCHB(\"ss\";[.A2])", Observed::Number(2.0)),
        ("=LEFTB([.A1];3)", Observed::Text("Aé".to_owned())),
        ("=RIGHTB([.A1];1)", Observed::Text(String::new())),
        ("=MIDB([.A1];3;2)", Observed::Text("é".to_owned())),
        (
            "=REPLACEB([.A1];3;2;\"X\")",
            Observed::Text("AX界🙂".to_owned()),
        ),
    ] {
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
    }
    for (source, expected) in [
        ("=LENB([.B4])", Observed::Number(0.0)),
        ("=LEFTB(\"Aé\";[.B4])", Observed::Text(String::new())),
        (
            "=REPLACEB(\"Aé\";1;1;[.B4])",
            Observed::Text("é".to_owned()),
        ),
    ] {
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
    }
    for source in [
        "=FINDB(COMPLEX(1;2);\"x\")",
        "=FINDB(\"x\";COMPLEX(1;2))",
        "=FINDB(\"x\";\"x\";COMPLEX(1;2))",
        "=LEFTB(COMPLEX(1;2);1)",
        "=LEFTB(\"x\";COMPLEX(1;2))",
        "=LENB(COMPLEX(1;2))",
        "=MIDB(COMPLEX(1;2);1;1)",
        "=MIDB(\"x\";COMPLEX(1;2);1)",
        "=MIDB(\"x\";1;COMPLEX(1;2))",
        "=REPLACEB(COMPLEX(1;2);1;1;\"x\")",
        "=REPLACEB(\"x\";COMPLEX(1;2);1;\"x\")",
        "=REPLACEB(\"x\";1;COMPLEX(1;2);\"x\")",
        "=REPLACEB(\"x\";1;1;COMPLEX(1;2))",
        "=RIGHTB(\"x\";COMPLEX(1;2))",
        "=SEARCHB(COMPLEX(1;2);\"x\")",
        "=SEARCHB(\"x\";COMPLEX(1;2))",
        "=SEARCHB(\"x\";\"x\";COMPLEX(1;2))",
    ] {
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }
    assert_value(
        "=LENB({\"A\";\"é\";\"界\"})",
        Observed::Array {
            rows: 1,
            columns: 3,
            cells: vec![
                CellObserved::Number(1.0),
                CellObserved::Number(2.0),
                CellObserved::Number(3.0),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=LEFTB({\"A\";\"é\"};1)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("A".to_owned()),
                CellObserved::Text(String::new()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=FINDB({\"é\";\"界\"};\"Aé界\")",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(2.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=SEARCHB({\"SS\";\"NEEDLE\"};{\"aß\";\"xNEEDLE\"})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(2.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=RIGHTB({\"Aé\";\"界\"};1)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text(String::new()),
                CellObserved::Text(String::new()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=MIDB({\"Aé\";\"界\"};2;2)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("é".to_owned()),
                CellObserved::Text(String::new()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=REPLACEB({\"Aé\";\"界\"};2;2;\"X\")",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("AX".to_owned()),
                CellObserved::Text("X界".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=LEFTB(\"Aé\";{1;2})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("A".to_owned()),
                CellObserved::Text("A".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=MIDB({\"Aé\";\"Aé\"};{1;2};{1;2})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("A".to_owned()),
                CellObserved::Text("é".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=REPLACEB({\"Aé\";\"Aé\"};{1;2};{1;2};\"X\")",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("Xé".to_owned()),
                CellObserved::Text("AX".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=REPLACEB({\"Aé\";\"Aé\"};{1;2};{1;2};{\"X\";\"YZ\"})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![
                CellObserved::Text("Xé".to_owned()),
                CellObserved::Text("AYZ".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=FINDB({\"A\";\"é\"};\"Aé\";{1;2})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    resolver.clear_reads();
    assert_value(
        "=IF(TRUE();MIDB([.A1:.A1];3;2);MIDB([Missing.A1];1;1))",
        Observed::Array {
            rows: 1,
            columns: 1,
            cells: vec![CellObserved::Text("é".to_owned())],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 1, "selected byte references only");

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};LEFTB(\"Aé界🙂\";MUNIT(2));LEFTB([Missing.A1];1))",
        Observed::Array {
            rows: 2,
            columns: 2,
            cells: vec![
                CellObserved::Text("A".to_owned()),
                CellObserved::Text(String::new()),
                CellObserved::Text(String::new()),
                CellObserved::Text("A".to_owned()),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "lazy byte branch skips Missing reference"
    );
}

#[test]
fn value_byte_shapes_and_error_precedence_are_explicit() {
    let resolver = ByteResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-byte-text-shapes");
    let error = value(
        "=LENB([.A1]~[.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a scalar ReferenceList refusal is a formula error");
    assert_eq!(error, Observed::Error(ScalarError::Value));
    assert_eq!(resolver.reads(), 0);

    let error = value(
        "=LENB([Main.A1:Archive.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("3-D reference is a typed capability refusal");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);

    resolver.clear_reads();
    let error = value(
        "=LENB([.A4:.A5])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("provider failure supersedes a retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);
}

#[test]
fn computed_text_reducers_preserve_each_projected_reference_coordinate() {
    let resolver = ByteResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-byte-projected-reducer");
    for (expression, expected, reads) in [
        ("AVERAGE(LENB([.A1:.A2]))", [10.0, 3.0], 2),
        // SUM evaluates its computed argument in matrix context.
        ("SUM(LENB([.A1:.A2]))", [13.0, 13.0], 4),
        ("AVERAGE(LEN([.A1:.A2]))", [4.0, 2.0], 2),
    ] {
        resolver.clear_reads();
        let source = format!("=IF({{TRUE()|TRUE()}};{expression};0)");
        assert_value(
            &source,
            Observed::Array {
                rows: 2,
                columns: 1,
                cells: expected.into_iter().map(CellObserved::Number).collect(),
            },
            &resolver,
            &execution,
            Mode::Matrix,
        );
        assert_eq!(resolver.reads(), reads, "{source}: argument-context reads");
    }
}
