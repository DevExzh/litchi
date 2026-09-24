//! Semantic coverage for the ODF 1.4 value-inspection and conversion family.
//!
//! The fixture keeps physically empty cells, empty Text, typed formula errors,
//! logicals, numbers, and borrowed text distinct.  This is deliberate: the
//! inspection functions must see those distinctions before ordinary scalar
//! coercion can turn an empty cell into zero.

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
enum InspectionCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct InspectionResolver {
    rows: usize,
    columns: usize,
    cells: Vec<InspectionCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    archive_a1_unsupported: Cell<bool>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl InspectionResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 16,
            columns: 12,
            cells: (0..192).map(|_| InspectionCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            archive_a1_unsupported: Cell::new(false),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };

        // Main.A1 is physically empty; Main.A2 contains empty Text.  Keeping
        // these in adjacent cells makes the ISBLANK/ISTEXT distinction visible
        // through both scalar and matrix APIs.
        resolver.set(0, 0, InspectionCell::Empty);
        resolver.set(1, 0, InspectionCell::Text(String::new()));
        resolver.set(2, 0, InspectionCell::Number(42.5));
        resolver.set(3, 0, InspectionCell::Logical(true));
        resolver.set(4, 0, InspectionCell::Text("1234.5".to_owned()));
        resolver.set(5, 0, InspectionCell::Error(ScalarError::NotAvailable));
        resolver.set(6, 0, InspectionCell::Error(ScalarError::DivisionByZero));
        resolver.set(7, 0, InspectionCell::Error(ScalarError::Value));

        resolver.set(0, 1, InspectionCell::Text("1,234.50".to_owned()));
        resolver.set(1, 1, InspectionCell::Text("12,5".to_owned()));
        resolver.set(2, 1, InspectionCell::Text("2006-05-21".to_owned()));
        resolver.set(3, 1, InspectionCell::Text("2:00".to_owned()));
        resolver.set(4, 1, InspectionCell::Text("2 1/2".to_owned()));
        resolver.set(5, 1, InspectionCell::Text("not numeric".to_owned()));

        resolver.set(0, 2, InspectionCell::Number(2.9));
        resolver.set(1, 2, InspectionCell::Number(-3.9));
        resolver.set(0, 6, InspectionCell::Number(1.0));
        resolver.set(1, 6, InspectionCell::Number(2.0));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: InspectionCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
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

    fn set_archive_a1_unsupported(&self) {
        self.archive_a1_unsupported.set(true);
    }
}

impl Resolver for InspectionResolver {
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
            match sheet {
                "Data" => return Ok(CellRead::Number(7.0)),
                "Archive" => {
                    return Ok(if self.archive_a1_unsupported.get() {
                        CellRead::Unsupported
                    } else {
                        CellRead::Text("archive")
                    });
                },
                _ => {},
            }
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            InspectionCell::Empty => CellRead::Empty,
            InspectionCell::Number(value) => CellRead::Number(*value),
            InspectionCell::Logical(value) => CellRead::Logical(*value),
            InspectionCell::Text(value) => CellRead::Text(value.as_str()),
            InspectionCell::Error(error) => CellRead::Error(*error),
            InspectionCell::Unsupported => CellRead::Unsupported,
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
    Number(f64),
    Logical(bool),
    Text(String),
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
        ScalarValue::Number(value) => ScalarObserved::Number(*value),
        ScalarValue::Logical(value) => ScalarObserved::Logical(*value),
        ScalarValue::Text(value) => ScalarObserved::Text(value.to_string()),
        ScalarValue::Error(error) => ScalarObserved::Error(*error),
        ScalarValue::Complex(_) => ScalarObserved::Other,
        _ => ScalarObserved::Other,
    })
}

#[derive(Debug, PartialEq)]
enum CellObserved {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Other,
}

#[derive(Debug, PartialEq)]
enum Observed {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
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
        Some(Value::Number(value)) => CellObserved::Number(value),
        Some(Value::Logical(value)) => CellObserved::Logical(value),
        Some(Value::Text(value)) => CellObserved::Text(value.to_owned()),
        Some(Value::Error(error)) => CellObserved::Error(error),
        Some(_) | None => CellObserved::Other,
    }
}

fn observe(result: &value::Evaluated<'_>) -> Observed {
    match result.value() {
        Value::Empty => Observed::Empty,
        Value::Number(value) => Observed::Number(value),
        Value::Logical(value) => Observed::Logical(value),
        Value::Text(value) => Observed::Text((*value).to_owned()),
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
    resolver: &InspectionResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    value_at(
        source,
        resolver,
        execution,
        Position::new("Main", 0, 0),
        mode,
        limits,
    )
}

fn value_at(
    source: &str,
    resolver: &InspectionResolver,
    execution: &ExecutionContext,
    position: Position<'_>,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, position).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(observe(&result))
}

fn assert_scalar_number(source: &str, expected: f64, execution: &ExecutionContext) {
    assert_eq!(
        scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}")),
        ScalarObserved::Number(expected),
        "{source}"
    );
}

fn assert_scalar_logical(source: &str, expected: bool, execution: &ExecutionContext) {
    assert_eq!(
        scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}")),
        ScalarObserved::Logical(expected),
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
    resolver: &InspectionResolver,
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
fn scalar_inspection_functions_preserve_types_and_formula_errors() {
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-scalar-types");

    for (source, expected) in [
        ("=ERROR.TYPE(#NULL!)", 1.0),
        ("=ERROR.TYPE(#DIV/0!)", 2.0),
        ("=ERROR.TYPE(#VALUE!)", 3.0),
        ("=ERROR.TYPE(#REF!)", 4.0),
        ("=ERROR.TYPE(#NAME?)", 5.0),
        ("=ERROR.TYPE(#NUM!)", 6.0),
        ("=ERROR.TYPE(#N/A)", 7.0),
        ("=TYPE(42.5)", 1.0),
        ("=TYPE(\"text\")", 2.0),
        ("=TYPE(TRUE())", 4.0),
        ("=TYPE(#N/A)", 16.0),
    ] {
        assert_scalar_number(source, expected, &execution);
    }
    for (source, expected) in [
        ("=ISBLANK(\"\")", false),
        ("=ISERR(#N/A)", false),
        ("=ISERR(#DIV/0!)", true),
        ("=ISERROR(#N/A)", true),
        ("=ISERROR(1)", false),
        ("=ISLOGICAL(TRUE())", true),
        ("=ISLOGICAL(1)", false),
        ("=ISNA(#N/A)", true),
        ("=ISNA(#DIV/0!)", false),
        ("=ISNONTEXT(\"text\")", false),
        ("=ISNONTEXT(1)", true),
        ("=ISNUMBER(1)", true),
        ("=ISNUMBER(TRUE())", false),
        ("=ISTEXT(\"\")", true),
        ("=ISTEXT(1)", false),
        ("=ISEVEN(2.9)", true),
        ("=ISODD(-3.9)", true),
        ("=ISEVEN(TRUE())", false),
        ("=ISODD(TRUE())", true),
        ("=ISEVEN(FALSE())", true),
        ("=ISODD(FALSE())", false),
        ("=ISNUMBER(COMPLEX(1;2))", true),
        ("=ISTEXT(COMPLEX(1;2))", false),
        ("=ISNONTEXT(COMPLEX(1;2))", true),
    ] {
        assert_scalar_logical(source, expected, &execution);
    }

    assert_scalar_number("=N(42.5)", 42.5, &execution);
    assert_scalar_number("=N(TRUE())", 1.0, &execution);
    assert_scalar_number("=N(FALSE())", 0.0, &execution);
    assert_scalar_number("=N(\"text\")", 0.0, &execution);
    assert_scalar_error("=N(#N/A)", ScalarError::NotAvailable, &execution);
    assert_scalar_number("=TYPE(COMPLEX(1;2))", 1.0, &execution);
    assert_eq!(
        scalar("=N(COMPLEX(1;2))", &execution).expect("N complex"),
        ScalarObserved::Other
    );
    for source in [
        "=VALUE(COMPLEX(1;2))",
        "=NUMBERVALUE(COMPLEX(1;2);\".\";\",\")",
    ] {
        assert_scalar_error(source, ScalarError::Value, &execution);
    }
    assert_scalar_error("=NA()", ScalarError::NotAvailable, &execution);

    for source in [
        "=ERROR.TYPE(1)",
        "=ERROR.TYPE()",
        "=ISNUMBER()",
        "=ISEVEN()",
        "=ISODD()",
        "=N()",
        "=NA(1)",
        "=NUMBERVALUE()",
        "=TYPE()",
        "=VALUE()",
    ] {
        assert_scalar_error(source, ScalarError::Value, &execution);
    }
}

#[test]
fn scalar_converters_follow_numbervalue_and_value_profiles() {
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-conversions");

    for (source, expected) in [
        (r#"=NUMBERVALUE("1,234.5";".";",")"#, 1234.5),
        (r#"=NUMBERVALUE("1,234.5")"#, 1234.5),
        (r#"=NUMBERVALUE("1,234.5";;)"#, 1234.5),
        (r#"=NUMBERVALUE("1,234.5";".";)"#, 1234.5),
        (r#"=NUMBERVALUE("12,5";",";".")"#, 12.5),
        (r#"=NUMBERVALUE(".5";".";",")"#, 0.5),
        (r#"=NUMBERVALUE("50%%";".";",")"#, 0.005),
        (r#"=NUMBERVALUE("1.234,50%";",";".")"#, 12.345),
        (r#"=NUMBERVALUE(" 1 234,5 ";",";" ")"#, 1234.5),
        (r#"=VALUE("123")"#, 123.0),
        (r#"=VALUE("+1.2e2")"#, 120.0),
        (r#"=VALUE("50%")"#, 0.5),
        (r#"=VALUE("$1,234.50")"#, 1234.5),
        (
            r#"=VALUE("1,000,000,000,000,000,128")"#,
            1.0000000000000001e18,
        ),
        (r#"=VALUE("1e3")"#, 1000.0),
        (r#"=VALUE("2 1/2")"#, 2.5),
        (r#"=VALUE("2:00")"#, 1.0 / 12.0),
        (r#"=VALUE("2006-05-21")"#, 38_858.0),
        (r#"=VALUE("2006-05-21 12:00")"#, 38_858.5),
        (r#"=VALUE("2006-05-21T12:00:00")"#, 38_858.5),
        (r#"=VALUE("5/21/2006")"#, 38_858.0),
        (r#"=VALUE("5/21/06")"#, 38_858.0),
        (r#"=VALUE("5-21-2006")"#, 38_858.0),
        (r#"=VALUE("Oct 29, 2006")"#, 39_019.0),
        (r#"=VALUE("29 Oct 2006")"#, 39_019.0),
        (r#"=VALUE("October 29, 2006")"#, 39_019.0),
        (r#"=VALUE("29 October 2006")"#, 39_019.0),
        (
            r#"=VALUE("12:34:56.5")"#,
            12.0 / 24.0 + 34.0 / 1440.0 + 56.5 / 86400.0,
        ),
    ] {
        assert_scalar_number(source, expected, &execution);
    }
    let mut long_grouped = String::from("0");
    for _ in 0..400 {
        long_grouped.push_str(",000");
    }
    long_grouped.push_str(",001.000000000000000111022302462515654042363166809082031251");
    let long_grouped_source = format!("=VALUE(\"{long_grouped}\")");
    assert_scalar_number(&long_grouped_source, 1.0000000000000002, &execution);
    for source in [
        r#"=NUMBERVALUE("1";"..";",")"#,
        r#"=NUMBERVALUE("1";"";",")"#,
        r#"=NUMBERVALUE("1,2";",";",")"#,
        r#"=NUMBERVALUE("not numeric";".";",")"#,
        r#"=NUMBERVALUE("1.2,3";".";",")"#,
        r#"=NUMBERVALUE("")"#,
        r#"=VALUE("not numeric")"#,
        r#"=VALUE("")"#,
        r#"=VALUE("2023-02-29")"#,
        r#"=VALUE("24:00")"#,
    ] {
        assert_scalar_error(source, ScalarError::Value, &execution);
    }
    assert_scalar_error(
        r#"=NUMBERVALUE("1e999";".";",")"#,
        ScalarError::Number,
        &execution,
    );
    assert_scalar_error(r#"=VALUE("1e999")"#, ScalarError::Number, &execution);
    assert_scalar_error(r#"=VALUE("1899-12-29")"#, ScalarError::Number, &execution);
    assert_scalar_error(
        r#"=NUMBERVALUE(#N/A;".";",")"#,
        ScalarError::NotAvailable,
        &execution,
    );
    assert_scalar_error(r#"=VALUE(#N/A)"#, ScalarError::NotAvailable, &execution);
}

#[test]
fn value_inspection_distinguishes_empty_text_errors_and_broadcasts() {
    let resolver = InspectionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-value-types");

    for (source, expected) in [
        ("=ISBLANK([.A1])", Observed::Logical(true)),
        ("=ISBLANK([.A2])", Observed::Logical(false)),
        ("=ISBLANK([.A6])", Observed::Logical(false)),
        ("=ISERROR([.A6])", Observed::Logical(true)),
        ("=ISERR([.A7])", Observed::Logical(true)),
        ("=ISNA([.A6])", Observed::Logical(true)),
        ("=ISNUMBER([.A3])", Observed::Logical(true)),
        ("=ISLOGICAL([.A4])", Observed::Logical(true)),
        ("=ISTEXT([.A2])", Observed::Logical(true)),
        ("=ISNONTEXT([.A1])", Observed::Logical(true)),
        ("=ISEVEN([.C1])", Observed::Logical(true)),
        ("=ISODD([.C2])", Observed::Logical(true)),
        ("=TYPE([.A3])", Observed::Number(1.0)),
        ("=TYPE([.A2])", Observed::Number(2.0)),
        ("=TYPE([.A1])", Observed::Number(1.0)),
        ("=TYPE([.A4])", Observed::Number(4.0)),
        ("=TYPE([.A6])", Observed::Number(16.0)),
        ("=N([.A1])", Observed::Number(0.0)),
        ("=N([.A4])", Observed::Number(1.0)),
        ("=N([.A6])", Observed::Error(ScalarError::NotAvailable)),
        ("=NA()", Observed::Error(ScalarError::NotAvailable)),
        ("=NUMBERVALUE([.B2];\",\";\".\")", Observed::Number(12.5)),
        (
            "=NUMBERVALUE([.A1];\".\";\",\")",
            Observed::Error(ScalarError::Value),
        ),
        ("=VALUE([.B1])", Observed::Number(1234.5)),
        ("=VALUE([.B3])", Observed::Number(38_858.0)),
        ("=VALUE([.B4])", Observed::Number(1.0 / 12.0)),
        ("=VALUE([.B5])", Observed::Number(2.5)),
        ("=VALUE([.A1])", Observed::Number(0.0)),
        ("=VALUE([.A2])", Observed::Error(ScalarError::Value)),
    ] {
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
    }
    for (source, expected) in [
        ("=NUMBERVALUE(\"1,234.5\";;)", Observed::Number(1234.5)),
        ("=NUMBERVALUE(\"1,234.5\";\".\";)", Observed::Number(1234.5)),
        (
            "=NUMBERVALUE(\"1,234.5\";\"\";\",\")",
            Observed::Error(ScalarError::Value),
        ),
    ] {
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
    }

    assert_value(
        "=ISNUMBER([.A3:.A4])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(false)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ISBLANK([.A1:.A2])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(false)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ISERROR([.A6:.A7])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(true)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ISNA([.A6:.A7])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(false)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ISEVEN([.C1:.C2])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(false)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        value_at(
            "=N([.A3:.A4])",
            &resolver,
            &execution,
            Position::new("Main", 2, 0),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("N uses the explicit A3 intersection"),
        Observed::Number(42.5)
    );
    assert_value(
        "=VALUE({\"123\"|\"50%\"})",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(123.0), CellObserved::Number(0.5)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=NUMBERVALUE({\"1,2\"|\"3,4\"};\",\";\".\")",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.2), CellObserved::Number(3.4)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    for source in [
        "=NUMBERVALUE([.A1];\"..\";\",\")",
        "=NUMBERVALUE([.A1];\".\";\".\")",
    ] {
        resolver.clear_reads();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} rejects separators before reading its source reference"
        );
    }
    resolver.clear_reads();
    assert_value(
        "=NUMBERVALUE([.A1:.A2];\"..\";\",\")",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "invalid separators refuse a known matrix reference before reads"
    );
    assert_value(
        "=TYPE({1|\"text\"})",
        Observed::Number(64.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=N({42.5|1})",
        Observed::Number(42.5),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=N({0|1})",
        Observed::Number(0.0),
        &resolver,
        &execution,
        Mode::Matrix,
    );

    assert_value(
        "=TYPE(ISNUMBER([.A3:.A4]))",
        Observed::Number(4.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=TYPE(ISNUMBER([.A3:.A4]))",
        Observed::Number(64.0),
        &resolver,
        &execution,
        Mode::Matrix,
    );

    resolver.clear_reads();
    assert_value(
        "=TYPE([.A3:.A4])",
        Observed::Number(64.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "TYPE scans an admitted rectangular reference"
    );

    resolver.clear_reads();
    assert_value(
        "=TYPE([.A6:.A7])",
        Observed::Number(64.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "TYPE scans referenced formula errors before returning Array"
    );

    resolver.clear_reads();
    assert_value(
        "=TYPE([.A1]~[.A2])",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "TYPE rejects a known ReferenceList before reads"
    );

    resolver.clear_reads();
    assert_value(
        "=N([.A1]~[.A2])",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "N rejects a known ReferenceList before reads"
    );
}

#[test]
fn three_dimensional_inspection_selects_planes_by_consumer_contract() {
    let resolver = InspectionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-three-dimensional");

    resolver.clear_reads();
    assert_value(
        "=TYPE([Main.A1:Archive.A1])",
        Observed::Number(64.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        3,
        "TYPE scans every plane of an admitted 3-D reference"
    );

    resolver.clear_reads();
    assert_value(
        "=N([Main.A1:Archive.A1])",
        Observed::Number(0.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "N uses the current Main plane rather than flattening a 3-D reference"
    );

    resolver.clear_reads();
    assert_eq!(
        value_at(
            "=N([Main.A1:Archive.A1])",
            &resolver,
            &execution,
            Position::new("Data", 0, 0),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("Data-plane N"),
        Observed::Number(7.0)
    );
    assert_eq!(resolver.reads(), 1, "N intersects the caller's Data plane");

    resolver.clear_reads();
    assert_eq!(
        value_at(
            "=N([Main.A1:Archive.A1])",
            &resolver,
            &execution,
            Position::new("Outside", 0, 0),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("outside-plane N formula value"),
        Observed::Error(ScalarError::NotAvailable)
    );
    assert_eq!(resolver.reads(), 0, "no current plane means no N cell read");

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()};ISNUMBER([Main.A1:Archive.A3]);FALSE())",
        Observed::Array {
            rows: 3,
            columns: 1,
            cells: vec![
                CellObserved::Logical(false),
                CellObserved::Logical(false),
                CellObserved::Logical(true),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        3,
        "a projected 3-D mapper widens to the current plane's 2-D shape"
    );

    let typed_resolver = InspectionResolver::standard();
    typed_resolver.set_archive_a1_unsupported();
    typed_resolver.clear_reads();
    let error = value(
        "=TYPE([Main.A1:Archive.A1])",
        &typed_resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("TYPE must expose typed failures from later 3-D planes");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(typed_resolver.reads(), 3);
}

#[test]
fn projected_inspection_branches_are_lazy_and_position_sensitive() {
    let resolver = InspectionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-projected");

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};ISNUMBER([.A3:.A4]);FALSE())",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(false)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "each projected inspection coordinate reads once"
    );

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};TYPE([.A1:.A3]);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(64.0), CellObserved::Number(64.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        3,
        "projected TYPE reuses the complete rectangular reference"
    );

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};TYPE(IF(TRUE();{1|2};{3|4}));0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(64.0), CellObserved::Number(64.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "TYPE classifies a computed array without resolver reads"
    );
    assert_value(
        "=IF({TRUE()|TRUE()};N(IF(TRUE();{1|2};{3|4}));0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(1.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};TYPE(MUNIT(N([.G1:.G2])));0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(64.0), CellObserved::Number(64.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "nested TYPE keeps MUNIT's position-sensitive scalar parameter"
    );

    let mut typed_resolver = InspectionResolver::standard();
    typed_resolver.set(0, 6, InspectionCell::Unsupported);
    typed_resolver.clear_reads();
    let error = value(
        "=IF({TRUE()|TRUE()};TYPE(MUNIT(N([.G1:.G2])));0)",
        &typed_resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("nested TYPE must preserve a typed MUNIT parameter failure");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(typed_resolver.reads(), 1);

    resolver.clear_reads();
    assert_value(
        "=IF(FALSE();ISNUMBER([Missing.A1]);TRUE())",
        Observed::Logical(true),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "unselected inspection branch stays lazy"
    );

    resolver.clear_reads();
    assert_value(
        "=IF({TRUE()|TRUE()};ISNUMBER(MUNIT(2));FALSE())",
        Observed::Array {
            rows: 2,
            columns: 2,
            cells: vec![
                CellObserved::Logical(true),
                CellObserved::Logical(true),
                CellObserved::Logical(true),
                CellObserved::Logical(true),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0, "MUNIT branch has no resolver reads");

    assert_value(
        "=IF({TRUE()|TRUE()};N({42.5|1});0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(42.5), CellObserved::Number(42.5)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=IF({TRUE()|TRUE()};N({0|1});0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(0.0), CellObserved::Number(0.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
}

#[test]
fn value_formula_errors_are_values_but_provider_failures_remain_typed() {
    let mut resolver = InspectionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-inspection-errors");

    assert_value(
        "=ERROR.TYPE([.A6])",
        Observed::Number(7.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=ISERROR([.A8])",
        Observed::Logical(true),
        &resolver,
        &execution,
        Mode::Scalar,
    );

    resolver.set(0, 0, InspectionCell::Error(ScalarError::NotAvailable));
    resolver.set(1, 0, InspectionCell::Unsupported);
    resolver.clear_reads();
    let error = value(
        "=ISNUMBER([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("typed provider failure supersedes a retained formula error");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);

    resolver.clear_reads();
    let error = value(
        "=TYPE([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("TYPE must continue its scan until a typed failure is visible");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 2);
}
