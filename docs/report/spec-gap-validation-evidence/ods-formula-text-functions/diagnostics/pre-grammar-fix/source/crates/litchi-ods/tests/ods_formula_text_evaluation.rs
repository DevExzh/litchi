//! Independent semantic coverage for the OpenFormula 1.4 §6.20 text family.
//!
//! The resolver deliberately exposes borrowed text, empty cells, formula
//! errors, 3-D planes, and ordered reference lists.  The tests exercise both
//! the resolver-free scalar façade and the value VM so that Unicode scalar
//! indexing, text conversion, matrix lifting, implicit intersection, and
//! lazy branch behavior remain visible at the public API boundary.

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
enum TextCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct TextResolver {
    rows: usize,
    columns: usize,
    cells: Vec<TextCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl TextResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 8,
            columns: 12,
            cells: (0..96).map(|_| TextCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };

        for (row, value) in [
            TextCell::Text("Alpha".to_owned()),
            TextCell::Text("é🌟".to_owned()),
            TextCell::Text("  a\t b \n".to_owned()),
            TextCell::Text("ＡＢＣ".to_owned()),
            TextCell::Text("ｶﾞ".to_owned()),
            TextCell::Text("needle".to_owned()),
            TextCell::Empty,
            TextCell::Error(ScalarError::NotAvailable),
        ]
        .into_iter()
        .enumerate()
        {
            resolver.set(row, 0, value);
        }
        for (row, value) in [
            TextCell::Text("needle".to_owned()),
            TextCell::Text("NEEDLE".to_owned()),
            TextCell::Text("xneedle".to_owned()),
            TextCell::Text("foo".to_owned()),
        ]
        .into_iter()
        .enumerate()
        {
            resolver.set(row, 1, value);
        }
        resolver.set(0, 2, TextCell::Number(12.5));
        resolver.set(1, 2, TextCell::Logical(true));
        resolver.set(2, 2, TextCell::Empty);
        resolver.set(3, 2, TextCell::Error(ScalarError::NotAvailable));
        resolver.set(0, 6, TextCell::Number(1.0));
        resolver.set(1, 6, TextCell::Number(2.0));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: TextCell) {
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

    fn set_source_versions(&mut self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions = Some((expected, observed));
        self.source_version_calls.set(0);
    }
}

impl Resolver for TextResolver {
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

        // Give each 3-D plane a stable, observable text value.
        if row == 0 && column == 0 {
            return Ok(match sheet {
                "Main" => CellRead::Text("Alpha"),
                "Data" => CellRead::Text("data"),
                "Archive" => CellRead::Text("archive"),
                _ => unreachable!(),
            });
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            TextCell::Empty => CellRead::Empty,
            TextCell::Number(value) => CellRead::Number(*value),
            TextCell::Logical(value) => CellRead::Logical(*value),
            TextCell::Text(value) => CellRead::Text(value.as_str()),
            TextCell::Error(error) => CellRead::Error(*error),
            TextCell::Unsupported => CellRead::Unsupported,
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
enum ValueObserved {
    Empty,
    Text(String),
    Number(f64),
    Logical(bool),
    Error(ScalarError),
    Array {
        rows: usize,
        columns: usize,
        cells: Vec<ValueObserved>,
    },
    Other,
}

fn observe_value(result: &value::Evaluated<'_>) -> ValueObserved {
    match result.value() {
        Value::Empty => ValueObserved::Empty,
        Value::Text(value) => ValueObserved::Text((*value).to_owned()),
        Value::Number(value) => ValueObserved::Number(value),
        Value::Logical(value) => ValueObserved::Logical(value),
        Value::Error(error) => ValueObserved::Error(error),
        Value::Array(array) => {
            let cells = (0..array.len())
                .map(|index| match array.get(index) {
                    Some(Value::Empty) => ValueObserved::Empty,
                    Some(Value::Text(value)) => ValueObserved::Text(value.to_owned()),
                    Some(Value::Number(value)) => ValueObserved::Number(value),
                    Some(Value::Logical(value)) => ValueObserved::Logical(value),
                    Some(Value::Error(error)) => ValueObserved::Error(error),
                    Some(Value::Array(_))
                    | Some(Value::Reference(_))
                    | Some(Value::ReferenceList(_))
                    | Some(Value::Complex(_))
                    | None => ValueObserved::Other,
                    Some(_) => ValueObserved::Other,
                })
                .collect();
            ValueObserved::Array {
                rows: array.shape().rows(),
                columns: array.shape().columns(),
                cells,
            }
        },
        Value::Complex(_) | Value::Reference(_) | Value::ReferenceList(_) => ValueObserved::Other,
        _ => ValueObserved::Other,
    }
}

fn value(
    source: &str,
    resolver: &TextResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<ValueObserved, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(observe_value(&result))
}

fn assert_scalar_text(source: &str, expected: &str, execution: &ExecutionContext) {
    let result = scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}"));
    match result {
        ScalarObserved::Text(actual) => assert_eq!(actual, expected, "{source}"),
        other => panic!("{source}: expected text {expected:?}, got {other:?}"),
    }
}

fn assert_scalar_number(source: &str, expected: f64, execution: &ExecutionContext) {
    let result = scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}"));
    match result {
        ScalarObserved::Number(actual) => assert_eq!(actual, expected, "{source}"),
        other => panic!("{source}: expected number {expected}, got {other:?}"),
    }
}

fn assert_scalar_logical(source: &str, expected: bool, execution: &ExecutionContext) {
    let result = scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}"));
    match result {
        ScalarObserved::Logical(actual) => assert_eq!(actual, expected, "{source}"),
        other => panic!("{source}: expected logical {expected}, got {other:?}"),
    }
}

fn assert_scalar_error(source: &str, expected: ScalarError, execution: &ExecutionContext) {
    let result = scalar(source, execution).unwrap_or_else(|error| panic!("{source}: {error}"));
    assert_eq!(result, ScalarObserved::Error(expected), "{source}");
}

fn assert_value_text(
    source: &str,
    expected: &str,
    resolver: &TextResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    let result = value(source, resolver, execution, mode, &Limits::default())
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
    match result {
        ValueObserved::Text(actual) => assert_eq!(actual, expected, "{source} in {mode:?}"),
        other => panic!("{source} in {mode:?}: expected text {expected:?}, got {other:?}"),
    }
}

fn assert_value_number(
    source: &str,
    expected: f64,
    resolver: &TextResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    let result = value(source, resolver, execution, mode, &Limits::default())
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
    match result {
        ValueObserved::Number(actual) => assert_eq!(actual, expected, "{source} in {mode:?}"),
        other => panic!("{source} in {mode:?}: expected number {expected}, got {other:?}"),
    }
}

fn assert_value_error(
    source: &str,
    expected: ScalarError,
    resolver: &TextResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    let result = value(source, resolver, execution, mode, &Limits::default())
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
    assert_eq!(
        result,
        ValueObserved::Error(expected),
        "{source} in {mode:?}"
    );
}

fn assert_value_text_array(
    source: &str,
    rows: usize,
    columns: usize,
    expected: &[&str],
    resolver: &TextResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    let result = value(source, resolver, execution, mode, &Limits::default())
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}"));
    let ValueObserved::Array {
        rows: actual_rows,
        columns: actual_columns,
        cells,
    } = result
    else {
        panic!("{source}: expected an array result");
    };
    assert_eq!((actual_rows, actual_columns), (rows, columns));
    assert_eq!(cells.len(), expected.len());
    for (index, expected) in expected.iter().copied().enumerate() {
        match &cells[index] {
            ValueObserved::Text(actual) => assert_eq!(actual, expected, "{source}[{index}]"),
            other => panic!("{source}[{index}]: expected text {expected:?}, got {other:?}"),
        }
    }
}

#[test]
fn scalar_text_family_covers_unicode_width_codepoints_and_character_slicing() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-unicode");

    for (source, expected) in [
        (r#"=ASC("ＡＢＣ")"#, "ABC"),
        (r#"=ASC("ガ")"#, "ｶﾞ"),
        (r#"=ASC("￥，。")"#, "\\,｡"),
        (r#"=JIS("ABC")"#, "ＡＢＣ"),
        (r#"=JIS("ｶﾞ")"#, "ガ"),
        (r#"=JIS("\,｡")"#, "￥，。"),
        (r#"=CHAR(65)"#, "A"),
        (r#"=UNICHAR(128512)"#, "😀"),
        (r#"=LEFT("é🌟";1)"#, "é"),
        (r#"=RIGHT("é🌟";1)"#, "🌟"),
        (r#"=MID("é🌟";2;1)"#, "🌟"),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
    for (source, expected) in [
        (r#"=CODE("A")"#, 65.0),
        (r#"=LEN("é🌟")"#, 2.0),
        (r#"=UNICODE("🌟")"#, 127_775.0),
    ] {
        assert_scalar_number(source, expected, &execution);
    }
}

#[test]
fn char_code_round_trip_covers_the_complete_selected_byte_profile() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-char-code-profile");
    let resolver = TextResolver::standard();

    for number in 1_u32..=255 {
        let expected = char::from_u32(number)
            .expect("the selected CHAR profile is Unicode-scalar compatible")
            .to_string();
        let char_source = format!("=CHAR({number})");
        let code_source = format!("=CODE(CHAR({number}))");
        assert_scalar_text(&char_source, &expected, &execution);
        assert_scalar_number(&code_source, f64::from(number), &execution);
        for mode in [Mode::Scalar, Mode::Matrix] {
            assert_value_text(&char_source, &expected, &resolver, &execution, mode);
            assert_value_number(&code_source, f64::from(number), &resolver, &execution, mode);
        }
    }

    for (source, expected) in [
        (r#"=CHAR(0)"#, ScalarError::Value),
        (r#"=CHAR(256)"#, ScalarError::Value),
        (r#"=CHAR(0.9)"#, ScalarError::Value),
    ] {
        assert_scalar_error(source, expected, &execution);
        for mode in [Mode::Scalar, Mode::Matrix] {
            assert_value_error(source, expected, &resolver, &execution, mode);
        }
    }
    assert_scalar_text(r#"=CHAR(255.9)"#, "ÿ", &execution);
    for mode in [Mode::Scalar, Mode::Matrix] {
        assert_value_text(r#"=CHAR(255.9)"#, "ÿ", &resolver, &execution, mode);
    }
}

#[test]
fn scalar_text_family_covers_case_search_cleanup_replacement_and_conversion() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-core");
    for (source, expected) in [
        (r#"=CLEAN(CONCATENATE("a";CHAR(1);"b"))"#, "ab"),
        (r#"=CONCATENATE("x";12.5)"#, "x12.5"),
        (r#"=LOWER("ÄBC")"#, "äbc"),
        (r#"=UPPER("ébc")"#, "ÉBC"),
        (r#"=PROPER("hello WORLD")"#, "Hello World"),
        (r#"=PROPER("l'école")"#, "L'École"),
        (r#"=REPLACE("abcdef";2;3;"X")"#, "aXef"),
        (r#"=REPLACE("abc";2;0;"X")"#, "aXbc"),
        (r#"=REPT("ab";3)"#, "ababab"),
        (r#"=SUBSTITUTE("a-b-a";"-";"+")"#, "a+b+a"),
        (r#"=SUBSTITUTE("a-b-a";"-";"+";2)"#, "a-b+a"),
        (r#"=TRIM(CONCATENATE("  a";CHAR(10);" b "))"#, "a b"),
        (r#"=TRIM("a	b")"#, "a	b"),
        (r#"=T("text")"#, "text"),
        (r#"=T(42)"#, ""),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
    assert_scalar_number(r#"=FIND("needle";"find needle")"#, 6.0, &execution);
    assert_scalar_number(r#"=SEARCH("NEEDLE";"find needle")"#, 6.0, &execution);
    assert_scalar_number(r#"=SEARCH("ss";"ß")"#, 1.0, &execution);
    assert_scalar_number(r#"=SEARCH("ß";"ss")"#, 1.0, &execution);
    assert_scalar_error(r#"=SEARCH("s";"ß")"#, ScalarError::Value, &execution);
    assert_scalar_number(r#"=SEARCH("ss";"sß")"#, 2.0, &execution);
    assert_scalar_logical(r#"=EXACT("Ab";"Ab")"#, true, &execution);
    assert_scalar_logical(r#"=EXACT("Ab";"ab")"#, false, &execution);
}

#[test]
fn scalar_text_catalog_smoke_covers_every_section_620_name() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-catalog");
    let sources = [
        r#"=ASC("Ａ")"#,
        r#"=CHAR(65)"#,
        r#"=CLEAN("a")"#,
        r#"=CODE("A")"#,
        r#"=CONCATENATE("a";"b")"#,
        r#"=DOLLAR(1)"#,
        r#"=EXACT("a";"a")"#,
        r#"=FIND("a";"a")"#,
        r#"=FIXED(1)"#,
        r#"=JIS("A")"#,
        r#"=LEFT("a")"#,
        r#"=LEN("a")"#,
        r#"=LOWER("A")"#,
        r#"=MID("a";1;1)"#,
        r#"=PROPER("a")"#,
        r#"=REPLACE("a";1;1;"b")"#,
        r#"=REPT("a";1)"#,
        r#"=RIGHT("a")"#,
        r#"=SEARCH("a";"A")"#,
        r#"=SUBSTITUTE("a";"a";"b")"#,
        r#"=T("a")"#,
        r#"=TEXT(1;"0")"#,
        r#"=TRIM("a")"#,
        r#"=UNICHAR(65)"#,
        r#"=UNICODE("A")"#,
        r#"=UPPER("a")"#,
    ];
    for source in sources {
        let result = scalar(source, &execution);
        assert!(
            !matches!(
                result,
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Function))
            ),
            "catalog entry was not dispatched: {source}"
        );
    }
}

#[test]
fn scalar_format_text_functions_use_the_documented_locale_independent_profile() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-format");
    for (source, expected) in [
        (r#"=DOLLAR(1234.5;2)"#, "$1,234.50"),
        (r#"=DOLLAR(12.345;1)"#, "$12.3"),
        (r#"=FIXED(1234.5;2)"#, "1,234.50"),
        (r#"=FIXED(12.345;1;TRUE())"#, "12.3"),
        (r#"=FIXED(1234.5;2;TRUE())"#, "1234.50"),
        (r#"=TEXT(1234.5;"0.00")"#, "1234.50"),
        (r#"=TEXT(12.345;"0.0")"#, "12.3"),
        (r#"=TEXT("abc";"@")"#, "abc"),
    ] {
        assert_scalar_text(source, expected, &execution);
    }
    for source in [
        r#"=DOLLAR(#REF!;#DIV/0!)"#,
        r#"=FIXED(#REF!;#DIV/0!;FALSE())"#,
        r#"=TEXT(#REF!;#DIV/0!)"#,
        r#"=TEXT(;#REF!)"#,
    ] {
        assert_scalar_error(source, ScalarError::Reference, &execution);
        let wrapped = format!(r#"=IFERROR({};"fallback")"#, source.trim_start_matches('='));
        assert_scalar_text(&wrapped, "fallback", &execution);
    }
    assert_scalar_error(r#"=TEXT(1;COMPLEX(1;2))"#, ScalarError::Value, &execution);
    assert_scalar_error(
        r#"=TEXT(#REF!;COMPLEX(1;2))"#,
        ScalarError::Reference,
        &execution,
    );
}

#[test]
fn scalar_text_optional_defaults_and_formula_errors_follow_function_profiles() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-optional");
    assert_scalar_text(r#"=LEFT("abc")"#, "a", &execution);
    assert_scalar_text(r#"=RIGHT("abc")"#, "c", &execution);
    assert_scalar_text(r#"=LEFT("abc";)"#, "a", &execution);
    assert_scalar_text(r#"=RIGHT("abc";)"#, "c", &execution);
    assert_scalar_text(r#"=LEFT("abc";2.9)"#, "ab", &execution);
    assert_scalar_text(r#"=MID("abc";1.9;1.9)"#, "a", &execution);
    for source in [
        r#"=LEFT("abc";1e308)"#,
        r#"=RIGHT("abc";1e308)"#,
        r#"=MID("abc";1;1e308)"#,
    ] {
        assert_scalar_text(source, "abc", &execution);
    }
    assert_scalar_text(r#"=REPT("";1e308)"#, "", &execution);
    assert_scalar_number(r#"=FIND("b";"abc")"#, 2.0, &execution);
    assert_scalar_number(r#"=SEARCH("b";"abc")"#, 2.0, &execution);
    assert_scalar_number(r#"=FIND("b";"abc";)"#, 2.0, &execution);
    assert_scalar_number(r#"=SEARCH("b";"abc";)"#, 2.0, &execution);
    assert_scalar_text(r#"=SUBSTITUTE("a-a";"-";"+")"#, "a+a", &execution);
    assert_scalar_error(r#"=MID("abc";0;1)"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=LEFT("abc";-0.2)"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=RIGHT("abc";-0.2)"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=MID("abc";0.9;1)"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=REPT("x";-1)"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=FIND("x";"abc")"#, ScalarError::Value, &execution);
    assert_scalar_error(r#"=UPPER(#N/A)"#, ScalarError::NotAvailable, &execution);
    assert_scalar_error(
        r#"=CONCATENATE("ok";#N/A)"#,
        ScalarError::NotAvailable,
        &execution,
    );
    assert_scalar_error(r#"=LEFT(#N/A;)"#, ScalarError::NotAvailable, &execution);
    assert_scalar_error(r#"=FIND("a";#N/A;)"#, ScalarError::NotAvailable, &execution);
    assert_scalar_error(
        r#"=SUBSTITUTE(#N/A;"a";"b";)"#,
        ScalarError::NotAvailable,
        &execution,
    );
    assert_scalar_text(
        r#"=IFERROR(UPPER(#N/A);"fallback")"#,
        "fallback",
        &execution,
    );
}

#[test]
fn value_text_arrays_lift_and_broadcast_without_flattening_concatenate_inputs() {
    let resolver = TextResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-text-arrays");
    assert_value_text_array(
        r#"=UPPER({"a"|"b"})"#,
        2,
        1,
        &["A", "B"],
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value_text(
        r#"=UPPER({"a"|"b"})"#,
        "A",
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value_text_array(
        r#"=CONCATENATE({"a"|"b"};{"!";"?"})"#,
        2,
        2,
        &["a!", "a?", "b!", "b?"],
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value_text_array(
        r#"=SUBSTITUTE({"a"|"b"};"a";{"x";"y"})"#,
        2,
        2,
        &["x", "y", "b", "b"],
        &resolver,
        &execution,
        Mode::Matrix,
    );
}

#[test]
fn value_text_references_stream_cells_in_order_and_keep_scalar_intersection() {
    let resolver = TextResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-text-references");
    assert_value_text_array(
        "=UPPER([.A1:.A2])",
        2,
        1,
        &["ALPHA", "É🌟"],
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value_text(
        "=UPPER([.A1:.A2])",
        "ALPHA",
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value_text_array(
        r#"=CONCATENATE([.A1:.A2];"!")"#,
        2,
        1,
        &["Alpha!", "é🌟!"],
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        5,
        "each reference argument cell is read once"
    );
}

#[test]
fn value_text_empty_reference_parameters_use_their_declared_text_profiles() {
    let (_budget, _cancellation, execution) = execution("ods-formula-text-empty-parameters");
    for (source, expected) in [
        (r#"=SUBSTITUTE([.A7];"a";"b")"#, ""),
        (r#"=SUBSTITUTE("a";[.A7];"b")"#, "a"),
        (r#"=SUBSTITUTE("a";"a";[.A7])"#, ""),
        (r#"=REPLACE("abc";2;1;[.A7])"#, "ac"),
    ] {
        let resolver = TextResolver::standard();
        assert_value_text(source, expected, &resolver, &execution, Mode::Scalar);
    }
    for source in [r#"=FIND("a";[.A7])"#, r#"=SEARCH("a";[.A7])"#] {
        let resolver = TextResolver::standard();
        assert_value_error(
            source,
            ScalarError::Value,
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }

    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, error) in [
            (r#"=TEXT([.A7];"0")"#, ScalarError::Value),
            (r#"=TEXT([.A7];#REF!)"#, ScalarError::Reference),
        ] {
            let resolver = TextResolver::standard();
            let result = value(source, &resolver, &execution, mode, &Limits::default())
                .unwrap_or_else(|failure| panic!("{source} in {mode:?}: {failure}"));
            let expected = match mode {
                Mode::Scalar => ValueObserved::Error(error),
                Mode::Matrix => ValueObserved::Array {
                    rows: 1,
                    columns: 1,
                    cells: vec![ValueObserved::Error(error)],
                },
                _ => panic!("unsupported text evaluation mode"),
            };
            assert_eq!(result, expected, "{source} in {mode:?}");
        }
    }

    let mut resolver = TextResolver::standard();
    resolver.set(0, 3, TextCell::Unsupported);
    let error = value(
        r#"=TEXT([.A7];[.D1])"#,
        &resolver,
        &execution,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("a later provider failure supersedes TEXT's Empty type refusal");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
}

#[test]
fn value_text_three_dimensional_reference_and_reference_list_are_refused_read_free() {
    let resolver = TextResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-text-shapes");
    let error = value(
        "=UPPER([Main.A1:Archive.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("a 3-D cuboid is outside the 2-D text mapper profile");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);
    resolver.clear_reads();
    assert_value_error(
        "=UPPER([.A1]~[.A2])",
        ScalarError::Value,
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0, "a rejected ReferenceList is read-free");
}

#[test]
fn value_text_any_and_t_preserve_empty_nontext_and_formula_error_cells() {
    let resolver = TextResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-text-any");
    let result = value(
        r#"=T({"text"|42})"#,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("T array");
    let ValueObserved::Array {
        rows,
        columns,
        cells,
    } = result
    else {
        panic!("T array result");
    };
    assert_eq!((rows, columns), (2, 1));
    assert_eq!(cells[0], ValueObserved::Text("text".to_owned()));
    assert_eq!(cells[1], ValueObserved::Text(String::new()));

    assert_eq!(
        value(
            "=T([.A8])",
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect("T error cell"),
        ValueObserved::Array {
            rows: 1,
            columns: 1,
            cells: vec![ValueObserved::Error(ScalarError::NotAvailable)],
        }
    );
    assert_value_text(
        "=CONCATENATE([.A1];[.C1])",
        "Alpha12.5",
        &resolver,
        &execution,
        Mode::Scalar,
    );
}

#[test]
fn projected_if_keeps_text_shape_and_skips_unselected_reference_reads() {
    let resolver = TextResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-text-if");
    let result = value(
        "=IF(TRUE();UPPER([.A1:.A2]);UPPER([Missing.A1:.Z100]))",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("selected text branch");
    let ValueObserved::Array {
        rows,
        columns,
        cells,
    } = result
    else {
        panic!("projected text array");
    };
    assert_eq!((rows, columns), (2, 1));
    assert_eq!(cells[0], ValueObserved::Text("ALPHA".to_owned()));
    assert_eq!(cells[1], ValueObserved::Text("É🌟".to_owned()));
    assert_eq!(resolver.reads(), 2, "the unselected branch stays lazy");

    let resolver = TextResolver::standard();
    let result = value(
        r#"=IF(FALSE();UPPER([Missing.A1]);"ok")"#,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("unselected missing reference");
    assert_eq!(result, ValueObserved::Text("ok".to_owned()));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn value_text_provider_errors_and_source_fences_remain_typed() {
    let mut resolver = TextResolver::standard();
    resolver.set(1, 0, TextCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-text-provider");
    let error = value(
        "=UPPER([.A1:.A2])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("unsupported provider value must escape text conversion");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));

    let mut resolver = TextResolver::standard();
    resolver.set_source_versions(
        SourceVersion::new(0x5445_5854, 0),
        SourceVersion::new(0x5445_5854, 1),
    );
    let error = value(
        "=UPPER([.A1])",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("source change must reject text publication");
    assert!(matches!(error, EvaluationFailure::SourceChanged { .. }));
}
