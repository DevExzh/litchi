//! Deterministic scalar formatting for ODF 1.4 `DOLLAR`, `FIXED`, and `TEXT`.
//!
//! ODF deliberately leaves the number-format grammar of `TEXT` implementation
//! defined.  This profile documents the grammar in the parser below instead
//! of consulting the process locale: literals, quoted literals, the numeric
//! placeholders `0`, `#`, and `?`, grouping and decimal punctuation, percent,
//! scientific notation, simple fractions, and Gregorian date/time fields are
//! accepted.  An unrecognised format character is a formula `#VALUE!` error.
//! Alphabetic literals in a text section are accepted when the section also
//! contains `@`; numeric sections require quoting for alphabetic literals.
//! Numeric format precision is limited by the caller's text/work budgets;
//! `DOLLAR` and `FIXED` likewise use the caller's text-byte limit for positive
//! `D`.  Fraction denominators use one through six `?` places in this profile,
//! with a denominator in `1..=10^places-1` (the denominator has at most six
//! placeholder digits).
//! Date serials use the non-leap-1900 epoch 1899-12-30 and English month and
//! weekday names.
//! `DOLLAR` and `FIXED` use the same decimal rounding kernel and a fixed
//! en-US presentation (`$`, `,`, and `.`).

use super::super::rounding::round_nearest;
use super::super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, NumberText, ScalarError, TextValue,
    WorkingValue, local_limit,
};
use super::core;
use litchi_core::Resource;
use std::fmt::Write as FmtWrite;

const MAX_FORMAT_SECTIONS: usize = 4;
const DECIMAL_DIGIT_CAPACITY: usize = 1024;
const MAX_FRACTION_DENOMINATOR_DIGITS: usize = 6;
const SECONDS_PER_DAY: f64 = 86_400.0;
const SERIAL_EPOCH_TO_UNIX_DAYS: i64 = 25_569;

#[derive(Clone, Copy)]
enum Function {
    Dollar,
    Fixed,
    Text,
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("DOLLAR") {
        Some(Function::Dollar)
    } else if name.eq_ignore_ascii_case("FIXED") {
        Some(Function::Fixed)
    } else if name.eq_ignore_ascii_case("TEXT") {
        Some(Function::Text)
    } else {
        None
    }
}

pub(super) fn is_format_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply one of the scalar formatting functions after eager argument
/// evaluation.  The argument tail is inspected in source order so a source
/// formula error retains precedence over a format-domain error.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Dollar) => apply_dollar(evaluator, node),
        Some(Function::Fixed) => apply_fixed(evaluator, node),
        Some(Function::Text) => apply_text(evaluator, node),
        None => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        )),
    }
}

fn missing(node: Node<'_>, index: usize) -> bool {
    index >= node.child_count() || node.child(index).is_some_and(|child| child.is_missing())
}

fn apply_dollar<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=2).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = core::first_formula_error(evaluator, node, count)? {
        core::discard_tail(evaluator, count)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if reject_complex_arguments(evaluator, count)? {
        return Ok(());
    }

    let decimals_value = if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number_value = evaluator.pop_value()?;
    if missing(node, 0) {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Value));
    }
    let number = match super::super::to_number(number_value, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let decimals = match optional_decimal(evaluator, decimals_value, missing(node, 1), 2.0)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    match format_dollar(evaluator, number, decimals)? {
        Ok(text) => evaluator.push_value(WorkingValue::Text(text)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn apply_fixed<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=3).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = core::first_formula_error(evaluator, node, count)? {
        core::discard_tail(evaluator, count)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if reject_complex_arguments(evaluator, count)? {
        return Ok(());
    }

    let omit_value = if count == 3 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let decimals_value = if count >= 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number_value = evaluator.pop_value()?;
    if missing(node, 0) {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Value));
    }
    let number = match super::super::to_number(number_value, evaluator)? {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let decimals = match optional_decimal(evaluator, decimals_value, missing(node, 1), 2.0)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let omit = match optional_logical(evaluator, omit_value, missing(node, 2), false)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    match format_fixed(evaluator, number, decimals, omit)? {
        Ok(text) => evaluator.push_value(WorkingValue::Text(text)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn apply_text<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 2 {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = core::first_formula_error(evaluator, node, 2)? {
        core::discard_tail(evaluator, 2)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if reject_complex_arguments(evaluator, 2)? {
        return Ok(());
    }
    let format_value = evaluator.pop_value()?;
    let value = evaluator.pop_value()?;
    if missing(node, 1) {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Value));
    }
    let format_text = super::super::to_text(format_value, evaluator)?;
    let format_code = format_text.text.as_ref();
    let result = format_value_result(evaluator, value, format_code)?;
    drop(format_text);
    match result {
        Ok(text) => evaluator.push_value(WorkingValue::Text(text)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn reject_complex_arguments(
    evaluator: &mut Evaluator<'_, '_, '_>,
    count: usize,
) -> EvaluationResult<bool> {
    if !core::has_complex_argument(evaluator, count) {
        return Ok(false);
    }
    core::discard_tail(evaluator, count)?;
    evaluator.push_value(WorkingValue::Error(ScalarError::Value))?;
    Ok(true)
}

fn format_value_result<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
    format_code: &str,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    evaluator.charge_bytes(format_code.len())?;
    let mut checks = FormatChecks::new(evaluator.context.execution);
    let parsed = FormatProgram::parse(format_code, &mut checks);
    let program = match checks.finish(parsed)? {
        Ok(program) => program,
        Err(error) => return Ok(Err(error)),
    };
    match value {
        WorkingValue::Error(error) => Ok(Err(error)),
        WorkingValue::Number(number) => {
            if !number.is_finite() {
                return Ok(Err(ScalarError::Number));
            }
            format_number_program(evaluator, number, &program)
        },
        WorkingValue::Logical(value) => {
            format_text_program(evaluator, if value { "TRUE" } else { "FALSE" }, &program)
        },
        WorkingValue::Text(value) => {
            let source = value.text.as_ref();
            if program.is_identity_text() {
                evaluator.charge_bytes(source.len())?;
                return Ok(Ok(value));
            }
            format_text_program(evaluator, source, &program)
        },
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
    }
}

fn optional_decimal<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: Option<WorkingValue<'a>>,
    empty: bool,
    default: f64,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if empty {
        return Ok(Ok(default));
    }
    match value {
        Some(value) => match super::super::to_integer(value, evaluator)? {
            Ok(value) if value.is_finite() => Ok(Ok(value)),
            Ok(_) => Ok(Err(ScalarError::Number)),
            Err(error) => Ok(Err(error)),
        },
        None => Ok(Err(ScalarError::Value)),
    }
}

fn optional_logical<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: Option<WorkingValue<'a>>,
    empty: bool,
    default: bool,
) -> EvaluationResult<Result<bool, ScalarError>> {
    if empty {
        return Ok(Ok(default));
    }
    match value {
        Some(WorkingValue::Logical(value)) => Ok(Ok(value)),
        Some(WorkingValue::Number(value)) if value.is_finite() => Ok(Ok(value != 0.0)),
        Some(WorkingValue::Text(value)) => {
            evaluator.charge_bytes(value.text.len())?;
            Ok(Err(ScalarError::Value))
        },
        Some(WorkingValue::Error(error)) => Ok(Err(error)),
        Some(WorkingValue::Complex(_)) => Ok(Err(ScalarError::Value)),
        Some(WorkingValue::Number(_)) => Ok(Err(ScalarError::Number)),
        None => Ok(Err(ScalarError::Value)),
    }
}

fn format_dollar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    number: f64,
    decimals: f64,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    format_fixed_inner(evaluator, number, decimals, false, true)
}

fn format_fixed<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    number: f64,
    decimals: f64,
    omit_separators: bool,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    format_fixed_inner(evaluator, number, decimals, omit_separators, false)
}

fn format_fixed_inner<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    number: f64,
    decimals: f64,
    omit_separators: bool,
    currency: bool,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    evaluator.charge_work(1)?;
    let decimals = decimals.trunc();
    if !number.is_finite() || !decimals.is_finite() {
        return Ok(Err(ScalarError::Number));
    }
    let rounded = match round_nearest(number, decimals) {
        Ok(value) if value.is_finite() => value,
        Ok(_) => return Ok(Err(ScalarError::Number)),
        Err(error) => return Ok(Err(error)),
    };
    let places = positive_places(evaluator, decimals)?;
    let absolute = rounded.abs();
    let representation = match DecimalRepresentation::from_number(absolute) {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    let negative = rounded.is_sign_negative() && rounded != 0.0;
    let mut render = |sink: &mut dyn Sink| {
        render_fixed_value(
            sink,
            &representation,
            places,
            !omit_separators,
            negative,
            currency,
        )
    };
    make_text(evaluator, &mut render)
}

fn positive_places(evaluator: &Evaluator<'_, '_, '_>, decimals: f64) -> EvaluationResult<usize> {
    if decimals <= 0.0 {
        return Ok(0);
    }
    if decimals > evaluator.limits.max_text_bytes as f64 {
        return Err(local_limit(
            Resource::Memory,
            decimals as u64,
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        ));
    }
    usize::try_from(decimals as u64).map_err(|_| {
        local_limit(
            Resource::Memory,
            u64::MAX,
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        )
    })
}

#[derive(Clone, Copy)]
struct DecimalRepresentation {
    digits: [u8; DECIMAL_DIGIT_CAPACITY],
    length: usize,
    decimal_position: i32,
}

impl DecimalRepresentation {
    fn from_number(value: f64) -> Result<Self, ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        let mut text = NumberText::new();
        write!(&mut text, "{value}").map_err(|_| ScalarError::Number)?;
        let text = text.as_str().map_err(|_| ScalarError::Number)?;
        let bytes = text.as_bytes();
        let mut digits = [0_u8; DECIMAL_DIGIT_CAPACITY];
        let mut length = 0usize;
        let mut before_decimal = 0usize;
        let mut seen_decimal = false;
        let mut exponent = 0_i32;
        let mut cursor = 0usize;
        while cursor < bytes.len() {
            match bytes[cursor] {
                b'+' | b'-' if cursor == 0 => cursor += 1,
                b'.' => {
                    seen_decimal = true;
                    cursor += 1;
                },
                b'e' | b'E' => {
                    let rest = std::str::from_utf8(&bytes[cursor + 1..])
                        .map_err(|_| ScalarError::Number)?;
                    exponent = rest.parse::<i32>().map_err(|_| ScalarError::Number)?;
                    break;
                },
                byte if byte.is_ascii_digit() => {
                    if length == digits.len() {
                        return Err(ScalarError::Number);
                    }
                    digits[length] = byte;
                    length += 1;
                    if !seen_decimal {
                        before_decimal = before_decimal.saturating_add(1);
                    }
                    cursor += 1;
                },
                _ => return Err(ScalarError::Number),
            }
        }
        if length == 0 {
            return Err(ScalarError::Number);
        }
        let decimal_position = i32::try_from(before_decimal)
            .ok()
            .and_then(|position| position.checked_add(exponent))
            .ok_or(ScalarError::Number)?;
        Ok(Self {
            digits,
            length,
            decimal_position,
        })
    }

    fn digit_at(&self, position: i32) -> u8 {
        if position < 0 {
            return b'0';
        }
        let index = usize::try_from(position).unwrap_or(usize::MAX);
        if index < self.length {
            self.digits[index]
        } else {
            b'0'
        }
    }

    fn integer_len(&self) -> usize {
        if self.decimal_position > 0 {
            usize::try_from(self.decimal_position).unwrap_or(usize::MAX)
        } else {
            1
        }
    }
}

trait Sink {
    fn push_str(&mut self, value: &str) -> Result<(), ScalarError>;
    fn push_char(&mut self, value: char) -> Result<(), ScalarError>;
    fn len(&self) -> usize;
    fn checkpoint(&mut self) -> Result<(), ScalarError>;
}

/// Rendering and grammar scans can consume input without emitting output.
/// Preserve typed execution failures across the formatter's formula-error API.
struct FormatChecks<'a> {
    execution: &'a litchi_core::ExecutionContext,
    iterations: usize,
    failure: Option<EvaluationFailure>,
}

impl<'a> FormatChecks<'a> {
    fn new(execution: &'a litchi_core::ExecutionContext) -> Self {
        Self {
            execution,
            iterations: 0,
            failure: None,
        }
    }

    fn check(&mut self) -> Result<(), ScalarError> {
        if self.failure.is_some() {
            return Err(ScalarError::Number);
        }
        if let Err(error) = self.execution.check() {
            self.failure = Some(super::super::map_execution_error(error));
            return Err(ScalarError::Number);
        }
        Ok(())
    }

    fn checkpoint(&mut self) -> Result<(), ScalarError> {
        let check = self.iterations & 4095 == 0;
        self.iterations = self.iterations.wrapping_add(1);
        if check { self.check() } else { Ok(()) }
    }

    fn finish<T>(self, result: Result<T, ScalarError>) -> EvaluationResult<Result<T, ScalarError>> {
        if let Some(error) = self.failure {
            return Err(error);
        }
        self.execution
            .check()
            .map_err(super::super::map_execution_error)?;
        Ok(result)
    }
}

struct CountSink<'borrow, 'expr, 'ctx, 'exec> {
    length: usize,
    evaluator: &'borrow mut Evaluator<'expr, 'ctx, 'exec>,
    failure: Option<EvaluationFailure>,
    checks: FormatChecks<'exec>,
}

impl CountSink<'_, '_, '_, '_> {
    fn admit(&mut self, bytes: usize) -> Result<(), ScalarError> {
        if self.failure.is_some() {
            return Err(ScalarError::Number);
        }
        if let Err(error) = self.evaluator.charge_bytes(bytes) {
            self.failure = Some(error);
            return Err(ScalarError::Number);
        }
        let length = self.length.checked_add(bytes);
        if length.is_none_or(|length| length > self.evaluator.limits.max_text_bytes) {
            self.failure = Some(core::text_limit(
                self.evaluator,
                length.unwrap_or(usize::MAX),
            ));
            return Err(ScalarError::Number);
        }
        self.length = length.ok_or(ScalarError::Number)?;
        Ok(())
    }
}

impl Sink for CountSink<'_, '_, '_, '_> {
    fn push_str(&mut self, value: &str) -> Result<(), ScalarError> {
        self.admit(value.len())
    }

    fn push_char(&mut self, value: char) -> Result<(), ScalarError> {
        self.admit(value.len_utf8())
    }

    fn len(&self) -> usize {
        self.length
    }

    fn checkpoint(&mut self) -> Result<(), ScalarError> {
        self.checks.checkpoint()
    }
}

struct StringSink<'a, 'exec> {
    output: &'a mut String,
    checks: FormatChecks<'exec>,
    admitted_length: usize,
}

impl StringSink<'_, '_> {
    fn admit(&mut self, bytes: usize) -> Result<(), ScalarError> {
        if self
            .output
            .len()
            .checked_add(bytes)
            .is_none_or(|length| length > self.admitted_length)
        {
            self.checks.failure = Some(EvaluationFailure::InvalidExpression(
                "formatted text exceeded its admitted size",
            ));
            return Err(ScalarError::Number);
        }
        Ok(())
    }
}

impl Sink for StringSink<'_, '_> {
    fn push_str(&mut self, value: &str) -> Result<(), ScalarError> {
        self.checks.check()?;
        self.admit(value.len())?;
        self.output.push_str(value);
        self.checks.check()
    }

    fn push_char(&mut self, value: char) -> Result<(), ScalarError> {
        self.checkpoint()?;
        self.admit(value.len_utf8())?;
        self.output.push(value);
        Ok(())
    }

    fn len(&self) -> usize {
        self.output.len()
    }

    fn checkpoint(&mut self) -> Result<(), ScalarError> {
        self.checks.checkpoint()
    }
}

fn make_text<'a, F>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    render: &mut F,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>>
where
    F: FnMut(&mut dyn Sink) -> Result<(), ScalarError>,
{
    let execution = evaluator.context.execution;
    let mut count = CountSink {
        length: 0,
        evaluator,
        failure: None,
        checks: FormatChecks::new(execution),
    };
    let rendered = render(&mut count);
    let failure = count.failure.take().or_else(|| count.checks.failure.take());
    let length = count.len();
    drop(count);
    if let Some(error) = failure {
        return Err(error);
    }
    if let Err(error) = rendered {
        execution
            .check()
            .map_err(super::super::map_execution_error)?;
        return Ok(Err(error));
    }
    if length > evaluator.limits.max_text_bytes {
        return Err(core::text_limit(evaluator, length));
    }
    let reservation = evaluator.reserve_storage(length, "formula scalar formatted text")?;
    let mut output = String::new();
    if length != 0 {
        output
            .try_reserve_exact(length)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula scalar formatted text",
                source,
            })?;
    }
    let mut sink = StringSink {
        output: &mut output,
        checks: FormatChecks::new(execution),
        admitted_length: length,
    };
    let rendered = render(&mut sink);
    if let Some(error) = sink.checks.failure.take() {
        return Err(error);
    }
    if let Err(error) = rendered {
        execution
            .check()
            .map_err(super::super::map_execution_error)?;
        return Ok(Err(error));
    }
    if sink.len() != length {
        return Err(EvaluationFailure::InvalidExpression(
            "formatted text size changed between passes",
        ));
    }
    execution
        .check()
        .map_err(super::super::map_execution_error)?;
    Ok(Ok(TextValue::owned(output, reservation)))
}

fn render_fixed_value(
    sink: &mut dyn Sink,
    representation: &DecimalRepresentation,
    places: usize,
    grouping: bool,
    negative: bool,
    currency: bool,
) -> Result<(), ScalarError> {
    if currency && negative {
        sink.push_char('(')?;
    } else if negative {
        sink.push_char('-')?;
    }
    if currency {
        sink.push_char('$')?;
    }
    let integer_len = representation.integer_len();
    for index in 0..integer_len {
        if grouping && index != 0 && (integer_len - index).is_multiple_of(3) {
            sink.push_char(',')?;
        }
        let position = i32::try_from(index).map_err(|_| ScalarError::Number)?;
        sink.push_char(char::from(representation.digit_at(position)))?;
    }
    if places != 0 {
        sink.push_char('.')?;
        for index in 0..places {
            let position = representation
                .decimal_position
                .checked_add(i32::try_from(index).map_err(|_| ScalarError::Number)?)
                .ok_or(ScalarError::Number)?;
            sink.push_char(char::from(representation.digit_at(position)))?;
        }
    }
    if currency && negative {
        sink.push_char(')')?;
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum SectionKind {
    Literal,
    Numeric,
    Date,
    Text,
}

#[derive(Clone, Copy)]
struct FormatSection<'a> {
    source: &'a str,
    kind: SectionKind,
    first_placeholder: usize,
    last_placeholder: usize,
    percent: bool,
}

impl<'a> FormatSection<'a> {
    const fn empty(source: &'a str) -> Self {
        Self {
            source,
            kind: SectionKind::Literal,
            first_placeholder: 0,
            last_placeholder: 0,
            percent: false,
        }
    }
}

struct FormatProgram<'a> {
    sections: [FormatSection<'a>; MAX_FORMAT_SECTIONS],
    count: usize,
}

impl<'a> FormatProgram<'a> {
    fn parse(source: &'a str, checks: &mut FormatChecks<'_>) -> Result<Self, ScalarError> {
        if source.is_empty() {
            return Err(ScalarError::Value);
        }
        let mut sections = [FormatSection::empty(""); MAX_FORMAT_SECTIONS];
        let mut count = 0usize;
        let mut start = 0usize;
        let mut cursor = 0usize;
        let bytes = source.as_bytes();
        let mut quoted = false;
        while cursor < bytes.len() {
            checks.checkpoint()?;
            match bytes[cursor] {
                b'"' => {
                    if quoted && bytes.get(cursor + 1) == Some(&b'"') {
                        cursor += 2;
                    } else {
                        quoted = !quoted;
                        cursor += 1;
                    }
                },
                b'\\' if !quoted => {
                    cursor = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
                },
                b';' if !quoted => {
                    if count == MAX_FORMAT_SECTIONS {
                        return Err(ScalarError::Value);
                    }
                    sections[count] = parse_section(&source[start..cursor], checks)?;
                    count += 1;
                    cursor += 1;
                    start = cursor;
                },
                _ => cursor = next_char_end(source, cursor).ok_or(ScalarError::Value)?,
            }
        }
        if quoted || count == MAX_FORMAT_SECTIONS {
            return Err(ScalarError::Value);
        }
        sections[count] = parse_section(&source[start..], checks)?;
        count += 1;
        Ok(Self { sections, count })
    }

    fn section(&self, index: usize) -> FormatSection<'a> {
        self.sections[index.min(self.count.saturating_sub(1))]
    }

    fn is_identity_text(&self) -> bool {
        self.count == 1
            && matches!(self.sections[0].kind, SectionKind::Text)
            && self.sections[0].source == "@"
    }
}

fn next_char_end(source: &str, start: usize) -> Option<usize> {
    if start >= source.len() {
        None
    } else {
        Some(start + source[start..].chars().next()?.len_utf8())
    }
}

fn has_unquoted_text_marker(
    source: &str,
    checks: &mut FormatChecks<'_>,
) -> Result<bool, ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = 0usize;
    let mut quoted = false;
    while cursor < bytes.len() {
        checks.checkpoint()?;
        let character = match source[cursor..].chars().next() {
            Some(character) => character,
            None => return Ok(false),
        };
        match character {
            '"' => {
                if quoted && bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                } else {
                    quoted = !quoted;
                    cursor += 1;
                }
            },
            '\\' | '_' if !quoted => {
                cursor = next_char_end(source, cursor + 1).unwrap_or(source.len());
            },
            '@' if !quoted => return Ok(true),
            _ => cursor += character.len_utf8(),
        }
    }
    Ok(false)
}

fn parse_section<'a>(
    source: &'a str,
    checks: &mut FormatChecks<'_>,
) -> Result<FormatSection<'a>, ScalarError> {
    let mut cursor = 0usize;
    let mut quoted = false;
    let mut first_placeholder = None;
    let mut last_placeholder = 0usize;
    let mut numeric = false;
    let mut date = false;
    let mut text = has_unquoted_text_marker(source, checks)?;
    let mut unknown_alpha = false;
    let mut percent = false;
    let mut percent_count = 0usize;
    let bytes = source.as_bytes();
    while cursor < bytes.len() {
        checks.checkpoint()?;
        let start = cursor;
        let character = source[cursor..].chars().next().ok_or(ScalarError::Value)?;
        let width = character.len_utf8();
        if quoted {
            if character == '"' {
                if bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                } else {
                    quoted = false;
                    cursor += 1;
                }
            } else {
                cursor += width;
            }
            continue;
        }
        match character {
            '"' => {
                quoted = true;
                cursor += 1;
            },
            '\\' => {
                cursor = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
            },
            '_' => {
                cursor = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
            },
            '*' => return Err(ScalarError::Value),
            '@' => {
                text = true;
                cursor += 1;
            },
            '%' => {
                percent = true;
                percent_count = percent_count.saturating_add(1);
                cursor += 1;
            },
            '0' | '#' | '?' => {
                numeric = true;
                if first_placeholder.is_none() {
                    first_placeholder = Some(start);
                }
                last_placeholder = cursor + 1;
                cursor += 1;
            },
            '/' if numeric => {
                last_placeholder = cursor + 1;
                cursor += 1;
            },
            'E' | 'e' => {
                if text {
                    unknown_alpha = true;
                    cursor += width;
                    continue;
                }
                let mut look = cursor + 1;
                if bytes
                    .get(look)
                    .is_some_and(|byte| *byte == b'+' || *byte == b'-')
                {
                    look += 1;
                }
                let exponent_start = look;
                while bytes
                    .get(look)
                    .is_some_and(|byte| matches!(*byte, b'0' | b'#' | b'?'))
                {
                    checks.checkpoint()?;
                    look += 1;
                }
                if exponent_start == look {
                    if numeric {
                        return Err(ScalarError::Value);
                    }
                    unknown_alpha = true;
                    cursor += 1;
                    continue;
                }
                numeric = true;
                if first_placeholder.is_none() {
                    first_placeholder = Some(start);
                }
                last_placeholder = look;
                cursor = look;
            },
            'A' | 'a' | 'P' | 'p' if !text && starts_am_pm(source, cursor) => {
                date = true;
                cursor += 5;
            },
            'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => {
                if text {
                    unknown_alpha = true;
                    cursor += width;
                    continue;
                }
                let end = same_letter_end(source, cursor, character);
                let length = source[cursor..end].chars().count();
                let valid = match character.to_ascii_lowercase() {
                    'y' => matches!(length, 1 | 2 | 4),
                    'm' | 'd' => (1..=4).contains(&length),
                    'h' | 's' => (1..=2).contains(&length),
                    _ => false,
                };
                if !valid {
                    return Err(ScalarError::Value);
                }
                date = true;
                cursor = end;
            },
            character if character.is_ascii_alphabetic() => {
                unknown_alpha = true;
                cursor += width;
            },
            _ => cursor += width,
        }
    }
    if quoted
        || (numeric && (date || text || unknown_alpha))
        || (date && (text || unknown_alpha))
        || (unknown_alpha && !text && !numeric && !date)
        || (numeric && percent_count > 1)
        || (date && (percent || has_invalid_date_punctuation(source, checks)?))
    {
        return Err(ScalarError::Value);
    }
    let (kind, first, last) = if date {
        (SectionKind::Date, 0, 0)
    } else if numeric {
        (
            SectionKind::Numeric,
            first_placeholder.ok_or(ScalarError::Value)?,
            last_placeholder,
        )
    } else if text {
        (SectionKind::Text, 0, 0)
    } else {
        (SectionKind::Literal, 0, 0)
    };
    Ok(FormatSection {
        source,
        kind,
        first_placeholder: first,
        last_placeholder: last,
        percent,
    })
}

fn has_invalid_date_punctuation(
    source: &str,
    checks: &mut FormatChecks<'_>,
) -> Result<bool, ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = 0usize;
    let mut quoted = false;
    while cursor < bytes.len() {
        checks.checkpoint()?;
        let character = match source[cursor..].chars().next() {
            Some(character) => character,
            None => return Ok(true),
        };
        if quoted {
            if character == '"' {
                if bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                } else {
                    quoted = false;
                    cursor += 1;
                }
            } else {
                cursor += character.len_utf8();
            }
            continue;
        }
        match character {
            '"' => {
                quoted = true;
                cursor += 1;
            },
            '\\' | '_' => {
                cursor = next_char_end(source, cursor + 1).unwrap_or(source.len());
            },
            '\t' => return Ok(true),
            '-' | '/' | ':' | ' ' => cursor += character.len_utf8(),
            'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' | 'A' | 'a' | 'P' | 'p' => {
                cursor += character.len_utf8()
            },
            _ => return Ok(true),
        }
    }
    Ok(quoted)
}

fn same_letter_end(source: &str, start: usize, character: char) -> usize {
    let mut cursor = start;
    let mut length = 0usize;
    while cursor < source.len() && length < 5 {
        let Some(next) = source[cursor..].chars().next() else {
            break;
        };
        if !next.eq_ignore_ascii_case(&character) {
            break;
        }
        cursor += next.len_utf8();
        length += 1;
    }
    cursor
}

fn starts_am_pm(source: &str, start: usize) -> bool {
    source
        .get(start..start.saturating_add(5))
        .is_some_and(|value| value.eq_ignore_ascii_case("AM/PM"))
}

fn format_number_program<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    number: f64,
    program: &FormatProgram<'_>,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    let negative = number.is_sign_negative() && number != 0.0;
    let section_index = if program.count == 1 {
        0
    } else if negative {
        1
    } else if number == 0.0 && program.count >= 3 {
        2
    } else {
        0
    };
    let section = program.section(section_index);
    if matches!(section.kind, SectionKind::Literal) {
        let mut render = |sink: &mut dyn Sink| render_literal(sink, section.source, None);
        return make_text(evaluator, &mut render);
    }
    if matches!(section.kind, SectionKind::Text) {
        let mut number_text = NumberText::new();
        if write!(&mut number_text, "{number}").is_err() {
            return Ok(Err(ScalarError::Number));
        }
        let source = number_text
            .as_str()
            .map_err(|_| EvaluationFailure::InvalidExpression("invalid general number format"))?;
        evaluator.charge_bytes(source.len())?;
        return format_text_section(evaluator, source, section);
    }
    let absolute = number.abs();
    match section.kind {
        SectionKind::Numeric => {
            let mut checks = FormatChecks::new(evaluator.context.execution);
            let parsed = NumericSpec::parse(
                &section.source[section.first_placeholder..section.last_placeholder],
                &mut checks,
            );
            let mut spec = match checks.finish(parsed)? {
                Ok(spec) => spec,
                Err(error) => return Ok(Err(error)),
            };
            spec.percent |= section.percent;
            if let Some(digits) = spec.fraction_denominator_digits {
                if digits > MAX_FRACTION_DENOMINATOR_DIGITS {
                    return Ok(Err(ScalarError::Value));
                }
            }
            let computed_fraction = if let Some(fraction) = spec.fraction {
                let value = if spec.percent {
                    absolute * 100.0
                } else {
                    absolute
                };
                match fraction_parts(evaluator, value, fraction.denominator_digits)? {
                    Ok(parts) => Some(parts),
                    Err(error) => return Ok(Err(error)),
                }
            } else {
                None
            };
            let mut render = |sink: &mut dyn Sink| {
                render_literal(sink, &section.source[..section.first_placeholder], None)?;
                render_numeric_core(
                    sink,
                    absolute,
                    &spec,
                    negative && program.count == 1,
                    computed_fraction.as_ref(),
                )?;
                render_literal(sink, &section.source[section.last_placeholder..], None)
            };
            make_text(evaluator, &mut render)
        },
        SectionKind::Date => {
            let date = match date_parts(number) {
                Ok(date) => date,
                Err(error) => return Ok(Err(error)),
            };
            let mut render = |sink: &mut dyn Sink| render_date(sink, section.source, &date);
            make_text(evaluator, &mut render)
        },
        SectionKind::Literal | SectionKind::Text => Ok(Err(ScalarError::Value)),
    }
}

fn format_text_program<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    source: &str,
    program: &FormatProgram<'_>,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    evaluator.charge_bytes(source.len())?;
    let section = if program.count == MAX_FORMAT_SECTIONS {
        program.section(3)
    } else {
        program.section(0)
    };
    format_text_section(evaluator, source, section)
}

fn format_text_section<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    source: &str,
    section: FormatSection<'_>,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    let mut render = |sink: &mut dyn Sink| match section.kind {
        SectionKind::Text => render_literal(sink, section.source, Some(source)),
        SectionKind::Literal => render_literal(sink, section.source, None),
        SectionKind::Numeric | SectionKind::Date => Err(ScalarError::Value),
    };
    make_text(evaluator, &mut render)
}

fn render_literal(
    sink: &mut dyn Sink,
    source: &str,
    substitution: Option<&str>,
) -> Result<(), ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = 0usize;
    let mut quoted = false;
    while cursor < bytes.len() {
        sink.checkpoint()?;
        let character = source[cursor..].chars().next().ok_or(ScalarError::Value)?;
        match character {
            '"' => {
                if quoted && bytes.get(cursor + 1) == Some(&b'"') {
                    sink.push_char('"')?;
                    cursor += 2;
                } else {
                    quoted = !quoted;
                    cursor += 1;
                }
            },
            '\\' if !quoted => {
                let next = source[cursor + 1..]
                    .chars()
                    .next()
                    .ok_or(ScalarError::Value)?;
                sink.push_char(next)?;
                cursor += 1 + next.len_utf8();
            },
            '_' if !quoted => {
                let next = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
                sink.push_char(' ')?;
                cursor = next;
            },
            '@' if !quoted => {
                if let Some(value) = substitution {
                    sink.push_str(value)?;
                } else {
                    return Err(ScalarError::Value);
                }
                cursor += 1;
            },
            _ => {
                sink.push_char(character)?;
                cursor += character.len_utf8();
            },
        }
    }
    if quoted {
        Err(ScalarError::Value)
    } else {
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct ExponentSpec<'a> {
    marker: char,
    plus: bool,
    width: usize,
    pattern: &'a str,
}

#[derive(Clone, Copy)]
struct FractionSpec<'a> {
    denominator_digits: usize,
    pre_slash_pattern: &'a str,
    denominator_pattern: &'a str,
    whole_pattern: &'a str,
    numerator_pattern: &'a str,
    mixed: bool,
}

#[derive(Clone, Copy)]
struct NumericSpec<'a> {
    integer_pattern: &'a str,
    integer_required: usize,
    fraction_pattern: &'a str,
    fraction_total: usize,
    grouping: bool,
    percent: bool,
    exponent: Option<ExponentSpec<'a>>,
    fraction: Option<FractionSpec<'a>>,
    fraction_denominator_digits: Option<usize>,
}

fn integer_pattern<'a>(
    source: &'a str,
    checks: &mut FormatChecks<'_>,
) -> Result<&'a str, ScalarError> {
    let bytes = source.as_bytes();
    let mut end = bytes.len();
    for (index, byte) in bytes.iter().enumerate() {
        checks.checkpoint()?;
        if matches!(*byte, b'.' | b'/' | b'E' | b'e') {
            end = index;
            break;
        }
    }
    Ok(&source[..end])
}

fn fraction_pattern<'a>(
    source: &'a str,
    checks: &mut FormatChecks<'_>,
) -> Result<&'a str, ScalarError> {
    let bytes = source.as_bytes();
    let mut decimal = None;
    for (index, byte) in bytes.iter().enumerate() {
        checks.checkpoint()?;
        if *byte == b'.' {
            decimal = Some(index);
            break;
        }
    }
    let Some(decimal) = decimal else {
        return Ok("");
    };
    let mut end = bytes.len();
    for (index, byte) in bytes.iter().enumerate().skip(decimal + 1) {
        checks.checkpoint()?;
        if matches!(*byte, b'/' | b'E' | b'e') {
            end = index;
            break;
        }
    }
    Ok(&source[decimal + 1..end])
}

impl<'a> NumericSpec<'a> {
    fn parse(source: &'a str, checks: &mut FormatChecks<'_>) -> Result<Self, ScalarError> {
        let bytes = source.as_bytes();
        let mut cursor = 0usize;
        let mut decimal = false;
        let mut exponent = None;
        let mut fraction = None;
        let mut integer_required = 0usize;
        let mut integer_total = 0usize;
        let mut fraction_total = 0usize;
        let mut grouping = false;
        let mut percent = false;
        let mut in_exponent = false;
        let mut in_denominator = false;
        let mut denominator_digits = 0usize;
        let mut slash_position = None;
        let mut space_position = None;
        while cursor < bytes.len() {
            checks.checkpoint()?;
            match bytes[cursor] {
                b'0' | b'#' | b'?' => {
                    if in_exponent {
                        return Err(ScalarError::Value);
                    } else if in_denominator {
                        denominator_digits = denominator_digits.saturating_add(1);
                        if denominator_digits > MAX_FRACTION_DENOMINATOR_DIGITS {
                            return Err(ScalarError::Value);
                        }
                    } else if decimal {
                        fraction_total = fraction_total.saturating_add(1);
                    } else {
                        integer_total = integer_total.saturating_add(1);
                        if bytes[cursor] == b'0' {
                            integer_required = integer_required.saturating_add(1);
                        }
                    }
                    cursor += 1;
                },
                b'.' if !in_exponent && !in_denominator => {
                    if decimal {
                        return Err(ScalarError::Value);
                    }
                    decimal = true;
                    cursor += 1;
                },
                b',' if !in_exponent && !in_denominator && !decimal => {
                    if cursor == 0
                        || !matches!(
                            bytes.get(cursor.saturating_sub(1)),
                            Some(b'0' | b'#' | b'?')
                        )
                        || !matches!(bytes.get(cursor + 1), Some(b'0' | b'#' | b'?'))
                        || !matches!(bytes.get(cursor + 2), Some(b'0' | b'#' | b'?'))
                        || !matches!(bytes.get(cursor + 3), Some(b'0' | b'#' | b'?'))
                        || matches!(bytes.get(cursor + 4), Some(b'0' | b'#' | b'?'))
                    {
                        return Err(ScalarError::Value);
                    }
                    grouping = true;
                    cursor += 1;
                },
                b'%' => {
                    if percent {
                        return Err(ScalarError::Value);
                    }
                    percent = true;
                    cursor += 1;
                },
                b'E' | b'e' if !in_exponent && !in_denominator => {
                    let marker = char::from(bytes[cursor]);
                    cursor += 1;
                    let plus = match bytes.get(cursor) {
                        Some(b'+') => {
                            cursor += 1;
                            true
                        },
                        Some(b'-') => {
                            cursor += 1;
                            false
                        },
                        _ => false,
                    };
                    let start = cursor;
                    while bytes
                        .get(cursor)
                        .is_some_and(|byte| matches!(*byte, b'0' | b'#' | b'?'))
                    {
                        checks.checkpoint()?;
                        cursor += 1;
                    }
                    if start == cursor {
                        return Err(ScalarError::Value);
                    }
                    exponent = Some(ExponentSpec {
                        marker,
                        plus,
                        width: cursor - start,
                        pattern: &source[start..cursor],
                    });
                    in_exponent = true;
                },
                b'/' if !in_exponent => {
                    if fraction.is_some() || integer_total == 0 || decimal || exponent.is_some() {
                        return Err(ScalarError::Value);
                    }
                    fraction = Some(FractionSpec {
                        denominator_digits: 0,
                        pre_slash_pattern: "",
                        denominator_pattern: "",
                        whole_pattern: "",
                        numerator_pattern: "",
                        mixed: false,
                    });
                    in_denominator = true;
                    slash_position = Some(cursor);
                    cursor += 1;
                },
                b' ' => {
                    if decimal || in_exponent || in_denominator || space_position.is_some() {
                        return Err(ScalarError::Value);
                    }
                    space_position = Some(cursor);
                    cursor += 1;
                },
                b'\t' => return Err(ScalarError::Value),
                b'+' | b'-' => return Err(ScalarError::Value),
                _ => return Err(ScalarError::Value),
            }
        }
        if integer_total == 0 && fraction.is_none() {
            return Err(ScalarError::Value);
        }
        if space_position.is_some() && fraction.is_none() {
            return Err(ScalarError::Value);
        }
        if let Some(value) = fraction.as_mut() {
            value.denominator_digits = denominator_digits;
            if denominator_digits == 0 {
                return Err(ScalarError::Value);
            }
            let slash = slash_position.ok_or(ScalarError::Value)?;
            value.pre_slash_pattern = &source[..slash];
            value.denominator_pattern = &source[slash + 1..];
            let mut separator = None;
            for (index, byte) in value.pre_slash_pattern.bytes().enumerate() {
                checks.checkpoint()?;
                if byte == b' ' {
                    separator = Some(index);
                }
            }
            if let Some(separator) = separator {
                let whole = &value.pre_slash_pattern[..separator];
                let numerator = &value.pre_slash_pattern[separator + 1..];
                if whole.is_empty() || numerator.is_empty() {
                    return Err(ScalarError::Value);
                }
                value.whole_pattern = whole;
                value.numerator_pattern = numerator;
                value.mixed = true;
            }
        }
        Ok(Self {
            integer_pattern: integer_pattern(source, checks)?,
            integer_required,
            fraction_pattern: fraction_pattern(source, checks)?,
            fraction_total,
            grouping,
            percent,
            exponent,
            fraction,
            fraction_denominator_digits: fraction.map(|value| value.denominator_digits),
        })
    }
}

fn render_numeric_core(
    sink: &mut dyn Sink,
    number: f64,
    spec: &NumericSpec<'_>,
    negative: bool,
    fraction_parts: Option<&FractionParts>,
) -> Result<(), ScalarError> {
    let mut value = number;
    if spec.percent {
        value *= 100.0;
    }
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    if spec.fraction.is_some() {
        return render_fraction(
            sink,
            fraction_parts.ok_or(ScalarError::Number)?,
            spec.fraction.as_ref().ok_or(ScalarError::Number)?,
            negative,
        );
    }
    if let Some(exponent) = spec.exponent {
        return render_scientific(sink, value, spec, exponent, negative);
    }
    render_pattern_fixed(sink, value, spec, negative)
}

fn render_pattern_fixed(
    sink: &mut dyn Sink,
    value: f64,
    spec: &NumericSpec<'_>,
    negative: bool,
) -> Result<(), ScalarError> {
    let places = spec.fraction_total;
    let rounded = round_nearest(value, places as f64)?;
    if !rounded.is_finite() {
        return Err(ScalarError::Number);
    }
    let representation = DecimalRepresentation::from_number(rounded.abs())?;
    render_optional_fixed(sink, &representation, spec, negative && rounded != 0.0)
}

fn render_optional_fixed(
    sink: &mut dyn Sink,
    representation: &DecimalRepresentation,
    spec: &NumericSpec<'_>,
    negative: bool,
) -> Result<(), ScalarError> {
    if negative {
        sink.push_char('-')?;
    }
    let actual_integer_len = representation.integer_len();
    let integer_is_zero = actual_integer_len == 1 && representation.digit_at(0) == b'0';
    let actual_integer_len = if integer_is_zero && spec.integer_required == 0 {
        0
    } else {
        actual_integer_len
    };
    render_integer_pattern(sink, representation, spec, actual_integer_len)?;
    let mut nonzero_fraction = 0usize;
    for index in 0..spec.fraction_total {
        sink.checkpoint()?;
        let position = representation
            .decimal_position
            .checked_add(i32::try_from(index).map_err(|_| ScalarError::Number)?)
            .ok_or(ScalarError::Number)?;
        if representation.digit_at(position) != b'0' {
            nonzero_fraction = index.saturating_add(1);
        }
    }
    let mut fraction_alignment = false;
    for byte in spec.fraction_pattern.bytes() {
        sink.checkpoint()?;
        if matches!(byte, b'0' | b'?') {
            fraction_alignment = true;
        }
    }
    if nonzero_fraction != 0 || fraction_alignment {
        sink.push_char('.')?;
        let mut token_index = 0usize;
        for token in spec.fraction_pattern.bytes() {
            sink.checkpoint()?;
            if token == b'%' {
                sink.push_char('%')?;
                continue;
            }
            if !matches!(token, b'0' | b'#' | b'?') {
                continue;
            }
            let position = representation
                .decimal_position
                .checked_add(i32::try_from(token_index).map_err(|_| ScalarError::Number)?)
                .ok_or(ScalarError::Number)?;
            token_index = token_index.saturating_add(1);
            let digit = representation.digit_at(position);
            if token_index <= nonzero_fraction {
                sink.push_char(char::from(digit))?;
            } else {
                match token {
                    b'0' => sink.push_char('0')?,
                    b'?' => sink.push_char(' ')?,
                    b'#' => {},
                    _ => unreachable!(),
                }
            }
        }
    } else {
        for token in spec.fraction_pattern.bytes() {
            sink.checkpoint()?;
            if token == b'%' {
                sink.push_char('%')?;
            }
        }
    }
    Ok(())
}

fn render_integer_pattern(
    sink: &mut dyn Sink,
    representation: &DecimalRepresentation,
    spec: &NumericSpec<'_>,
    actual_integer_len: usize,
) -> Result<(), ScalarError> {
    let mut slots = 0usize;
    for byte in spec.integer_pattern.bytes() {
        sink.checkpoint()?;
        if matches!(byte, b'0' | b'#' | b'?') {
            slots = slots.saturating_add(1);
        }
    }
    let leading_missing = slots.saturating_sub(actual_integer_len);
    let extra_digits = actual_integer_len.saturating_sub(slots);
    let mut missing_zeroes = 0usize;
    let mut slot = 0usize;
    for token in spec.integer_pattern.bytes() {
        sink.checkpoint()?;
        if !matches!(token, b'0' | b'#' | b'?') {
            continue;
        }
        if slot < leading_missing && token == b'0' {
            missing_zeroes = missing_zeroes.saturating_add(1);
        }
        slot += 1;
    }
    let digit_count = actual_integer_len.saturating_add(missing_zeroes);
    let mut emitted_digits = 0usize;
    for position in 0..extra_digits {
        sink.checkpoint()?;
        push_grouped_digit(
            sink,
            representation.digit_at(i32::try_from(position).map_err(|_| ScalarError::Number)?),
            spec.grouping,
            &mut emitted_digits,
            digit_count,
        )?;
    }
    let mut token_slot = 0usize;
    for token in spec.integer_pattern.bytes() {
        sink.checkpoint()?;
        if token == b'%' {
            sink.push_char('%')?;
            continue;
        }
        if !matches!(token, b'0' | b'#' | b'?') {
            continue;
        }
        if token_slot < leading_missing {
            match token {
                b'0' => {
                    push_grouped_digit(sink, b'0', spec.grouping, &mut emitted_digits, digit_count)?
                },
                b'?' => sink.push_char(' ')?,
                b'#' => {},
                _ => unreachable!(),
            }
        } else {
            let position = extra_digits.saturating_add(token_slot - leading_missing);
            push_grouped_digit(
                sink,
                representation.digit_at(i32::try_from(position).map_err(|_| ScalarError::Number)?),
                spec.grouping,
                &mut emitted_digits,
                digit_count,
            )?;
        }
        token_slot += 1;
    }
    Ok(())
}

fn push_grouped_digit(
    sink: &mut dyn Sink,
    digit: u8,
    grouping: bool,
    emitted_digits: &mut usize,
    digit_count: usize,
) -> Result<(), ScalarError> {
    if grouping
        && *emitted_digits != 0
        && digit_count
            .saturating_sub(*emitted_digits)
            .is_multiple_of(3)
    {
        sink.push_char(',')?;
    }
    sink.push_char(char::from(digit))?;
    *emitted_digits = (*emitted_digits).saturating_add(1);
    Ok(())
}

fn render_scientific(
    sink: &mut dyn Sink,
    value: f64,
    spec: &NumericSpec<'_>,
    exponent: ExponentSpec<'_>,
    negative: bool,
) -> Result<(), ScalarError> {
    let magnitude = value.abs();
    let (mut mantissa, mut power) = scientific_components(magnitude)?;
    mantissa = round_nearest(mantissa, spec.fraction_total as f64)?;
    if mantissa >= 10.0 {
        mantissa /= 10.0;
        power = power.checked_add(1).ok_or(ScalarError::Number)?;
    }
    let representation = DecimalRepresentation::from_number(mantissa)?;
    let mantissa_spec = NumericSpec {
        integer_pattern: spec.integer_pattern,
        integer_required: spec.integer_required.max(1),
        fraction_pattern: spec.fraction_pattern,
        fraction_total: spec.fraction_total,
        grouping: false,
        percent: false,
        exponent: None,
        fraction: None,
        fraction_denominator_digits: None,
    };
    render_optional_fixed(
        sink,
        &representation,
        &mantissa_spec,
        negative && mantissa != 0.0,
    )?;
    sink.push_char(exponent.marker)?;
    if power < 0 {
        sink.push_char('-')?;
    } else if exponent.plus {
        sink.push_char('+')?;
    }
    let mut digits = [0_u8; 16];
    let mut length = 0usize;
    let mut magnitude = power.unsigned_abs();
    while magnitude != 0 {
        sink.checkpoint()?;
        if length == digits.len() {
            return Err(ScalarError::Number);
        }
        digits[length] = b'0' + u8::try_from(magnitude % 10).map_err(|_| ScalarError::Number)?;
        length += 1;
        magnitude /= 10;
    }
    if length == 0 {
        length = 1;
        digits[0] = b'0';
    }
    let width = exponent.width.max(length);
    let missing = width.saturating_sub(length);
    let mut token_index = 0usize;
    for token in exponent.pattern.bytes() {
        sink.checkpoint()?;
        if token_index == missing {
            break;
        }
        if !matches!(token, b'0' | b'#' | b'?') {
            continue;
        }
        token_index += 1;
        match token {
            b'0' => sink.push_char('0')?,
            b'?' => sink.push_char(' ')?,
            b'#' => {},
            _ => unreachable!(),
        }
    }
    for digit in digits[..length].iter().rev() {
        sink.checkpoint()?;
        sink.push_char(char::from(*digit))?;
    }
    Ok(())
}

fn scientific_components(value: f64) -> Result<(f64, i32), ScalarError> {
    if !value.is_finite() || value == 0.0 {
        return Ok((0.0, 0));
    }
    let representation = DecimalRepresentation::from_number(value)?;
    let first = (0..representation.length)
        .find(|&index| representation.digits[index] != b'0')
        .ok_or(ScalarError::Number)?;
    let first_index = first;
    let first = i32::try_from(first_index).map_err(|_| ScalarError::Number)?;
    let power = representation
        .decimal_position
        .checked_sub(first)
        .and_then(|value| value.checked_sub(1))
        .ok_or(ScalarError::Number)?;
    let mut mantissa = 0.0;
    let mut place = 1.0;
    for index in first_index..representation.length {
        if index != first_index {
            place *= 0.1;
        }
        mantissa += f64::from(representation.digits[index] - b'0') * place;
    }
    if !mantissa.is_finite() || !(1.0..10.0).contains(&mantissa) {
        return Err(ScalarError::Number);
    }
    Ok((mantissa, power))
}

#[derive(Clone, Copy)]
struct FractionParts {
    whole: f64,
    numerator: u64,
    denominator: u64,
}

fn fraction_parts(
    evaluator: &mut Evaluator<'_, '_, '_>,
    value: f64,
    denominator_digits: usize,
) -> EvaluationResult<Result<FractionParts, ScalarError>> {
    if !value.is_finite() || denominator_digits > MAX_FRACTION_DENOMINATOR_DIGITS {
        return Ok(Err(ScalarError::Number));
    }
    let magnitude = value.abs();
    let power = match u32::try_from(denominator_digits) {
        Ok(power) => power,
        Err(_) => return Ok(Err(ScalarError::Number)),
    };
    let max_denominator = match 10_u64
        .checked_pow(power)
        .and_then(|value| value.checked_sub(1))
    {
        Some(value) if value != 0 => value,
        _ => return Ok(Err(ScalarError::Number)),
    };
    let whole = magnitude.floor();
    let fractional = magnitude - whole;
    let (mut numerator, mut denominator) =
        match super::fraction::nearest(fractional, max_denominator, |work| {
            evaluator.charge_work(work)
        })? {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
    let mut whole = whole;
    if numerator == denominator {
        numerator = 0;
        denominator = 1;
        whole += 1.0;
    }
    Ok(Ok(FractionParts {
        whole,
        numerator,
        denominator,
    }))
}

fn render_fraction(
    sink: &mut dyn Sink,
    parts: &FractionParts,
    fraction: &FractionSpec<'_>,
    negative: bool,
) -> Result<(), ScalarError> {
    let whole = integer_from_f64(parts.whole)?;
    if negative && (whole != 0 || parts.numerator != 0) {
        sink.push_char('-')?;
    }
    if fraction.mixed {
        render_slot_pattern(sink, fraction.whole_pattern, whole, whole == 0)?;
        if parts.numerator != 0 {
            sink.push_char(' ')?;
            render_slot_pattern(sink, fraction.numerator_pattern, parts.numerator, true)?;
            sink.push_char('/')?;
            render_slot_pattern(sink, fraction.denominator_pattern, parts.denominator, false)?;
        } else {
            render_percent_literals(sink, fraction.numerator_pattern)?;
            render_percent_literals(sink, fraction.denominator_pattern)?;
        }
    } else {
        let improper = u128::from(whole)
            .checked_mul(u128::from(parts.denominator))
            .and_then(|value| value.checked_add(u128::from(parts.numerator)))
            .ok_or(ScalarError::Number)?;
        let improper = u64::try_from(improper).map_err(|_| ScalarError::Number)?;
        if improper != 0 {
            render_slot_pattern(sink, fraction.pre_slash_pattern, improper, false)?;
            sink.push_char('/')?;
            render_slot_pattern(sink, fraction.denominator_pattern, parts.denominator, false)?;
        } else {
            let mut has_required = false;
            for byte in fraction.pre_slash_pattern.bytes() {
                sink.checkpoint()?;
                if byte == b'0' {
                    has_required = true;
                }
            }
            if has_required {
                render_slot_pattern(sink, fraction.pre_slash_pattern, 0, false)?;
            } else {
                render_percent_literals(sink, fraction.pre_slash_pattern)?;
            }
            render_percent_literals(sink, fraction.denominator_pattern)?;
        }
    }
    Ok(())
}

fn integer_from_f64(value: f64) -> Result<u64, ScalarError> {
    if !value.is_finite() || value < 0.0 || value >= u64::MAX as f64 || value.fract() != 0.0 {
        return Err(ScalarError::Number);
    }
    Ok(value as u64)
}

fn render_slot_pattern(
    sink: &mut dyn Sink,
    pattern: &str,
    value: u64,
    omit_zero: bool,
) -> Result<(), ScalarError> {
    let mut digits = [0_u8; 20];
    let mut length = 0usize;
    let mut remaining = value;
    if remaining == 0 {
        digits[0] = b'0';
        length = 1;
    } else {
        while remaining != 0 {
            if length == digits.len() {
                return Err(ScalarError::Number);
            }
            digits[length] =
                b'0' + u8::try_from(remaining % 10).map_err(|_| ScalarError::Number)?;
            length += 1;
            remaining /= 10;
        }
    }
    let mut has_required = false;
    for byte in pattern.bytes() {
        sink.checkpoint()?;
        if byte == b'0' {
            has_required = true;
        }
    }
    if omit_zero && value == 0 && !has_required {
        length = 0;
    }
    let mut slots = 0usize;
    for byte in pattern.bytes() {
        sink.checkpoint()?;
        if matches!(byte, b'0' | b'#' | b'?') {
            slots = slots.saturating_add(1);
        }
    }
    let extra = length.saturating_sub(slots);
    for index in 0..extra {
        sink.checkpoint()?;
        sink.push_char(char::from(digits[length - index - 1]))?;
    }
    let leading_missing = slots.saturating_sub(length);
    let mut slot = 0usize;
    for byte in pattern.bytes() {
        sink.checkpoint()?;
        match byte {
            b'0' | b'#' | b'?' => {
                if slot < leading_missing {
                    match byte {
                        b'0' => sink.push_char('0')?,
                        b'?' => sink.push_char(' ')?,
                        b'#' => {},
                        _ => unreachable!(),
                    }
                } else {
                    let index = extra.saturating_add(slot - leading_missing);
                    sink.push_char(char::from(digits[length - index - 1]))?;
                }
                slot += 1;
            },
            b'%' => sink.push_char('%')?,
            byte => sink.push_char(char::from(byte))?,
        }
    }
    Ok(())
}

fn render_percent_literals(sink: &mut dyn Sink, pattern: &str) -> Result<(), ScalarError> {
    for byte in pattern.bytes() {
        sink.checkpoint()?;
        if byte == b'%' {
            sink.push_char('%')?;
        }
    }
    Ok(())
}

fn write_unsigned(sink: &mut dyn Sink, mut value: u64) -> Result<(), ScalarError> {
    let mut digits = [0_u8; 20];
    let mut length = 0usize;
    if value == 0 {
        return sink.push_char('0');
    }
    while value != 0 {
        digits[length] = b'0' + u8::try_from(value % 10).map_err(|_| ScalarError::Number)?;
        length += 1;
        value /= 10;
    }
    for digit in digits[..length].iter().rev() {
        sink.push_char(char::from(*digit))?;
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct DateParts {
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    weekday: usize,
}

fn date_parts(value: f64) -> Result<DateParts, ScalarError> {
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    let days = value.floor();
    let max_days = 3_000_000.0;
    if !(-max_days..=max_days).contains(&days) {
        return Err(ScalarError::Number);
    }
    let mut day_count = days as i64;
    let fraction = value - days;
    let mut seconds = (fraction * SECONDS_PER_DAY).round() as i64;
    if seconds >= 86_400 {
        seconds -= 86_400;
        day_count = day_count.checked_add(1).ok_or(ScalarError::Number)?;
    }
    let unix_days = day_count
        .checked_sub(SERIAL_EPOCH_TO_UNIX_DAYS)
        .ok_or(ScalarError::Number)?;
    let (year, month, day) = civil_from_days(unix_days);
    if !(1899..=9999).contains(&year) {
        return Err(ScalarError::Number);
    }
    let hour = u32::try_from(seconds / 3_600).map_err(|_| ScalarError::Number)?;
    let minute = u32::try_from((seconds % 3_600) / 60).map_err(|_| ScalarError::Number)?;
    let second = u32::try_from(seconds % 60).map_err(|_| ScalarError::Number)?;
    let weekday =
        usize::try_from((unix_days + 4).rem_euclid(7)).map_err(|_| ScalarError::Number)?;
    Ok(DateParts {
        year,
        month,
        day,
        hour,
        minute,
        second,
        weekday,
    })
}

// Howard Hinnant's proleptic-Gregorian civil-date conversion, expressed in
// terms of days since 1970-01-01 and kept integer-only for deterministic TEXT.
fn civil_from_days(days: i64) -> (i32, u32, u32) {
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if month <= 2 { 1 } else { 0 };
    (
        i32::try_from(year).unwrap_or(i32::MAX),
        u32::try_from(month).unwrap_or(u32::MAX),
        u32::try_from(day).unwrap_or(u32::MAX),
    )
}

fn render_date(sink: &mut dyn Sink, source: &str, date: &DateParts) -> Result<(), ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = 0usize;
    let mut quoted = false;
    let has_am_pm = has_am_pm(sink, source)?;
    let mut previous_field = None;
    while cursor < bytes.len() {
        sink.checkpoint()?;
        let character = source[cursor..].chars().next().ok_or(ScalarError::Value)?;
        if quoted {
            if character == '"' {
                if bytes.get(cursor + 1) == Some(&b'"') {
                    sink.push_char('"')?;
                    cursor += 2;
                } else {
                    quoted = false;
                    cursor += 1;
                }
            } else {
                sink.push_char(character)?;
                cursor += character.len_utf8();
            }
            continue;
        }
        match character {
            '"' => {
                quoted = true;
                cursor += 1;
            },
            '\\' => {
                let next = source[cursor + 1..]
                    .chars()
                    .next()
                    .ok_or(ScalarError::Value)?;
                sink.push_char(next)?;
                cursor += 1 + next.len_utf8();
            },
            '_' => {
                let next = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
                sink.push_char(' ')?;
                cursor = next;
            },
            'A' | 'a' | 'P' | 'p' if starts_am_pm(source, cursor) => {
                let value = if date.hour < 12 { "AM" } else { "PM" };
                sink.push_str(value)?;
                cursor += 5;
            },
            'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => {
                let end = same_letter_end(source, cursor, character);
                let length = source[cursor..end].chars().count();
                let following_field = if character.eq_ignore_ascii_case(&'m') {
                    next_date_field(sink, source, end)?
                } else {
                    None
                };
                render_date_token(
                    sink,
                    previous_field,
                    following_field,
                    character,
                    length,
                    date,
                    has_am_pm,
                )?;
                previous_field = Some(character.to_ascii_lowercase());
                cursor = end;
            },
            _ => {
                sink.push_char(character)?;
                cursor += character.len_utf8();
            },
        }
    }
    if quoted {
        Err(ScalarError::Value)
    } else {
        Ok(())
    }
}

fn render_date_token(
    sink: &mut dyn Sink,
    previous_field: Option<char>,
    following_field: Option<char>,
    character: char,
    length: usize,
    date: &DateParts,
    has_am_pm: bool,
) -> Result<(), ScalarError> {
    match character.to_ascii_lowercase() {
        'y' => match length {
            1 => write_unsigned(
                sink,
                u64::try_from(date.year).map_err(|_| ScalarError::Number)?,
            )?,
            2 => write_padded(
                sink,
                u64::try_from(date.year.rem_euclid(100)).map_err(|_| ScalarError::Number)?,
                2,
            )?,
            4 => write_padded(
                sink,
                u64::try_from(date.year).map_err(|_| ScalarError::Number)?,
                4,
            )?,
            _ => return Err(ScalarError::Value),
        },
        'm' => {
            if date_token_is_minute(previous_field, following_field) {
                if length == 1 {
                    write_unsigned(sink, u64::from(date.minute))?;
                } else {
                    write_padded(sink, u64::from(date.minute), 2)?;
                }
            } else {
                match length {
                    1 => write_unsigned(sink, u64::from(date.month))?,
                    2 => write_padded(sink, u64::from(date.month), 2)?,
                    3 => sink.push_str(MONTH_SHORT[date.month.saturating_sub(1) as usize])?,
                    4 => sink.push_str(MONTH_LONG[date.month.saturating_sub(1) as usize])?,
                    _ => return Err(ScalarError::Value),
                }
            }
        },
        'd' => match length {
            1 => write_unsigned(sink, u64::from(date.day))?,
            2 => write_padded(sink, u64::from(date.day), 2)?,
            3 => sink.push_str(WEEKDAY_SHORT[date.weekday])?,
            4 => sink.push_str(WEEKDAY_LONG[date.weekday])?,
            _ => return Err(ScalarError::Value),
        },
        'h' => {
            let hour = if has_am_pm {
                let value = date.hour % 12;
                if value == 0 { 12 } else { value }
            } else {
                date.hour
            };
            if length == 1 {
                write_unsigned(sink, u64::from(hour))?;
            } else {
                write_padded(sink, u64::from(hour), 2)?;
            }
        },
        's' => {
            if length == 1 {
                write_unsigned(sink, u64::from(date.second))?;
            } else {
                write_padded(sink, u64::from(date.second), 2)?;
            }
        },
        _ => return Err(ScalarError::Value),
    }
    Ok(())
}

fn has_am_pm(sink: &mut dyn Sink, source: &str) -> Result<bool, ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = 0usize;
    let mut quoted = false;
    while cursor < bytes.len() {
        sink.checkpoint()?;
        let character = match source[cursor..].chars().next() {
            Some(character) => character,
            None => return Ok(false),
        };
        if quoted {
            if character == '"' {
                if bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                } else {
                    quoted = false;
                    cursor += 1;
                }
            } else {
                cursor += character.len_utf8();
            }
            continue;
        }
        match character {
            '"' => {
                quoted = true;
                cursor += 1;
            },
            '\\' | '_' => {
                cursor = next_char_end(source, cursor + 1).unwrap_or(source.len());
            },
            'A' | 'a' | 'P' | 'p' if starts_am_pm(source, cursor) => return Ok(true),
            _ => cursor += character.len_utf8(),
        }
    }
    Ok(false)
}

fn next_date_field(
    sink: &mut dyn Sink,
    source: &str,
    start: usize,
) -> Result<Option<char>, ScalarError> {
    let bytes = source.as_bytes();
    let mut cursor = start;
    let mut quoted = false;
    while cursor < bytes.len() {
        sink.checkpoint()?;
        let character = source[cursor..].chars().next().ok_or(ScalarError::Value)?;
        if quoted {
            if character == '"' {
                if bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                } else {
                    quoted = false;
                    cursor += 1;
                }
            } else {
                cursor += character.len_utf8();
            }
            continue;
        }
        match character {
            '"' => {
                quoted = true;
                cursor += 1;
            },
            '\\' | '_' => {
                cursor = next_char_end(source, cursor + 1).ok_or(ScalarError::Value)?;
            },
            'A' | 'a' | 'P' | 'p' if starts_am_pm(source, cursor) => cursor += 5,
            'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => {
                return Ok(Some(character.to_ascii_lowercase()));
            },
            _ => cursor += character.len_utf8(),
        }
    }
    Ok(None)
}

fn date_token_is_minute(previous_field: Option<char>, following_field: Option<char>) -> bool {
    previous_field.is_some_and(|value| matches!(value, 'h' | 's'))
        || following_field.is_some_and(|value| matches!(value, 'h' | 's'))
}

fn write_padded(sink: &mut dyn Sink, value: u64, width: usize) -> Result<(), ScalarError> {
    let mut digits = [0_u8; 20];
    let mut length = 0usize;
    let mut value = value;
    if value == 0 {
        digits[0] = b'0';
        length = 1;
    } else {
        while value != 0 {
            digits[length] = b'0' + u8::try_from(value % 10).map_err(|_| ScalarError::Number)?;
            length += 1;
            value /= 10;
        }
    }
    for _ in length..width {
        sink.push_char('0')?;
    }
    for digit in digits[..length].iter().rev() {
        sink.push_char(char::from(*digit))?;
    }
    Ok(())
}

const MONTH_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
const MONTH_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const WEEKDAY_SHORT: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const WEEKDAY_LONG: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];
