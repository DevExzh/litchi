//! OpenFormula value inspection and text-to-number kernels.
//!
//! This module deliberately sits beside the ordinary scalar coercion bridge.
//! `Input::Empty` must reach the predicates before conversion to zero, while
//! formula errors remain values for the inspection predicates and
//! `ERROR.TYPE`/`TYPE`.  The value evaluator uses the same catalog and raw
//! kernel after it has selected one streamed element from an array or
//! reference.

pub(super) mod parse_value;

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextValue, UnsupportedKind,
    WorkingValue, map_execution_error,
};

/// An argument before the ordinary scalar bridge erases an empty cell.
///
/// `Missing` is used by the value evaluator for an omitted optional argument;
/// it is kept separate from an empty cell so NUMBERVALUE can apply its
/// defaults without treating a supplied empty separator as omitted.
pub(super) enum Input<'a> {
    Empty,
    Missing,
    Value(WorkingValue<'a>),
}

/// The shared catalog entry for the value-inspection family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Function {
    ErrorType,
    IsBlank,
    IsErr,
    IsError,
    IsEven,
    IsLogical,
    IsNa,
    IsNonText,
    IsNumber,
    IsOdd,
    IsText,
    N,
    Na,
    NumberValue,
    Type,
    Value,
}

impl Function {
    /// Look up a function without allocating or normalizing its spelling.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        // Most scalar functions never enter this family.  Keep the common
        // dispatch path to one byte test before the case-insensitive catalog
        // comparisons below.
        if !matches!(
            name.as_bytes().first().copied(),
            Some(b'E' | b'e' | b'I' | b'i' | b'N' | b'n' | b'T' | b't' | b'V' | b'v')
        ) {
            return None;
        }
        Some(if name.eq_ignore_ascii_case("ERROR.TYPE") {
            Self::ErrorType
        } else if name.eq_ignore_ascii_case("ISBLANK") {
            Self::IsBlank
        } else if name.eq_ignore_ascii_case("ISERR") {
            Self::IsErr
        } else if name.eq_ignore_ascii_case("ISERROR") {
            Self::IsError
        } else if name.eq_ignore_ascii_case("ISEVEN") {
            Self::IsEven
        } else if name.eq_ignore_ascii_case("ISLOGICAL") {
            Self::IsLogical
        } else if name.eq_ignore_ascii_case("ISNA") {
            Self::IsNa
        } else if name.eq_ignore_ascii_case("ISNONTEXT") {
            Self::IsNonText
        } else if name.eq_ignore_ascii_case("ISNUMBER") {
            Self::IsNumber
        } else if name.eq_ignore_ascii_case("ISODD") {
            Self::IsOdd
        } else if name.eq_ignore_ascii_case("ISTEXT") {
            Self::IsText
        } else if name.eq_ignore_ascii_case("N") {
            Self::N
        } else if name.eq_ignore_ascii_case("NA") {
            Self::Na
        } else if name.eq_ignore_ascii_case("NUMBERVALUE") {
            Self::NumberValue
        } else if name.eq_ignore_ascii_case("TYPE") {
            Self::Type
        } else if name.eq_ignore_ascii_case("VALUE") {
            Self::Value
        } else {
            return None;
        })
    }

    /// Return whether this function accepts the supplied number of arguments.
    pub(super) const fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::NumberValue => count >= 1 && count <= 3,
            Self::Na => count == 0,
            Self::ErrorType
            | Self::IsBlank
            | Self::IsErr
            | Self::IsError
            | Self::IsEven
            | Self::IsLogical
            | Self::IsNa
            | Self::IsNonText
            | Self::IsNumber
            | Self::IsOdd
            | Self::IsText
            | Self::N
            | Self::Type
            | Self::Value => count == 1,
        }
    }

    const fn inspects_formula_error(self) -> bool {
        matches!(
            self,
            Self::ErrorType
                | Self::IsBlank
                | Self::IsErr
                | Self::IsError
                | Self::IsLogical
                | Self::IsNa
                | Self::IsNonText
                | Self::IsNumber
                | Self::IsText
                | Self::Type
        )
    }

    /// Whether a formula error is an input value for this function rather
    /// than an error to propagate before conversion.
    pub(super) const fn inspects_errors(self) -> bool {
        self.inspects_formula_error()
    }
}

pub(super) fn is_inspection_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

/// Apply a function from the scalar evaluator's already-evaluated value stack.
/// Values are popped right-to-left, then restored to source order for error
/// precedence and optional-argument handling.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = Function::from_name(name)
        .ok_or(EvaluationFailure::Unsupported(UnsupportedKind::Function))?;
    let count = node.child_count();
    if !function.valid_arity(count) {
        return evaluator.finish_invalid_arity(node);
    }

    let mut inputs: [Option<Input<'a>>; 3] = std::array::from_fn(|_| None);
    for index in (0..count).rev() {
        // The scalar scheduler visits a `Kind::Missing` child and pushes a
        // formula #VALUE placeholder so the eager stack stays balanced.  Keep
        // that pop, but restore the AST distinction needed by optional
        // NUMBERVALUE separators before entering the raw kernel.
        let value = evaluator.pop_value()?;
        inputs[index] = Some(
            if node.child(index).is_some_and(|child| child.is_missing()) {
                Input::Missing
            } else {
                Input::Value(value)
            },
        );
    }
    let result = apply_inputs(
        evaluator,
        function,
        inputs.into_iter().take(count).map(|input| match input {
            Some(input) => input,
            None => unreachable!("inspection argument storage was incomplete"),
        }),
    )?;
    evaluator.push_value(result)
}

/// Apply a catalog entry to raw values supplied by the scalar or value VM.
///
/// The iterator is consumed into a fixed three-slot stack.  This is enough for
/// NUMBERVALUE's optional separators and avoids a second heap-backed argument
/// vector at the scalar/value seam.
pub(super) fn apply_inputs<'a, I>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    function: Function,
    arguments: I,
) -> EvaluationResult<WorkingValue<'a>>
where
    I: IntoIterator<Item = Input<'a>>,
{
    evaluator.charge_work(1)?;
    let mut inputs: [Option<Input<'a>>; 3] = std::array::from_fn(|_| None);
    let mut count = 0usize;
    for input in arguments {
        if count == inputs.len() {
            return Err(EvaluationFailure::InvalidExpression(
                "inspection function received too many arguments",
            ));
        }
        inputs[count] = Some(input);
        count += 1;
    }
    if !function.valid_arity(count) {
        return Ok(WorkingValue::Error(ScalarError::Value));
    }

    // The predicates intentionally inspect formula errors.  Every converting
    // function gets the ordinary leftmost-error rule before conversion, so a
    // generated conversion error cannot hide an admitted formula error.
    if !function.inspects_errors()
        && let Some(error) = first_formula_error(&inputs, count)
    {
        return Ok(WorkingValue::Error(error));
    }

    match function {
        Function::ErrorType
        | Function::IsBlank
        | Function::IsErr
        | Function::IsError
        | Function::IsLogical
        | Function::IsNa
        | Function::IsNonText
        | Function::IsNumber
        | Function::IsText
        | Function::Type => apply_raw(function, inputs[0].as_ref()),
        Function::IsEven | Function::IsOdd => {
            let input = inputs[0]
                .take()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "parity argument storage was incomplete",
                ))?;
            apply_parity(evaluator, function, input)
        },
        Function::N => {
            let input = inputs[0]
                .take()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "N argument storage was incomplete",
                ))?;
            apply_n(evaluator, input)
        },
        Function::Na => Ok(WorkingValue::Error(ScalarError::NotAvailable)),
        Function::NumberValue => apply_number_value(evaluator, inputs, count),
        Function::Value => {
            let input = inputs[0]
                .take()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "VALUE argument storage was incomplete",
                ))?;
            apply_value(evaluator, input)
        },
    }
}

fn first_formula_error(inputs: &[Option<Input<'_>>; 3], count: usize) -> Option<ScalarError> {
    inputs[..count]
        .iter()
        .find_map(|input| match input.as_ref()? {
            Input::Value(WorkingValue::Error(error)) => Some(*error),
            Input::Empty | Input::Missing | Input::Value(_) => None,
        })
}

fn apply_raw<'a>(
    function: Function,
    input: Option<&Input<'_>>,
) -> EvaluationResult<WorkingValue<'a>> {
    // This helper only returns static literals.  The caller immediately
    // coerces the lifetime to the evaluator's result lifetime; no source text
    // is manufactured by the raw predicates.
    let input = input.ok_or(EvaluationFailure::InvalidExpression(
        "inspection argument storage was incomplete",
    ))?;
    let result = match function {
        Function::ErrorType => match input {
            Input::Value(WorkingValue::Error(error)) => {
                WorkingValue::Number(error_type_number(*error))
            },
            Input::Missing => WorkingValue::Error(ScalarError::Value),
            Input::Empty | Input::Value(_) => WorkingValue::Error(ScalarError::Value),
        },
        Function::IsBlank => WorkingValue::Logical(matches!(input, Input::Empty)),
        Function::IsErr => WorkingValue::Logical(matches!(
            input,
            Input::Value(WorkingValue::Error(error)) if *error != ScalarError::NotAvailable
        )),
        Function::IsError => {
            WorkingValue::Logical(matches!(input, Input::Value(WorkingValue::Error(_))))
        },
        Function::IsLogical => {
            WorkingValue::Logical(matches!(input, Input::Value(WorkingValue::Logical(_))))
        },
        Function::IsNa => WorkingValue::Logical(matches!(
            input,
            Input::Value(WorkingValue::Error(ScalarError::NotAvailable))
        )),
        Function::IsNonText => {
            WorkingValue::Logical(!matches!(input, Input::Value(WorkingValue::Text(_))))
        },
        Function::IsNumber => WorkingValue::Logical(matches!(
            input,
            Input::Value(WorkingValue::Number(_) | WorkingValue::Complex(_))
        )),
        Function::IsText => {
            WorkingValue::Logical(matches!(input, Input::Value(WorkingValue::Text(_))))
        },
        Function::Type => WorkingValue::Number(type_number(input)),
        Function::IsEven
        | Function::IsOdd
        | Function::N
        | Function::Na
        | Function::NumberValue
        | Function::Value => {
            return Err(EvaluationFailure::InvalidExpression(
                "converting inspection function reached raw kernel",
            ));
        },
    };
    Ok(result)
}

fn error_type_number(error: ScalarError) -> f64 {
    match error {
        ScalarError::Null => 1.0,
        ScalarError::DivisionByZero => 2.0,
        ScalarError::Value => 3.0,
        ScalarError::Reference => 4.0,
        ScalarError::Name => 5.0,
        ScalarError::Number => 6.0,
        ScalarError::NotAvailable => 7.0,
    }
}

fn type_number(input: &Input<'_>) -> f64 {
    match input {
        // The bounded profile uses the spreadsheet-compatible Number code for
        // a blank scalar.  Reference/array descriptors are classified by the
        // value VM before this scalar kernel is entered.
        Input::Empty => 1.0,
        Input::Missing => 16.0,
        Input::Value(WorkingValue::Number(_) | WorkingValue::Complex(_)) => 1.0,
        Input::Value(WorkingValue::Text(_)) => 2.0,
        Input::Value(WorkingValue::Logical(_)) => 4.0,
        Input::Value(WorkingValue::Error(_)) => 16.0,
    }
}

fn apply_n<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    input: Input<'a>,
) -> EvaluationResult<WorkingValue<'a>> {
    Ok(match input {
        Input::Empty => WorkingValue::Number(0.0),
        Input::Missing => WorkingValue::Error(ScalarError::Value),
        Input::Value(WorkingValue::Number(value)) => WorkingValue::Number(value),
        Input::Value(WorkingValue::Logical(value)) => {
            WorkingValue::Number(if value { 1.0 } else { 0.0 })
        },
        Input::Value(WorkingValue::Text(text)) => {
            evaluator.charge_bytes(text.text.len())?;
            WorkingValue::Number(0.0)
        },
        // §4.11.2 includes Complex in Number.  Preserve it as a scalar value;
        // the public profile can still distinguish it from a real Number.
        Input::Value(WorkingValue::Complex(value)) => WorkingValue::Complex(value),
        Input::Value(WorkingValue::Error(error)) => WorkingValue::Error(error),
    })
}

fn apply_parity<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    function: Function,
    input: Input<'a>,
) -> EvaluationResult<WorkingValue<'a>> {
    let value = match input {
        // Conversion to Number maps an empty referenced cell to zero (§6.3.5).
        Input::Empty => 0.0,
        Input::Missing => return Ok(WorkingValue::Error(ScalarError::Value)),
        Input::Value(value) => match parity_number(evaluator, value)? {
            Ok(value) => value,
            Err(error) => return Ok(WorkingValue::Error(error)),
        },
    };
    let truncated = value.trunc();
    let even = truncated.rem_euclid(2.0) == 0.0;
    Ok(WorkingValue::Logical(if function == Function::IsEven {
        even
    } else {
        !even
    }))
}

fn parity_number<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match value {
        WorkingValue::Number(value) if value.is_finite() => Ok(Ok(value)),
        WorkingValue::Number(_) => Ok(Err(ScalarError::Number)),
        WorkingValue::Logical(value) => Ok(Ok(if value { 1.0 } else { 0.0 })),
        WorkingValue::Text(text) => {
            evaluator.charge_bytes(text.text.len())?;
            match fast_float2::parse::<f64, _>(text.text.as_ref()) {
                Ok(value) if value.is_finite() => Ok(Ok(value)),
                Ok(_) => Ok(Err(ScalarError::Number)),
                Err(_) if valid_xsd_float(evaluator, text.text.as_ref())? => {
                    Ok(Err(ScalarError::Number))
                },
                Err(_) => Ok(Err(ScalarError::Value)),
            }
        },
        WorkingValue::Error(error) => Ok(Err(error)),
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
    }
}

fn number_value_text_argument<'a>(input: Input<'a>) -> Result<TextValue<'a>, ScalarError> {
    match input {
        Input::Value(WorkingValue::Text(text)) => Ok(text),
        Input::Value(WorkingValue::Error(error)) => Err(error),
        Input::Empty | Input::Missing => Ok(TextValue::borrowed("")),
        Input::Value(
            WorkingValue::Number(_) | WorkingValue::Logical(_) | WorkingValue::Complex(_),
        ) => Err(ScalarError::Value),
    }
}

/// Validate NUMBERVALUE's separator metadata without touching its source
/// text.  The value VM uses this predicate for zero-read descriptor
/// preflight when both separator arguments are already known.
pub(super) fn valid_number_value_separators(decimal: &str, group: &str) -> bool {
    decimal.chars().count() == 1 && !group.contains(decimal)
}

fn apply_number_value<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    mut inputs: [Option<Input<'a>>; 3],
    count: usize,
) -> EvaluationResult<WorkingValue<'a>> {
    let text = match number_value_text_argument(inputs[0].take().ok_or(
        EvaluationFailure::InvalidExpression("NUMBERVALUE text argument storage was incomplete"),
    )?) {
        Ok(text) => text,
        Err(error) => return Ok(WorkingValue::Error(error)),
    };
    let decimal = if count >= 2 {
        match inputs[1]
            .take()
            .ok_or(EvaluationFailure::InvalidExpression(
                "NUMBERVALUE decimal separator storage was incomplete",
            ))? {
            Input::Missing => TextValue::borrowed("."),
            input => match number_value_text_argument(input) {
                Ok(value) => value,
                Err(error) => return Ok(WorkingValue::Error(error)),
            },
        }
    } else {
        TextValue::borrowed(".")
    };
    if count >= 2 {
        evaluator.charge_bytes(decimal.text.len())?;
    }
    let group = if count >= 3 {
        match inputs[2]
            .take()
            .ok_or(EvaluationFailure::InvalidExpression(
                "NUMBERVALUE group separator storage was incomplete",
            ))? {
            Input::Missing => TextValue::borrowed(","),
            input => match number_value_text_argument(input) {
                Ok(value) => value,
                Err(error) => return Ok(WorkingValue::Error(error)),
            },
        }
    } else {
        TextValue::borrowed(",")
    };
    if count >= 3 {
        evaluator.charge_bytes(group.text.len())?;
    }

    let decimal = decimal.text.as_ref();
    let group = group.text.as_ref();
    if !valid_number_value_separators(decimal, group) {
        return Ok(WorkingValue::Error(ScalarError::Value));
    }
    let decimal = decimal
        .chars()
        .next()
        .ok_or(EvaluationFailure::InvalidExpression(
            "validated NUMBERVALUE separator was empty",
        ))?;
    parse_number_value(evaluator, text.text.as_ref(), decimal, group)
}

fn parse_number_value<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    source: &str,
    decimal: char,
    group: &str,
) -> EvaluationResult<WorkingValue<'a>> {
    evaluator.charge_bytes(source.len())?;
    let first_decimal = source.find(decimal);
    let capacity = source
        .len()
        .checked_add(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "NUMBERVALUE text length overflow",
        ))?;
    let reservation = evaluator.reserve_storage(capacity, "formula NUMBERVALUE scratch")?;
    let mut normalized = String::new();
    if capacity != 0 {
        normalized
            .try_reserve_exact(capacity)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula NUMBERVALUE scratch",
                source,
            })?;
    }

    // `str::split` uses a linear substring searcher and removes each
    // non-overlapping group occurrence in one pass.  The old character loop
    // called `starts_with(group)` at every source position, which made a
    // repeated long group prefix quadratic in the separator length.  Keep
    // the decimal split separate because group separators after the first
    // decimal are part of the lexical value and must remain for validation.
    let mut processed = 0usize;
    let mut next_check = 4096usize;
    if let Some(decimal_offset) = first_decimal {
        let (prefix, remainder) = source.split_at(decimal_offset);
        append_group_filtered(
            evaluator,
            &mut normalized,
            prefix,
            group,
            &mut processed,
            &mut next_check,
        )?;
        append_normalized_segment(
            evaluator,
            &mut normalized,
            ".",
            decimal.len_utf8(),
            &mut processed,
            &mut next_check,
        )?;
        let suffix = &remainder[decimal.len_utf8()..];
        append_normalized_segment(
            evaluator,
            &mut normalized,
            suffix,
            suffix.len(),
            &mut processed,
            &mut next_check,
        )?;
    } else {
        append_group_filtered(
            evaluator,
            &mut normalized,
            source,
            group,
            &mut processed,
            &mut next_check,
        )?;
    }

    if normalized.starts_with('.') {
        normalized.insert(0, '0');
    }
    let mut percent_count = 0u32;
    while normalized.ends_with('%') {
        normalized.pop();
        percent_count = percent_count.saturating_add(1);
        if percent_count % 4096 == 0 {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
        }
    }

    evaluator
        .context
        .execution
        .check()
        .map_err(map_execution_error)?;
    let result = if !valid_xsd_float(evaluator, &normalized)? {
        WorkingValue::Error(ScalarError::Value)
    } else {
        match fast_float2::parse::<f64, _>(&normalized) {
            Ok(value) if value.is_finite() => {
                let mut value = value;
                for index in 0..percent_count {
                    if index % 4096 == 0 {
                        evaluator
                            .context
                            .execution
                            .check()
                            .map_err(map_execution_error)?;
                    }
                    value /= 100.0;
                }
                if value.is_finite() {
                    WorkingValue::Number(value)
                } else {
                    WorkingValue::Error(ScalarError::Number)
                }
            },
            Ok(_) => WorkingValue::Error(ScalarError::Number),
            // The lexical gate has already accepted the xsd:float form. A
            // parser failure here means it is outside the finite Number
            // domain (INF, NaN, or an overflowing exponent).
            Err(_) => WorkingValue::Error(ScalarError::Number),
        }
    };
    drop(normalized);
    drop(reservation);
    Ok(result)
}

/// Append a source segment while keeping cancellation checks tied to source
/// bytes rather than output bytes.  Group separators are omitted from the
/// output but still count as scanned input for the periodic check.
fn append_normalized_segment<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    normalized: &mut String,
    segment: &str,
    source_bytes: usize,
    processed: &mut usize,
    next_check: &mut usize,
) -> EvaluationResult<()> {
    let mut offset = 0usize;
    while offset < segment.len() {
        let absolute_offset =
            processed
                .checked_add(offset)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "NUMBERVALUE source length overflow",
                ))?;
        if absolute_offset >= *next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            *next_check = absolute_offset.saturating_add(4096);
        }
        let character =
            segment[offset..]
                .chars()
                .next()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "NUMBERVALUE source was not valid UTF-8",
                ))?;
        let width = character.len_utf8();
        if !matches!(character, ' ' | '\t' | '\n' | '\r') {
            normalized.push(character);
        }
        offset += width;
    }
    *processed =
        processed
            .checked_add(source_bytes)
            .ok_or(EvaluationFailure::InvalidExpression(
                "NUMBERVALUE source length overflow",
            ))?;
    if *processed >= *next_check {
        evaluator
            .context
            .execution
            .check()
            .map_err(map_execution_error)?;
        *next_check = processed.saturating_add(4096);
    }
    Ok(())
}

/// Append `source` after removing every non-overlapping `group` occurrence.
/// An empty group is a valid NUMBERVALUE profile and means no removal.
fn append_group_filtered<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    normalized: &mut String,
    source: &str,
    group: &str,
    processed: &mut usize,
    next_check: &mut usize,
) -> EvaluationResult<()> {
    if group.is_empty() {
        return append_normalized_segment(
            evaluator,
            normalized,
            source,
            source.len(),
            processed,
            next_check,
        );
    }

    let mut parts = source.split(group).peekable();
    while let Some(part) = parts.next() {
        let separator_bytes = if parts.peek().is_some() {
            group.len()
        } else {
            0
        };
        let source_bytes =
            part.len()
                .checked_add(separator_bytes)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "NUMBERVALUE source length overflow",
                ))?;
        append_normalized_segment(
            evaluator,
            normalized,
            part,
            source_bytes,
            processed,
            next_check,
        )?;
    }
    Ok(())
}

/// Validate the xsd:float lexical subset used by NUMBERVALUE.
/// `INF`, `-INF`, and `NaN` are valid lexical spellings but are mapped to
/// `#NUM!` after validation because this evaluator publishes finite Numbers.
fn valid_xsd_float(evaluator: &mut Evaluator<'_, '_, '_>, value: &str) -> EvaluationResult<bool> {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return Ok(true);
    }
    let bytes = value.as_bytes();
    if bytes.is_empty() {
        return Ok(false);
    }
    let mut index = usize::from(matches!(bytes[0], b'+' | b'-'));
    if index == bytes.len() {
        return Ok(false);
    }
    let mut digits_before = 0usize;
    let mut next_check = 4096usize;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        if index >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = index.saturating_add(4096);
        }
        digits_before += 1;
        index += 1;
    }
    let mut digits_after = 0usize;
    if index < bytes.len() && bytes[index] == b'.' {
        index += 1;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            if index >= next_check {
                evaluator
                    .context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = index.saturating_add(4096);
            }
            digits_after += 1;
            index += 1;
        }
    }
    if digits_before == 0 && digits_after == 0 {
        return Ok(false);
    }
    if index < bytes.len() && matches!(bytes[index], b'e' | b'E') {
        index += 1;
        if index < bytes.len() && matches!(bytes[index], b'+' | b'-') {
            index += 1;
        }
        let exponent_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            if index >= next_check {
                evaluator
                    .context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = index.saturating_add(4096);
            }
            index += 1;
        }
        if exponent_start == index {
            return Ok(false);
        }
    }
    Ok(index == bytes.len())
}

fn apply_value<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    input: Input<'a>,
) -> EvaluationResult<WorkingValue<'a>> {
    let text = match input {
        Input::Empty => return Ok(WorkingValue::Number(0.0)),
        Input::Missing => return Ok(WorkingValue::Error(ScalarError::Value)),
        Input::Value(WorkingValue::Error(error)) => {
            return Ok(WorkingValue::Error(error));
        },
        Input::Value(WorkingValue::Text(text)) => text,
        Input::Value(
            WorkingValue::Number(_) | WorkingValue::Logical(_) | WorkingValue::Complex(_),
        ) => return Ok(WorkingValue::Error(ScalarError::Value)),
    };
    evaluator.charge_bytes(text.text.len())?;
    evaluator
        .context
        .execution
        .check()
        .map_err(map_execution_error)?;
    let parsed_value = if parse_value::needs_scratch(text.text.as_ref()) {
        // Grouped VALUE input needs a complete normalized byte buffer for a
        // correctly rounded conversion.  Reserve the buffer against the
        // evaluator's memory budget before constructing it, and release both
        // the allocation and its lease before the final cancellation fence.
        let scratch_bytes = text.text.len();
        let reservation = evaluator.reserve_storage(scratch_bytes, "formula VALUE scratch")?;
        let mut scratch = Vec::new();
        scratch.try_reserve_exact(scratch_bytes).map_err(|source| {
            EvaluationFailure::Allocation {
                resource: "formula VALUE scratch",
                source,
            }
        })?;
        scratch.resize(scratch_bytes, 0);
        evaluator
            .context
            .execution
            .check()
            .map_err(map_execution_error)?;
        let parsed = parse_value::parse_value_with_scratch(text.text.as_ref(), &mut scratch);
        drop(scratch);
        drop(reservation);
        evaluator
            .context
            .execution
            .check()
            .map_err(map_execution_error)?;
        parsed
    } else {
        parse_value::parse_value(text.text.as_ref())
    };
    let parsed = match parsed_value {
        Ok(value) if value.is_finite() => Ok(WorkingValue::Number(value)),
        Ok(_) => Ok(WorkingValue::Error(ScalarError::Number)),
        Err(error) => Ok(WorkingValue::Error(error)),
    };
    evaluator
        .context
        .execution
        .check()
        .map_err(map_execution_error)?;
    parsed
}
