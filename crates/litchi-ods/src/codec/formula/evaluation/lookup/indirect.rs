//! Bounded lexical adapters for `INDIRECT`.
//!
//! The reference parser in `formula::reference` owns the OpenFormula A1
//! grammar.  This module only adapts the text accepted by `INDIRECT` to that
//! parser and translates the small R1C1 coordinate grammar.  It never resolves
//! a reference or borrows any text from the normalization scratch buffer.

use super::super::{EvaluationFailure, EvaluationResult, Evaluator, ScalarError};
use crate::codec::formula::reference::{self, Reference};
use litchi_core::{Error, Reservation, Resource};
use std::fmt::Write as _;
use std::{borrow::Cow, mem::size_of};

/// Parsed text retained until the value evaluator turns it into a resolver
/// descriptor.  `requires_origin` is set for syntactically valid relative
/// R1C1 coordinates when the caller did not supply a position.  Keeping that
/// state separate from `Reference::Error` lets the scalar profile report a
/// capability refusal rather than inventing a position and changing the
/// meaning of the text.
#[derive(Debug)]
pub(crate) struct ParsedReferenceText {
    pub(super) reference: Option<Reference>,
    pub(super) requires_origin: bool,
    // Keep the budget token after the parser-owned reference.  The reference
    // and all of its component strings are dropped before this token.
    pub(super) reservation: Option<Reservation>,
}

struct NormalizedInput<'a> {
    text: Cow<'a, str>,
    reservation: Option<Reservation>,
}

impl ParsedReferenceText {
    /// Transfer the parser-owned reference and its accounting token to the
    /// value-side descriptor adapter. The reference is returned first so a
    /// caller can drop it before releasing the reservation.
    pub(crate) fn into_parts(self) -> (Option<Reference>, bool, Option<Reservation>) {
        (self.reference, self.requires_origin, self.reservation)
    }
}

/// Parse one INDIRECT text with a concrete zero-based origin for relative
/// R1C1 coordinates.  The value evaluator supplies the current position here.
pub(crate) fn parse_indirect<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
    a1: bool,
    base_row: usize,
    base_column: usize,
) -> EvaluationResult<Result<ParsedReferenceText, ScalarError>> {
    parse_indirect_at(evaluator, text, a1, Some((base_row, base_column)))
}

/// Parse one INDIRECT text without a current position.  Absolute A1/R1C1
/// syntax is still checked, while valid relative R1C1 text is retained as a
/// context-dependent parse result.
pub(crate) fn parse_indirect_without_origin<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
    a1: bool,
) -> EvaluationResult<Result<ParsedReferenceText, ScalarError>> {
    parse_indirect_at(evaluator, text, a1, None)
}

fn parse_indirect_at<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
    a1: bool,
    origin: Option<(usize, usize)>,
) -> EvaluationResult<Result<ParsedReferenceText, ScalarError>> {
    evaluator.charge_bytes(text.len())?;
    if text.len() > evaluator.limits.max_text_bytes() {
        return Err(super::super::local_limit(
            Resource::Memory,
            u64::try_from(text.len()).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes()).unwrap_or(u64::MAX),
        ));
    }

    let trimmed = trim_formula_text(text, evaluator)?;
    if trimmed.is_empty() {
        return Ok(Err(ScalarError::Reference));
    }

    let (mut normalized, requires_origin) = if a1 {
        (normalize_a1(trimmed, evaluator)?, false)
    } else {
        let normalized = match normalize_r1c1(trimmed, origin, evaluator)? {
            Ok(normalized) => normalized,
            Err(error) => return Ok(Err(error)),
        };
        let NormalizedR1C1 {
            text,
            requires_origin,
            reservation,
        } = normalized;
        (
            NormalizedInput {
                text: Cow::Owned(text),
                reservation,
            },
            requires_origin,
        )
    };

    // Parser-owned names and subtable vectors are bounded by the normalized
    // input size. The finite envelope is intentionally derived from this text
    // rather than reserving the global reference-component maximum for every
    // call. The normalization lease remains live while the parser borrows its
    // text, and is dropped before the parser-owned reference is published.
    let component_count = normalized.text.len().saturating_add(1);
    let parser_bytes = normalized
        .text
        .len()
        .checked_add(
            component_count
                .checked_mul(3)
                .and_then(|value| value.checked_mul(size_of::<reference::Subtable>()))
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDIRECT parser component reservation overflows",
                ))?,
        )
        .and_then(|value| value.checked_add(256))
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDIRECT parser reservation overflows",
        ))?;
    let reservation = evaluator.reserve_storage(parser_bytes, "formula INDIRECT parser")?;

    let limits = reference::Limits::default()
        .with_max_bytes(evaluator.limits.max_text_bytes())
        .with_max_name_bytes(evaluator.limits.max_text_bytes())
        .with_max_components(
            reference::DEFAULT_MAX_REFERENCE_COMPONENTS.min(trimmed.len().saturating_add(2)),
        );
    let parsed_reference = reference::parse_body(normalized.text.as_ref(), &limits);
    // The canonical parser owns the complete grammar. Fence its bounded scan
    // at both ends so cancellation/resource policy supersedes a formula parse
    // error discovered inside that parser.
    evaluator.charge_work(0)?;
    let reference = match parsed_reference {
        Ok(reference) => reference,
        Err(error) => return map_reference_error(error),
    };
    if reference.is_error() {
        return Ok(Err(ScalarError::Reference));
    }

    // The parser has finished borrowing the normalization buffer. Drop it
    // before returning the retained Reference and its parser reservation.
    let scratch_reservation = normalized.reservation.take();
    drop(normalized);
    drop(scratch_reservation);

    Ok(Ok(ParsedReferenceText {
        reference: Some(reference),
        requires_origin,
        reservation: Some(reservation),
    }))
}

fn map_reference_error(error: Error) -> EvaluationResult<Result<ParsedReferenceText, ScalarError>> {
    match error {
        Error::ResourceLimit(limit) => Err(EvaluationFailure::ResourceLimit(limit)),
        Error::Allocation { resource, source } => {
            Err(EvaluationFailure::Allocation { resource, source })
        },
        Error::InvalidFormat(_) => Ok(Err(ScalarError::Reference)),
        Error::Unsupported(_) => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
        Error::SourceChanged { expected, observed } => {
            Err(EvaluationFailure::SourceChanged { expected, observed })
        },
        // The canonical parser currently produces only InvalidFormat,
        // ResourceLimit, and Allocation. Preserve any future host failure as
        // a typed capability refusal rather than converting it to #REF.
        _ => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn trim_formula_text<'a>(
    value: &'a str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<&'a str> {
    let bytes = value.as_bytes();
    let mut start = 0usize;
    let mut polls_remaining = 0usize;
    while start < bytes.len() && matches!(bytes[start], b' ' | b'\t' | b'\r' | b'\n') {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        start += 1;
    }
    let mut end = bytes.len();
    polls_remaining = 0;
    while end > start && matches!(bytes[end - 1], b' ' | b'\t' | b'\r' | b'\n') {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        end -= 1;
    }
    Ok(&value[start..end])
}

/// Adapt A1 text to the canonical body grammar.  The canonical parser uses a
/// leading `.` for the current sheet, inherited range endpoints use `:.`, and
/// sheet separators are `.`.  All scans honor quoted sheet names and source
/// IRIs, so `!` in a quoted component is never rewritten.
fn normalize_a1<'a>(
    text: &'a str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<NormalizedInput<'a>> {
    if text.eq_ignore_ascii_case("#REF!") {
        return Ok(NormalizedInput {
            text: Cow::Borrowed(text),
            reservation: None,
        });
    }
    let body = text
        .strip_prefix('[')
        .and_then(|value| value.strip_suffix(']'))
        .unwrap_or(text);
    let colon = top_level_byte(body, b':', evaluator)?;
    let first = colon.map_or(body, |index| &body[..index]);
    let second = colon.map(|index| &body[index + 1..]);
    let has_source_prefix = has_unquoted_byte(first, b'#', evaluator)?;
    let first_has_locator = has_coordinate_separator(first, evaluator)?;
    let second_has_locator = match second {
        Some(value) => has_coordinate_separator(value, evaluator)?,
        None => false,
    };
    let needs_leading = !first_has_locator && !has_source_prefix && !first.starts_with('.');
    let needs_inherited = match second {
        Some(value) => {
            !second_has_locator && !starts_with_dot_after_formula_space(value, evaluator)?
        },
        None => false,
    };
    let has_bang = has_unquoted_byte(body, b'!', evaluator)?;
    let source_without_dot = has_source_prefix && !first_has_locator;
    if !needs_leading && !needs_inherited && !has_bang && !source_without_dot {
        // The body is already in canonical shape.  Parsing it directly avoids
        // a second transient string for the common `.A1` case.
        return Ok(NormalizedInput {
            text: Cow::Borrowed(body),
            reservation: None,
        });
    }

    let extra = usize::from(needs_leading)
        .saturating_add(usize::from(needs_inherited))
        .saturating_add(usize::from(source_without_dot))
        .saturating_add(body.len());
    let capacity = body
        .len()
        .checked_add(extra)
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDIRECT A1 normalization length overflows",
        ))?;
    let reservation = evaluator.reserve_storage(capacity, "formula INDIRECT A1 normalization")?;
    let mut normalized = String::new();
    normalized
        .try_reserve_exact(capacity)
        .map_err(|source| EvaluationFailure::Allocation {
            resource: "formula INDIRECT A1 normalization",
            source,
        })?;

    if needs_leading {
        normalized.push('.');
    }
    let mut quote = false;
    let bytes = body.as_bytes();
    let mut offset = 0usize;
    let mut polls_remaining = 0usize;
    while offset < body.len() {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        let character =
            body[offset..]
                .chars()
                .next()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDIRECT A1 text is not valid UTF-8",
                ))?;
        let width = character.len_utf8();
        if character == '\'' {
            normalized.push(character);
            if bytes.get(offset + width) == Some(&b'\'') {
                normalized.push('\'');
                offset = offset.saturating_add(width * 2);
                continue;
            }
            quote = !quote;
        } else if character == '!' && !quote {
            normalized.push('.');
        } else {
            normalized.push(character);
        }
        offset = offset.saturating_add(width);
    }
    if needs_inherited {
        let colon_offset = top_level_byte(normalized.as_str(), b':', evaluator)?.ok_or(
            EvaluationFailure::InvalidExpression("INDIRECT A1 inherited endpoint disappeared"),
        )?;
        normalized.insert(colon_offset + 1, '.');
    }
    if source_without_dot {
        let hash_offset = top_level_byte(normalized.as_str(), b'#', evaluator)?.ok_or(
            EvaluationFailure::InvalidExpression("INDIRECT source prefix disappeared"),
        )?;
        normalized.insert(hash_offset + 1, '.');
    }
    if normalized.len() > evaluator.limits.max_text_bytes() {
        drop(normalized);
        return Err(super::super::local_limit(
            Resource::Memory,
            u64::try_from(normalized_len_hint(body, extra)).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes()).unwrap_or(u64::MAX),
        ));
    }
    // Keep the scratch token with the String until the parser has finished
    // borrowing it. The caller drops this pair before releasing the parser's
    // own reservation on the retained Reference.
    Ok(NormalizedInput {
        text: Cow::Owned(normalized),
        reservation: Some(reservation),
    })
}

fn normalized_len_hint(body: &str, extra: usize) -> usize {
    body.len().saturating_add(extra)
}

fn starts_with_dot_after_formula_space(
    value: &str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<bool> {
    let bytes = value.as_bytes();
    let mut offset = 0usize;
    let mut polls_remaining = 0usize;
    while offset < bytes.len() && matches!(bytes[offset], b' ' | b'\t' | b'\r' | b'\n') {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        offset += 1;
    }
    evaluator.charge_work(0)?;
    Ok(bytes.get(offset) == Some(&b'.'))
}

fn has_coordinate_separator(
    value: &str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<bool> {
    Ok(value.starts_with('.')
        || has_unquoted_byte(value, b'.', evaluator)?
        || has_unquoted_byte(value, b'!', evaluator)?)
}

fn has_unquoted_byte(
    value: &str,
    sought: u8,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<bool> {
    let bytes = value.as_bytes();
    let mut quote = false;
    let mut index = 0usize;
    let mut polls_remaining = 0usize;
    while index < bytes.len() {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        match bytes[index] {
            b'\'' => {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index = index.saturating_add(2);
                    continue;
                }
                quote = !quote;
            },
            byte if byte == sought && !quote => return Ok(true),
            _ => {},
        }
        index = index.saturating_add(1);
    }
    Ok(false)
}

fn top_level_byte(
    value: &str,
    sought: u8,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Option<usize>> {
    let bytes = value.as_bytes();
    let mut quote = false;
    let mut index = 0usize;
    let mut polls_remaining = 0usize;
    while index < bytes.len() {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        match bytes[index] {
            b'\'' => {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index = index.saturating_add(2);
                    continue;
                }
                quote = !quote;
            },
            byte if byte == sought && !quote => return Ok(Some(index)),
            _ => {},
        }
        index = index.saturating_add(1);
    }
    Ok(None)
}

struct NormalizedR1C1 {
    text: String,
    requires_origin: bool,
    reservation: Option<Reservation>,
}

fn normalize_r1c1(
    text: &str,
    origin: Option<(usize, usize)>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<NormalizedR1C1, ScalarError>> {
    let (sheet, coordinates) = match split_r1c1_sheet(text, evaluator)? {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    let colon = top_level_byte(coordinates, b':', evaluator)?;
    let first_text = colon.map_or(coordinates, |index| &coordinates[..index]);
    let second_text = colon.map(|index| &coordinates[index + 1..]);
    let first = match parse_r1c1_endpoint(first_text, origin, evaluator)? {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    let second = match second_text {
        Some(value) => match parse_r1c1_endpoint(value, origin, evaluator)? {
            Ok(value) => Some(value),
            Err(error) => return Ok(Err(error)),
        },
        None => None,
    };
    if let Some(second) = second.as_ref()
        && first.kind() != second.kind()
    {
        return Ok(Err(ScalarError::Reference));
    }
    let requires_origin =
        first.requires_origin() || second.as_ref().is_some_and(|value| value.requires_origin());
    let capacity = text
        .len()
        .checked_add(16)
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDIRECT R1C1 normalization length overflows",
        ))?;
    let reservation = evaluator.reserve_storage(capacity, "formula INDIRECT R1C1 normalization")?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| EvaluationFailure::Allocation {
            resource: "formula INDIRECT R1C1 normalization",
            source,
        })?;
    output.push('.');
    if let Some(sheet) = sheet {
        output.clear();
        output.push_str(sheet);
        output.push('.');
    }
    write_r1c1_endpoint(&mut output, &first)?;
    if let Some(second) = second {
        output.push(':');
        output.push('.');
        write_r1c1_endpoint(&mut output, &second)?;
    }
    if output.len() > evaluator.limits.max_text_bytes() {
        drop(output);
        return Err(super::super::local_limit(
            Resource::Memory,
            u64::try_from(text.len()).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes()).unwrap_or(u64::MAX),
        ));
    }
    evaluator.charge_work(0)?;
    Ok(Ok(NormalizedR1C1 {
        text: output,
        requires_origin,
        reservation: Some(reservation),
    }))
}

fn split_r1c1_sheet<'text>(
    text: &'text str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<(Option<&'text str>, &'text str), ScalarError>> {
    let bytes = text.as_bytes();
    let mut quote = false;
    let mut index = 0usize;
    let mut polls_remaining = 0usize;
    while index < bytes.len() {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        match bytes[index] {
            b'\'' => {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index = index.saturating_add(2);
                    continue;
                }
                quote = !quote;
            },
            b'!' | b'.' if !quote => {
                if index == 0 || index + 1 >= bytes.len() {
                    return Ok(Err(ScalarError::Reference));
                }
                return Ok(Ok((Some(&text[..index]), &text[index + 1..])));
            },
            _ => {},
        }
        index = index.saturating_add(1);
    }
    if quote {
        return Ok(Err(ScalarError::Reference));
    }
    Ok(Ok((None, text)))
}

#[derive(Clone, Copy)]
enum AxisValue {
    Absolute(u64),
    Relative(i64),
}

#[derive(Clone, Copy)]
enum R1C1Endpoint {
    Cell {
        row: usize,
        column: usize,
        requires_origin: bool,
    },
    Row {
        row: usize,
        requires_origin: bool,
    },
    Column {
        column: usize,
        requires_origin: bool,
    },
}

impl R1C1Endpoint {
    fn kind(self) -> u8 {
        match self {
            Self::Cell { .. } => 0,
            Self::Row { .. } => 1,
            Self::Column { .. } => 2,
        }
    }

    fn requires_origin(self) -> bool {
        match self {
            Self::Cell {
                requires_origin, ..
            }
            | Self::Row {
                requires_origin, ..
            }
            | Self::Column {
                requires_origin, ..
            } => requires_origin,
        }
    }
}

fn parse_r1c1_endpoint(
    value: &str,
    origin: Option<(usize, usize)>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<R1C1Endpoint, ScalarError>> {
    let value = trim_formula_text(value, evaluator)?;
    if value.is_empty() {
        return Ok(Err(ScalarError::Reference));
    }
    let bytes = value.as_bytes();
    let mut offset = 0usize;
    let row = if bytes.first() == Some(&b'R') {
        offset += 1;
        Some(match parse_axis(value, &mut offset, evaluator)? {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        })
    } else {
        None
    };
    let column = if bytes.get(offset) == Some(&b'C') {
        offset += 1;
        Some(match parse_axis(value, &mut offset, evaluator)? {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        })
    } else {
        None
    };
    if offset != value.len() || (row.is_none() && column.is_none()) {
        return Ok(Err(ScalarError::Reference));
    }
    match (row, column) {
        (Some(row), Some(column)) => {
            let (row, row_relative) = match resolve_axis(row, origin.map(|value| value.0)) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            let (column, column_relative) = match resolve_axis(column, origin.map(|value| value.1))
            {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            Ok(Ok(R1C1Endpoint::Cell {
                row,
                column,
                requires_origin: row_relative || column_relative,
            }))
        },
        (Some(row), None) => {
            let (row, relative) = match resolve_axis(row, origin.map(|value| value.0)) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            Ok(Ok(R1C1Endpoint::Row {
                row,
                requires_origin: relative,
            }))
        },
        (None, Some(column)) => {
            let (column, relative) = match resolve_axis(column, origin.map(|value| value.1)) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            Ok(Ok(R1C1Endpoint::Column {
                column,
                requires_origin: relative,
            }))
        },
        (None, None) => Ok(Err(ScalarError::Reference)),
    }
}

fn parse_axis(
    value: &str,
    offset: &mut usize,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<AxisValue, ScalarError>> {
    let bytes = value.as_bytes();
    if bytes.get(*offset) == Some(&b'[') {
        *offset += 1;
        let start = *offset;
        let mut polls_remaining = 0usize;
        while bytes.get(*offset).is_some_and(u8::is_ascii_digit)
            || (*offset == start
                && bytes
                    .get(*offset)
                    .is_some_and(|byte| *byte == b'+' || *byte == b'-'))
        {
            if polls_remaining == 0 {
                evaluator.charge_work(0)?;
                polls_remaining = 4096;
            }
            polls_remaining -= 1;
            *offset += 1;
        }
        if *offset == start || bytes.get(*offset) != Some(&b']') {
            return Ok(Err(ScalarError::Reference));
        }
        let number = match value[start..*offset].parse::<i64>() {
            Ok(number) => number,
            Err(_) => return Ok(Err(ScalarError::Reference)),
        };
        *offset += 1;
        return Ok(Ok(AxisValue::Relative(number)));
    }
    let start = *offset;
    let mut polls_remaining = 0usize;
    while bytes.get(*offset).is_some_and(u8::is_ascii_digit) {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        *offset += 1;
    }
    if *offset == start {
        return Ok(Ok(AxisValue::Relative(0)));
    }
    let number = match value[start..*offset].parse::<u64>() {
        Ok(number) => number,
        Err(_) => return Ok(Err(ScalarError::Reference)),
    };
    if number == 0 {
        return Ok(Err(ScalarError::Reference));
    }
    Ok(Ok(AxisValue::Absolute(number)))
}

fn resolve_axis(value: AxisValue, origin: Option<usize>) -> Result<(usize, bool), ScalarError> {
    match value {
        AxisValue::Absolute(value) => usize::try_from(value)
            .map(|value| (value, false))
            .map_err(|_| ScalarError::Reference),
        AxisValue::Relative(offset) => {
            let Some(origin) = origin else {
                return Ok((1, true));
            };
            let coordinate = if offset >= 0 {
                origin.checked_add(offset as usize)
            } else {
                origin.checked_sub(offset.unsigned_abs() as usize)
            }
            .and_then(|value| value.checked_add(1))
            .ok_or(ScalarError::Reference)?;
            Ok((coordinate, false))
        },
    }
}

fn write_r1c1_endpoint(output: &mut String, endpoint: &R1C1Endpoint) -> EvaluationResult<()> {
    match *endpoint {
        R1C1Endpoint::Cell { row, column, .. } => {
            write_column(output, column)?;
            write!(output, "{row}").map_err(|_| {
                EvaluationFailure::InvalidExpression("INDIRECT coordinate formatting failed")
            })
        },
        R1C1Endpoint::Row { row, .. } => write!(output, "{row}")
            .map_err(|_| EvaluationFailure::InvalidExpression("INDIRECT row formatting failed")),
        R1C1Endpoint::Column { column, .. } => {
            output.push('$');
            write_column(output, column)
        },
    }
}

fn write_column(output: &mut String, mut column: usize) -> EvaluationResult<()> {
    if column == 0 {
        return Err(EvaluationFailure::InvalidExpression(
            "INDIRECT column is zero",
        ));
    }
    let mut letters = [0_u8; 32];
    let mut length = 0usize;
    while column != 0 {
        if length == letters.len() {
            return Err(EvaluationFailure::InvalidExpression(
                "INDIRECT column label exceeds fixed width",
            ));
        }
        column -= 1;
        letters[length] = b'A'
            + u8::try_from(column % 26)
                .map_err(|_| EvaluationFailure::InvalidExpression("INDIRECT column overflows"))?;
        length += 1;
        column /= 26;
    }
    for letter in letters[..length].iter().rev() {
        output.push(char::from(*letter));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::SourceVersion;

    #[test]
    fn typed_parser_failures_do_not_become_formula_errors() {
        let unsupported = map_reference_error(Error::Unsupported("host reference".into()));
        assert!(matches!(
            unsupported,
            Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference
            ))
        ));

        let expected = SourceVersion::new(11, 3);
        let observed = SourceVersion::new(11, 4);
        let changed = map_reference_error(Error::SourceChanged { expected, observed });
        assert!(matches!(
            changed,
            Err(EvaluationFailure::SourceChanged {
                expected: actual_expected,
                observed: actual_observed,
            }) if actual_expected == expected && actual_observed == observed
        ));

        let malformed = map_reference_error(Error::InvalidFormat("bad reference".into()));
        assert!(matches!(malformed, Ok(Err(ScalarError::Reference))));

        let future_host_failure = map_reference_error(Error::Other("future host error".into()));
        assert!(matches!(
            future_host_failure,
            Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference
            ))
        ));
    }
}
