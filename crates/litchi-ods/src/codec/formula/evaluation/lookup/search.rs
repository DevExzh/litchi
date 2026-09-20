//! Shared lookup comparison primitives.
//!
//! The value VM owns traversal, shape checks, and read accounting.  This child
//! only compares already selected scalar values, so sorted and exact lookup
//! implementations can share one ordering rule without copying text or
//! retaining a search vector.

use super::super::{EvaluationFailure, EvaluationResult, ScalarError};
use std::cmp::Ordering;

/// Scalar kinds that participate in lookup ordering.
///
/// Empty cells are intentionally kept separate from Text.  The value VM
/// decides the function-specific Empty conversion before calling the search
/// kernel; silently treating an Empty as `""` here would make the same helper
/// unsuitable for MATCH and the three table lookups.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum SearchValue<'a> {
    Number(f64),
    Text(&'a str),
    Logical(bool),
}

impl SearchValue<'_> {
    #[inline]
    fn rank(self) -> Option<u8> {
        match self {
            Self::Number(_) => Some(0),
            Self::Text(_) => Some(1),
            Self::Logical(_) => Some(2),
        }
    }
}

/// Compare two lookup values using the ODF mixed-type order.
///
/// Numbers precede Text and Text precedes Logical. Text comparison uses the
/// evaluator's pinned Unicode 17 full C+F case fold and never allocates a
/// folded string. Formula errors are values for the caller's precedence pass,
/// but cannot be compared as search keys and are returned as formula errors.
pub(crate) fn compare_values<F>(
    left: SearchValue<'_>,
    right: SearchValue<'_>,
    mut charge: F,
) -> EvaluationResult<Result<Ordering, ScalarError>>
where
    F: FnMut(u64) -> EvaluationResult<()>,
{
    let Some(left_rank) = left.rank() else {
        return Ok(Err(ScalarError::Value));
    };
    let Some(right_rank) = right.rank() else {
        return Ok(Err(ScalarError::Value));
    };
    if left_rank != right_rank {
        return Ok(Ok(left_rank.cmp(&right_rank)));
    }

    let result = match (left, right) {
        (SearchValue::Number(left), SearchValue::Number(right)) => {
            if !left.is_finite() || !right.is_finite() {
                return Ok(Err(ScalarError::Number));
            }
            left.partial_cmp(&right).ok_or(ScalarError::Number)
        },
        (SearchValue::Text(left), SearchValue::Text(right)) => {
            compare_text(left, right, &mut charge)?
        },
        (SearchValue::Logical(left), SearchValue::Logical(right)) => Ok(left.cmp(&right)),
        _ => Err(ScalarError::Value),
    };
    Ok(result)
}

/// Compare lookup Text without allocating a case-folded copy.
pub(crate) fn compare_text<F>(
    left: &str,
    right: &str,
    charge: &mut F,
) -> EvaluationResult<Result<Ordering, ScalarError>>
where
    F: FnMut(u64) -> EvaluationResult<()>,
{
    let mut left = Folded::new(left);
    let mut right = Folded::new(right);
    loop {
        char_charge(charge)?;
        let left = left.next()?;
        char_charge(charge)?;
        let right = right.next()?;
        match (left, right) {
            (None, None) => return Ok(Ok(Ordering::Equal)),
            (None, Some(_)) => return Ok(Ok(Ordering::Less)),
            (Some(_), None) => return Ok(Ok(Ordering::Greater)),
            (Some(left), Some(right)) => match left.cmp(&right) {
                Ordering::Equal => {},
                ordering => return Ok(Ok(ordering)),
            },
        }
    }
}

fn char_charge<F>(charge: &mut F) -> EvaluationResult<()>
where
    F: FnMut(u64) -> EvaluationResult<()>,
{
    charge(1)
}

/// Return whether a candidate of the selected type is permitted for an
/// approximate search. The ODF rules reject the Number/Text crossover in one
/// direction depending on the search direction; this is separate from the
/// total ordering so callers can select a duplicate before applying the rule.
pub(crate) fn approximate_type_compatible(
    search: SearchValue<'_>,
    candidate: SearchValue<'_>,
    descending: bool,
) -> bool {
    match (search, candidate) {
        (SearchValue::Text(_), SearchValue::Number(_)) if !descending => false,
        (SearchValue::Number(_), SearchValue::Text(_)) if descending => false,
        _ => true,
    }
}

/// Check an exact lookup match after the caller has filtered formula errors.
pub(crate) fn exact_match<F>(
    left: SearchValue<'_>,
    right: SearchValue<'_>,
    charge: F,
) -> EvaluationResult<Result<bool, ScalarError>>
where
    F: FnMut(u64) -> EvaluationResult<()>,
{
    Ok(compare_values(left, right, charge)?.map(|ordering| ordering == Ordering::Equal))
}

struct Folded<'a> {
    source: std::str::Chars<'a>,
    mapped: [char; 3],
    length: usize,
    position: usize,
}

impl<'a> Folded<'a> {
    fn new(source: &'a str) -> Self {
        Self {
            source: source.chars(),
            mapped: ['\0'; 3],
            length: 0,
            position: 0,
        }
    }

    fn next(&mut self) -> EvaluationResult<Option<char>> {
        if self.position < self.length {
            let value = self.mapped[self.position];
            self.position += 1;
            return Ok(Some(value));
        }
        let Some(character) = self.source.next() else {
            return Ok(None);
        };
        let mut length = 0usize;
        for folded in super::super::text::case_fold(character) {
            if length == self.mapped.len() {
                return Err(EvaluationFailure::InvalidExpression(
                    "lookup Unicode case fold exceeded fixed mapping width",
                ));
            }
            self.mapped[length] = folded;
            length += 1;
        }
        self.length = length;
        self.position = 0;
        self.next()
    }
}
