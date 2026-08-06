//! Bounded, dependency-free table edit plans.
//!
//! This module validates cell mutations before a format adapter touches an
//! archive. It owns the coordinate uniqueness and allocation boundary while
//! leaving package identifiers, formula graphs, and wire mutation in the
//! concrete format crate.

use super::Position;
use crate::cell::{Type, Update, Value};
use std::collections::HashSet;
use std::fmt;

/// Upper bound selected by a caller for one transactional cell batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Budget {
    max_updates: usize,
}

impl Budget {
    /// Creates a batch budget without allocating.
    #[must_use]
    pub const fn new(max_updates: usize) -> Self {
        Self { max_updates }
    }

    /// Returns the maximum number of updates accepted by this budget.
    #[must_use]
    pub const fn max_updates(self) -> usize {
        self.max_updates
    }
}

/// An owned, validated set of unique cell mutations.
#[derive(Debug, Clone, PartialEq)]
pub struct Batch {
    updates: Box<[Update]>,
    positions: Box<[(usize, usize)]>,
}

impl Batch {
    /// Collects a unique cell batch within the caller-supplied allocation
    /// budget.
    ///
    /// The iterator's lower size hint is used only for bounded preallocation;
    /// every subsequent reservation is fallible and capped by `budget`.
    ///
    /// # Errors
    ///
    /// Returns a typed error for duplicate coordinates, unsupported values,
    /// non-finite numeric values, coordinate overflow, allocation failure, or
    /// a batch larger than the supplied budget.
    pub fn collect<I>(updates: I, budget: Budget) -> Result<Self>
    where
        I: IntoIterator<Item = Update>,
    {
        let updates_iter = updates.into_iter();
        let lower_bound = updates_iter.size_hint().0;
        if lower_bound > budget.max_updates {
            return Err(Error::LimitExceeded {
                requested: lower_bound,
                maximum: budget.max_updates,
            });
        }

        let mut values = Vec::new();
        reserve(&mut values, lower_bound, "table edit values")?;
        let mut positions = Vec::new();
        reserve(&mut positions, lower_bound, "table edit coordinates")?;
        let mut seen = HashSet::new();
        seen.try_reserve(lower_bound)
            .map_err(|_error| Error::Allocation {
                resource: "table edit coordinate set",
                amount: lower_bound,
            })?;

        for update in updates_iter {
            if values.len() >= budget.max_updates {
                return Err(Error::LimitExceeded {
                    requested: budget.max_updates.saturating_add(1),
                    maximum: budget.max_updates,
                });
            }

            let position =
                Position::try_from_usize(update.row, update.column).map_err(|_error| {
                    Error::CoordinateOverflow {
                        row: update.row,
                        column: update.column,
                    }
                })?;
            seen.try_reserve(1).map_err(|_error| Error::Allocation {
                resource: "table edit coordinate set",
                amount: 1,
            })?;
            if !seen.insert(position) {
                return Err(Error::DuplicatePosition { position });
            }
            validate_value(&update.value)?;

            reserve(&mut values, 1, "table edit values")?;
            reserve(&mut positions, 1, "table edit coordinates")?;
            values.push(update);
            positions.push((position.row() as usize, position.column() as usize));
        }

        Ok(Self {
            updates: values.into_boxed_slice(),
            positions: positions.into_boxed_slice(),
        })
    }

    /// Returns the number of mutations in this batch.
    #[must_use]
    pub fn len(&self) -> usize {
        self.updates.len()
    }

    /// Returns whether this batch contains no mutations.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.updates.is_empty()
    }

    /// Splits the validated updates from their compact cache-refresh
    /// coordinates without cloning either collection.
    #[must_use]
    pub fn into_parts(self) -> (Box<[Update]>, Box<[(usize, usize)]>) {
        (self.updates, self.positions)
    }
}

/// Failure while constructing a bounded cell-edit batch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A coordinate cannot be represented by the compact table position.
    CoordinateOverflow { row: usize, column: usize },
    /// A coordinate occurs more than once in the batch.
    DuplicatePosition { position: Position },
    /// The batch exceeds its caller-supplied bound.
    LimitExceeded { requested: usize, maximum: usize },
    /// A collection reservation failed before the batch was completed.
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    /// Formula and error values require format-owned construction context.
    UnsupportedValue { kind: Type },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::CoordinateOverflow { row, column } => {
                write!(
                    formatter,
                    "table coordinate ({row}, {column}) overflows u32"
                )
            },
            Self::DuplicatePosition { position } => write!(
                formatter,
                "table edit batch repeats coordinate ({}, {})",
                position.row(),
                position.column()
            ),
            Self::LimitExceeded { requested, maximum } => write!(
                formatter,
                "table edit batch requests {requested} updates, budget is {maximum}"
            ),
            Self::Allocation { resource, amount } => {
                write!(
                    formatter,
                    "table edit allocation failed for {resource}: {amount}"
                )
            },
            Self::UnsupportedValue { kind } => write!(
                formatter,
                "table edit value kind {} requires format-owned construction",
                kind.name()
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for bounded table-edit operations.
pub type Result<T> = std::result::Result<T, Error>;

fn validate_value(value: &Value) -> Result<()> {
    match value {
        Value::Empty
        | Value::Text(_)
        | Value::Boolean(_)
        | Value::Number(_)
        | Value::Date(_)
        | Value::Duration(_) => Ok(()),
        Value::Formula(_) | Value::Error(_) => Err(Error::UnsupportedValue {
            kind: value.cell_type(),
        }),
    }
}

fn reserve<T>(values: &mut Vec<T>, additional: usize, resource: &'static str) -> Result<()> {
    values
        .try_reserve(additional)
        .map_err(|_error| Error::Allocation {
            resource,
            amount: additional,
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn collects_unique_updates_and_releases_owned_parts() {
        let batch = Batch::collect(
            [
                Update::new(2, 3, Value::number(42.0).expect("finite test number")),
                Update::clear(0, 1),
            ],
            Budget::new(2),
        )
        .unwrap_or_else(|error| panic!("unexpected batch failure: {error}"));

        assert_eq!(batch.len(), 2);
        let (updates, positions) = batch.into_parts();
        assert_eq!(updates.len(), 2);
        assert_eq!(positions.as_ref(), &[(2, 3), (0, 1)]);
    }

    #[test]
    fn rejects_duplicates_and_budget_overflow() {
        let duplicate = Batch::collect(
            [
                Update::clear(1, 1),
                Update::new(1, 1, Value::Text("duplicate".to_owned())),
            ],
            Budget::new(2),
        );
        assert!(matches!(duplicate, Err(Error::DuplicatePosition { .. })));

        let oversized = Batch::collect([Update::clear(0, 0), Update::clear(0, 1)], Budget::new(1));
        assert!(matches!(oversized, Err(Error::LimitExceeded { .. })));
    }

    #[test]
    fn rejects_unsupported_values_and_requires_finite_scalar_construction() {
        let formula = Batch::collect(
            [Update::new(0, 0, Value::Formula("SUM(A1)".to_owned()))],
            Budget::new(1),
        );
        assert!(matches!(
            formula,
            Err(Error::UnsupportedValue {
                kind: Type::Formula
            })
        ));

        assert!(crate::cell::FiniteF64::new(f64::NAN).is_err());
    }
}
