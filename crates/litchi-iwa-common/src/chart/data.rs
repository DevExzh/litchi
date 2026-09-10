//! Archive-free rectangular data shared by iWork chart owners.
//!
//! The concrete format crates own chart archives and their wire adapters. This
//! module owns only the validated row and column labels plus the optional
//! numeric values exchanged at that boundary.

/// A validation failure while constructing [`ChartData`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum DataError {
    /// No row labels were supplied.
    #[error("chart data must contain at least one row")]
    EmptyRows,
    /// No column labels were supplied.
    #[error("chart data must contain at least one column")]
    EmptyColumns,
    /// The number of value rows does not match the number of row labels.
    #[error("chart values contain {actual} rows, expected {expected}")]
    RowCountMismatch {
        /// Number of row labels supplied.
        expected: usize,
        /// Number of value rows supplied.
        actual: usize,
    },
    /// One value row does not match the number of column labels.
    #[error("chart values row {row} contains {actual} columns, expected {expected}")]
    ColumnCountMismatch {
        /// Zero-based index of the row whose width is invalid.
        row: usize,
        /// Number of column labels supplied.
        expected: usize,
        /// Number of values in the invalid row.
        actual: usize,
    },
    /// A value is not representable as a finite chart number.
    #[error("chart value at row {row}, column {column} must be finite")]
    NonFiniteValue {
        /// Zero-based row containing the invalid value.
        row: usize,
        /// Zero-based column containing the invalid value.
        column: usize,
    },
}

/// A rectangular chart data grid with optional numeric values.
///
/// Missing values are represented by [`None`]. Every present number is
/// finite, and the value grid is non-empty, rectangular, and dimensioned by
/// [`Self::row_names`] and [`Self::column_names`].
#[derive(Debug, Clone, PartialEq)]
pub struct ChartData {
    row_names: Vec<String>,
    column_names: Vec<String>,
    values: Vec<Vec<Option<f64>>>,
}

impl ChartData {
    /// Validate and construct a non-empty rectangular chart grid.
    ///
    /// The constructor takes ownership of all supplied labels and values. It
    /// does not normalize labels or silently replace missing values.
    ///
    /// # Errors
    ///
    /// Returns a [`DataError`] when an axis is empty, the value grid has the
    /// wrong dimensions, or a present number is NaN or infinite.
    pub fn new(
        row_names: Vec<String>,
        column_names: Vec<String>,
        values: Vec<Vec<Option<f64>>>,
    ) -> Result<Self, DataError> {
        if row_names.is_empty() {
            return Err(DataError::EmptyRows);
        }
        if column_names.is_empty() {
            return Err(DataError::EmptyColumns);
        }
        if values.len() != row_names.len() {
            return Err(DataError::RowCountMismatch {
                expected: row_names.len(),
                actual: values.len(),
            });
        }

        for (row, values) in values.iter().enumerate() {
            if values.len() != column_names.len() {
                return Err(DataError::ColumnCountMismatch {
                    row,
                    expected: column_names.len(),
                    actual: values.len(),
                });
            }
            for (column, value) in values.iter().enumerate() {
                if value.is_some_and(|value| !value.is_finite()) {
                    return Err(DataError::NonFiniteValue { row, column });
                }
            }
        }

        Ok(Self {
            row_names,
            column_names,
            values,
        })
    }

    /// Borrow the row labels.
    #[must_use]
    pub fn row_names(&self) -> &[String] {
        &self.row_names
    }

    /// Borrow the column labels.
    #[must_use]
    pub fn column_names(&self) -> &[String] {
        &self.column_names
    }

    /// Borrow the row-major numeric values.
    #[must_use]
    pub fn values(&self) -> &[Vec<Option<f64>>] {
        &self.values
    }

    /// Return the owned labels and row-major values for a source builder.
    #[must_use]
    pub fn into_parts(self) -> (Vec<String>, Vec<String>, Vec<Vec<Option<f64>>>) {
        (self.row_names, self.column_names, self.values)
    }
}

#[cfg(test)]
mod tests {
    use super::{ChartData, DataError};

    fn labels() -> (Vec<String>, Vec<String>) {
        (
            vec![String::from("North"), String::from("South")],
            vec![String::from("Q1"), String::from("Q2")],
        )
    }

    #[test]
    fn accepts_rectangular_finite_data_and_preserves_missing_values() {
        let (rows, columns) = labels();
        let data = ChartData::new(
            rows.clone(),
            columns.clone(),
            vec![vec![Some(4.0), None], vec![Some(-8.5), Some(12.0)]],
        )
        .expect("valid chart data");

        assert_eq!(data.row_names(), rows.as_slice());
        assert_eq!(data.column_names(), columns.as_slice());
        assert_eq!(data.values()[0][1], None);
        assert_eq!(data.values()[1][0], Some(-8.5));
    }

    #[test]
    fn rejects_empty_axes_with_typed_errors() {
        let (_, columns) = labels();
        assert_eq!(
            ChartData::new(Vec::new(), columns.clone(), Vec::new()),
            Err(DataError::EmptyRows)
        );

        let (rows, _) = labels();
        assert_eq!(
            ChartData::new(rows, Vec::new(), Vec::new()),
            Err(DataError::EmptyColumns)
        );
    }

    #[test]
    fn rejects_missing_or_extra_value_rows() {
        let (rows, columns) = labels();
        assert_eq!(
            ChartData::new(
                rows.clone(),
                columns.clone(),
                vec![vec![Some(1.0), Some(2.0)]]
            ),
            Err(DataError::RowCountMismatch {
                expected: 2,
                actual: 1,
            })
        );

        assert_eq!(
            ChartData::new(
                rows,
                columns,
                vec![
                    vec![Some(1.0), Some(2.0)],
                    vec![Some(3.0), Some(4.0)],
                    vec![Some(5.0), Some(6.0)],
                ],
            ),
            Err(DataError::RowCountMismatch {
                expected: 2,
                actual: 3,
            })
        );
    }

    #[test]
    fn rejects_ragged_rows_and_reports_the_first_bad_row() {
        let (rows, columns) = labels();
        assert_eq!(
            ChartData::new(
                rows,
                columns,
                vec![vec![Some(1.0), Some(2.0)], vec![Some(3.0)]],
            ),
            Err(DataError::ColumnCountMismatch {
                row: 1,
                expected: 2,
                actual: 1,
            })
        );
    }

    #[test]
    fn rejects_nan_and_both_infinities_with_coordinates() {
        let cases = [f64::NAN, f64::INFINITY, f64::NEG_INFINITY];
        for value in cases {
            let (rows, columns) = labels();
            assert_eq!(
                ChartData::new(
                    rows,
                    columns,
                    vec![vec![Some(0.0), Some(value)], vec![None, Some(1.0)]],
                ),
                Err(DataError::NonFiniteValue { row: 0, column: 1 })
            );
        }
    }

    #[test]
    fn into_parts_returns_original_owned_components() {
        let (rows, columns) = labels();
        let values = vec![vec![Some(1.0), None], vec![None, Some(2.0)]];
        let data = ChartData::new(rows.clone(), columns.clone(), values.clone()).unwrap();

        assert_eq!(data.into_parts(), (rows, columns, values));
    }
}
