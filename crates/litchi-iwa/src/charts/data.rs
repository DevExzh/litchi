//! Validated, source-buildable chart data.

use crate::{Error, Result};

/// Rectangular chart data with optional numeric values.
///
/// Missing values are represented by `None`; non-finite numbers are rejected
/// because Numbers cannot render them reliably.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartData {
    row_names: Vec<String>,
    column_names: Vec<String>,
    values: Vec<Vec<Option<f64>>>,
}

impl ChartData {
    /// Validate and construct a non-empty rectangular chart grid.
    pub fn new(
        row_names: Vec<String>,
        column_names: Vec<String>,
        values: Vec<Vec<Option<f64>>>,
    ) -> Result<Self> {
        if row_names.is_empty() || column_names.is_empty() {
            return Err(Error::ParseError(
                "chart data must contain at least one row and one column".to_owned(),
            ));
        }
        if values.len() != row_names.len()
            || values.iter().any(|row| row.len() != column_names.len())
        {
            return Err(Error::ParseError(
                "chart values must match the row and column label dimensions".to_owned(),
            ));
        }
        if values
            .iter()
            .flatten()
            .flatten()
            .any(|value| !value.is_finite())
        {
            return Err(Error::ParseError(
                "chart numeric values must be finite".to_owned(),
            ));
        }
        Ok(Self {
            row_names,
            column_names,
            values,
        })
    }

    /// Borrow the row labels.
    pub fn row_names(&self) -> &[String] {
        &self.row_names
    }

    /// Borrow the column labels.
    pub fn column_names(&self) -> &[String] {
        &self.column_names
    }

    /// Borrow the row-major numeric values.
    pub fn values(&self) -> &[Vec<Option<f64>>] {
        &self.values
    }

    pub(crate) fn into_parts(self) -> (Vec<String>, Vec<String>, Vec<Vec<Option<f64>>>) {
        (self.row_names, self.column_names, self.values)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_rectangular_finite_data() {
        let data = ChartData::new(
            vec!["North".to_owned(), "South".to_owned()],
            vec!["Q1".to_owned(), "Q2".to_owned()],
            vec![vec![Some(4.0), None], vec![Some(8.0), Some(12.0)]],
        )
        .unwrap();
        assert_eq!(data.values().len(), 2);

        assert!(
            ChartData::new(
                vec!["North".to_owned()],
                vec!["Q1".to_owned()],
                vec![vec![Some(f64::NAN)]],
            )
            .is_err()
        );
        assert!(
            ChartData::new(
                vec!["North".to_owned()],
                vec!["Q1".to_owned(), "Q2".to_owned()],
                vec![vec![Some(1.0)]],
            )
            .is_err()
        );
    }
}
