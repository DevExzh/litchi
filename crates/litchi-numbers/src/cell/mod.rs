//! Numbers cell vocabulary.
//!
//! Native Binary Numbers Cell (BNC) storage is an implementation detail. Use
//! the semantic [`Value`](crate::cell::Value) and
//! [`data_format`](crate::cell::data_format) APIs instead of depending on its
//! byte layout.

/// Checked, archive-free cell display formats.
pub mod data_format;
pub(crate) use litchi_numbers_wire as wire;

use std::fmt;

use litchi_iwa_common::formula::FormulaCachedValue;
pub use litchi_iwa_common::formula::{FiniteF64, FiniteF64Error};

/// Seconds between the Unix epoch and Apple's 2001-01-01 UTC epoch.
pub const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;

/// A typed value stored in a Numbers cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum Value {
    /// No materialized value.
    #[default]
    Empty,
    /// User-entered text.
    Text(String),
    /// Numeric value.
    Number(FiniteF64),
    /// Boolean value.
    Boolean(bool),
    /// Seconds since Apple's 2001-01-01 UTC epoch.
    Date(FiniteF64),
    /// Duration in seconds.
    Duration(FiniteF64),
    /// Formula source or rendered formula expression.
    Formula(String),
    /// Producer-reported cell error text.
    Error(String),
}

impl Value {
    /// Constructs a finite numeric value.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when `value` is NaN or infinite.
    pub fn number(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Number)
    }

    /// Constructs a finite Apple-epoch date value.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when `value` is NaN or infinite.
    pub fn date(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Date)
    }

    /// Constructs a finite duration measured in seconds.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when `value` is NaN or infinite.
    pub fn duration(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Duration)
    }

    /// Constructs a Numbers date from Unix epoch seconds.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when the converted Apple-epoch value is not
    /// finite.
    pub fn date_from_unix_seconds(unix_seconds: f64) -> Result<Self, FiniteF64Error> {
        Self::date(unix_seconds - APPLE_EPOCH_UNIX_OFFSET_SECONDS)
    }

    /// Converts a Numbers date to Unix epoch seconds.
    #[must_use]
    pub fn date_as_unix_seconds(&self) -> Option<f64> {
        match self {
            Self::Date(seconds) => {
                let unix_seconds = seconds.get() + APPLE_EPOCH_UNIX_OFFSET_SECONDS;
                unix_seconds.is_finite().then_some(unix_seconds)
            },
            Self::Empty
            | Self::Text(_)
            | Self::Number(_)
            | Self::Boolean(_)
            | Self::Duration(_)
            | Self::Formula(_)
            | Self::Error(_) => None,
        }
    }

    /// Returns whether the value is an explicit or implicit empty cell.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    /// Returns the semantic kind of this value.
    #[must_use]
    pub const fn cell_type(&self) -> Type {
        match self {
            Self::Empty => Type::Empty,
            Self::Text(_) => Type::Text,
            Self::Number(_) => Type::Number,
            Self::Boolean(_) => Type::Boolean,
            Self::Date(_) => Type::Date,
            Self::Duration(_) => Type::Duration,
            Self::Formula(_) => Type::Formula,
            Self::Error(_) => Type::Error,
        }
    }

    /// Returns the display text used by CSV and text projections.
    #[must_use]
    pub fn as_text(&self) -> String {
        match self {
            Self::Empty => String::new(),
            Self::Text(value) | Self::Formula(value) => value.clone(),
            Self::Number(value) | Self::Date(value) | Self::Duration(value) => {
                value.get().to_string()
            },
            Self::Boolean(value) => value.to_string(),
            Self::Error(value) => format!("ERROR: {value}"),
        }
    }

    /// Converts numeric-compatible values without allocating.
    #[must_use]
    pub fn as_number(&self) -> Option<f64> {
        match self {
            Self::Number(value) | Self::Date(value) | Self::Duration(value) => Some(value.get()),
            Self::Text(value) => value
                .parse::<f64>()
                .ok()
                .filter(|parsed| parsed.is_finite()),
            Self::Boolean(value) => Some(if *value { 1.0 } else { 0.0 }),
            Self::Empty | Self::Formula(_) | Self::Error(_) => None,
        }
    }

    /// Converts boolean-compatible values without allocating.
    #[must_use]
    pub fn as_boolean(&self) -> Option<bool> {
        match self {
            Self::Boolean(value) => Some(*value),
            Self::Number(value) => Some(value.get() != 0.0),
            Self::Text(value) => {
                if value.eq_ignore_ascii_case("true")
                    || value.eq_ignore_ascii_case("yes")
                    || value == "1"
                {
                    Some(true)
                } else if value.eq_ignore_ascii_case("false")
                    || value.eq_ignore_ascii_case("no")
                    || value == "0"
                {
                    Some(false)
                } else {
                    None
                }
            },
            Self::Empty | Self::Date(_) | Self::Duration(_) | Self::Formula(_) | Self::Error(_) => {
                None
            },
        }
    }
}

impl From<FormulaCachedValue> for Value {
    fn from(value: FormulaCachedValue) -> Self {
        match value {
            FormulaCachedValue::Number(number) => Self::Number(number),
            FormulaCachedValue::Text(text) => Self::Text(text),
            FormulaCachedValue::Boolean(boolean) => Self::Boolean(boolean),
            FormulaCachedValue::Date(date) => Self::Date(date),
            FormulaCachedValue::Duration(duration) => Self::Duration(duration),
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let text = self.as_text();
        if text.contains([',', '"', '\n']) {
            write!(formatter, "\"{}\"", text.replace('"', "\"\""))
        } else {
            formatter.write_str(&text)
        }
    }
}

/// The semantic kind of a Numbers cell value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Type {
    /// No materialized value.
    Empty,
    /// User-entered text.
    Text,
    /// Numeric value.
    Number,
    /// Boolean value.
    Boolean,
    /// Date value.
    Date,
    /// Duration value.
    Duration,
    /// Formula value.
    Formula,
    /// Error value.
    Error,
}

impl Type {
    /// Returns a stable human-readable name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Empty => "Empty",
            Self::Text => "Text",
            Self::Number => "Number",
            Self::Boolean => "Boolean",
            Self::Date => "Date",
            Self::Duration => "Duration",
            Self::Formula => "Formula",
            Self::Error => "Error",
        }
    }
}

/// One typed mutation in a transactional cell batch.
#[derive(Debug, Clone, PartialEq)]
pub struct Update {
    /// Zero-based row coordinate.
    pub row: usize,
    /// Zero-based column coordinate.
    pub column: usize,
    /// Final value for the coordinate.
    pub value: Value,
}

impl Update {
    /// Creates an update for a zero-based coordinate.
    #[must_use]
    pub const fn new(row: usize, column: usize, value: Value) -> Self {
        Self { row, column, value }
    }

    /// Creates an update that explicitly clears a coordinate.
    #[must_use]
    pub const fn clear(row: usize, column: usize) -> Self {
        Self::new(row, column, Value::Empty)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn values_report_types_and_convert_without_surprises() {
        let empty = Value::default();
        assert!(empty.is_empty());
        assert_eq!(empty.cell_type(), Type::Empty);

        let date = Value::date_from_unix_seconds(APPLE_EPOCH_UNIX_OFFSET_SECONDS + 123.5)
            .expect("finite date should construct");
        assert_eq!(date.cell_type(), Type::Date);
        assert_eq!(
            date.date_as_unix_seconds(),
            Some(APPLE_EPOCH_UNIX_OFFSET_SECONDS + 123.5)
        );
        assert_eq!(Value::Text("true".to_owned()).as_boolean(), Some(true));
        assert_eq!(Value::Text("123.45".to_owned()).as_number(), Some(123.45));
    }

    #[test]
    fn display_escapes_csv_values() {
        assert_eq!(Value::Text("Simple".to_owned()).to_string(), "Simple");
        assert_eq!(
            Value::Text("Hello, World".to_owned()).to_string(),
            "\"Hello, World\""
        );
        assert_eq!(
            Value::Text("Say \"Hi\"".to_owned()).to_string(),
            "\"Say \"\"Hi\"\"\""
        );
    }

    #[test]
    fn updates_are_typed_and_clear_explicitly() {
        let update = Update::new(
            2,
            3,
            Value::number(42.0).expect("finite number should construct"),
        );
        assert_eq!(update.row, 2);
        assert_eq!(update.column, 3);
        assert_eq!(Update::clear(2, 3).value, Value::Empty);
    }

    #[test]
    fn scalar_constructors_reject_non_finite_input() {
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(Value::number(value).is_err());
            assert!(Value::date(value).is_err());
            assert!(Value::duration(value).is_err());
            assert!(Value::date_from_unix_seconds(value).is_err());
        }
    }
}
