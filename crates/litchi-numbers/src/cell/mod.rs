//! Numbers cell vocabulary.
//!
//! Native Binary Numbers Cell (BNC) storage is an implementation detail. Use
//! the semantic [`Value`](crate::cell::Value) and
//! [`data_format`](crate::cell::data_format) APIs instead of depending on its
//! byte layout.

/// Checked, archive-free cell display formats.
pub mod data_format;
/// Native BNC adapters stay crate-private.  Their wire/common finite scalar
/// is converted to the Numbers-owned [`FiniteF64`] at this boundary.
pub(crate) mod wire {
    use core::ops::{Deref, DerefMut};

    use litchi_numbers_wire as native;

    use super::FiniteF64;

    #[cfg(test)]
    pub(crate) use native::decimal128_le;
    pub(crate) use native::{ClearValue, Error, RewritePlan, StoredValue};

    /// A finite scalar accepted by the private BNC rewrite adapter.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub(crate) enum ScalarValue {
        String(u32),
        RichText(u32),
        Number(FiniteF64),
        Boolean(bool),
        Date(FiniteF64),
        Duration(FiniteF64),
    }

    impl ScalarValue {
        fn into_native(self) -> native::ScalarValue {
            match self {
                Self::String(identifier) => native::ScalarValue::String(identifier),
                Self::RichText(identifier) => native::ScalarValue::RichText(identifier),
                Self::Number(value) => native::ScalarValue::Number(to_native(value)),
                Self::Boolean(value) => native::ScalarValue::Boolean(value),
                Self::Date(value) => native::ScalarValue::Date(to_native(value)),
                Self::Duration(value) => native::ScalarValue::Duration(to_native(value)),
            }
        }
    }

    /// A decoded BNC scalar represented in the Numbers semantic vocabulary.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub(crate) enum CachedScalar {
        Number(FiniteF64),
        Boolean(bool),
        Date(FiniteF64),
        Duration(FiniteF64),
        Unsupported(u8),
    }

    impl CachedScalar {
        fn from_native(value: native::CachedScalar) -> Self {
            match value {
                native::CachedScalar::Number(value) => Self::Number(from_native(value)),
                native::CachedScalar::Boolean(value) => Self::Boolean(value),
                native::CachedScalar::Date(value) => Self::Date(from_native(value)),
                native::CachedScalar::Duration(value) => Self::Duration(from_native(value)),
                native::CachedScalar::Unsupported(value) => Self::Unsupported(value),
            }
        }

        fn into_native(self) -> native::CachedScalar {
            match self {
                Self::Number(value) => native::CachedScalar::Number(to_native(value)),
                Self::Boolean(value) => native::CachedScalar::Boolean(value),
                Self::Date(value) => native::CachedScalar::Date(to_native(value)),
                Self::Duration(value) => native::CachedScalar::Duration(to_native(value)),
                Self::Unsupported(value) => native::CachedScalar::Unsupported(value),
            }
        }
    }

    /// Owned BNC cell adapter with semantic scalar conversion at the seam.
    pub(crate) struct BncCell(native::BncCell);

    impl Deref for BncCell {
        type Target = native::BncCell;

        fn deref(&self) -> &Self::Target {
            &self.0
        }
    }

    impl DerefMut for BncCell {
        fn deref_mut(&mut self) -> &mut Self::Target {
            &mut self.0
        }
    }

    impl BncCell {
        pub(crate) fn parse(data: &[u8]) -> Result<Self, Error> {
            native::BncCell::parse(data).map(Self)
        }

        #[cfg(test)]
        #[must_use]
        pub(crate) fn minimal() -> Self {
            Self(native::BncCell::minimal())
        }

        #[cfg(test)]
        pub(crate) fn cached_scalar(&self) -> Result<Option<CachedScalar>, Error> {
            self.0
                .cached_scalar()
                .map(|value| value.map(CachedScalar::from_native))
        }
    }

    /// Borrowed BNC view adapter with semantic scalar conversion at the seam.
    pub(crate) struct BncCellView<'a>(native::BncCellView<'a>);

    impl<'a> BncCellView<'a> {
        pub(crate) fn parse(data: &'a [u8]) -> Result<Self, Error> {
            native::BncCellView::parse(data).map(Self)
        }

        #[must_use]
        pub(crate) fn stored_value(&self) -> StoredValue {
            self.0.stored_value()
        }

        #[must_use]
        pub(crate) fn cached_scalar(&self) -> Option<CachedScalar> {
            self.0.cached_scalar().map(CachedScalar::from_native)
        }

        #[must_use]
        pub(crate) fn formula_text_key(&self) -> Option<u32> {
            self.0.formula_text_key()
        }

        #[must_use]
        pub(crate) fn scalar_equals(&self, value: ScalarValue) -> bool {
            self.0.scalar_equals(value.into_native())
        }

        pub(crate) fn plan_scalar_rewrite(&self, value: ScalarValue) -> Result<RewritePlan, Error> {
            self.0.plan_scalar_rewrite(value.into_native())
        }

        pub(crate) fn plan_formula_rewrite(
            &self,
            identifier: u32,
            cache: Option<ScalarValue>,
        ) -> Result<RewritePlan, Error> {
            self.0
                .plan_formula_rewrite(identifier, cache.map(ScalarValue::into_native))
        }

        pub(crate) fn plan_formula_cache_rewrite(
            &self,
            cache: CachedScalar,
        ) -> Result<RewritePlan, Error> {
            self.0.plan_formula_cache_rewrite(cache.into_native())
        }

        pub(crate) fn plan_clear_value(&self, retain_empty: bool) -> Result<RewritePlan, Error> {
            self.0.plan_clear_value(retain_empty)
        }

        pub(crate) fn rewrite_scalar_with_limit(
            &self,
            value: ScalarValue,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_scalar_with_limit(value.into_native(), max_output_bytes)
        }

        pub(crate) fn clear_value_with_limit(
            &self,
            max_output_bytes: usize,
        ) -> Result<ClearValue, Error> {
            self.0.clear_value_with_limit(max_output_bytes)
        }

        pub(crate) fn formula_cache_equals(&self, value: CachedScalar) -> bool {
            self.0.formula_cache_equals(value.into_native())
        }

        pub(crate) fn formula_value_equals(
            &self,
            identifier: u32,
            value: ScalarValue,
        ) -> Result<bool, Error> {
            self.0.formula_value_equals(identifier, value.into_native())
        }

        pub(crate) fn rewrite_formula_with_limit(
            &self,
            identifier: u32,
            cache: ScalarValue,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_with_limit(identifier, cache.into_native(), max_output_bytes)
        }

        pub(crate) fn rewrite_formula_without_cache_with_limit(
            &self,
            identifier: u32,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_without_cache_with_limit(identifier, max_output_bytes)
        }

        pub(crate) fn rewrite_formula_cache_with_limit(
            &self,
            cache: CachedScalar,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_cache_with_limit(cache.into_native(), max_output_bytes)
        }

        #[must_use]
        pub(crate) fn formula_error_identifier(&self) -> Option<u32> {
            self.0.formula_error_identifier()
        }

        #[must_use]
        pub(crate) fn comment_identifier(&self) -> Option<u32> {
            self.0.comment_identifier()
        }
    }

    fn from_native(value: litchi_iwa_common::formula::FiniteF64) -> FiniteF64 {
        FiniteF64::new(value.get()).expect("wire parser guarantees finite scalar")
    }

    fn to_native(value: FiniteF64) -> litchi_iwa_common::formula::FiniteF64 {
        litchi_iwa_common::formula::FiniteF64::new(value.get())
            .expect("Numbers scalar invariant guarantees finite value")
    }
}

use std::fmt;

/// Failure returned when a Numbers semantic scalar is not finite.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[error("semantic scalar must be finite")]
pub struct FiniteF64Error;

/// A compact finite `f64` owned by the Numbers semantic API.
///
/// The inner value is private so public cell and formula values cannot be
/// constructed with NaN or infinity. Use [`FiniteF64::new`] or
/// [`TryFrom::try_from`] at an input boundary, and [`FiniteF64::get`] when a
/// native `f64` is required. Native Numbers adapters convert this value to
/// their private wire/common representation only at the archive boundary.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct FiniteF64(f64);

impl FiniteF64 {
    /// Construct a finite Numbers semantic scalar.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] for NaN or either infinity.
    pub const fn new(value: f64) -> Result<Self, FiniteF64Error> {
        if value.is_finite() {
            Ok(Self(value))
        } else {
            Err(FiniteF64Error)
        }
    }

    /// Return the finite scalar as a native `f64`.
    #[must_use]
    pub const fn get(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for FiniteF64 {
    type Error = FiniteF64Error;

    fn try_from(value: f64) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<FiniteF64> for f64 {
    fn from(value: FiniteF64) -> Self {
        value.get()
    }
}

/// Seconds between the Unix epoch and Apple's 2001-01-01 UTC epoch.
pub const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;

/// A typed value stored in a Numbers cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum Value {
    /// An empty semantic value; physical presence is represented separately.
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
        FiniteF64::new(value)
            .map(Self::Number)
            .map_err(|_error| FiniteF64Error)
    }

    /// Constructs a finite Apple-epoch date value.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when `value` is NaN or infinite.
    pub fn date(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value)
            .map(Self::Date)
            .map_err(|_error| FiniteF64Error)
    }

    /// Constructs a finite duration measured in seconds.
    ///
    /// # Errors
    ///
    /// Returns [`FiniteF64Error`] when `value` is NaN or infinite.
    pub fn duration(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value)
            .map(Self::Duration)
            .map_err(|_error| FiniteF64Error)
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
    /// An empty semantic value; physical presence is represented separately.
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

    #[test]
    fn finite_scalar_round_trips_through_owned_conversions() {
        use std::mem::size_of;

        let scalar = FiniteF64::try_from(3.5).expect("finite scalar should construct");
        assert_eq!(scalar.get(), 3.5);
        assert_eq!(f64::from(scalar), 3.5);
        assert_eq!(size_of::<FiniteF64>(), size_of::<f64>());
    }
}
