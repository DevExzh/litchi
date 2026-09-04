//! Dependency-free custom cell-format values.
//!
//! Semantic values retain exact user content through their getters, while
//! their [`fmt::Debug`] implementations expose only bounded shape metadata.

/// Selector-first package transactions for an existing cell's document-level
/// Custom display format.
pub mod transaction {
    pub use crate::package::table_cell_custom_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}

use std::fmt;

/// Maximum UTF-8 bytes retained by a custom name or pattern.
pub const MAX_PATTERN_BYTES: usize = 4 * 1_024;
/// Maximum UTF-8 bytes retained by a custom format name.
pub const MAX_NAME_BYTES: usize = 255;
/// Maximum ordered rules retained by a custom Number format.
pub const MAX_RULES: usize = 32;

/// Errors returned by checked custom-format constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// A required value is empty.
    Empty { field: &'static str },
    /// A value exceeds its bounded UTF-8 budget.
    TooLong {
        field: &'static str,
        length: usize,
        maximum: usize,
    },
    /// A value contains a control character.
    ContainsControl { field: &'static str, index: usize },
    /// A name has surrounding whitespace.
    SurroundingWhitespace { field: &'static str },
    /// A Number pattern has no digit placeholder.
    MissingNumberPlaceholder,
    /// A Date & Time pattern has no date or time field.
    MissingDateTimeField,
    /// A condition threshold is not finite.
    NonFiniteThreshold,
    /// A custom Number format has too many rules.
    TooManyRules { actual: usize, maximum: usize },
    /// A custom Number rule buffer could not be allocated.
    Allocation { amount: usize },
    /// A custom Number format repeats a condition.
    DuplicateCondition,
    /// A literal-only Text format has no literal content.
    EmptyLiteral,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty { field } => write!(formatter, "{field} cannot be empty"),
            Self::TooLong {
                field,
                length,
                maximum,
            } => write!(formatter, "{field} is {length} bytes; maximum is {maximum}"),
            Self::ContainsControl { field, index } => {
                write!(
                    formatter,
                    "{field} contains a control character at index {index}"
                )
            },
            Self::SurroundingWhitespace { field } => {
                write!(formatter, "{field} cannot start or end with whitespace")
            },
            Self::MissingNumberPlaceholder => {
                formatter.write_str("custom Number pattern needs a '#' or '0' placeholder")
            },
            Self::MissingDateTimeField => {
                formatter.write_str("custom Date & Time pattern needs a date or time field")
            },
            Self::NonFiniteThreshold => {
                formatter.write_str("custom Number threshold must be finite")
            },
            Self::TooManyRules { actual, maximum } => {
                write!(
                    formatter,
                    "custom Number has {actual} rules; maximum is {maximum}"
                )
            },
            Self::Allocation { amount } => {
                write!(formatter, "custom Number could not allocate {amount} rules")
            },
            Self::DuplicateCondition => formatter.write_str("custom Number repeats a condition"),
            Self::EmptyLiteral => formatter.write_str("literal custom Text cannot be empty"),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked custom-format constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated custom-format name.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Name(String);

impl fmt::Debug for Name {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Name")
            .field("byte_len", &self.0.len())
            .finish()
    }
}

impl Name {
    /// Validates a borrowed name before allocating.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the name is empty, too long, contains a
    /// control character, or has surrounding whitespace.
    pub fn new(value: &str) -> Result<Self> {
        validate_visible(value, "custom format name", MAX_NAME_BYTES, false)?;
        if value.trim() != value {
            return Err(Error::SurroundingWhitespace {
                field: "custom format name",
            });
        }
        Ok(Self(value.to_owned()))
    }

    /// Validates and adopts an owned name.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the name is empty, too long, contains a
    /// control character, or has surrounding whitespace.
    pub fn from_owned(value: String) -> Result<Self> {
        validate_visible(&value, "custom format name", MAX_NAME_BYTES, false)?;
        if value.trim() != value {
            return Err(Error::SurroundingWhitespace {
                field: "custom format name",
            });
        }
        Ok(Self(value))
    }

    /// Convenience constructor for callers with an owned or borrowed input.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the name is empty, too long, contains a
    /// control character, or has surrounding whitespace.
    pub fn try_new(value: impl AsRef<str> + Into<String>) -> Result<Self> {
        let borrowed = value.as_ref();
        validate_visible(borrowed, "custom format name", MAX_NAME_BYTES, false)?;
        if borrowed.trim() != borrowed {
            return Err(Error::SurroundingWhitespace {
                field: "custom format name",
            });
        }
        Self::from_owned(value.into())
    }

    /// Borrows the validated name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A validated custom Number pattern.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct NumberPattern(String);

impl fmt::Debug for NumberPattern {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NumberPattern")
            .field("byte_len", &self.0.len())
            .finish()
    }
}

impl NumberPattern {
    /// Validates a pattern containing a digit placeholder.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no `#` or `0` placeholder.
    pub fn new(value: &str) -> Result<Self> {
        validate_visible(value, "custom Number pattern", MAX_PATTERN_BYTES, false)?;
        if !value
            .chars()
            .any(|character| matches!(character, '#' | '0'))
        {
            return Err(Error::MissingNumberPlaceholder);
        }
        Ok(Self(value.to_owned()))
    }

    /// Validates and adopts an owned pattern.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no `#` or `0` placeholder.
    pub fn from_owned(value: String) -> Result<Self> {
        validate_visible(&value, "custom Number pattern", MAX_PATTERN_BYTES, false)?;
        if !value
            .chars()
            .any(|character| matches!(character, '#' | '0'))
        {
            return Err(Error::MissingNumberPlaceholder);
        }
        Ok(Self(value))
    }

    /// Convenience constructor for an owned or borrowed input.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no `#` or `0` placeholder.
    pub fn try_new(value: impl AsRef<str> + Into<String>) -> Result<Self> {
        let borrowed = value.as_ref();
        validate_visible(borrowed, "custom Number pattern", MAX_PATTERN_BYTES, false)?;
        if !borrowed
            .chars()
            .any(|character| matches!(character, '#' | '0'))
        {
            return Err(Error::MissingNumberPlaceholder);
        }
        Self::from_owned(value.into())
    }

    /// Borrows the exact pattern.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A finite custom Number condition threshold.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ConditionValue(u64);

impl fmt::Debug for ConditionValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ConditionValue")
            .field("finite", &true)
            .finish()
    }
}

impl ConditionValue {
    /// Validates and stores a finite threshold without preserving a negative zero.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFiniteThreshold`] when `value` is not finite.
    pub fn try_new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::NonFiniteThreshold);
        }
        Ok(Self(if value == 0.0 { 0.0 } else { value }.to_bits()))
    }

    /// Returns the finite threshold.
    #[must_use]
    pub fn value(self) -> f64 {
        f64::from_bits(self.0)
    }
}

impl TryFrom<f64> for ConditionValue {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::try_new(value)
    }
}

/// Comparison performed by a custom Number rule.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum Condition {
    /// Match values equal to the threshold.
    EqualTo(ConditionValue),
    /// Match values below the threshold.
    LessThan(ConditionValue),
    /// Match values at or below the threshold.
    LessThanOrEqualTo(ConditionValue),
    /// Match values above the threshold.
    GreaterThan(ConditionValue),
    /// Match values at or above the threshold.
    GreaterThanOrEqualTo(ConditionValue),
}

impl fmt::Debug for Condition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let kind = match self {
            Self::EqualTo(_) => "EqualTo",
            Self::LessThan(_) => "LessThan",
            Self::LessThanOrEqualTo(_) => "LessThanOrEqualTo",
            Self::GreaterThan(_) => "GreaterThan",
            Self::GreaterThanOrEqualTo(_) => "GreaterThanOrEqualTo",
        };
        formatter
            .debug_struct("Condition")
            .field("kind", &kind)
            .field("threshold_present", &true)
            .finish()
    }
}

impl Condition {
    /// Returns the comparison threshold.
    #[must_use]
    pub const fn threshold(self) -> ConditionValue {
        match self {
            Self::EqualTo(value)
            | Self::LessThan(value)
            | Self::LessThanOrEqualTo(value)
            | Self::GreaterThan(value)
            | Self::GreaterThanOrEqualTo(value) => value,
        }
    }
}

/// One ordered conditional presentation rule.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct NumberRule {
    condition: Condition,
    pattern: NumberPattern,
}

impl fmt::Debug for NumberRule {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NumberRule")
            .field("condition", &self.condition)
            .field("pattern", &self.pattern)
            .finish()
    }
}

impl NumberRule {
    /// Constructs a rule from checked values.
    #[must_use]
    pub const fn new(condition: Condition, pattern: NumberPattern) -> Self {
        Self { condition, pattern }
    }

    /// Returns the condition.
    #[must_use]
    pub const fn condition(&self) -> Condition {
        self.condition
    }

    /// Returns the rule pattern.
    #[must_use]
    pub const fn pattern(&self) -> &NumberPattern {
        &self.pattern
    }
}

/// A custom Number format with an ordered conditional rule list.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Number {
    name: Name,
    default_pattern: NumberPattern,
    rules: Vec<NumberRule>,
}

impl fmt::Debug for Number {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Number")
            .field("name", &self.name)
            .field("default_pattern", &self.default_pattern)
            .field("rule_count", &self.rules.len())
            .finish()
    }
}

impl Number {
    /// Constructs a custom Number format without rules.
    #[must_use]
    pub const fn new(name: Name, default_pattern: NumberPattern) -> Self {
        Self {
            name,
            default_pattern,
            rules: Vec::new(),
        }
    }

    /// Constructs a custom Number format with checked ordered rules.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManyRules`] when the input exceeds
    /// [`MAX_RULES`], [`Error::Allocation`] when the bounded rule buffer
    /// cannot be reserved, or [`Error::DuplicateCondition`] when two rules
    /// use the same condition. Over-limit input takes precedence over a
    /// duplicate condition, and only the first over-limit item is consumed.
    pub fn try_with_rules(
        name: Name,
        default_pattern: NumberPattern,
        rules: impl IntoIterator<Item = NumberRule>,
    ) -> Result<Self> {
        let iterator = rules.into_iter();
        let (lower_bound, _) = iterator.size_hint();
        let initial_capacity = lower_bound.min(MAX_RULES);
        let mut collected_rules: Vec<NumberRule> = Vec::new();
        if initial_capacity != 0 {
            collected_rules
                .try_reserve_exact(initial_capacity)
                .map_err(|_| Error::Allocation {
                    amount: initial_capacity,
                })?;
        }

        // Retain the old precedence: an over-limit input wins over a
        // duplicate condition, even when the duplicate appears early.  A
        // duplicate therefore gets recorded while iteration continues until
        // either the iterator ends or the first over-limit item is observed.
        let mut duplicate_condition = false;
        for (seen_count, rule) in iterator.enumerate() {
            if seen_count == MAX_RULES {
                return Err(Error::TooManyRules {
                    actual: MAX_RULES + 1,
                    maximum: MAX_RULES,
                });
            }

            if duplicate_condition
                || collected_rules
                    .iter()
                    .any(|existing| existing.condition == rule.condition)
            {
                duplicate_condition = true;
                continue;
            }

            if collected_rules.len() == collected_rules.capacity() {
                collected_rules
                    .try_reserve(1)
                    .map_err(|_| Error::Allocation { amount: 1 })?;
            }
            collected_rules.push(rule);
        }
        if duplicate_condition {
            return Err(Error::DuplicateCondition);
        }
        Ok(Self {
            name,
            default_pattern,
            rules: collected_rules,
        })
    }

    /// Returns the user-visible name.
    #[must_use]
    pub const fn name(&self) -> &Name {
        &self.name
    }

    /// Returns the fallback pattern.
    #[must_use]
    pub const fn default_pattern(&self) -> &NumberPattern {
        &self.default_pattern
    }

    /// Returns conditional rules in native evaluation order.
    #[must_use]
    pub fn rules(&self) -> &[NumberRule] {
        &self.rules
    }
}

/// A validated custom Date & Time pattern.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct DateTimePattern(String);

impl fmt::Debug for DateTimePattern {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DateTimePattern")
            .field("byte_len", &self.0.len())
            .finish()
    }
}

impl DateTimePattern {
    /// Validates a pattern containing at least one date or time field.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no supported date or time field.
    pub fn new(value: &str) -> Result<Self> {
        validate_visible(
            value,
            "custom Date & Time pattern",
            MAX_PATTERN_BYTES,
            false,
        )?;
        if !value.chars().any(|character| {
            matches!(
                character,
                'G' | 'y'
                    | 'Y'
                    | 'M'
                    | 'L'
                    | 'w'
                    | 'W'
                    | 'D'
                    | 'd'
                    | 'F'
                    | 'E'
                    | 'e'
                    | 'a'
                    | 'h'
                    | 'H'
                    | 'K'
                    | 'k'
                    | 'm'
                    | 's'
                    | 'S'
            )
        }) {
            return Err(Error::MissingDateTimeField);
        }
        Ok(Self(value.to_owned()))
    }

    /// Validates and adopts an owned pattern.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no supported date or time field.
    pub fn from_owned(value: String) -> Result<Self> {
        validate_visible(
            &value,
            "custom Date & Time pattern",
            MAX_PATTERN_BYTES,
            false,
        )?;
        if !value.chars().any(|character| {
            matches!(
                character,
                'G' | 'y'
                    | 'Y'
                    | 'M'
                    | 'L'
                    | 'w'
                    | 'W'
                    | 'D'
                    | 'd'
                    | 'F'
                    | 'E'
                    | 'e'
                    | 'a'
                    | 'h'
                    | 'H'
                    | 'K'
                    | 'k'
                    | 'm'
                    | 's'
                    | 'S'
            )
        }) {
            return Err(Error::MissingDateTimeField);
        }
        Ok(Self(value))
    }

    /// Convenience constructor for an owned or borrowed input.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the pattern is empty, too long, contains a
    /// control character, or has no supported date or time field.
    pub fn try_new(value: impl AsRef<str> + Into<String>) -> Result<Self> {
        let borrowed = value.as_ref();
        validate_visible(
            borrowed,
            "custom Date & Time pattern",
            MAX_PATTERN_BYTES,
            false,
        )?;
        if !borrowed.chars().any(|character| {
            matches!(
                character,
                'G' | 'y'
                    | 'Y'
                    | 'M'
                    | 'L'
                    | 'w'
                    | 'W'
                    | 'D'
                    | 'd'
                    | 'F'
                    | 'E'
                    | 'e'
                    | 'a'
                    | 'h'
                    | 'H'
                    | 'K'
                    | 'k'
                    | 'm'
                    | 's'
                    | 'S'
            )
        }) {
            return Err(Error::MissingDateTimeField);
        }
        Self::from_owned(value.into())
    }

    /// Borrows the exact ICU-style pattern.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A custom Date & Time format.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct DateTime {
    name: Name,
    pattern: DateTimePattern,
}

impl fmt::Debug for DateTime {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DateTime")
            .field("name", &self.name)
            .field("pattern", &self.pattern)
            .finish()
    }
}

impl DateTime {
    /// Constructs a custom Date & Time format from validated values.
    #[must_use]
    pub const fn new(name: Name, pattern: DateTimePattern) -> Self {
        Self { name, pattern }
    }

    /// Returns the user-visible name.
    #[must_use]
    pub const fn name(&self) -> &Name {
        &self.name
    }

    /// Returns the ICU-style presentation pattern.
    #[must_use]
    pub const fn pattern(&self) -> &DateTimePattern {
        &self.pattern
    }
}

/// A custom Text format with optional cell text and literal affixes.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Text {
    name: Name,
    prefix: String,
    suffix: String,
    includes_cell: bool,
}

impl fmt::Debug for Text {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Text")
            .field("name", &self.name)
            .field("prefix_byte_len", &self.prefix.len())
            .field("suffix_byte_len", &self.suffix.len())
            .field("includes_cell_text", &self.includes_cell)
            .finish()
    }
}

impl Text {
    /// Constructs a format placing the cell text between two affixes.
    ///
    /// # Errors
    ///
    /// Returns a typed error when either affix is too long or contains a
    /// control character.
    pub fn try_new(
        name: Name,
        prefix: impl AsRef<str> + Into<String>,
        suffix: impl AsRef<str> + Into<String>,
    ) -> Result<Self> {
        Self::try_with_cell_text(name, prefix, suffix, true)
    }

    /// Constructs a literal-only custom Text format.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the literal is empty, too long, or contains
    /// a control character.
    pub fn try_literal(name: Name, literal: impl AsRef<str> + Into<String>) -> Result<Self> {
        Self::try_with_cell_text(name, literal, String::new(), false)
    }

    fn try_with_cell_text<P, S>(
        name: Name,
        prefix: P,
        suffix: S,
        includes_cell: bool,
    ) -> Result<Self>
    where
        P: AsRef<str> + Into<String>,
        S: AsRef<str> + Into<String>,
    {
        // Validate both borrowed values before consuming either one.  Apart
        // from avoiding an unnecessary allocation for rejected input, this
        // keeps the error deterministic when the second affix is invalid.
        let prefix_ref = prefix.as_ref();
        let suffix_ref = suffix.as_ref();
        validate_affix(prefix_ref)?;
        validate_affix(suffix_ref)?;
        let encoded_bytes = prefix_ref
            .len()
            .checked_add(suffix_ref.len())
            .and_then(|length| length.checked_add(usize::from(includes_cell)))
            .ok_or(Error::TooLong {
                field: "custom Text pattern",
                length: usize::MAX,
                maximum: MAX_PATTERN_BYTES,
            })?;
        if encoded_bytes > MAX_PATTERN_BYTES {
            return Err(Error::TooLong {
                field: "custom Text pattern",
                length: encoded_bytes,
                maximum: MAX_PATTERN_BYTES,
            });
        }
        if !includes_cell && prefix_ref.is_empty() && suffix_ref.is_empty() {
            return Err(Error::EmptyLiteral);
        }
        Ok(Self {
            name,
            prefix: prefix.into(),
            suffix: suffix.into(),
            includes_cell,
        })
    }

    /// Returns the user-visible name.
    #[must_use]
    pub const fn name(&self) -> &Name {
        &self.name
    }

    /// Returns the literal before the cell text.
    #[must_use]
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Returns the literal after the cell text.
    #[must_use]
    pub fn suffix(&self) -> &str {
        &self.suffix
    }

    /// Whether the stored cell text appears in the rendered value.
    #[must_use]
    pub const fn includes_cell_text(&self) -> bool {
        self.includes_cell
    }
}

/// One of the three custom cell-format families.
#[derive(Clone, PartialEq, Eq, Hash)]
pub enum Custom {
    /// Custom numeric presentation.
    Number(Number),
    /// Custom text presentation.
    Text(Text),
    /// Custom date-and-time presentation.
    DateTime(DateTime),
}

impl fmt::Debug for Custom {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Number(number) => formatter
                .debug_struct("Custom")
                .field("family", &"Number")
                .field("name_byte_len", &number.name.0.len())
                .field("default_pattern_byte_len", &number.default_pattern.0.len())
                .field("rule_count", &number.rules.len())
                .finish(),
            Self::Text(text) => formatter
                .debug_struct("Custom")
                .field("family", &"Text")
                .field("name_byte_len", &text.name.0.len())
                .field("prefix_byte_len", &text.prefix.len())
                .field("suffix_byte_len", &text.suffix.len())
                .field("includes_cell_text", &text.includes_cell)
                .finish(),
            Self::DateTime(date_time) => formatter
                .debug_struct("Custom")
                .field("family", &"DateTime")
                .field("name_byte_len", &date_time.name.0.len())
                .field("pattern_byte_len", &date_time.pattern.0.len())
                .finish(),
        }
    }
}

impl From<Number> for Custom {
    fn from(value: Number) -> Self {
        Self::Number(value)
    }
}

impl From<Text> for Custom {
    fn from(value: Text) -> Self {
        Self::Text(value)
    }
}

impl From<DateTime> for Custom {
    fn from(value: DateTime) -> Self {
        Self::DateTime(value)
    }
}

fn validate_visible(
    value: &str,
    field: &'static str,
    maximum: usize,
    allow_empty: bool,
) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return Err(Error::Empty { field });
    }
    if value.len() > maximum {
        return Err(Error::TooLong {
            field,
            length: value.len(),
            maximum,
        });
    }
    if let Some((index, _)) = value
        .chars()
        .enumerate()
        .find(|(_, character)| character.is_control())
    {
        return Err(Error::ContainsControl { field, index });
    }
    Ok(())
}

fn validate_affix(value: &str) -> Result<()> {
    validate_visible(value, "custom Text affix", MAX_PATTERN_BYTES, true)
}

#[cfg(test)]
mod tests {
    use super::*;

    struct NoStringConversion<'a>(&'a str);

    impl AsRef<str> for NoStringConversion<'_> {
        fn as_ref(&self) -> &str {
            self.0
        }
    }

    impl From<NoStringConversion<'_>> for String {
        fn from(_value: NoStringConversion<'_>) -> Self {
            panic!("invalid custom input was converted")
        }
    }

    fn rule(index: usize) -> NumberRule {
        NumberRule::new(
            Condition::GreaterThan(ConditionValue::try_new(index as f64).unwrap()),
            NumberPattern::new("#,##0").unwrap(),
        )
    }

    struct DishonestRules {
        next: usize,
        count: usize,
    }

    impl Iterator for DishonestRules {
        type Item = NumberRule;

        fn next(&mut self) -> Option<Self::Item> {
            if self.next == self.count {
                return None;
            }
            let rule = rule(self.next);
            self.next += 1;
            Some(rule)
        }

        fn size_hint(&self) -> (usize, Option<usize>) {
            // Deliberately claim an impossible lower bound.  The constructor
            // must cap reservation and inspect the iterator itself.
            (usize::MAX, None)
        }
    }

    #[test]
    fn custom_values_reject_malformed_input() {
        assert!(matches!(Name::new(""), Err(Error::Empty { .. })));
        assert!(matches!(
            Name::new(" name"),
            Err(Error::SurroundingWhitespace { .. })
        ));
        assert!(matches!(
            NumberPattern::new("literal"),
            Err(Error::MissingNumberPlaceholder)
        ));
        assert_eq!(
            ConditionValue::try_new(f64::NAN),
            Err(Error::NonFiniteThreshold)
        );
        assert!(matches!(
            Text::try_literal(Name::new("Literal").unwrap(), ""),
            Err(Error::EmptyLiteral)
        ));
    }

    #[test]
    fn custom_values_preserve_rules_and_affixes() {
        let name = Name::new("Accounting").unwrap();
        let pattern = NumberPattern::new("#,##0").unwrap();
        let rule = NumberRule::new(
            Condition::LessThan(ConditionValue::try_new(0.0).unwrap()),
            NumberPattern::new("(#,##0)").unwrap(),
        );
        let number = Number::try_with_rules(name.clone(), pattern, [rule.clone()]).unwrap();
        assert_eq!(number.name().as_str(), "Accounting");
        assert_eq!(number.rules(), &[rule]);

        let text = Text::try_new(name, "ID: ", "").unwrap();
        assert_eq!(text.prefix(), "ID: ");
        assert_eq!(text.suffix(), "");
        assert!(text.includes_cell_text());
    }

    #[test]
    fn borrowed_oversized_values_are_rejected_before_conversion() {
        let oversized_name = "n".repeat(MAX_NAME_BYTES + 1);
        assert!(matches!(
            Name::try_new(NoStringConversion(&oversized_name)),
            Err(Error::TooLong {
                field: "custom format name",
                ..
            })
        ));

        let oversized_number_pattern = format!("{}#", "n".repeat(MAX_PATTERN_BYTES));
        assert!(matches!(
            NumberPattern::try_new(NoStringConversion(&oversized_number_pattern)),
            Err(Error::TooLong {
                field: "custom Number pattern",
                ..
            })
        ));

        let oversized_date_time_pattern = format!("{}y", "n".repeat(MAX_PATTERN_BYTES));
        assert!(matches!(
            DateTimePattern::try_new(NoStringConversion(&oversized_date_time_pattern)),
            Err(Error::TooLong {
                field: "custom Date & Time pattern",
                ..
            })
        ));

        let oversized_affix = "a".repeat(MAX_PATTERN_BYTES + 1);
        assert!(matches!(
            Text::try_new(
                Name::new("Text").unwrap(),
                NoStringConversion("prefix"),
                NoStringConversion(&oversized_affix),
            ),
            Err(Error::TooLong {
                field: "custom Text affix",
                ..
            })
        ));
    }

    #[test]
    fn number_rules_stop_at_the_first_over_limit_item_even_when_infinite() {
        let name = Name::new("Accounting").unwrap();
        let pattern = NumberPattern::new("#,##0").unwrap();
        let repeated = rule(0);

        assert_eq!(
            Number::try_with_rules(name, pattern, std::iter::repeat(repeated)),
            Err(Error::TooManyRules {
                actual: MAX_RULES + 1,
                maximum: MAX_RULES,
            })
        );
    }

    #[test]
    fn number_rules_do_not_trust_a_dishonest_size_hint() {
        let name = Name::new("Accounting").unwrap();
        let pattern = NumberPattern::new("#,##0").unwrap();

        assert_eq!(
            Number::try_with_rules(
                name,
                pattern,
                DishonestRules {
                    next: 0,
                    count: MAX_RULES + 1,
                },
            ),
            Err(Error::TooManyRules {
                actual: MAX_RULES + 1,
                maximum: MAX_RULES,
            })
        );
    }

    #[test]
    fn number_rules_accept_exactly_the_bounded_rule_count() {
        let name = Name::new("Accounting").unwrap();
        let pattern = NumberPattern::new("#,##0").unwrap();
        let number = Number::try_with_rules(name, pattern, (0..MAX_RULES).map(rule)).unwrap();

        assert_eq!(number.rules().len(), MAX_RULES);
    }

    #[test]
    fn number_rule_overflow_takes_precedence_over_an_early_duplicate() {
        let name = Name::new("Accounting").unwrap();
        let pattern = NumberPattern::new("#,##0").unwrap();
        let repeated = rule(0);
        let rules = std::iter::once(repeated.clone()).chain(std::iter::repeat(repeated));

        assert_eq!(
            Number::try_with_rules(name, pattern, rules),
            Err(Error::TooManyRules {
                actual: MAX_RULES + 1,
                maximum: MAX_RULES,
            })
        );
    }

    #[test]
    fn custom_debug_redacts_all_user_content_and_thresholds() {
        const NAME_MARKER: &str = "custom-name-secret-marker";
        const DEFAULT_MARKER: &str = "custom-default-secret-marker#";
        const RULE_MARKER: &str = "custom-rule-secret-marker#";
        const PREFIX_MARKER: &str = "custom-prefix-secret-marker";
        const SUFFIX_MARKER: &str = "custom-suffix-secret-marker";
        const DATE_TIME_MARKER: &str = "custom-date-secret-marker-yyyy";
        const THRESHOLD: f64 = 987_654_321.25;

        let name = Name::new(NAME_MARKER).unwrap();
        let default_pattern = NumberPattern::new(DEFAULT_MARKER).unwrap();
        let threshold = ConditionValue::try_new(THRESHOLD).unwrap();
        let condition = Condition::GreaterThan(threshold);
        let rule = NumberRule::new(condition, NumberPattern::new(RULE_MARKER).unwrap());
        let number =
            Number::try_with_rules(name.clone(), default_pattern.clone(), [rule.clone()]).unwrap();
        let text = Text::try_new(name.clone(), PREFIX_MARKER, SUFFIX_MARKER).unwrap();
        let date_time = DateTime::new(name, DateTimePattern::new(DATE_TIME_MARKER).unwrap());
        let error = Name::new("custom-error-secret-marker\u{0007}").unwrap_err();

        for rendered in [
            format!("{number:?}"),
            format!("{:?}", Custom::Number(number.clone())),
            format!("{text:?}"),
            format!("{:?}", Custom::Text(text.clone())),
            format!("{date_time:?}"),
            format!("{:?}", Custom::DateTime(date_time.clone())),
            format!("{rule:?}"),
            format!("{default_pattern:?}"),
            format!("{threshold:?}"),
            format!("{condition:?}"),
            format!("{error:?}"),
            format!("{error}"),
        ] {
            for marker in [
                NAME_MARKER,
                DEFAULT_MARKER,
                RULE_MARKER,
                PREFIX_MARKER,
                SUFFIX_MARKER,
                DATE_TIME_MARKER,
                "987654321.25",
            ] {
                assert!(
                    !rendered.contains(marker),
                    "custom debug/error output leaked {marker:?}: {rendered}"
                );
            }
        }
    }
}
