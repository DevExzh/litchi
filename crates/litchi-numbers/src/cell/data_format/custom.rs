//! Dependency-free custom cell-format values.

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
            Self::DuplicateCondition => formatter.write_str("custom Number repeats a condition"),
            Self::EmptyLiteral => formatter.write_str("literal custom Text cannot be empty"),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked custom-format constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated custom-format name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Name(String);

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
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        Self::from_owned(value.into())
    }

    /// Borrows the validated name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A validated custom Number pattern.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NumberPattern(String);

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
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        Self::from_owned(value.into())
    }

    /// Borrows the exact pattern.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A finite custom Number condition threshold.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ConditionValue(u64);

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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NumberRule {
    condition: Condition,
    pattern: NumberPattern,
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
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Number {
    name: Name,
    default_pattern: NumberPattern,
    rules: Vec<NumberRule>,
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
    /// [`MAX_RULES`], or [`Error::DuplicateCondition`] when two rules
    /// use the same condition.
    pub fn try_with_rules(
        name: Name,
        default_pattern: NumberPattern,
        rules: impl IntoIterator<Item = NumberRule>,
    ) -> Result<Self> {
        let collected_rules = rules.into_iter().collect::<Vec<_>>();
        if collected_rules.len() > MAX_RULES {
            return Err(Error::TooManyRules {
                actual: collected_rules.len(),
                maximum: MAX_RULES,
            });
        }
        for (index, rule) in collected_rules.iter().enumerate() {
            if collected_rules[..index]
                .iter()
                .any(|existing| existing.condition == rule.condition)
            {
                return Err(Error::DuplicateCondition);
            }
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
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct DateTimePattern(String);

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
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        Self::from_owned(value.into())
    }

    /// Borrows the exact ICU-style pattern.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A custom Date & Time format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct DateTime {
    name: Name,
    pattern: DateTimePattern,
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
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Text {
    name: Name,
    prefix: String,
    suffix: String,
    includes_cell: bool,
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
        prefix: impl Into<String>,
        suffix: impl Into<String>,
    ) -> Result<Self> {
        Self::try_with_cell_text(name, prefix, suffix, true)
    }

    /// Constructs a literal-only custom Text format.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the literal is empty, too long, or contains
    /// a control character.
    pub fn try_literal(name: Name, literal: impl Into<String>) -> Result<Self> {
        Self::try_with_cell_text(name, literal, String::new(), false)
    }

    fn try_with_cell_text(
        name: Name,
        prefix: impl Into<String>,
        suffix: impl Into<String>,
        includes_cell: bool,
    ) -> Result<Self> {
        let prefix_text = prefix.into();
        let suffix_text = suffix.into();
        validate_affix(&prefix_text)?;
        validate_affix(&suffix_text)?;
        let encoded_bytes = prefix_text
            .len()
            .checked_add(suffix_text.len())
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
        if !includes_cell && prefix_text.is_empty() && suffix_text.is_empty() {
            return Err(Error::EmptyLiteral);
        }
        Ok(Self {
            name,
            prefix: prefix_text,
            suffix: suffix_text,
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
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Custom {
    /// Custom numeric presentation.
    Number(Number),
    /// Custom text presentation.
    Text(Text),
    /// Custom date-and-time presentation.
    DateTime(DateTime),
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
}
