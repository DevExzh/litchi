//! Strongly typed custom table-cell formats.

use crate::{Error, Result};

const MAXIMUM_CUSTOM_FORMAT_NAME_BYTES: usize = 255;
const MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES: usize = 4_096;
const MAXIMUM_CUSTOM_NUMBER_RULES: usize = 32;
pub(crate) const NATIVE_CUSTOM_TEXT_VALUE_TOKEN: char = '\u{e421}';

/// User-visible name of a custom table-cell format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomFormatName(String);

impl TableCellCustomFormatName {
    /// Validate a non-empty, bounded name accepted by the iWork inspector.
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_visible_text(
            &value,
            "custom table-cell format name",
            MAXIMUM_CUSTOM_FORMAT_NAME_BYTES,
            false,
        )?;
        if value.trim() != value {
            return Err(Error::InvalidFormat(
                "custom table-cell format name cannot start or end with whitespace".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the validated name.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Validated native pattern for a custom Number format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomNumberPattern(String);

impl TableCellCustomNumberPattern {
    /// Validate a Number pattern containing at least one digit placeholder.
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_visible_text(
            &value,
            "custom Number pattern",
            MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES,
            false,
        )?;
        if !value
            .chars()
            .any(|character| matches!(character, '#' | '0'))
        {
            return Err(Error::InvalidFormat(
                "custom Number pattern must contain a '#' or '0' digit placeholder".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the validated native pattern.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Finite threshold used by one custom Number rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellCustomNumberConditionValue(u64);

impl TableCellCustomNumberConditionValue {
    /// Validate and construct a finite rule threshold.
    pub fn try_new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::InvalidFormat(
                "custom Number rule threshold must be finite".to_owned(),
            ));
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        Ok(Self(normalized.to_bits()))
    }

    /// Return the finite threshold.
    pub fn value(self) -> f64 {
        f64::from_bits(self.0)
    }
}

impl TryFrom<f64> for TableCellCustomNumberConditionValue {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::try_new(value)
    }
}

/// Comparison performed by one custom Number rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableCellCustomNumberCondition {
    /// Match values equal to the threshold.
    EqualTo(TableCellCustomNumberConditionValue),
    /// Match values below the threshold.
    LessThan(TableCellCustomNumberConditionValue),
    /// Match values at or below the threshold.
    LessThanOrEqualTo(TableCellCustomNumberConditionValue),
    /// Match values above the threshold.
    GreaterThan(TableCellCustomNumberConditionValue),
    /// Match values at or above the threshold.
    GreaterThanOrEqualTo(TableCellCustomNumberConditionValue),
}

impl TableCellCustomNumberCondition {
    /// Return the comparison threshold.
    pub const fn threshold(self) -> TableCellCustomNumberConditionValue {
        match self {
            Self::EqualTo(value)
            | Self::LessThan(value)
            | Self::LessThanOrEqualTo(value)
            | Self::GreaterThan(value)
            | Self::GreaterThanOrEqualTo(value) => value,
        }
    }
}

/// Conditional presentation rule in a custom Number format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomNumberRule {
    condition: TableCellCustomNumberCondition,
    pattern: TableCellCustomNumberPattern,
}

impl TableCellCustomNumberRule {
    /// Construct a rule from a typed comparison and validated pattern.
    pub const fn new(
        condition: TableCellCustomNumberCondition,
        pattern: TableCellCustomNumberPattern,
    ) -> Self {
        Self { condition, pattern }
    }

    /// Return the comparison.
    pub const fn condition(&self) -> TableCellCustomNumberCondition {
        self.condition
    }

    /// Return the presentation used when the comparison matches.
    pub const fn pattern(&self) -> &TableCellCustomNumberPattern {
        &self.pattern
    }
}

/// Custom Number format with a fallback pattern and ordered rules.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomNumberFormat {
    name: TableCellCustomFormatName,
    default_pattern: TableCellCustomNumberPattern,
    rules: Vec<TableCellCustomNumberRule>,
}

impl TableCellCustomNumberFormat {
    /// Construct a custom Number format without conditional rules.
    pub const fn new(
        name: TableCellCustomFormatName,
        default_pattern: TableCellCustomNumberPattern,
    ) -> Self {
        Self {
            name,
            default_pattern,
            rules: Vec::new(),
        }
    }

    /// Construct and validate an ordered rule set.
    pub fn try_with_rules(
        name: TableCellCustomFormatName,
        default_pattern: TableCellCustomNumberPattern,
        rules: impl IntoIterator<Item = TableCellCustomNumberRule>,
    ) -> Result<Self> {
        let rules = rules.into_iter().collect::<Vec<_>>();
        if rules.len() > MAXIMUM_CUSTOM_NUMBER_RULES {
            return Err(Error::InvalidFormat(format!(
                "custom Number format cannot contain more than {MAXIMUM_CUSTOM_NUMBER_RULES} rules"
            )));
        }
        for (index, rule) in rules.iter().enumerate() {
            if rules[..index]
                .iter()
                .any(|existing| existing.condition == rule.condition)
            {
                return Err(Error::InvalidFormat(
                    "custom Number format cannot contain duplicate conditions".to_owned(),
                ));
            }
        }
        Ok(Self {
            name,
            default_pattern,
            rules,
        })
    }

    /// Return the user-visible name.
    pub const fn name(&self) -> &TableCellCustomFormatName {
        &self.name
    }

    /// Return the fallback pattern.
    pub const fn default_pattern(&self) -> &TableCellCustomNumberPattern {
        &self.default_pattern
    }

    /// Return conditional rules in native evaluation order.
    pub fn rules(&self) -> &[TableCellCustomNumberRule] {
        &self.rules
    }
}

/// Validated native pattern for a custom Date & Time format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomDateTimePattern(String);

impl TableCellCustomDateTimePattern {
    /// Validate an ICU-style pattern containing a date or time field.
    pub fn try_new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_visible_text(
            &value,
            "custom Date & Time pattern",
            MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES,
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
            return Err(Error::InvalidFormat(
                "custom Date & Time pattern must contain a date or time field".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the validated ICU-style pattern.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Custom Date & Time format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomDateTimeFormat {
    name: TableCellCustomFormatName,
    pattern: TableCellCustomDateTimePattern,
}

impl TableCellCustomDateTimeFormat {
    /// Construct a custom Date & Time format.
    pub const fn new(
        name: TableCellCustomFormatName,
        pattern: TableCellCustomDateTimePattern,
    ) -> Self {
        Self { name, pattern }
    }

    /// Return the user-visible name.
    pub const fn name(&self) -> &TableCellCustomFormatName {
        &self.name
    }

    /// Return the ICU-style presentation pattern.
    pub const fn pattern(&self) -> &TableCellCustomDateTimePattern {
        &self.pattern
    }
}

/// Custom Text format with optional cell text and literal affixes.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellCustomTextFormat {
    name: TableCellCustomFormatName,
    prefix: String,
    suffix: String,
    includes_cell_text: bool,
}

impl TableCellCustomTextFormat {
    /// Construct a format that places the cell text between literal affixes.
    pub fn try_new(
        name: TableCellCustomFormatName,
        prefix: impl Into<String>,
        suffix: impl Into<String>,
    ) -> Result<Self> {
        Self::try_with_cell_text(name, prefix, suffix, true)
    }

    /// Construct a literal-only custom Text format.
    pub fn try_literal(
        name: TableCellCustomFormatName,
        literal: impl Into<String>,
    ) -> Result<Self> {
        Self::try_with_cell_text(name, literal, String::new(), false)
    }

    fn try_with_cell_text(
        name: TableCellCustomFormatName,
        prefix: impl Into<String>,
        suffix: impl Into<String>,
        includes_cell_text: bool,
    ) -> Result<Self> {
        let prefix = prefix.into();
        let suffix = suffix.into();
        validate_text_affix(&prefix)?;
        validate_text_affix(&suffix)?;
        let encoded_bytes = prefix
            .len()
            .checked_add(suffix.len())
            .and_then(|length| length.checked_add(usize::from(includes_cell_text)))
            .ok_or_else(|| Error::InvalidFormat("custom Text pattern is too large".to_owned()))?;
        if encoded_bytes > MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES {
            return Err(Error::InvalidFormat(format!(
                "custom Text pattern cannot exceed {MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES} bytes"
            )));
        }
        if !includes_cell_text && prefix.is_empty() && suffix.is_empty() {
            return Err(Error::InvalidFormat(
                "literal custom Text format cannot be empty".to_owned(),
            ));
        }
        Ok(Self {
            name,
            prefix,
            suffix,
            includes_cell_text,
        })
    }

    /// Return the user-visible name.
    pub const fn name(&self) -> &TableCellCustomFormatName {
        &self.name
    }

    /// Return the literal before the cell text.
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Return the literal after the cell text.
    pub fn suffix(&self) -> &str {
        &self.suffix
    }

    /// Whether the stored cell text appears in the rendered value.
    pub const fn includes_cell_text(&self) -> bool {
        self.includes_cell_text
    }

    pub(crate) fn native_pattern(&self) -> String {
        let mut pattern = String::with_capacity(
            self.prefix.len() + self.suffix.len() + usize::from(self.includes_cell_text) * 3,
        );
        pattern.push_str(&self.prefix);
        if self.includes_cell_text {
            pattern.push(NATIVE_CUSTOM_TEXT_VALUE_TOKEN);
        }
        pattern.push_str(&self.suffix);
        pattern
    }
}

/// One of iWork's three custom table-cell format families.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TableCellCustomFormat {
    /// Custom numeric pattern with optional conditional rules.
    Number(TableCellCustomNumberFormat),
    /// Custom Text pattern with literal affixes.
    Text(TableCellCustomTextFormat),
    /// Custom Date & Time pattern.
    DateTime(TableCellCustomDateTimeFormat),
}

impl From<TableCellCustomNumberFormat> for TableCellCustomFormat {
    fn from(value: TableCellCustomNumberFormat) -> Self {
        Self::Number(value)
    }
}

impl From<TableCellCustomTextFormat> for TableCellCustomFormat {
    fn from(value: TableCellCustomTextFormat) -> Self {
        Self::Text(value)
    }
}

impl From<TableCellCustomDateTimeFormat> for TableCellCustomFormat {
    fn from(value: TableCellCustomDateTimeFormat) -> Self {
        Self::DateTime(value)
    }
}

fn validate_text_affix(value: &str) -> Result<()> {
    validate_visible_text(
        value,
        "custom Text affix",
        MAXIMUM_CUSTOM_FORMAT_PATTERN_BYTES,
        true,
    )?;
    if value.contains(NATIVE_CUSTOM_TEXT_VALUE_TOKEN) {
        return Err(Error::InvalidFormat(
            "custom Text affix cannot contain iWork's private cell-value token".to_owned(),
        ));
    }
    Ok(())
}

fn validate_visible_text(
    value: &str,
    label: &str,
    maximum_bytes: usize,
    allow_empty: bool,
) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} cannot be empty")));
    }
    if value.len() > maximum_bytes {
        return Err(Error::InvalidFormat(format!(
            "{label} cannot exceed {maximum_bytes} bytes"
        )));
    }
    if value.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(format!(
            "{label} cannot contain control characters"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn custom_formats_are_bounded_typed_and_composable() {
        let name = TableCellCustomFormatName::try_new("Grouped Integer").unwrap();
        let pattern = TableCellCustomNumberPattern::try_new("#,###").unwrap();
        let threshold = TableCellCustomNumberConditionValue::try_new(0.0).unwrap();
        let rule = TableCellCustomNumberRule::new(
            TableCellCustomNumberCondition::LessThan(threshold),
            pattern.clone(),
        );
        let number = TableCellCustomNumberFormat::try_with_rules(name, pattern, [rule]).unwrap();
        assert_eq!(number.rules().len(), 1);
        assert!(TableCellCustomNumberConditionValue::try_new(f64::NAN).is_err());
        assert!(TableCellCustomNumberPattern::try_new("literal").is_err());

        let text = TableCellCustomTextFormat::try_new(
            TableCellCustomFormatName::try_new("Identifier").unwrap(),
            "ID: ",
            "",
        )
        .unwrap();
        assert_eq!(
            text.native_pattern(),
            format!("ID: {NATIVE_CUSTOM_TEXT_VALUE_TOKEN}")
        );

        assert!(TableCellCustomDateTimePattern::try_new("---").is_err());
        assert!(TableCellCustomFormatName::try_new(" trailing ").is_err());
    }
}
