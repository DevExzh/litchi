//! Strictly typed conditional-highlight rules shared by iWork table editors.

use crate::shapes::RgbaColor;
use crate::{Error, Result};

/// Finite numeric operand used by a conditional-highlight comparison.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellConditionalHighlightNumber(f64);

impl TableCellConditionalHighlightNumber {
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "iWork conditional-highlight numbers must be finite".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    pub const fn get(self) -> f64 {
        self.0
    }
}

/// Ordered finite bounds used by a range conditional-highlight rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellConditionalHighlightRange {
    lower: TableCellConditionalHighlightNumber,
    upper: TableCellConditionalHighlightNumber,
}

impl TableCellConditionalHighlightRange {
    pub fn new(
        lower: TableCellConditionalHighlightNumber,
        upper: TableCellConditionalHighlightNumber,
    ) -> Result<Self> {
        if lower.get() > upper.get() {
            return Err(Error::ParseError(
                "iWork conditional-highlight range bounds must be ordered".to_owned(),
            ));
        }
        Ok(Self { lower, upper })
    }

    pub const fn lower(self) -> TableCellConditionalHighlightNumber {
        self.lower
    }

    pub const fn upper(self) -> TableCellConditionalHighlightNumber {
        self.upper
    }
}

/// Numeric condition evaluated against the cell carrying the rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TableCellConditionalHighlightCondition {
    EqualTo(TableCellConditionalHighlightNumber),
    NotEqualTo(TableCellConditionalHighlightNumber),
    GreaterThan(TableCellConditionalHighlightNumber),
    GreaterThanOrEqualTo(TableCellConditionalHighlightNumber),
    LessThan(TableCellConditionalHighlightNumber),
    LessThanOrEqualTo(TableCellConditionalHighlightNumber),
    Between(TableCellConditionalHighlightRange),
    NotBetween(TableCellConditionalHighlightRange),
}

impl TableCellConditionalHighlightCondition {
    pub const fn single_operand(self) -> Option<TableCellConditionalHighlightNumber> {
        match self {
            Self::EqualTo(value)
            | Self::NotEqualTo(value)
            | Self::GreaterThan(value)
            | Self::GreaterThanOrEqualTo(value)
            | Self::LessThan(value)
            | Self::LessThanOrEqualTo(value) => Some(value),
            Self::Between(_) | Self::NotBetween(_) => None,
        }
    }

    pub const fn range(self) -> Option<TableCellConditionalHighlightRange> {
        match self {
            Self::Between(range) | Self::NotBetween(range) => Some(range),
            Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_) => None,
        }
    }
}

/// Visual overrides applied when a conditional-highlight rule matches.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellConditionalHighlightStyle {
    fill: Option<RgbaColor>,
    text_color: Option<RgbaColor>,
    bold: bool,
}

impl TableCellConditionalHighlightStyle {
    /// Construct a non-empty highlight style.
    pub fn new(fill: Option<RgbaColor>, text_color: Option<RgbaColor>, bold: bool) -> Result<Self> {
        if fill.is_none() && text_color.is_none() && !bold {
            return Err(Error::ParseError(
                "an iWork conditional-highlight style must override at least one property"
                    .to_owned(),
            ));
        }
        Ok(Self {
            fill,
            text_color,
            bold,
        })
    }

    pub const fn with_fill(fill: RgbaColor) -> Self {
        Self {
            fill: Some(fill),
            text_color: None,
            bold: false,
        }
    }

    pub const fn with_text_color(text_color: RgbaColor) -> Self {
        Self {
            fill: None,
            text_color: Some(text_color),
            bold: false,
        }
    }

    pub const fn fill(self) -> Option<RgbaColor> {
        self.fill
    }

    pub const fn text_color(self) -> Option<RgbaColor> {
        self.text_color
    }

    pub const fn bold(self) -> bool {
        self.bold
    }
}

/// One ordered conditional-highlight condition and its visual style.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellConditionalHighlightRule {
    pub condition: TableCellConditionalHighlightCondition,
    pub style: TableCellConditionalHighlightStyle,
}

impl TableCellConditionalHighlightRule {
    pub const fn new(
        condition: TableCellConditionalHighlightCondition,
        style: TableCellConditionalHighlightStyle,
    ) -> Self {
        Self { condition, style }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::{RgbColorSpace, RgbaColor};

    #[test]
    fn numeric_operands_and_styles_reject_empty_or_non_finite_values() {
        assert!(TableCellConditionalHighlightNumber::new(f64::NAN).is_err());
        assert!(TableCellConditionalHighlightNumber::new(f64::INFINITY).is_err());
        assert!(TableCellConditionalHighlightStyle::new(None, None, false).is_err());
        let lower = TableCellConditionalHighlightNumber::new(7.0).unwrap();
        let upper = TableCellConditionalHighlightNumber::new(3.0).unwrap();
        assert!(TableCellConditionalHighlightRange::new(lower, upper).is_err());

        let color = RgbaColor::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
        assert_eq!(
            TableCellConditionalHighlightStyle::with_fill(color).fill(),
            Some(color)
        );
    }
}
