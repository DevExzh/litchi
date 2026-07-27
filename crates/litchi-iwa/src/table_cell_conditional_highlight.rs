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

/// Numeric condition evaluated against the cell carrying the rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TableCellConditionalHighlightCondition {
    EqualTo(TableCellConditionalHighlightNumber),
    NotEqualTo(TableCellConditionalHighlightNumber),
    GreaterThan(TableCellConditionalHighlightNumber),
    GreaterThanOrEqualTo(TableCellConditionalHighlightNumber),
    LessThan(TableCellConditionalHighlightNumber),
    LessThanOrEqualTo(TableCellConditionalHighlightNumber),
}

impl TableCellConditionalHighlightCondition {
    pub const fn operand(self) -> TableCellConditionalHighlightNumber {
        match self {
            Self::EqualTo(value)
            | Self::NotEqualTo(value)
            | Self::GreaterThan(value)
            | Self::GreaterThanOrEqualTo(value)
            | Self::LessThan(value)
            | Self::LessThanOrEqualTo(value) => value,
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

        let color = RgbaColor::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
        assert_eq!(
            TableCellConditionalHighlightStyle::with_fill(color).fill(),
            Some(color)
        );
    }
}
