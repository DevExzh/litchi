//! Typed text layout shared by native iWork table cells.

use crate::{Error, Result};

/// Whether text remains on one line or wraps inside its table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellTextWrap {
    #[default]
    Unwrapped,
    Wrapped,
}

/// Vertical placement of text inside a native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellVerticalAlignment {
    #[default]
    Top,
    Middle,
    Bottom,
}

/// Finite, non-negative distance from one cell edge, measured in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TableCellInset(f32);

impl TableCellInset {
    pub const ZERO: Self = Self(0.0);

    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ParseError(
                "iWork table-cell inset must be finite and non-negative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Independently typed text insets for all four cell edges.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellInsets {
    left: TableCellInset,
    top: TableCellInset,
    right: TableCellInset,
    bottom: TableCellInset,
}

impl TableCellInsets {
    pub const ZERO: Self = Self::uniform(TableCellInset::ZERO);

    pub const fn new(
        left: TableCellInset,
        top: TableCellInset,
        right: TableCellInset,
        bottom: TableCellInset,
    ) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    pub const fn uniform(inset: TableCellInset) -> Self {
        Self::new(inset, inset, inset, inset)
    }

    pub const fn left(self) -> TableCellInset {
        self.left
    }

    pub const fn top(self) -> TableCellInset {
        self.top
    }

    pub const fn right(self) -> TableCellInset {
        self.right
    }

    pub const fn bottom(self) -> TableCellInset {
        self.bottom
    }
}

impl Default for TableCellInsets {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Effective composable text layout for one native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct TableCellLayout {
    text_wrap: TableCellTextWrap,
    vertical_alignment: TableCellVerticalAlignment,
    insets: TableCellInsets,
}

impl TableCellLayout {
    pub const fn new(
        text_wrap: TableCellTextWrap,
        vertical_alignment: TableCellVerticalAlignment,
        insets: TableCellInsets,
    ) -> Self {
        Self {
            text_wrap,
            vertical_alignment,
            insets,
        }
    }

    pub const fn text_wrap(self) -> TableCellTextWrap {
        self.text_wrap
    }

    pub const fn vertical_alignment(self) -> TableCellVerticalAlignment {
        self.vertical_alignment
    }

    pub const fn insets(self) -> TableCellInsets {
        self.insets
    }

    pub const fn with_text_wrap(mut self, text_wrap: TableCellTextWrap) -> Self {
        self.text_wrap = text_wrap;
        self
    }

    pub const fn with_vertical_alignment(
        mut self,
        vertical_alignment: TableCellVerticalAlignment,
    ) -> Self {
        self.vertical_alignment = vertical_alignment;
        self
    }

    pub const fn with_insets(mut self, insets: TableCellInsets) -> Self {
        self.insets = insets;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inset_values_are_strict_and_layout_is_composable() {
        assert!(TableCellInset::from_points(-0.01).is_err());
        assert!(TableCellInset::from_points(f32::NAN).is_err());
        assert!(TableCellInset::from_points(f32::INFINITY).is_err());

        let inset = TableCellInset::from_points(6.0).unwrap();
        let layout = TableCellLayout::default()
            .with_text_wrap(TableCellTextWrap::Wrapped)
            .with_vertical_alignment(TableCellVerticalAlignment::Middle)
            .with_insets(TableCellInsets::uniform(inset));
        assert_eq!(layout.insets().left().points(), 6.0);
        assert_eq!(
            layout.vertical_alignment(),
            TableCellVerticalAlignment::Middle
        );
    }
}
