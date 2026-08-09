//! Typed worksheet page-break values.

use crate::error::{Result, invalid};

/// Maximum number of horizontal breaks accepted by desktop Office.
pub const MAX_HORIZONTAL_BREAKS: usize = 1_022;
/// Maximum number of vertical breaks accepted by desktop Office.
pub const MAX_VERTICAL_BREAKS: usize = 1_023;

const LAST_ROW: u32 = 1_048_575;
const LAST_COLUMN: u32 = 16_383;

/// Worksheet axis owned by one page-break collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    /// Breaks between rows (`rowBreaks`).
    Horizontal,
    /// Breaks between columns (`colBreaks`).
    Vertical,
}

/// One checked `SpreadsheetML` page break.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Break {
    id: u32,
    minimum: u32,
    maximum: u32,
    manual: bool,
    pivot: bool,
}

impl Break {
    /// Create a page break with an inclusive perpendicular span.
    ///
    /// Axis-specific worksheet bounds are checked when the value enters a
    /// [`Collection`].
    ///
    /// # Errors
    ///
    /// Returns an error when `minimum` exceeds `maximum`.
    pub fn new(id: u32, minimum: u32, maximum: u32) -> Result<Self> {
        if minimum > maximum {
            return Err(invalid("page-break minimum exceeds maximum"));
        }
        Ok(Self {
            id,
            minimum,
            maximum,
            manual: false,
            pivot: false,
        })
    }

    /// Zero-based row or column immediately before the break.
    #[must_use]
    pub const fn id(self) -> u32 {
        self.id
    }

    /// Inclusive perpendicular start coordinate.
    #[must_use]
    pub const fn minimum(self) -> u32 {
        self.minimum
    }

    /// Inclusive perpendicular end coordinate.
    #[must_use]
    pub const fn maximum(self) -> u32 {
        self.maximum
    }

    /// Whether this is an authored manual break.
    #[must_use]
    pub const fn is_manual(self) -> bool {
        self.manual
    }

    /// Whether this break belongs to a `PivotTable` area.
    #[must_use]
    pub const fn is_pivot(self) -> bool {
        self.pivot
    }

    /// Set the manual-break marker.
    #[must_use]
    pub const fn with_manual(mut self, manual: bool) -> Self {
        self.manual = manual;
        self
    }

    /// Set the PivotTable-break marker.
    #[must_use]
    pub const fn with_pivot(mut self, pivot: bool) -> Self {
        self.pivot = pivot;
        self
    }

    pub(super) fn validate(self, axis: Axis) -> Result<()> {
        if self.minimum > self.maximum {
            return Err(invalid("page-break minimum exceeds maximum"));
        }
        let (maximum_id, maximum_span) = match axis {
            Axis::Horizontal => (LAST_ROW, LAST_COLUMN),
            Axis::Vertical => (LAST_COLUMN, LAST_ROW),
        };
        if self.id > maximum_id || self.maximum > maximum_span {
            return Err(invalid("page break exceeds the worksheet grid"));
        }
        Ok(())
    }
}

/// One axis-specific, ordered page-break collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Collection {
    axis: Axis,
    breaks: Box<[Break]>,
}

impl Collection {
    /// Build horizontal row breaks in source order.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-grid break or too many breaks.
    pub fn horizontal(values: impl IntoIterator<Item = Break>) -> Result<Self> {
        Self::collect(Axis::Horizontal, values)
    }

    /// Build vertical column breaks in source order.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-grid break or too many breaks.
    pub fn vertical(values: impl IntoIterator<Item = Break>) -> Result<Self> {
        Self::collect(Axis::Vertical, values)
    }

    fn collect(axis: Axis, values: impl IntoIterator<Item = Break>) -> Result<Self> {
        let mut breaks = Vec::new();
        for value in values {
            let maximum = match axis {
                Axis::Horizontal => MAX_HORIZONTAL_BREAKS,
                Axis::Vertical => MAX_VERTICAL_BREAKS,
            };
            if breaks.len() >= maximum {
                return Err(invalid("page-break collection exceeds the Office limit"));
            }
            value.validate(axis)?;
            breaks.push(value);
        }
        Ok(Self {
            axis,
            breaks: breaks.into_boxed_slice(),
        })
    }

    /// Axis owned by this collection.
    #[must_use]
    pub const fn axis(&self) -> Axis {
        self.axis
    }

    /// Ordered page breaks.
    #[must_use]
    pub fn breaks(&self) -> &[Break] {
        &self.breaks
    }

    /// Number of page breaks.
    #[must_use]
    pub fn len(&self) -> usize {
        self.breaks.len()
    }

    /// Whether no page breaks are present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.breaks.is_empty()
    }

    /// Number of authored manual breaks.
    #[must_use]
    pub fn manual_break_count(&self) -> usize {
        self.breaks.iter().filter(|value| value.is_manual()).count()
    }
}

/// Direct worksheet horizontal and vertical page-break state.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageBreaks {
    horizontal: Option<Collection>,
    vertical: Option<Collection>,
}

impl PageBreaks {
    /// Create an absent page-break state.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            horizontal: None,
            vertical: None,
        }
    }

    /// Authored `rowBreaks`, including an explicitly empty collection.
    #[must_use]
    pub const fn horizontal(&self) -> Option<&Collection> {
        self.horizontal.as_ref()
    }

    /// Authored `colBreaks`, including an explicitly empty collection.
    #[must_use]
    pub const fn vertical(&self) -> Option<&Collection> {
        self.vertical.as_ref()
    }

    /// Replace the horizontal collection.
    ///
    /// # Errors
    ///
    /// Returns an error when the collection is not horizontal.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<()> {
        if collection.axis != Axis::Horizontal {
            return Err(invalid(
                "horizontal page breaks require a horizontal collection",
            ));
        }
        self.horizontal = Some(collection);
        Ok(())
    }

    /// Replace the vertical collection.
    ///
    /// # Errors
    ///
    /// Returns an error when the collection is not vertical.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<()> {
        if collection.axis != Axis::Vertical {
            return Err(invalid(
                "vertical page breaks require a vertical collection",
            ));
        }
        self.vertical = Some(collection);
        Ok(())
    }

    /// Remove the authored horizontal collection.
    pub fn remove_horizontal(&mut self) -> bool {
        self.horizontal.take().is_some()
    }

    /// Remove the authored vertical collection.
    pub fn remove_vertical(&mut self) -> bool {
        self.vertical.take().is_some()
    }

    /// Remove both collections.
    pub fn clear(&mut self) -> bool {
        self.remove_horizontal() | self.remove_vertical()
    }

    pub(super) fn validate(&self) -> Result<()> {
        if let Some(horizontal) = &self.horizontal {
            if horizontal.axis != Axis::Horizontal {
                return Err(invalid("invalid horizontal page-break axis"));
            }
            for value in horizontal.breaks() {
                value.validate(Axis::Horizontal)?;
            }
        }
        if let Some(vertical) = &self.vertical {
            if vertical.axis != Axis::Vertical {
                return Err(invalid("invalid vertical page-break axis"));
            }
            for value in vertical.breaks() {
                value.validate(Axis::Vertical)?;
            }
        }
        Ok(())
    }
}
