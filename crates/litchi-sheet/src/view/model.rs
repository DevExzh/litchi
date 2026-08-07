//! Semantic worksheet-view model.

use thiserror::Error as ThisError;

use crate::{Cell, Column, Rect, Row};

/// The presentation mode of a worksheet.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Mode {
    /// Display cells normally.
    #[default]
    Normal,
    /// Show page boundaries while preserving the cell grid.
    PageBreakPreview,
    /// Present the worksheet in page-layout form.
    PageLayout,
}

/// One of the four worksheet pane positions.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Position {
    /// The lower-right pane.
    BottomRight,
    /// The upper-right pane.
    TopRight,
    /// The lower-left pane.
    BottomLeft,
    /// The upper-left pane.
    #[default]
    TopLeft,
}

/// The behavior of worksheet pane boundaries.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum State {
    /// Pane boundaries can be moved freely.
    #[default]
    Split,
    /// Pane boundaries remain fixed.
    Frozen,
    /// A fixed pane boundary also has a movable split.
    FrozenSplit,
}

/// Checked workbook-view collection index.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[must_use]
pub struct Window(u32);

impl Window {
    /// Create a workbook-view collection index.
    #[inline]
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the zero-based workbook-view collection index.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl From<u32> for Window {
    fn from(value: u32) -> Self {
        Self::new(value)
    }
}

impl From<Window> for u32 {
    fn from(value: Window) -> Self {
        value.get()
    }
}

/// Checked worksheet color index in the standard palette range.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[must_use]
pub struct Color(u8);

impl Color {
    /// The default worksheet color index.
    pub const DEFAULT: Self = Self(64);

    /// Validate a worksheet color index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Color`] when `value` is outside `0..=64`.
    #[inline]
    pub const fn new(value: u8) -> Result<Self, Error> {
        if value <= 64 {
            Ok(Self(value))
        } else {
            Err(Error::Color { value })
        }
    }

    /// Return the validated color index.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for Color {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u8> for Color {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Color> for u8 {
    fn from(value: Color) -> Self {
        value.get()
    }
}

/// Checked worksheet scale expressed as a percentage.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[must_use]
pub struct Scale(u16);

impl Scale {
    /// The default scale percentage.
    pub const DEFAULT: Self = Self(100);

    /// Validate a scale percentage.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Scale`] when `value` is outside `10..=400`.
    #[inline]
    pub const fn new(value: u16) -> Result<Self, Error> {
        if value >= 10 && value <= 400 {
            Ok(Self(value))
        } else {
            Err(Error::Scale { value })
        }
    }

    /// Return the validated scale percentage.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl Default for Scale {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u16> for Scale {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Scale> for u16 {
    fn from(value: Scale) -> Self {
        value.get()
    }
}

/// Checked finite, nonnegative pane-boundary position.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[must_use]
pub struct Split(f64);

impl Split {
    /// No displacement from the pane origin.
    pub const ZERO: Self = Self(0.0);

    /// Validate a pane-boundary position.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Split`] when `value` is negative, infinite, or not a
    /// number.
    #[inline]
    pub fn new(value: f64) -> Result<Self, Error> {
        if value.is_finite() && value >= 0.0 {
            Ok(Self(value))
        } else {
            Err(Error::Split { value })
        }
    }

    /// Return the validated pane-boundary position.
    #[inline]
    #[must_use]
    pub const fn get(self) -> f64 {
        self.0
    }
}

impl Default for Split {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f64> for Split {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Split> for f64 {
    fn from(value: Split) -> Self {
        value.get()
    }
}

/// Visibility settings grouped for a worksheet view.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the flags form one cohesive worksheet visibility configuration"
)]
pub struct Display {
    /// Whether worksheet window protection is enabled.
    pub window_protection: bool,
    /// Whether formulas are displayed instead of calculated values.
    pub show_formulas: bool,
    /// Whether worksheet grid lines are visible.
    pub grid_lines: bool,
    /// Whether row and column headings are visible.
    pub row_column_headers: bool,
    /// Whether cells with zero values display those values.
    pub zero_values: bool,
    /// Whether the worksheet is displayed right-to-left.
    pub right_to_left: bool,
    /// Whether the ruler is visible in page-layout presentation.
    pub ruler: bool,
    /// Whether outline controls are visible.
    pub outline_symbols: bool,
    /// Whether the default grid color is used.
    pub default_grid_color: bool,
    /// Whether page-layout whitespace is visible.
    pub white_space: bool,
}

impl Default for Display {
    fn default() -> Self {
        Self {
            window_protection: false,
            show_formulas: false,
            grid_lines: true,
            row_column_headers: true,
            zero_values: true,
            right_to_left: false,
            ruler: true,
            outline_symbols: true,
            default_grid_color: true,
            white_space: true,
        }
    }
}

/// Current and remembered worksheet scale percentages.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct Zoom {
    /// The scale currently applied to the worksheet.
    pub current: Scale,
    /// Scale remembered for normal presentation.
    pub normal: Option<Scale>,
    /// Scale remembered for page-layout presentation.
    pub page_layout: Option<Scale>,
    /// Scale remembered for page-break preview presentation.
    pub page_break_preview: Option<Scale>,
}

impl Default for Zoom {
    fn default() -> Self {
        Self {
            current: Scale::DEFAULT,
            normal: None,
            page_layout: None,
            page_break_preview: None,
        }
    }
}

/// Pane boundary configuration for a worksheet view.
#[derive(Debug, Clone, Copy, PartialEq)]
#[must_use]
pub struct Pane {
    /// Position of the pane this configuration addresses.
    pub position: Position,
    /// How pane boundaries behave.
    pub state: State,
    /// Horizontal pane-boundary position, when present.
    pub horizontal: Option<Split>,
    /// Vertical pane-boundary position, when present.
    pub vertical: Option<Split>,
    /// First visible cell in this pane.
    pub top_left: Cell,
}

impl Default for Pane {
    fn default() -> Self {
        Self {
            position: Position::TopLeft,
            state: State::Split,
            horizontal: None,
            vertical: None,
            top_left: Cell::new(Row::FIRST, Column::FIRST),
        }
    }
}

/// Non-empty selected ranges and their active range.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Selection {
    position: Position,
    active_cell: Cell,
    active_range: usize,
    ranges: Vec<Rect>,
}

impl Selection {
    /// Create a selection with one active range.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptySelection`] when `ranges` is empty, or
    /// [`Error::ActiveRange`] when `active_range` does not index `ranges`.
    pub fn new(
        position: Position,
        active_cell: Cell,
        active_range: usize,
        ranges: Vec<Rect>,
    ) -> Result<Self, Error> {
        if ranges.is_empty() {
            return Err(Error::EmptySelection);
        }
        if active_range >= ranges.len() {
            return Err(Error::ActiveRange {
                active_range,
                range_count: ranges.len(),
            });
        }
        Ok(Self {
            position,
            active_cell,
            active_range,
            ranges,
        })
    }

    /// Return the position of the pane containing this selection.
    #[inline]
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the active cell address.
    #[inline]
    #[must_use]
    pub const fn active_cell(&self) -> Cell {
        self.active_cell
    }

    /// Return the index of the active range.
    #[inline]
    #[must_use]
    pub const fn active_range(&self) -> usize {
        self.active_range
    }

    /// Return the selected ranges.
    #[inline]
    #[must_use]
    pub fn ranges(&self) -> &[Rect] {
        &self.ranges
    }
}

impl Default for Selection {
    fn default() -> Self {
        let cell = Cell::new(Row::FIRST, Column::FIRST);
        Self {
            position: Position::TopLeft,
            active_cell: cell,
            active_range: 0,
            ranges: vec![Rect::single(cell)],
        }
    }
}

/// Complete format-neutral state for one worksheet view.
#[derive(Debug, Clone, PartialEq)]
#[must_use]
pub struct View {
    /// Associated workbook-view collection index.
    pub window: Window,
    /// Worksheet presentation mode.
    pub mode: Mode,
    /// Worksheet color index.
    pub color: Color,
    /// Visibility settings.
    pub display: Display,
    /// Current and remembered scale percentages.
    pub zoom: Zoom,
    /// First visible worksheet cell outside of a pane configuration.
    pub origin: Cell,
    /// Pane-boundary configuration.
    pub pane: Option<Pane>,
    /// Selections in the worksheet panes.
    pub selections: Vec<Selection>,
    /// Whether this worksheet's tab is selected.
    pub tab_selected: bool,
}

impl Default for View {
    fn default() -> Self {
        Self {
            window: Window::default(),
            mode: Mode::Normal,
            color: Color::DEFAULT,
            display: Display::default(),
            zoom: Zoom::default(),
            origin: Cell::new(Row::FIRST, Column::FIRST),
            pane: None,
            selections: vec![Selection::default()],
            tab_selected: false,
        }
    }
}

/// Invalid format-neutral worksheet-view state.
#[derive(Debug, Clone, PartialEq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The color index lies outside the standard palette range.
    #[error("color index {value} is outside 0..=64")]
    Color {
        /// The rejected color index.
        value: u8,
    },
    /// The scale percentage lies outside the supported range.
    #[error("scale {value} is outside 10..=400")]
    Scale {
        /// The rejected scale percentage.
        value: u16,
    },
    /// The pane-boundary position is not finite and nonnegative.
    #[error("split {value} must be finite and nonnegative")]
    Split {
        /// The rejected pane-boundary position.
        value: f64,
    },
    /// A selection needs at least one range.
    #[error("selection must contain at least one range")]
    EmptySelection,
    /// The active range does not index the selection ranges.
    #[error("active range {active_range} is outside {range_count} selection ranges")]
    ActiveRange {
        /// The rejected active-range index.
        active_range: usize,
        /// Number of selection ranges supplied.
        range_count: usize,
    },
}
