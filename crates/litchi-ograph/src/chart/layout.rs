//! Checked chart view, growth, and position values.

/// Chart-window zoom as a positive fraction in the specification range
/// `1/10..=4`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Zoom {
    numerator: u16,
    denominator: u16,
}

impl Zoom {
    /// One hundred percent zoom.
    pub const ONE: Self = Self {
        numerator: 1,
        denominator: 1,
    };

    /// Creates a checked zoom fraction.
    #[must_use]
    pub const fn new(numerator: u16, denominator: u16) -> Option<Self> {
        if numerator == 0
            || denominator == 0
            || numerator > i16::MAX as u16
            || denominator > i16::MAX as u16
        {
            return None;
        }
        let wide_numerator = numerator as u32;
        let wide_denominator = denominator as u32;
        if wide_numerator * 10 < wide_denominator || wide_numerator > wide_denominator * 4 {
            return None;
        }
        Some(Self {
            numerator,
            denominator,
        })
    }

    /// Fraction numerator.
    #[must_use]
    pub const fn numerator(self) -> u16 {
        self.numerator
    }

    /// Fraction denominator.
    #[must_use]
    pub const fn denominator(self) -> u16 {
        self.denominator
    }
}

impl Default for Zoom {
    fn default() -> Self {
        Self::ONE
    }
}

/// Signed 16.16 fixed-point chart coordinate.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Fixed(i32);

impl Fixed {
    /// Fixed-point value `1.0`.
    pub const ONE: Self = Self(1 << 16);

    /// Preserves one signed 16.16 wire value.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the signed 16.16 wire value.
    #[must_use]
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Horizontal and vertical plot-area growth used for font scaling.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Growth {
    /// Horizontal growth.
    pub x: Fixed,
    /// Vertical growth.
    pub y: Fixed,
}

impl Growth {
    /// No growth in either direction.
    pub const ONE: Self = Self {
        x: Fixed::ONE,
        y: Fixed::ONE,
    };
}

impl Default for Growth {
    fn default() -> Self {
        Self::ONE
    }
}

/// Interpretation of coordinates in a `Pos` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Mode {
    /// Relative to the chart, in points.
    ChartPoints = 0,
    /// Absolute width and height, in points; valid only for the lower-right mode.
    Absolute = 1,
    /// Interpretation is selected by the owning chart collection.
    Parent = 2,
    /// Offset from the default position in thousandths of plot-area size.
    DefaultOffset = 3,
    /// Relative to the chart in SPRC units.
    Chart = 5,
}

impl Mode {
    pub(super) const fn from_raw(value: u16) -> Option<Self> {
        match value {
            0 => Some(Self::ChartPoints),
            1 => Some(Self::Absolute),
            2 => Some(Self::Parent),
            3 => Some(Self::DefaultOffset),
            5 => Some(Self::Chart),
            _ => None,
        }
    }
}

/// Position owned by an axis-parent, legend, or attached-label collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Pos {
    top_left: Mode,
    bottom_right: Mode,
    x: i16,
    y: i16,
    width: i16,
    height: i16,
}

impl Pos {
    /// Creates the mandatory primary-axis-group plot position.
    #[must_use]
    pub const fn plot(x: i16, y: i16, width: i16, height: i16) -> Self {
        Self {
            top_left: Mode::Parent,
            bottom_right: Mode::Parent,
            x,
            y,
            width,
            height,
        }
    }

    pub(super) const fn parsed(
        top_left: Mode,
        bottom_right: Mode,
        x: i16,
        y: i16,
        width: i16,
        height: i16,
    ) -> Self {
        Self {
            top_left,
            bottom_right,
            x,
            y,
            width,
            height,
        }
    }

    /// Upper-left positioning mode.
    #[must_use]
    pub const fn top_left(self) -> Mode {
        self.top_left
    }

    /// Lower-right positioning mode.
    #[must_use]
    pub const fn bottom_right(self) -> Mode {
        self.bottom_right
    }

    /// Horizontal position.
    #[must_use]
    pub const fn x(self) -> i16 {
        self.x
    }

    /// Vertical position.
    #[must_use]
    pub const fn y(self) -> i16 {
        self.y
    }

    /// Width or second horizontal coordinate, according to the modes.
    #[must_use]
    pub const fn width(self) -> i16 {
        self.width
    }

    /// Height or second vertical coordinate, according to the modes.
    #[must_use]
    pub const fn height(self) -> i16 {
        self.height
    }

    pub(super) const fn is_plot(self) -> bool {
        matches!(self.top_left, Mode::Parent) && matches!(self.bottom_right, Mode::Parent)
    }
}

impl Default for Pos {
    fn default() -> Self {
        Self::plot(0, 0, 4_000, 4_000)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic by design"
    )]
    use super::*;

    #[test]
    fn zoom_bounds_are_checked_without_floating_point() {
        assert_eq!(Zoom::new(1, 10), Some(Zoom::new(1, 10).unwrap()));
        assert!(Zoom::new(1, 11).is_none());
        assert!(Zoom::new(4, 1).is_some());
        assert!(Zoom::new(41, 10).is_none());
        assert!(Zoom::new(0, 1).is_none());
        assert!(Zoom::new(1, 0).is_none());
    }

    #[test]
    fn plot_position_owns_the_required_modes() {
        let pos = Pos::plot(1, 2, 3, 4);
        assert!(pos.is_plot());
        assert_eq!((pos.x(), pos.y(), pos.width(), pos.height()), (1, 2, 3, 4));
    }
}
