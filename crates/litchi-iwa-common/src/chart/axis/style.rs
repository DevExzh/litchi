//! Archive-free style values for native chart axes.

use super::TickMarkLocation;

/// Visibility of one chart-axis feature.
///
/// Native defaults are owned by the concrete adapter because they differ by
/// feature. This value therefore intentionally has no `Default` impl.
#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Visibility {
    /// The feature is hidden.
    Hidden,
    /// The feature is visible.
    Visible,
}

impl Visibility {
    /// Return whether the feature is visible.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self, Self::Visible)
    }

    /// Return whether the feature is hidden.
    #[must_use]
    pub const fn is_hidden(self) -> bool {
        matches!(self, Self::Hidden)
    }
}

impl From<bool> for Visibility {
    fn from(value: bool) -> Self {
        if value { Self::Visible } else { Self::Hidden }
    }
}

impl From<Visibility> for bool {
    fn from(value: Visibility) -> Self {
        value.is_visible()
    }
}

/// Visibility of the major and minor gridline families on one chart axis.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Gridlines {
    /// Visibility of major gridlines.
    pub major: Visibility,
    /// Visibility of minor gridlines.
    pub minor: Visibility,
}

impl Gridlines {
    /// Construct both gridline visibility values.
    #[must_use]
    pub const fn new(major: Visibility, minor: Visibility) -> Self {
        Self { major, minor }
    }
}

impl Default for Gridlines {
    fn default() -> Self {
        Self::new(Visibility::Hidden, Visibility::Hidden)
    }
}

/// Minor tick-mark visibility and major tick-mark placement for one axis.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct TickMarks {
    /// Visibility of minor tick marks.
    pub minor: Visibility,
    /// Placement of major tick marks.
    pub location: TickMarkLocation,
}

impl TickMarks {
    /// Construct minor tick-mark and major tick-mark settings.
    #[must_use]
    pub const fn new(minor: Visibility, location: TickMarkLocation) -> Self {
        Self { minor, location }
    }
}

impl Default for TickMarks {
    fn default() -> Self {
        Self::new(Visibility::Visible, TickMarkLocation::Centered)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Gridlines, TickMarks, Visibility};
    use crate::chart::axis::TickMarkLocation;

    #[test]
    fn visibility_is_compact_and_round_trips_booleans() {
        assert_eq!(size_of::<Visibility>(), 1);
        assert_eq!(Visibility::from(false), Visibility::Hidden);
        assert_eq!(Visibility::from(true), Visibility::Visible);
        assert!(!bool::from(Visibility::Hidden));
        assert!(bool::from(Visibility::Visible));
        assert!(Visibility::Visible.is_visible());
        assert!(Visibility::Hidden.is_hidden());
    }

    #[test]
    fn gridline_and_tick_mark_values_are_compact_and_explicit() {
        assert_eq!(size_of::<Gridlines>(), 2);
        assert_eq!(size_of::<TickMarks>(), 12);
        assert_eq!(Gridlines::default().major, Visibility::Hidden);
        assert_eq!(Gridlines::default().minor, Visibility::Hidden);
        assert_eq!(TickMarks::default().minor, Visibility::Visible);
        assert_eq!(TickMarks::default().location, TickMarkLocation::Centered);
        assert_eq!(
            TickMarks::new(Visibility::Hidden, TickMarkLocation::Outside),
            TickMarks {
                minor: Visibility::Hidden,
                location: TickMarkLocation::Outside,
            }
        );
    }
}
