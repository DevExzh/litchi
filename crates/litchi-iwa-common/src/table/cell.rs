//! Cell-level table vocabulary independent of archive and application models.

use crate::shape::stroke::Stroke;

pub mod conditional_highlight;
pub mod layout;
pub mod number_format;
pub mod value;

/// One edge of a native table cell.
#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum BorderSide {
    /// The cell's left edge.
    Left,
    /// The cell's right edge.
    Right,
    /// The cell's top edge.
    Top,
    /// The cell's bottom edge.
    Bottom,
}

impl BorderSide {
    /// The four cell edges in stable wire-independent order.
    pub const ALL: [Self; 4] = [Self::Left, Self::Right, Self::Top, Self::Bottom];

    /// Returns the zero-based compact index for this edge.
    #[must_use]
    pub const fn index(self) -> usize {
        match self {
            Self::Left => 0,
            Self::Right => 1,
            Self::Top => 2,
            Self::Bottom => 3,
        }
    }

    /// Returns the geometrically opposite edge.
    #[must_use]
    pub const fn opposite(self) -> Self {
        match self {
            Self::Left => Self::Right,
            Self::Right => Self::Left,
            Self::Top => Self::Bottom,
            Self::Bottom => Self::Top,
        }
    }
}

/// Explicit stroke overrides for the four edges of one table cell.
///
/// A missing edge means that the table style supplies the edge, or that a
/// later native stroke run explicitly clears it.  Native stroke sidecars and
/// package identifiers remain owned by the concrete iWork adapters.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Borders {
    /// Explicit left-edge stroke, if present.
    pub left: Option<Stroke>,
    /// Explicit right-edge stroke, if present.
    pub right: Option<Stroke>,
    /// Explicit top-edge stroke, if present.
    pub top: Option<Stroke>,
    /// Explicit bottom-edge stroke, if present.
    pub bottom: Option<Stroke>,
}

impl Borders {
    /// Return the explicit stroke for one cell edge.
    #[must_use]
    pub const fn get(self, side: BorderSide) -> Option<Stroke> {
        match side {
            BorderSide::Left => self.left,
            BorderSide::Right => self.right,
            BorderSide::Top => self.top,
            BorderSide::Bottom => self.bottom,
        }
    }

    /// Set the explicit stroke for one cell edge.
    ///
    /// This mutator is intentionally small so archive adapters can populate
    /// the common value without taking ownership of their native sidecars.
    pub fn set(&mut self, side: BorderSide, stroke: Option<Stroke>) {
        match side {
            BorderSide::Left => self.left = stroke,
            BorderSide::Right => self.right = stroke,
            BorderSide::Top => self.top = stroke,
            BorderSide::Bottom => self.bottom = stroke,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use crate::color::Rgba;
    use crate::shape::stroke::{Pattern, Width};

    use super::{BorderSide, Borders};

    #[test]
    fn edges_have_compact_stable_order() {
        assert_eq!(size_of::<BorderSide>(), 1);
        assert_eq!(BorderSide::ALL.map(BorderSide::index), [0, 1, 2, 3]);
        assert_eq!(BorderSide::Left.opposite(), BorderSide::Right);
        assert_eq!(BorderSide::Top.opposite(), BorderSide::Bottom);
    }

    #[test]
    fn borders_default_and_side_access_round_trip() {
        let mut borders = Borders::default();
        assert_eq!(
            BorderSide::ALL.map(|side| borders.get(side)),
            [None, None, None, None]
        );

        let stroke = crate::shape::stroke::Stroke::new(Rgba::black(), Width::ONE, Pattern::Solid);
        for side in BorderSide::ALL {
            borders.set(side, Some(stroke));
            assert_eq!(borders.get(side), Some(stroke));
        }

        borders.set(BorderSide::Right, None);
        assert_eq!(borders.get(BorderSide::Right), None);
        assert_eq!(borders.left, Some(stroke));
        assert_eq!(borders.top, Some(stroke));
        assert_eq!(borders.bottom, Some(stroke));
    }
}
