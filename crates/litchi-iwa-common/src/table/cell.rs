//! Cell-level table vocabulary independent of archive and application models.

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

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::BorderSide;

    #[test]
    fn edges_have_compact_stable_order() {
        assert_eq!(size_of::<BorderSide>(), 1);
        assert_eq!(BorderSide::ALL.map(BorderSide::index), [0, 1, 2, 3]);
        assert_eq!(BorderSide::Left.opposite(), BorderSide::Right);
        assert_eq!(BorderSide::Top.opposite(), BorderSide::Bottom);
    }
}
