//! Archive-free paragraph-border controls.

const ALL_BITS: u8 = 0b1111;
const DEFAULT_OFFSET_POINTS: f32 = 6.0;

/// Validation failures for paragraph-border semantic values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A border offset is not finite.
    NonFinite,
    /// A border offset is negative.
    Negative,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NonFinite => formatter.write_str("paragraph-border offset must be finite"),
            Self::Negative => formatter.write_str("paragraph-border offset must be nonnegative"),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for paragraph-border semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// Selected edges of a paragraph layout box.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Sides(u8);

impl Sides {
    /// No paragraph edges.
    pub const NONE: Self = Self(0);
    /// The top edge.
    pub const TOP: Self = Self(1 << 0);
    /// The bottom edge.
    pub const BOTTOM: Self = Self(1 << 1);
    /// The left edge.
    pub const LEFT: Self = Self(1 << 2);
    /// The right edge.
    pub const RIGHT: Self = Self(1 << 3);
    /// All four paragraph edges.
    pub const ALL: Self = Self(ALL_BITS);

    /// Construct any combination of paragraph-border edges.
    #[must_use]
    pub const fn new(top: bool, bottom: bool, left: bool, right: bool) -> Self {
        Self(
            (if top { Self::TOP.0 } else { 0 })
                | (if bottom { Self::BOTTOM.0 } else { 0 })
                | (if left { Self::LEFT.0 } else { 0 })
                | (if right { Self::RIGHT.0 } else { 0 }),
        )
    }

    /// Return whether all edges in `side` are selected.
    #[must_use]
    pub const fn contains(self, side: Self) -> bool {
        self.0 & side.0 == side.0
    }

    /// Return whether no edges are selected.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.0 == 0
    }

    /// Return whether all four edges are selected.
    #[must_use]
    pub const fn is_all(self) -> bool {
        self.0 == ALL_BITS
    }
}

impl std::ops::BitOr for Sides {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

/// Inspector-visible gap between paragraph text and its border, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Offset(f32);

impl Offset {
    /// The six-point offset used when an iWork paragraph border is enabled.
    pub const DEFAULT: Self = Self(DEFAULT_OFFSET_POINTS);

    /// Construct a finite, nonnegative paragraph-border offset.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinity and
    /// [`Error::Negative`] for a negative offset.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::NonFinite);
        }
        if points < 0.0 {
            return Err(Error::Negative);
        }
        Ok(Self(if points == 0.0 { 0.0 } else { points }))
    }

    /// Return the offset in typographic points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl Default for Offset {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sides_are_compact_and_composable() {
        let top_left = Sides::TOP | Sides::LEFT;
        assert!(top_left.contains(Sides::TOP));
        assert!(top_left.contains(Sides::LEFT));
        assert!(!top_left.contains(Sides::RIGHT));
        assert!(Sides::ALL.is_all());
        assert!(Sides::NONE.is_empty());
    }

    #[test]
    fn offsets_are_finite_and_nonnegative() {
        assert_eq!(Offset::from_points(9.0).unwrap().points(), 9.0);
        assert_eq!(Offset::from_points(-0.1), Err(Error::Negative));
        assert_eq!(Offset::from_points(f32::NAN), Err(Error::NonFinite));
        assert_eq!(Offset::from_points(-0.0).unwrap().points(), 0.0);
    }
}
