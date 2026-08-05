//! Allocation-free frame layout values shared by iWork text containers.

/// Validation failures for a text-frame inset.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The supplied distance was NaN or infinite.
    NonFinite,
    /// The supplied distance was less than zero.
    Negative,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::NonFinite => "text-frame inset must be finite",
            Self::Negative => "text-frame inset must be non-negative",
        })
    }
}

impl std::error::Error for Error {}

/// Result type for frame-layout construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Vertical placement of text inside a shape or ordinary text box.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum VerticalAlignment {
    /// Place text at the top edge of the content box.
    #[default]
    Top,
    /// Center text vertically in the content box.
    Middle,
    /// Place text at the bottom edge of the content box.
    Bottom,
    /// Distribute text through the content box.
    Justified,
}

/// Whether iWork should shrink text until it fits inside the shape.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum AutoSize {
    /// Keep the authored text size.
    #[default]
    Fixed,
    /// Shrink text until it fits the content box.
    ShrinkToFit,
}

/// Finite, non-negative distance from one text-frame edge, measured in points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Inset(f32);

impl Inset {
    /// A zero-point inset.
    pub const ZERO: Self = Self(0.0);

    /// Construct an inset after validating its finite, non-negative domain.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinite input and
    /// [`Error::Negative`] for a finite value below zero.
    #[must_use = "use the validated inset or handle its validation error"]
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::NonFinite);
        }
        if points < 0.0 {
            return Err(Error::Negative);
        }
        // Preserve the native scalar exactly; IEEE negative zero is still a
        // valid zero distance and must not be normalized during a read/write.
        Ok(Self(points))
    }

    /// Return the distance in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Independently typed text insets for all four frame edges.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Insets {
    left: Inset,
    top: Inset,
    right: Inset,
    bottom: Inset,
}

impl Insets {
    /// Zero inset on every edge.
    pub const ZERO: Self = Self::uniform(Inset::ZERO);

    /// Construct insets in left, top, right, bottom order.
    #[must_use]
    pub const fn new(left: Inset, top: Inset, right: Inset, bottom: Inset) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Construct equal insets on all four edges.
    #[must_use]
    pub const fn uniform(inset: Inset) -> Self {
        Self::new(inset, inset, inset, inset)
    }

    /// Return the left inset.
    #[must_use]
    pub const fn left(self) -> Inset {
        self.left
    }

    /// Return the top inset.
    #[must_use]
    pub const fn top(self) -> Inset {
        self.top
    }

    /// Return the right inset.
    #[must_use]
    pub const fn right(self) -> Inset {
        self.right
    }

    /// Return the bottom inset.
    #[must_use]
    pub const fn bottom(self) -> Inset {
        self.bottom
    }
}

impl Default for Insets {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Composable frame-level text layout stored in a native shape style.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Layout {
    vertical_alignment: VerticalAlignment,
    auto_size: AutoSize,
    insets: Insets,
}

impl Layout {
    /// Construct a complete frame layout.
    #[must_use]
    pub const fn new(
        vertical_alignment: VerticalAlignment,
        insets: Insets,
        auto_size: AutoSize,
    ) -> Self {
        Self {
            vertical_alignment,
            auto_size,
            insets,
        }
    }

    /// Return the vertical-alignment mode.
    #[must_use]
    pub const fn vertical_alignment(self) -> VerticalAlignment {
        self.vertical_alignment
    }

    /// Return the four edge insets.
    #[must_use]
    pub const fn insets(self) -> Insets {
        self.insets
    }

    /// Return the autosizing mode.
    #[must_use]
    pub const fn auto_size(self) -> AutoSize {
        self.auto_size
    }

    /// Return a layout with a different vertical-alignment mode.
    #[must_use]
    pub const fn with_vertical_alignment(mut self, vertical_alignment: VerticalAlignment) -> Self {
        self.vertical_alignment = vertical_alignment;
        self
    }

    /// Return a layout with different edge insets.
    #[must_use]
    pub const fn with_insets(mut self, insets: Insets) -> Self {
        self.insets = insets;
        self
    }

    /// Return a layout with a different autosizing mode.
    #[must_use]
    pub const fn with_auto_size(mut self, auto_size: AutoSize) -> Self {
        self.auto_size = auto_size;
        self
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{AutoSize, Error, Inset, Insets, Layout, VerticalAlignment};

    #[test]
    fn layout_values_are_compact_and_composable() {
        assert_eq!(size_of::<Inset>(), 4);
        assert_eq!(size_of::<Insets>(), 16);
        assert_eq!(size_of::<Layout>(), 20);
        assert_eq!(align_of::<Inset>(), 4);
        assert_eq!(align_of::<Insets>(), 4);
        assert_eq!(align_of::<Layout>(), 4);

        let inset = Inset::from_points(12.0).unwrap_or_else(|_| panic!("valid inset"));
        let layout = Layout::default()
            .with_vertical_alignment(VerticalAlignment::Middle)
            .with_insets(Insets::uniform(inset))
            .with_auto_size(AutoSize::ShrinkToFit);
        assert_eq!(layout.insets(), Insets::uniform(inset));
        assert_eq!(layout.vertical_alignment(), VerticalAlignment::Middle);
        assert_eq!(layout.auto_size(), AutoSize::ShrinkToFit);
    }

    #[test]
    fn inset_values_reject_non_finite_and_negative_input() {
        assert_eq!(Inset::from_points(-0.01), Err(Error::Negative));
        assert_eq!(Inset::from_points(f32::NAN), Err(Error::NonFinite));
        assert_eq!(Inset::from_points(f32::INFINITY), Err(Error::NonFinite));
        assert_eq!(Inset::from_points(f32::NEG_INFINITY), Err(Error::NonFinite));
        assert_eq!(
            Inset::from_points(-0.0)
                .unwrap_or_else(|_| panic!("negative zero is a valid zero distance"))
                .points()
                .to_bits(),
            (-0.0f32).to_bits()
        );
    }
}
