//! Archive-free chart-series value-label vocabulary.

/// Visibility of one chart series' data value labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Visibility {
    /// Data value labels are hidden.
    Hidden,
    /// Data value labels are visible.
    Visible,
}

impl Visibility {
    /// Return whether data value labels are visible.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self, Self::Visible)
    }
}

impl From<bool> for Visibility {
    fn from(visible: bool) -> Self {
        if visible { Self::Visible } else { Self::Hidden }
    }
}

impl From<Visibility> for bool {
    fn from(visibility: Visibility) -> Self {
        visibility.is_visible()
    }
}

/// Zero-based index of one series in native chart-series order.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Index(usize);

impl Index {
    /// Construct a zero-based series index.
    #[must_use]
    pub const fn from_zero_based(index: usize) -> Self {
        Self(index)
    }

    /// Return the zero-based series index.
    #[must_use]
    pub const fn zero_based(self) -> usize {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Index, Visibility};

    #[test]
    fn visibility_round_trips_to_boolean() {
        assert_eq!(size_of::<Visibility>(), 1);
        assert!(!Visibility::Hidden.is_visible());
        assert!(Visibility::Visible.is_visible());
        assert_eq!(Visibility::from(false), Visibility::Hidden);
        assert_eq!(Visibility::from(true), Visibility::Visible);
        assert!(!bool::from(Visibility::Hidden));
        assert!(bool::from(Visibility::Visible));
    }

    #[test]
    fn index_is_compact_and_zero_based() {
        assert_eq!(size_of::<Index>(), size_of::<usize>());
        let index = Index::from_zero_based(7);
        assert_eq!(index.zero_based(), 7);
        assert!(Index::from_zero_based(1) < index);
    }
}
