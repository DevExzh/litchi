//! Semantic selectors and style values for chart axes.

/// One of the standard axes exposed by an iWork chart formatter.
#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Axis {
    /// The category or horizontal axis.
    Category,
    /// The primary value or vertical axis.
    Value,
}

impl Axis {
    /// The axes in stable formatter order.
    pub const ALL: [Self; 2] = [Self::Category, Self::Value];

    /// Return the stable semantic name used in diagnostics.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Category => "category",
            Self::Value => "value",
        }
    }
}

/// Where a chart draws its major tick marks.
///
/// The concrete IWA owner maps this value to the native integer stored in the
/// chart archive. The `Unsupported` variant keeps an unrecognized native value
/// explicit so a read-modify-write operation cannot silently change it.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum TickMarkLocation {
    /// Do not draw major tick marks.
    None,
    /// Draw tick marks toward the plot area.
    Inside,
    /// Draw tick marks centered on the axis line.
    #[default]
    Centered,
    /// Draw tick marks away from the plot area.
    Outside,
    /// Preserve an unrecognized native value.
    Unsupported(i32),
}

impl TickMarkLocation {
    /// Decode the integer used by native iWork chart archives.
    #[must_use]
    pub const fn from_native(native_value: i32) -> Self {
        match native_value {
            0 => Self::None,
            1 => Self::Inside,
            2 => Self::Centered,
            3 => Self::Outside,
            other => Self::Unsupported(other),
        }
    }

    /// Return the integer used by native iWork chart archives.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Inside => 1,
            Self::Centered => 2,
            Self::Outside => 3,
            Self::Unsupported(value) => value,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Axis, TickMarkLocation};

    #[test]
    fn vocabulary_is_compact_and_deterministic() {
        assert_eq!(size_of::<Axis>(), 1);
        assert_eq!(Axis::ALL, [Axis::Category, Axis::Value]);
        assert_eq!(Axis::Category.as_str(), "category");
        assert_eq!(Axis::Value.as_str(), "value");
        assert_eq!(TickMarkLocation::default(), TickMarkLocation::Centered);
        assert_eq!(TickMarkLocation::from_native(0), TickMarkLocation::None);
        assert_eq!(TickMarkLocation::from_native(3), TickMarkLocation::Outside);
        assert_eq!(
            TickMarkLocation::from_native(99),
            TickMarkLocation::Unsupported(99)
        );
        assert_eq!(TickMarkLocation::Unsupported(99).native_value(), 99);
        assert_eq!(size_of::<TickMarkLocation>(), 8);
    }
}
