//! Archive-free values for three-dimensional charts.

const RECTANGLE_NATIVE_VALUE: i32 = 0;
const CYLINDER_NATIVE_VALUE: i32 = 1;

/// The bar or column geometry used by a native three-dimensional chart.
///
/// The native integer is retained directly so the value remains four bytes
/// and unrecognized future values survive a read-modify-write cycle without
/// allocation or lossy fallback.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct BarShape(i32);

#[allow(
    non_upper_case_globals,
    reason = "PascalCase associated constants are the focused BarShape API"
)]
impl BarShape {
    /// Rectangular prisms, the native default.
    pub const Rectangle: Self = Self(RECTANGLE_NATIVE_VALUE);
    /// Circular cylinders.
    pub const Cylinder: Self = Self(CYLINDER_NATIVE_VALUE);

    /// Decode the integer stored by a native iWork archive.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the integer used by a native iWork archive.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Whether this value is not one of the known native bar shapes.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.0 != RECTANGLE_NATIVE_VALUE && self.0 != CYLINDER_NATIVE_VALUE
    }
}

impl std::fmt::Debug for BarShape {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match *self {
            Self::Rectangle => formatter.write_str("Rectangle"),
            Self::Cylinder => formatter.write_str("Cylinder"),
            Self(value) => formatter.debug_tuple("Unsupported").field(&value).finish(),
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::BarShape;

    #[test]
    fn bar_shape_is_a_compact_lossless_value() {
        assert_eq!(size_of::<BarShape>(), size_of::<i32>());
        assert_eq!(align_of::<BarShape>(), align_of::<i32>());
        assert_eq!(BarShape::default(), BarShape::Rectangle);
        assert!(!BarShape::Rectangle.is_unsupported());
        assert!(!BarShape::Cylinder.is_unsupported());
    }

    #[test]
    fn known_values_have_stable_native_identifiers() {
        assert_eq!(BarShape::Rectangle.native_value(), 0);
        assert_eq!(BarShape::Cylinder.native_value(), 1);
        assert_eq!(BarShape::from_native(0), BarShape::Rectangle);
        assert_eq!(BarShape::from_native(1), BarShape::Cylinder);
    }

    #[test]
    fn unknown_values_round_trip_without_fallback() {
        for value in [i32::MIN, -1, 2, 9_001, i32::MAX] {
            let shape = BarShape::from_native(value);
            assert!(shape.is_unsupported());
            assert_eq!(shape.native_value(), value);
            assert_eq!(BarShape::from_native(shape.native_value()), shape);
        }
    }
}
