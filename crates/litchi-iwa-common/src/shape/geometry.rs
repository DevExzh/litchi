//! Dependency-free drawable geometry values.

/// A drawable position in document points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    /// Horizontal document coordinate.
    pub x: f32,
    /// Vertical document coordinate.
    pub y: f32,
}

/// A drawable size in document points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    /// Width in document points.
    pub width: f32,
    /// Height in document points.
    pub height: f32,
}

/// One native Arrange flip operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FlipAxis {
    /// Mirror the drawable around its vertical center line.
    Horizontal,
    /// Mirror the drawable around its horizontal center line.
    Vertical,
}

#[cfg(test)]
mod tests {
    use super::{FlipAxis, Point, Size};

    #[test]
    fn geometry_values_are_compact_and_copyable() {
        fn assert_copy<T: Copy>() {}

        assert_copy::<Point>();
        assert_copy::<Size>();
        assert_copy::<FlipAxis>();
    }

    #[test]
    fn geometry_values_preserve_document_units() {
        assert_eq!(Point { x: 12.0, y: 24.0 }, Point { x: 12.0, y: 24.0 });
        assert_eq!(
            Size {
                width: 96.0,
                height: 48.0,
            },
            Size {
                width: 96.0,
                height: 48.0,
            }
        );
    }
}
