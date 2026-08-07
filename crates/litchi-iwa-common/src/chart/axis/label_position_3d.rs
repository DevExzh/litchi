//! Archive-free 3D value-axis label positions.

const AUTOMATIC_NATIVE_VALUE: i32 = 1;
const LEADING_NATIVE_VALUE: i32 = 2;
const TRAILING_NATIVE_VALUE: i32 = 3;

/// Position of primary value-axis labels in a native 3D chart.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LabelPosition3d {
    /// Let iWork choose the side from chart orientation.
    #[default]
    Automatic,
    /// Show labels at Left, or Top for a horizontal bar chart.
    Leading,
    /// Show labels at Right, or Bottom for a horizontal bar chart.
    Trailing,
    /// Preserve an unrecognized future native value.
    Unsupported(i32),
}

impl LabelPosition3d {
    /// Decode the integer stored by native iWork archives.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        match value {
            AUTOMATIC_NATIVE_VALUE => Self::Automatic,
            LEADING_NATIVE_VALUE => Self::Leading,
            TRAILING_NATIVE_VALUE => Self::Trailing,
            other => Self::Unsupported(other),
        }
    }

    /// Return the integer used by native iWork archives.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Automatic => AUTOMATIC_NATIVE_VALUE,
            Self::Leading => LEADING_NATIVE_VALUE,
            Self::Trailing => TRAILING_NATIVE_VALUE,
            Self::Unsupported(value) => value,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::LabelPosition3d;

    #[test]
    fn positions_preserve_native_values() {
        for (value, position) in [
            (1, LabelPosition3d::Automatic),
            (2, LabelPosition3d::Leading),
            (3, LabelPosition3d::Trailing),
            (9_001, LabelPosition3d::Unsupported(9_001)),
        ] {
            assert_eq!(LabelPosition3d::from_native(value), position);
            assert_eq!(position.native_value(), value);
        }
    }
}
