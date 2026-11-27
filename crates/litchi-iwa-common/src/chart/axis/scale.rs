//! Archive-free numeric scales for native chart value axes.

/// Numeric scale used by a native chart value axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Scale {
    /// Plot values along a linear scale.
    #[default]
    Linear,
    /// Plot values along a logarithmic scale.
    Logarithmic,
    /// Preserve an unrecognized future native value.
    Unsupported(i32),
}

impl Scale {
    /// Decode the integer stored by native iWork archives.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        match value {
            1 => Self::Linear,
            2 => Self::Logarithmic,
            other => Self::Unsupported(other),
        }
    }

    /// Return the integer used by native iWork archives.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Linear => 1,
            Self::Logarithmic => 2,
            Self::Unsupported(value) => value,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Scale;

    #[test]
    fn scales_preserve_known_and_future_values() {
        for (value, scale) in [
            (1, Scale::Linear),
            (2, Scale::Logarithmic),
            (9_001, Scale::Unsupported(9_001)),
        ] {
            assert_eq!(Scale::from_native(value), scale);
            assert_eq!(scale.native_value(), value);
        }
        assert_eq!(Scale::default(), Scale::Linear);
    }
}
