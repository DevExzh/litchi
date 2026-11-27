//! Archive-free chart series orientation shared by iWork format owners.

const ROWS_NATIVE_VALUE: i32 = 1;
const COLUMNS_NATIVE_VALUE: i32 = 2;

/// Whether chart series are stored in rows or columns.
///
/// The native integer is retained directly so this value stays four bytes
/// while still preserving an unrecognized future value losslessly.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Direction(i32);

/// The recognized chart-series orientations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Each source row is one series.
    Rows,
    /// Each source column is one series.
    Columns,
}

impl Direction {
    /// Each source row is one series.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Rows: Self = Self(ROWS_NATIVE_VALUE);
    /// Each source column is one series.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Columns: Self = Self(COLUMNS_NATIVE_VALUE);

    /// Decode the integer stored by a native iWork archive.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the integer stored by a native iWork archive.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized semantic orientation, if this is a known value.
    #[must_use]
    pub const fn kind(self) -> Option<Kind> {
        match self.0 {
            ROWS_NATIVE_VALUE => Some(Kind::Rows),
            COLUMNS_NATIVE_VALUE => Some(Kind::Columns),
            _ => None,
        }
    }

    /// Whether this value is not one of the known row/column orientations.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.0 != ROWS_NATIVE_VALUE && self.0 != COLUMNS_NATIVE_VALUE
    }
}

impl std::fmt::Debug for Direction {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.0 {
            ROWS_NATIVE_VALUE => formatter.write_str("Rows"),
            COLUMNS_NATIVE_VALUE => formatter.write_str("Columns"),
            value => formatter.debug_tuple("Unsupported").field(&value).finish(),
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Direction, Kind};

    #[test]
    fn directions_preserve_known_and_future_values() {
        assert_eq!(size_of::<Direction>(), 4);
        for (value, direction) in [
            (1, Direction::Rows),
            (2, Direction::Columns),
            (9_001, Direction::from_native(9_001)),
        ] {
            assert_eq!(Direction::from_native(value), direction);
            assert_eq!(direction.native_value(), value);
        }
        assert_eq!(Direction::Rows.kind(), Some(Kind::Rows));
        assert_eq!(Direction::Columns.kind(), Some(Kind::Columns));
        assert!(Direction::from_native(9_001).is_unsupported());
        assert_eq!(Direction::from_native(9_001).kind(), None);
    }
}
