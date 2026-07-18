//! Strongly typed chart series orientation shared by the iWork applications.

use crate::protobuf::tsch;

/// Whether chart series are stored in rows or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartSeriesDirection {
    Rows,
    Columns,
    Unsupported(i32),
}

impl ChartSeriesDirection {
    pub(crate) const fn from_raw(value: i32) -> Self {
        match value {
            x if x == tsch::SeriesDirection::ByRow as i32 => Self::Rows,
            x if x == tsch::SeriesDirection::ByColumn as i32 => Self::Columns,
            value => Self::Unsupported(value),
        }
    }

    pub(crate) const fn into_raw(self) -> i32 {
        match self {
            Self::Rows => tsch::SeriesDirection::ByRow as i32,
            Self::Columns => tsch::SeriesDirection::ByColumn as i32,
            Self::Unsupported(value) => value,
        }
    }
}
