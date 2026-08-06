//! Stable identifiers, ranges, and BIFF total aggregation codes.

use super::super::{FEATURE11_RECORD_TYPE, invalid};
use crate::Result;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ListObjectId(u32);
impl ListObjectId {
    pub fn try_new(value: u32) -> Result<Self> {
        if value == 0 {
            Err(invalid(FEATURE11_RECORD_TYPE, "table id must be nonzero"))
        } else {
            Ok(Self(value))
        }
    }
    pub const fn value(self) -> u32 {
        self.0
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ListColumnId(u32);
impl ListColumnId {
    pub fn try_new(value: u32) -> Result<Self> {
        if value == 0 {
            Err(invalid(FEATURE11_RECORD_TYPE, "column id must be nonzero"))
        } else {
            Ok(Self(value))
        }
    }
    pub const fn value(self) -> u32 {
        self.0
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ListObjectRange {
    pub(in crate::list_object) first_row: u16,
    pub(in crate::list_object) last_row: u16,
    pub(in crate::list_object) first_column: u16,
    pub(in crate::list_object) last_column: u16,
}
impl ListObjectRange {
    pub fn try_new(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> Result<Self> {
        if first_row > last_row || first_column > last_column || last_column > 255 {
            Err(invalid(
                FEATURE11_RECORD_TYPE,
                "table range is reversed or outside BIFF8 columns",
            ))
        } else {
            Ok(Self {
                first_row,
                last_row,
                first_column,
                last_column,
            })
        }
    }
    pub const fn first_row(self) -> u16 {
        self.first_row
    }
    pub const fn last_row(self) -> u16 {
        self.last_row
    }
    pub const fn first_column(self) -> u16 {
        self.first_column
    }
    pub const fn last_column(self) -> u16 {
        self.last_column
    }
    pub const fn column_count(self) -> usize {
        (self.last_column - self.first_column + 1) as usize
    }
    pub const fn overlaps(self, other: Self) -> bool {
        self.first_row <= other.last_row
            && other.first_row <= self.last_row
            && self.first_column <= other.last_column
            && other.first_column <= self.last_column
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListTotalAggregation {
    None,
    Average,
    Count,
    CountNumbers,
    Max,
    Min,
    Sum,
    StandardDeviation,
    Variance,
    Custom,
}
impl ListTotalAggregation {
    pub(in crate::list_object) fn code(self) -> u32 {
        self as u32
    }
    pub(in crate::list_object) fn from_code(v: u32) -> Result<Self> {
        Ok(match v {
            0 => Self::None,
            1 => Self::Average,
            2 => Self::Count,
            3 => Self::CountNumbers,
            4 => Self::Max,
            5 => Self::Min,
            6 => Self::Sum,
            7 => Self::StandardDeviation,
            8 => Self::Variance,
            9 => Self::Custom,
            _ => return Err(invalid(FEATURE11_RECORD_TYPE, "invalid total aggregation")),
        })
    }
}
