//! Chart-group lines and up/down bars.

use super::format;
use crate::record::line;

/// Width of the gap between up or down bars (`0..=500`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Gap(u16);

impl Gap {
    pub const ZERO: Self = Self(0);

    pub const fn new(value: u16) -> Option<Self> {
        if value <= 500 {
            Some(Self(value))
        } else {
            None
        }
    }

    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Bar-series overlap percentage (`-100..=100`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Overlap(i16);

impl Overlap {
    pub const ZERO: Self = Self(0);

    pub const fn new(value: i16) -> Option<Self> {
        if value >= -100 && value <= 100 {
            Some(Self(value))
        } else {
            None
        }
    }

    pub const fn get(self) -> i16 {
        self.0
    }
}

/// Scatter/bubble size percentage (`0..=300`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct BubblePercent(u16);

impl BubblePercent {
    pub const ZERO: Self = Self(0);

    pub const fn new(value: u16) -> Option<Self> {
        if value <= 300 {
            Some(Self(value))
        } else {
            None
        }
    }

    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Bubble-size interpretation used when the bubble flag is enabled.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum BubbleKind {
    Area = 1,
    Width = 2,
}

/// One ordered chart-group line and its appearance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Line {
    pub kind: line::Kind,
    pub format: format::Line,
}

/// One complete up- or down-bar collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DropBar {
    pub gap: Gap,
    pub line: format::Line,
    pub area: format::Area,
}
