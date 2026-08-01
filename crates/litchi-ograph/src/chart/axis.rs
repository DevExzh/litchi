//! Axis scales, ticks, and lines.

use super::format;

/// Axis role in the chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    /// Category axis, or the horizontal axis in scatter charts.
    Category,
    /// Value axis, or the vertical axis in scatter charts.
    Value,
    /// Series axis used by three-dimensional charts.
    Series,
}

/// Numeric axis range and unit configuration.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Scale {
    /// Minimum axis value.
    pub min: f64,
    /// Maximum axis value.
    pub max: f64,
    /// Major unit.
    pub major: f64,
    /// Minor unit.
    pub minor: f64,
    /// Axis crossing value.
    pub crossing: f64,
    /// Raw ValueRange flags, including preserved producer bits.
    pub flags: u16,
}

/// Tick marks and labels for an axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Tick {
    /// Major tick-mark kind.
    pub major: u8,
    /// Minor tick-mark kind.
    pub minor: u8,
    /// Tick-label position.
    pub label: u8,
    /// Background mode.
    pub background: u8,
    /// Stored color bytes.
    pub color: [u8; 4],
    /// Raw Tick flags.
    pub flags: u16,
}

/// Semantic role of one axis line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum LineKind {
    /// Axis line itself.
    Axis,
    /// Major gridlines.
    MajorGrid,
    /// Minor gridlines.
    MinorGrid,
    /// Walls or floor in a three-dimensional plot.
    Wall,
}

/// One axis line and its optional appearance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Line {
    /// Line role.
    pub kind: LineKind,
    /// Required line appearance from the immediately following LineFormat.
    pub format: format::Line,
}

/// One chart axis.
#[derive(Debug, PartialEq)]
pub struct Axis {
    /// Axis role.
    pub kind: Kind,
    /// Optional numeric scale.
    pub scale: Option<Scale>,
    /// Optional tick configuration.
    pub tick: Option<Tick>,
    /// Ordered axis lines.
    pub lines: Vec<Line>,
}

impl Axis {
    /// Creates an axis with no explicit scale, ticks, or lines.
    pub const fn new(kind: Kind) -> Self {
        Self {
            kind,
            scale: None,
            tick: None,
            lines: Vec::new(),
        }
    }
}
