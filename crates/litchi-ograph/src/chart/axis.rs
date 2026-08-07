//! Axis scales, ticks, and lines.

use super::{format, layout};

/// Primary or secondary axis-parent identifier.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct ParentId(u8);

impl ParentId {
    /// Primary axis group.
    pub const PRIMARY: Self = Self(0);
    /// Secondary axis group.
    pub const SECONDARY: Self = Self(1);

    /// Creates a checked axis-parent identifier.
    #[must_use]
    pub const fn new(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::PRIMARY),
            1 => Some(Self::SECONDARY),
            _ => None,
        }
    }

    /// Raw primary-or-secondary index.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }

    pub(super) const fn index(self) -> usize {
        self.0 as usize
    }
}

/// One primary or secondary axis-parent collection and its mandatory plot
/// position. Axis and chart-group records remain ordered in the containing
/// chart model while collection ownership is validated by the codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Parent {
    id: ParentId,
    pos: layout::Pos,
}

impl Parent {
    /// Creates a primary axis-parent collection.
    #[must_use]
    pub const fn primary(pos: layout::Pos) -> Self {
        Self {
            id: ParentId::PRIMARY,
            pos,
        }
    }

    /// Creates a secondary axis-parent collection.
    #[must_use]
    pub const fn secondary(pos: layout::Pos) -> Self {
        Self {
            id: ParentId::SECONDARY,
            pos,
        }
    }

    /// Stable primary-or-secondary identifier.
    #[must_use]
    pub const fn id(self) -> ParentId {
        self.id
    }

    /// Whether this is the secondary axis group.
    #[must_use]
    pub const fn is_secondary(self) -> bool {
        matches!(self.id, ParentId::SECONDARY)
    }

    /// Mandatory axis-group plot position.
    #[must_use]
    pub const fn pos(self) -> layout::Pos {
        self.pos
    }
}

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
    /// Raw `ValueRange` flags, including preserved producer bits.
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
    /// Required line appearance from the immediately following `LineFormat`.
    pub format: format::Line,
}

/// One chart axis.
#[derive(Debug, PartialEq)]
pub struct Axis {
    /// Axis-parent collection that owns this axis.
    pub parent: ParentId,
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
    #[must_use]
    pub const fn new(kind: Kind) -> Self {
        Self::in_parent(kind, ParentId::PRIMARY)
    }

    /// Creates an axis owned by a selected primary or secondary group.
    #[must_use]
    pub const fn in_parent(kind: Kind, parent: ParentId) -> Self {
        Self {
            parent,
            kind,
            scale: None,
            tick: None,
            lines: Vec::new(),
        }
    }
}
