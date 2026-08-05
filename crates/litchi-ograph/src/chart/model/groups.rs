use super::super::{axis, group};
use super::context::Order;

/// Chart-family configuration attached to one group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Family {
    Line {
        flags: u16,
    },
    Bar {
        overlap: group::Overlap,
        gap: group::Gap,
        flags: u16,
    },
    Area {
        flags: u16,
    },
    Pie {
        rotation: u16,
        hole: u16,
        flags: u16,
    },
    Scatter {
        bubble_percent: group::BubblePercent,
        bubble_kind: group::BubbleKind,
        flags: u16,
    },
    Radar {
        filled: bool,
        flags: u16,
    },
    Surface {
        flags: u16,
    },
}

/// One ordered chart group.
#[derive(Debug, PartialEq, Eq)]
pub struct Group {
    /// Axis-parent collection that owns this group.
    pub parent: axis::ParentId,
    pub order: Order,
    pub vary_colors: bool,
    pub family: Family,
    /// Excel-mandatory written-but-unused CrtLink owned by this chart group.
    ///
    /// Standalone Graph preserves this record when present but does not require
    /// it without the unavailable normative chart-sheet grammar.
    pub link: crate::record::line::Link,
    pub lines: Vec<group::Line>,
    pub drop_bars: Vec<group::DropBar>,
}

impl Group {
    /// Primary line-chart group used by a new chart.
    pub const fn line() -> Self {
        Self {
            parent: axis::ParentId::PRIMARY,
            order: Order::ZERO,
            vary_colors: false,
            family: Family::Line { flags: 0 },
            link: crate::record::line::Link::new([0; 10]),
            lines: Vec::new(),
            drop_bars: Vec::new(),
        }
    }
}
