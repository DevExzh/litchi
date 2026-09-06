//! Archive-free Arrange-panel state shared by concrete iWork chart owners.

/// Archive-free state exposed by a chart's Arrange panel.
///
/// The value describes only interaction behavior for an existing chart. It
/// carries no native drawable identifier, archive name, protobuf message, or
/// other package identity. A concrete package owner resolves the selected
/// chart and applies these flags to its native drawable graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct ChartArrangement {
    locked: bool,
    constrain_proportions: bool,
}

impl ChartArrangement {
    /// Construct Arrange-panel state from its two native interaction flags.
    #[must_use]
    pub const fn new(locked: bool, constrain_proportions: bool) -> Self {
        Self {
            locked,
            constrain_proportions,
        }
    }

    /// Return whether the chart is locked against interactive editing.
    #[must_use]
    pub const fn locked(self) -> bool {
        self.locked
    }

    /// Return whether interactive resizing preserves the chart's aspect ratio.
    #[must_use]
    pub const fn constrain_proportions(self) -> bool {
        self.constrain_proportions
    }

    /// Return this state with the requested interactive lock.
    #[must_use]
    pub const fn with_locked(mut self, locked: bool) -> Self {
        self.locked = locked;
        self
    }

    /// Return this state with the requested aspect-ratio constraint.
    #[must_use]
    pub const fn with_constrain_proportions(mut self, constrain_proportions: bool) -> Self {
        self.constrain_proportions = constrain_proportions;
        self
    }
}
