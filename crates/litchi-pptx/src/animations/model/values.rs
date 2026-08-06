//! Compact value objects used by the animation semantic model.

/// Behavior of animated properties after a time node becomes inactive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fill {
    Remove,
    Freeze,
    Hold,
    Transition,
}

/// Policy controlling whether a completed time node can restart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Restart {
    Always,
    WhenNotActive,
    Never,
}

/// Repeat count for a time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Repeat {
    /// Count in OOXML thousandths, where `1000` means one iteration.
    Finite(u32),
    Indefinite,
}

/// Nonzero playback speed in thousandths of a percent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Speed(pub(in crate::animations) i32);

impl Speed {
    /// Return the encoded OOXML percentage value.
    pub const fn thousandths_percent(self) -> i32 {
        self.0
    }
}

/// Positive fixed percentage used for acceleration and deceleration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MotionFraction(pub(in crate::animations) u32);

impl MotionFraction {
    /// Return the encoded OOXML percentage value.
    pub const fn thousandths_percent(self) -> u32 {
        self.0
    }
}

/// Synchronization policy between a time node and its containing group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SyncBehavior {
    CanSlip,
    Locked,
    /// PowerPoint's assumed synchronization behavior.
    None,
}

/// Exact normalized time in the inclusive range `0..=1`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NormalizedTime {
    pub(in crate::animations) numerator: u64,
    pub(in crate::animations) scale: u64,
}

impl NormalizedTime {
    /// Exact numerator of the normalized decimal value.
    pub const fn numerator(self) -> u64 {
        self.numerator
    }

    /// Exact power-of-ten scale of the normalized decimal value.
    pub const fn scale(self) -> u64 {
        self.scale
    }
}

/// A source-time to warped-time mapping point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TimePoint {
    pub local_time: NormalizedTime,
    pub warped_time: NormalizedTime,
}

impl TimePoint {
    pub const fn new(local_time: NormalizedTime, warped_time: NormalizedTime) -> Self {
        Self {
            local_time,
            warped_time,
        }
    }
}

/// Bounded piecewise time-warp filter for a common time node.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeFilter {
    pub(in crate::animations) points: Box<[TimePoint]>,
}

/// Duration of a simple animation time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Duration {
    /// A finite duration in milliseconds.
    Finite(u32),
    /// An animation that has no finite duration.
    Indefinite,
}

impl Duration {
    /// Construct a finite duration in milliseconds.
    pub const fn milliseconds(value: u32) -> Self {
        Self::Finite(value)
    }

    /// Return the finite millisecond value, or `None` for an indefinite duration.
    pub const fn as_milliseconds(self) -> Option<u32> {
        match self {
            Self::Finite(value) => Some(value),
            Self::Indefinite => None,
        }
    }
}

impl From<u32> for Duration {
    fn from(value: u32) -> Self {
        Self::Finite(value)
    }
}

impl PartialEq<u32> for Duration {
    fn eq(&self, other: &u32) -> bool {
        matches!(self, Self::Finite(value) if value == other)
    }
}
