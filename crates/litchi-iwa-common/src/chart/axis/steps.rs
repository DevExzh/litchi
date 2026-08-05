//! Archive-free step counts for native chart value axes.

use std::num::NonZeroU32;

const MAX_NATIVE_AXIS_STEP_COUNT: u32 = i32::MAX as u32;

/// Validation failures for value-axis step counts.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A major-step count was zero.
    #[error("chart value-axis major step count must be positive")]
    MajorZero,
    /// A major-step count exceeded the native signed integer range.
    #[error("chart value-axis major step count exceeds {MAX_NATIVE_AXIS_STEP_COUNT}")]
    MajorOutOfRange,
    /// A minor-step count exceeded the native signed integer range.
    #[error("chart value-axis minor step count exceeds {MAX_NATIVE_AXIS_STEP_COUNT}")]
    MinorOutOfRange,
}

/// Result type for value-axis step construction.
pub type Result<T> = std::result::Result<T, Error>;

/// One positive number of major intervals in a value-axis scale.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct MajorStepCount(NonZeroU32);

impl MajorStepCount {
    /// One major interval.
    pub const ONE: Self = Self(NonZeroU32::MIN);

    /// Create a positive native major-step count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::MajorZero`] for zero or [`Error::MajorOutOfRange`]
    /// above the native signed integer limit.
    #[must_use = "use the validated step count or handle its validation error"]
    pub fn new(value: u32) -> Result<Self> {
        if value > MAX_NATIVE_AXIS_STEP_COUNT {
            return Err(Error::MajorOutOfRange);
        }
        NonZeroU32::new(value).map(Self).ok_or(Error::MajorZero)
    }

    /// Return the number shown in the native Major Steps control.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0.get()
    }
}

impl TryFrom<u32> for MajorStepCount {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// One non-negative number of minor intervals between major steps.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct MinorStepCount(u32);

impl MinorStepCount {
    /// Create a native minor-step count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::MinorOutOfRange`] above the native signed integer
    /// limit.
    #[must_use = "use the validated step count or handle its validation error"]
    pub fn new(value: u32) -> Result<Self> {
        if value > MAX_NATIVE_AXIS_STEP_COUNT {
            return Err(Error::MinorOutOfRange);
        }
        Ok(Self(value))
    }

    /// Return the number shown in the native Minor Steps control.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for MinorStepCount {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// Optional manual major and minor step settings for a value axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Steps {
    major: Option<MajorStepCount>,
    minor: Option<MinorStepCount>,
}

impl Steps {
    /// Build independent automatic or explicit step settings.
    #[must_use]
    pub const fn new(major: Option<MajorStepCount>, minor: Option<MinorStepCount>) -> Self {
        Self { major, minor }
    }

    /// Use native automatic major and minor step counts.
    #[must_use]
    pub const fn automatic() -> Self {
        Self::new(None, None)
    }

    /// Build fully manual step settings.
    #[must_use]
    pub const fn fixed(major: MajorStepCount, minor: MinorStepCount) -> Self {
        Self::new(Some(major), Some(minor))
    }

    /// Return the optional manual major-step count.
    #[must_use]
    pub const fn major(self) -> Option<MajorStepCount> {
        self.major
    }

    /// Return the optional manual minor-step count.
    #[must_use]
    pub const fn minor(self) -> Option<MinorStepCount> {
        self.minor
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, MajorStepCount, MinorStepCount, Steps};

    #[test]
    fn steps_are_compact_and_strict() {
        assert_eq!(size_of::<MajorStepCount>(), 4);
        assert_eq!(size_of::<MinorStepCount>(), 4);
        assert_eq!(size_of::<Steps>(), 12);
        assert_eq!(MajorStepCount::new(0), Err(Error::MajorZero));
        assert_eq!(MinorStepCount::new(0).unwrap().value(), 0);
        assert_eq!(Steps::automatic(), Steps::new(None, None));
    }
}
