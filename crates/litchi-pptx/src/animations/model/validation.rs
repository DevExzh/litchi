//! Invariant-enforcing constructors for animation values and builds.

use super::super::codec::MAX_TIME_FILTER_POINTS;
use super::super::invalid;
use super::{
    MotionFraction, NormalizedTime, ParagraphTemplate, Speed, TemplateTimeNode, TimeFilter,
};
use crate::Result;

impl ParagraphTemplate {
    /// Construct a paragraph template with a PowerPoint-supported level.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(level: u8, time_node: TemplateTimeNode) -> Result<Self> {
        if level > 9 {
            return Err(invalid("paragraph template level exceeds PowerPoint limit"));
        }
        Ok(Self { level, time_node })
    }
}

impl Speed {
    /// Construct a speed value. `PowerPoint` rejects zero speed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(thousandths_percent: i32) -> Result<Self> {
        if thousandths_percent == 0 {
            Err(invalid("animation speed must be nonzero"))
        } else {
            Ok(Self(thousandths_percent))
        }
    }
}

impl MotionFraction {
    /// Construct a value from thousandths of a percent (`100000` is 100%).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(thousandths_percent: u32) -> Result<Self> {
        if thousandths_percent > 100_000 {
            Err(invalid(
                "animation progression percentage exceeds 100 percent",
            ))
        } else {
            Ok(Self(thousandths_percent))
        }
    }
}

impl NormalizedTime {
    /// Construct a normalized time from millionths.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_millionths(value: u32) -> Result<Self> {
        if value > 1_000_000 {
            return Err(invalid("normalized time exceeds 1.0"));
        }
        Ok(Self::normalized(u64::from(value), 1_000_000))
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(in crate::animations) fn normalized(mut numerator: u64, mut scale: u64) -> Self {
        while scale > 1 && numerator.is_multiple_of(10) {
            numerator /= 10;
            scale /= 10;
        }
        Self { numerator, scale }
    }

    fn strictly_before(self, other: Self) -> bool {
        u128::from(self.numerator) * u128::from(other.scale)
            < u128::from(other.numerator) * u128::from(self.scale)
    }
}

impl TimeFilter {
    /// Construct a filter whose local times are strictly increasing.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(points: Vec<super::TimePoint>) -> Result<Self> {
        if points.is_empty() {
            return Err(invalid(
                "animation time filter must contain at least one point",
            ));
        }
        if points.len() > MAX_TIME_FILTER_POINTS {
            return Err(invalid(
                "animation time filter point count exceeds safety limit",
            ));
        }
        if points
            .windows(2)
            .any(|pair| !pair[0].local_time.strictly_before(pair[1].local_time))
        {
            return Err(invalid(
                "animation time filter local times must be strictly increasing",
            ));
        }
        Ok(Self {
            points: points.into_boxed_slice(),
        })
    }

    /// Mapping points in source-time order.
    #[must_use]
    pub fn points(&self) -> &[super::TimePoint] {
        &self.points
    }
}
