//! Archive-free visibility and geometry values for native pie and donut
//! charts.

const DATA_POINT_NAMES_MASK: u8 = 0b01;
const VALUES_MASK: u8 = 0b10;
const HIDDEN_LEADER_LINE: i32 = 0;
const VISIBLE_LEADER_LINE: i32 = 2;

/// Visibility of the two native label components for one pie or donut wedge.
///
/// The two boolean settings are packed into one byte. Native wire validation
/// remains in the concrete format package adapters; this value contains only
/// semantic state.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LabelVisibility(u8);

impl LabelVisibility {
    /// Native iWork defaults: values visible and data-point names hidden.
    pub const DEFAULT: Self = Self::new(false, true);
    /// Hide both label components.
    pub const HIDDEN: Self = Self::new(false, false);
    /// Show only data-point names.
    pub const DATA_POINT_NAMES_ONLY: Self = Self::new(true, false);
    /// Show only values.
    pub const VALUES_ONLY: Self = Self::DEFAULT;
    /// Show data-point names and values.
    pub const ALL: Self = Self::new(true, true);

    /// Construct an explicit label-visibility combination.
    #[must_use]
    pub const fn new(data_point_names_visible: bool, values_visible: bool) -> Self {
        let mut bits = 0;
        if data_point_names_visible {
            bits |= DATA_POINT_NAMES_MASK;
        }
        if values_visible {
            bits |= VALUES_MASK;
        }
        Self(bits)
    }

    /// Whether iWork renders the data-point name.
    #[must_use]
    pub const fn data_point_names_visible(self) -> bool {
        self.0 & DATA_POINT_NAMES_MASK != 0
    }

    /// Whether iWork renders the formatted value.
    #[must_use]
    pub const fn values_visible(self) -> bool {
        self.0 & VALUES_MASK != 0
    }
}

impl Default for LabelVisibility {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Recognized native leader-line states.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LeaderLineVisibilityKind {
    /// Do not draw a line between a label and its wedge.
    Hidden,
    /// Draw a line between a label and its wedge.
    Visible,
}

/// Whether iWork draws a leader line between a pie label and its wedge.
///
/// The native integer is retained directly so future values round-trip
/// without widening the semantic value beyond four bytes.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct LeaderLineVisibility(i32);

impl LeaderLineVisibility {
    /// Hide leader lines.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Hidden: Self = Self(HIDDEN_LEADER_LINE);
    /// Show leader lines, the native default.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Visible: Self = Self(VISIBLE_LEADER_LINE);

    /// Decode the integer stored by a native iWork archive.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the integer stored by a native iWork archive.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized native state, if known.
    #[must_use]
    pub const fn kind(self) -> Option<LeaderLineVisibilityKind> {
        match self.0 {
            HIDDEN_LEADER_LINE => Some(LeaderLineVisibilityKind::Hidden),
            VISIBLE_LEADER_LINE => Some(LeaderLineVisibilityKind::Visible),
            _ => None,
        }
    }

    /// Whether this value is not one of the known native states.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.kind().is_none()
    }
}

impl std::fmt::Debug for LeaderLineVisibility {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.0 {
            HIDDEN_LEADER_LINE => formatter.write_str("Hidden"),
            VISIBLE_LEADER_LINE => formatter.write_str("Visible"),
            value => formatter.debug_tuple("Unsupported").field(&value).finish(),
        }
    }
}

impl Default for LeaderLineVisibility {
    fn default() -> Self {
        Self::Visible
    }
}

const FULL_TURN_DEGREES: f32 = 360.0;
const PERCENT_SCALE: f32 = 100.0;
const MAXIMUM_WEDGE_EXPLOSION_PERCENT: f32 = 100.0;
const MINIMUM_LABEL_DISTANCE_PERCENT: f32 = 30.0;
const DEFAULT_LABEL_DISTANCE_PERCENT: f32 = 67.0;
const MAXIMUM_LABEL_DISTANCE_PERCENT: f32 = 200.0;

/// Validation failures of the pie and donut geometry values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The value was NaN or infinite.
    #[error("pie chart value must be finite")]
    NonFinite,
    /// The value was outside the native inspector's domain.
    #[error("pie chart value is outside the native inspector range")]
    OutOfRange,
}

/// Result type of the pie and donut geometry values.
pub type Result<T> = std::result::Result<T, Error>;

fn check(value: f32, valid: bool) -> Result<f32> {
    if !value.is_finite() {
        return Err(Error::NonFinite);
    }
    if !valid {
        return Err(Error::OutOfRange);
    }
    Ok(value)
}

/// Rotation of the first wedge of a pie or donut chart (the Wedges
/// inspector's `Rotation Angle`), in clockwise degrees in `0°..360°`.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct StartAngle(f32);

impl StartAngle {
    /// No rotation, the native default.
    pub const ZERO: Self = Self(0.0);
    /// A quarter turn.
    pub const QUARTER_TURN: Self = Self(90.0);
    /// A half turn.
    pub const HALF_TURN: Self = Self(180.0);
    /// Three quarters of a turn.
    pub const THREE_QUARTER_TURN: Self = Self(270.0);

    /// Construct an angle accepted by the native inspector.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `0°..360°`.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        check(degrees, (0.0..FULL_TURN_DEGREES).contains(&degrees)).map(Self)
    }

    /// The clockwise degrees displayed by iWork (also the stored value).
    #[must_use]
    pub const fn degrees(self) -> f32 {
        self.0
    }
}

impl Default for StartAngle {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for StartAngle {
    type Error = Error;

    fn try_from(degrees: f32) -> Result<Self> {
        Self::from_degrees(degrees)
    }
}

/// Distance of one wedge from the chart center (the Wedges inspector's
/// `Explode`), stored as a fraction in `0.0..=1.0` and shown in percent.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct WedgeExplosion(f32);

impl WedgeExplosion {
    /// The wedge at its unseparated position, the native default.
    pub const ZERO: Self = Self(0.0);
    /// The inspector's maximum distance.
    pub const MAXIMUM: Self = Self(1.0);

    /// Construct a wedge position from an inspector percentage in
    /// `0%..=100%`.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`].
    pub fn from_percent(percent: f32) -> Result<Self> {
        check(
            percent,
            (0.0..=MAXIMUM_WEDGE_EXPLOSION_PERCENT).contains(&percent),
        )
        .map(|percent| Self(percent / PERCENT_SCALE))
    }

    /// Construct a wedge position from its stored fraction.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `0.0..=1.0`.
    pub fn from_native(fraction: f32) -> Result<Self> {
        check(fraction, (0.0..=1.0).contains(&fraction)).map(Self)
    }

    /// The percentage displayed by iWork.
    #[must_use]
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_SCALE
    }

    /// The stored fraction.
    #[must_use]
    pub const fn native_fraction(self) -> f32 {
        self.0
    }
}

impl Default for WedgeExplosion {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for WedgeExplosion {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

/// Zero-based index of one wedge, in chart series order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct WedgeIndex(usize);

impl WedgeIndex {
    /// Construct a zero-based wedge index.
    #[must_use]
    pub const fn from_zero_based(index: usize) -> Self {
        Self(index)
    }

    /// The zero-based wedge index.
    #[must_use]
    pub const fn zero_based(self) -> usize {
        self.0
    }
}

/// Distance of a wedge's label from the chart center, in percent in
/// `30%..=200%` (stored unchanged).
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct LabelDistance(f32);

impl LabelDistance {
    /// The nearest position the inspector accepts.
    pub const MINIMUM: Self = Self(MINIMUM_LABEL_DISTANCE_PERCENT);
    /// The native position when no override is stored.
    pub const DEFAULT: Self = Self(DEFAULT_LABEL_DISTANCE_PERCENT);
    /// The farthest position the inspector accepts.
    pub const MAXIMUM: Self = Self(MAXIMUM_LABEL_DISTANCE_PERCENT);

    /// Construct a label distance from an inspector percentage.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `30%..=200%`.
    pub fn from_percent(percent: f32) -> Result<Self> {
        check(
            percent,
            (MINIMUM_LABEL_DISTANCE_PERCENT..=MAXIMUM_LABEL_DISTANCE_PERCENT).contains(&percent),
        )
        .map(Self)
    }

    /// The percentage displayed by iWork (also the stored value).
    #[must_use]
    pub const fn percent(self) -> f32 {
        self.0
    }
}

impl Default for LabelDistance {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<f32> for LabelDistance {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{LabelVisibility, LeaderLineVisibility, LeaderLineVisibilityKind};

    #[test]
    fn label_visibility_is_packed_and_exhaustive() {
        assert_eq!(size_of::<LabelVisibility>(), 1);
        assert_eq!(LabelVisibility::default(), LabelVisibility::VALUES_ONLY);
        assert!(!LabelVisibility::HIDDEN.values_visible());
        assert!(LabelVisibility::DATA_POINT_NAMES_ONLY.data_point_names_visible());
        assert!(LabelVisibility::ALL.values_visible());
        for names in [false, true] {
            for values in [false, true] {
                let visibility = LabelVisibility::new(names, values);
                assert_eq!(visibility.data_point_names_visible(), names);
                assert_eq!(visibility.values_visible(), values);
            }
        }
    }

    #[test]
    fn leader_line_visibility_is_compact_and_lossless() {
        assert_eq!(size_of::<LeaderLineVisibility>(), 4);
        assert_eq!(LeaderLineVisibility::Hidden.native_value(), 0);
        assert_eq!(LeaderLineVisibility::Visible.native_value(), 2);
        assert_eq!(
            LeaderLineVisibility::Visible.kind(),
            Some(LeaderLineVisibilityKind::Visible)
        );
        for native in [i32::MIN, -1, 1, 3, i32::MAX] {
            let value = LeaderLineVisibility::from_native(native);
            assert_eq!(value.native_value(), native);
            assert_eq!(value.kind(), None);
            assert!(value.is_unsupported());
        }
        assert!(!LeaderLineVisibility::Hidden.is_unsupported());
        assert!(!LeaderLineVisibility::Visible.is_unsupported());
    }
}

#[cfg(test)]
mod geometry_tests {
    use super::{Error, LabelDistance, StartAngle, WedgeExplosion, WedgeIndex};

    #[test]
    fn start_angles_are_half_open() {
        assert_eq!(StartAngle::default(), StartAngle::ZERO);
        assert_eq!(
            StartAngle::from_degrees(123.5).map(StartAngle::degrees),
            Ok(123.5)
        );
        assert_eq!(StartAngle::try_from(90.0), Ok(StartAngle::QUARTER_TURN));
        for invalid in [-0.1, 360.0, 400.0] {
            assert_eq!(StartAngle::from_degrees(invalid), Err(Error::OutOfRange));
        }
        for invalid in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY] {
            assert_eq!(StartAngle::from_degrees(invalid), Err(Error::NonFinite));
        }
    }

    #[test]
    fn wedge_explosions_store_fractions() {
        let half = WedgeExplosion::from_percent(50.0).unwrap();
        assert_eq!(half.native_fraction(), 0.5);
        assert_eq!(half.percent(), 50.0);
        assert_eq!(
            WedgeExplosion::from_native(1.0),
            Ok(WedgeExplosion::MAXIMUM)
        );
        assert_eq!(WedgeExplosion::from_percent(100.1), Err(Error::OutOfRange));
        assert_eq!(WedgeExplosion::from_native(-0.5), Err(Error::OutOfRange));
        assert_eq!(WedgeExplosion::default(), WedgeExplosion::ZERO);
        assert_eq!(WedgeIndex::from_zero_based(3).zero_based(), 3);
    }

    #[test]
    fn label_distances_are_inclusive_percentages() {
        assert_eq!(LabelDistance::default().percent(), 67.0);
        assert_eq!(
            LabelDistance::from_percent(30.0),
            Ok(LabelDistance::MINIMUM)
        );
        assert_eq!(
            LabelDistance::from_percent(200.0),
            Ok(LabelDistance::MAXIMUM)
        );
        assert_eq!(LabelDistance::from_percent(29.9), Err(Error::OutOfRange));
        assert_eq!(LabelDistance::try_from(f32::NAN), Err(Error::NonFinite));
    }
}
