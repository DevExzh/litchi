//! Archive-free visibility values for native pie and donut charts.

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
