//! Lossless chart-kind values shared by concrete iWork format owners.

/// A native iWork chart kind.
///
/// Known native identifiers are exposed as concise semantic constants. Any
/// other identifier remains lossless in the same four-byte value. The type
/// contains no protobuf, archive, or package state.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Kind(i32);

#[allow(
    non_upper_case_globals,
    reason = "PascalCase constants preserve the focused semantic Kind::Column2d API"
)]
impl Kind {
    /// No chart kind was specified by the producer.
    pub const Undefined: Self = Self(0);
    /// Two-dimensional column chart.
    pub const Column2d: Self = Self(1);
    /// Two-dimensional bar chart.
    pub const Bar2d: Self = Self(2);
    /// Two-dimensional line chart.
    pub const Line2d: Self = Self(3);
    /// Two-dimensional area chart.
    pub const Area2d: Self = Self(4);
    /// Two-dimensional pie chart.
    pub const Pie2d: Self = Self(5);
    /// Two-dimensional stacked column chart.
    pub const StackedColumn2d: Self = Self(6);
    /// Two-dimensional stacked bar chart.
    pub const StackedBar2d: Self = Self(7);
    /// Two-dimensional stacked area chart.
    pub const StackedArea2d: Self = Self(8);
    /// Two-dimensional scatter chart.
    pub const Scatter2d: Self = Self(9);
    /// Two-dimensional mixed chart.
    pub const Mixed2d: Self = Self(10);
    /// Two-dimensional two-axis chart.
    pub const TwoAxis2d: Self = Self(11);
    /// Three-dimensional column chart.
    pub const Column3d: Self = Self(12);
    /// Three-dimensional bar chart.
    pub const Bar3d: Self = Self(13);
    /// Three-dimensional line chart.
    pub const Line3d: Self = Self(14);
    /// Three-dimensional area chart.
    pub const Area3d: Self = Self(15);
    /// Three-dimensional pie chart.
    pub const Pie3d: Self = Self(16);
    /// Three-dimensional stacked column chart.
    pub const StackedColumn3d: Self = Self(17);
    /// Three-dimensional stacked bar chart.
    pub const StackedBar3d: Self = Self(18);
    /// Three-dimensional stacked area chart.
    pub const StackedArea3d: Self = Self(19);
    /// Multi-data two-dimensional column chart.
    pub const MultiDataColumn2d: Self = Self(20);
    /// Multi-data two-dimensional bar chart.
    pub const MultiDataBar2d: Self = Self(21);
    /// Two-dimensional bubble chart.
    pub const Bubble2d: Self = Self(22);
    /// Multi-data two-dimensional scatter chart.
    pub const MultiDataScatter2d: Self = Self(23);
    /// Multi-data two-dimensional bubble chart.
    pub const MultiDataBubble2d: Self = Self(24);
    /// Two-dimensional donut chart.
    pub const Donut2d: Self = Self(25);
    /// Three-dimensional donut chart.
    pub const Donut3d: Self = Self(26);
    /// Two-dimensional radar chart.
    pub const Radar2d: Self = Self(27);

    /// Construct a kind from its native identifier without losing unknown
    /// values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native chart-kind identifier used by iWork.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Whether this value is not one of the known native chart kinds.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.0 < 0 || self.0 > 27
    }

    /// Whether the 3D Scene inspector exposes rectangular/cylindrical bars.
    #[must_use]
    pub const fn supports_3d_bar_shape(self) -> bool {
        matches!(
            self,
            Self::Column3d | Self::Bar3d | Self::StackedColumn3d | Self::StackedBar3d
        )
    }

    /// Whether the chart exposes the native 3D Scene inspector.
    #[must_use]
    pub const fn supports_3d_scene(self) -> bool {
        matches!(
            self,
            Self::Column3d
                | Self::Bar3d
                | Self::Line3d
                | Self::Area3d
                | Self::Pie3d
                | Self::StackedColumn3d
                | Self::StackedBar3d
                | Self::StackedArea3d
                | Self::Donut3d
        )
    }

    /// Whether the chart has a native 3D depth ratio.
    #[must_use]
    pub const fn supports_3d_depth(self) -> bool {
        self.supports_3d_scene()
    }

    /// Whether the chart exposes the native 3D Lighting Style menu.
    #[must_use]
    pub const fn supports_3d_lighting_style(self) -> bool {
        self.supports_3d_scene()
    }

    /// Whether the chart exposes 3D primary value-axis label placement.
    #[must_use]
    pub const fn supports_3d_value_axis_label_position(self) -> bool {
        matches!(
            self,
            Self::Column3d
                | Self::Bar3d
                | Self::Line3d
                | Self::Area3d
                | Self::StackedColumn3d
                | Self::StackedBar3d
                | Self::StackedArea3d
        )
    }

    /// Whether the chart exposes the native 3D `Between Series` gap.
    #[must_use]
    pub const fn supports_3d_series_gap(self) -> bool {
        matches!(self, Self::Line3d | Self::Area3d)
    }

    /// Whether the chart exposes the native Wedges rotation control.
    #[must_use]
    pub const fn supports_pie_start_angle(self) -> bool {
        matches!(
            self,
            Self::Pie2d | Self::Pie3d | Self::Donut2d | Self::Donut3d
        )
    }

    /// Whether the chart exposes the native Segments inner-radius control.
    #[must_use]
    pub const fn supports_donut_inner_radius(self) -> bool {
        matches!(self, Self::Donut2d | Self::Donut3d)
    }

    /// Whether the chart exposes the native Radar Chart Grid Shape menu.
    #[must_use]
    pub const fn supports_radar_grid_shape(self) -> bool {
        matches!(self, Self::Radar2d)
    }

    /// Whether the chart exposes the native Radar series-style controls.
    #[must_use]
    pub const fn supports_radar_series_style(self) -> bool {
        matches!(self, Self::Radar2d)
    }

    /// Whether the chart exposes the native Radar rotation-angle control.
    #[must_use]
    pub const fn supports_radar_start_angle(self) -> bool {
        matches!(self, Self::Radar2d)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::Kind;

    #[test]
    fn kind_is_a_compact_lossless_value() {
        assert_eq!(size_of::<Kind>(), size_of::<i32>());
        for value in 0..=27 {
            assert_eq!(Kind::from_native(value).native_value(), value);
        }
    }

    #[test]
    fn unknown_values_remain_lossless() {
        for value in [i32::MIN, -1, 28, 9_001, i32::MAX] {
            let kind = Kind::from_native(value);
            assert!(kind.is_unsupported());
            assert_eq!(kind.native_value(), value);
        }
    }

    #[test]
    fn capability_predicates_match_native_support() {
        assert!(Kind::Radar2d.supports_radar_grid_shape());
        assert!(Kind::Radar2d.supports_radar_series_style());
        assert!(Kind::Radar2d.supports_radar_start_angle());
        assert!(Kind::Donut3d.supports_3d_scene());
        assert!(Kind::Donut3d.supports_pie_start_angle());
        assert!(Kind::Donut3d.supports_donut_inner_radius());
        assert!(!Kind::Donut3d.supports_3d_value_axis_label_position());
        assert!(!Kind::from_native(9_001).supports_3d_scene());
    }
}
