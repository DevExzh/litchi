//! Strongly typed chart kinds shared by Pages, Numbers, and Keynote.

use crate::protobuf::tsch;

/// A native iWork chart kind.
///
/// `Unsupported` preserves forward compatibility without turning a future
/// protobuf value into an untyped string or silently treating it as another
/// chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartKind {
    Undefined,
    Column2d,
    Bar2d,
    Line2d,
    Area2d,
    Pie2d,
    StackedColumn2d,
    StackedBar2d,
    StackedArea2d,
    Scatter2d,
    Mixed2d,
    TwoAxis2d,
    Column3d,
    Bar3d,
    Line3d,
    Area3d,
    Pie3d,
    StackedColumn3d,
    StackedBar3d,
    StackedArea3d,
    MultiDataColumn2d,
    MultiDataBar2d,
    Bubble2d,
    MultiDataScatter2d,
    MultiDataBubble2d,
    Donut2d,
    Donut3d,
    Radar2d,
    Unsupported(i32),
}

impl ChartKind {
    /// Whether the chart exposes the native Wedges rotation control.
    pub const fn supports_pie_start_angle(self) -> bool {
        matches!(
            self,
            Self::Pie2d | Self::Pie3d | Self::Donut2d | Self::Donut3d
        )
    }

    /// Decode the integer stored by the iWork protobuf schema.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            x if x == tsch::ChartType::UndefinedChartType as i32 => Self::Undefined,
            x if x == tsch::ChartType::ColumnChartType2D as i32 => Self::Column2d,
            x if x == tsch::ChartType::BarChartType2D as i32 => Self::Bar2d,
            x if x == tsch::ChartType::LineChartType2D as i32 => Self::Line2d,
            x if x == tsch::ChartType::AreaChartType2D as i32 => Self::Area2d,
            x if x == tsch::ChartType::PieChartType2D as i32 => Self::Pie2d,
            x if x == tsch::ChartType::StackedColumnChartType2D as i32 => Self::StackedColumn2d,
            x if x == tsch::ChartType::StackedBarChartType2D as i32 => Self::StackedBar2d,
            x if x == tsch::ChartType::StackedAreaChartType2D as i32 => Self::StackedArea2d,
            x if x == tsch::ChartType::ScatterChartType2D as i32 => Self::Scatter2d,
            x if x == tsch::ChartType::MixedChartType2D as i32 => Self::Mixed2d,
            x if x == tsch::ChartType::TwoAxisChartType2D as i32 => Self::TwoAxis2d,
            x if x == tsch::ChartType::ColumnChartType3D as i32 => Self::Column3d,
            x if x == tsch::ChartType::BarChartType3D as i32 => Self::Bar3d,
            x if x == tsch::ChartType::LineChartType3D as i32 => Self::Line3d,
            x if x == tsch::ChartType::AreaChartType3D as i32 => Self::Area3d,
            x if x == tsch::ChartType::PieChartType3D as i32 => Self::Pie3d,
            x if x == tsch::ChartType::StackedColumnChartType3D as i32 => Self::StackedColumn3d,
            x if x == tsch::ChartType::StackedBarChartType3D as i32 => Self::StackedBar3d,
            x if x == tsch::ChartType::StackedAreaChartType3D as i32 => Self::StackedArea3d,
            x if x == tsch::ChartType::MultiDataColumnChartType2D as i32 => Self::MultiDataColumn2d,
            x if x == tsch::ChartType::MultiDataBarChartType2D as i32 => Self::MultiDataBar2d,
            x if x == tsch::ChartType::BubbleChartType2D as i32 => Self::Bubble2d,
            x if x == tsch::ChartType::MultiDataScatterChartType2D as i32 => {
                Self::MultiDataScatter2d
            },
            x if x == tsch::ChartType::MultiDataBubbleChartType2D as i32 => Self::MultiDataBubble2d,
            x if x == tsch::ChartType::DonutChartType2D as i32 => Self::Donut2d,
            x if x == tsch::ChartType::DonutChartType3D as i32 => Self::Donut3d,
            x if x == tsch::ChartType::RadarChartType2D as i32 => Self::Radar2d,
            value => Self::Unsupported(value),
        }
    }

    /// Return the integer used by the iWork protobuf schema.
    pub const fn into_raw(self) -> i32 {
        match self {
            Self::Undefined => tsch::ChartType::UndefinedChartType as i32,
            Self::Column2d => tsch::ChartType::ColumnChartType2D as i32,
            Self::Bar2d => tsch::ChartType::BarChartType2D as i32,
            Self::Line2d => tsch::ChartType::LineChartType2D as i32,
            Self::Area2d => tsch::ChartType::AreaChartType2D as i32,
            Self::Pie2d => tsch::ChartType::PieChartType2D as i32,
            Self::StackedColumn2d => tsch::ChartType::StackedColumnChartType2D as i32,
            Self::StackedBar2d => tsch::ChartType::StackedBarChartType2D as i32,
            Self::StackedArea2d => tsch::ChartType::StackedAreaChartType2D as i32,
            Self::Scatter2d => tsch::ChartType::ScatterChartType2D as i32,
            Self::Mixed2d => tsch::ChartType::MixedChartType2D as i32,
            Self::TwoAxis2d => tsch::ChartType::TwoAxisChartType2D as i32,
            Self::Column3d => tsch::ChartType::ColumnChartType3D as i32,
            Self::Bar3d => tsch::ChartType::BarChartType3D as i32,
            Self::Line3d => tsch::ChartType::LineChartType3D as i32,
            Self::Area3d => tsch::ChartType::AreaChartType3D as i32,
            Self::Pie3d => tsch::ChartType::PieChartType3D as i32,
            Self::StackedColumn3d => tsch::ChartType::StackedColumnChartType3D as i32,
            Self::StackedBar3d => tsch::ChartType::StackedBarChartType3D as i32,
            Self::StackedArea3d => tsch::ChartType::StackedAreaChartType3D as i32,
            Self::MultiDataColumn2d => tsch::ChartType::MultiDataColumnChartType2D as i32,
            Self::MultiDataBar2d => tsch::ChartType::MultiDataBarChartType2D as i32,
            Self::Bubble2d => tsch::ChartType::BubbleChartType2D as i32,
            Self::MultiDataScatter2d => tsch::ChartType::MultiDataScatterChartType2D as i32,
            Self::MultiDataBubble2d => tsch::ChartType::MultiDataBubbleChartType2D as i32,
            Self::Donut2d => tsch::ChartType::DonutChartType2D as i32,
            Self::Donut3d => tsch::ChartType::DonutChartType3D as i32,
            Self::Radar2d => tsch::ChartType::RadarChartType2D as i32,
            Self::Unsupported(value) => value,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_known_protobuf_value_round_trips() {
        for value in 0..=tsch::ChartType::RadarChartType2D as i32 {
            assert_eq!(ChartKind::from_raw(value).into_raw(), value);
        }
    }

    #[test]
    fn future_values_remain_lossless() {
        const FUTURE_KIND: i32 = 9_001;
        assert_eq!(
            ChartKind::from_raw(FUTURE_KIND),
            ChartKind::Unsupported(FUTURE_KIND)
        );
        assert_eq!(ChartKind::Unsupported(FUTURE_KIND).into_raw(), FUTURE_KIND);
    }
}
