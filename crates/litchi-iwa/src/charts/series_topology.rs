//! Strict mapping from chart data axes to visible native series.
//!
//! Most chart kinds use one data row or column per series. Scatter and bubble
//! chart families expose one native series: their ordinary forms use `X + Y`
//! and `X + Y + size`, while their multi-data forms switch datasets
//! interactively instead of displaying multiple series at once.

use crate::charts::{ChartData, ChartKind, ChartSeriesDirection};
use crate::{Error, Result};

const SCATTER_COMPONENT_COUNT: usize = 2;
const BUBBLE_COMPONENT_COUNT: usize = 3;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChartSeriesTopology {
    Ordinary,
    SingleScatter,
    SingleBubble,
    InteractivePoint,
}

impl ChartSeriesTopology {
    const fn for_chart_kind(kind: ChartKind) -> Self {
        match kind {
            ChartKind::Scatter2d => Self::SingleScatter,
            ChartKind::Bubble2d => Self::SingleBubble,
            ChartKind::MultiDataScatter2d | ChartKind::MultiDataBubble2d => Self::InteractivePoint,
            _ => Self::Ordinary,
        }
    }

    const fn required_component_count(self) -> Option<usize> {
        match self {
            Self::SingleScatter => Some(SCATTER_COMPONENT_COUNT),
            Self::SingleBubble => Some(BUBBLE_COMPONENT_COUNT),
            Self::Ordinary | Self::InteractivePoint => None,
        }
    }

    const fn series_count(self, component_count: usize) -> usize {
        match self {
            Self::Ordinary => component_count,
            Self::SingleScatter | Self::SingleBubble | Self::InteractivePoint => 1,
        }
    }
}

/// Count visible native series from typed chart data and orientation.
pub(crate) fn chart_series_count(
    kind: ChartKind,
    direction: ChartSeriesDirection,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    let component_count = match direction {
        ChartSeriesDirection::Rows => data.row_names().len(),
        ChartSeriesDirection::Columns => data.column_names().len(),
        ChartSeriesDirection::Unsupported(value) => {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has unsupported series direction {value}"
            )));
        },
    };
    let topology = ChartSeriesTopology::for_chart_kind(kind);
    if let Some(required) = topology.required_component_count()
        && component_count != required
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {kind:?} data requires {required} components, got {component_count}"
        )));
    }
    let count = topology.series_count(component_count);
    if count == 0 {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no visible series"
        )));
    }
    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ordinary_and_interactive_point_topologies_count_visible_series() {
        let data = data_with_rows(6);
        assert_eq!(
            chart_series_count(
                ChartKind::Line2d,
                ChartSeriesDirection::Rows,
                &data,
                "test",
                1
            )
            .unwrap(),
            6
        );
        assert_eq!(
            chart_series_count(
                ChartKind::MultiDataScatter2d,
                ChartSeriesDirection::Rows,
                &data,
                "test",
                1
            )
            .unwrap(),
            1
        );
        assert_eq!(
            chart_series_count(
                ChartKind::MultiDataBubble2d,
                ChartSeriesDirection::Rows,
                &data,
                "test",
                1
            )
            .unwrap(),
            1
        );
    }

    #[test]
    fn single_scatter_and_bubble_topologies_count_one_series() {
        assert_eq!(
            chart_series_count(
                ChartKind::Scatter2d,
                ChartSeriesDirection::Rows,
                &data_with_rows(2),
                "test",
                1
            )
            .unwrap(),
            1
        );
        assert_eq!(
            chart_series_count(
                ChartKind::Bubble2d,
                ChartSeriesDirection::Rows,
                &data_with_rows(3),
                "test",
                1
            )
            .unwrap(),
            1
        );
    }

    #[test]
    fn ordinary_point_topologies_reject_malformed_component_counts() {
        for rows in [1, 3] {
            assert!(
                chart_series_count(
                    ChartKind::Scatter2d,
                    ChartSeriesDirection::Rows,
                    &data_with_rows(rows),
                    "test",
                    1
                )
                .is_err()
            );
        }
        for rows in [2, 4] {
            assert!(
                chart_series_count(
                    ChartKind::Bubble2d,
                    ChartSeriesDirection::Rows,
                    &data_with_rows(rows),
                    "test",
                    1
                )
                .is_err()
            );
        }
    }

    #[test]
    fn column_orientation_uses_column_components() {
        let data = ChartData::new(
            vec!["row".to_owned()],
            (0..4).map(|index| format!("column {index}")).collect(),
            vec![vec![Some(1.0); 4]],
        )
        .unwrap();
        assert_eq!(
            chart_series_count(
                ChartKind::MultiDataScatter2d,
                ChartSeriesDirection::Columns,
                &data,
                "test",
                1
            )
            .unwrap(),
            1
        );
    }

    fn data_with_rows(rows: usize) -> ChartData {
        ChartData::new(
            (0..rows).map(|index| format!("row {index}")).collect(),
            vec!["column".to_owned()],
            vec![vec![Some(1.0)]; rows],
        )
        .unwrap()
    }
}
