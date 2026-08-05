//! Strict mapping from chart data axes to visible native series.
//!
//! Most chart kinds use one data row or column per series. Scatter and bubble
//! chart families expose one native series: their ordinary forms use `X + Y`
//! and `X + Y + size`, while their multi-data forms switch datasets
//! interactively instead of displaying multiple series at once.

use crate::charts::{ChartData, ChartKind, Direction, DirectionKind};
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
    direction: Direction,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    let component_count = match direction.kind() {
        Some(DirectionKind::Rows) => data.row_names().len(),
        Some(DirectionKind::Columns) => data.column_names().len(),
        None => {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has unsupported series direction {}",
                direction.native_value()
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
            chart_series_count(ChartKind::Line2d, Direction::Rows, &data, "test", 1).unwrap(),
            6
        );
        assert_eq!(
            chart_series_count(
                ChartKind::MultiDataScatter2d,
                Direction::Rows,
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
                Direction::Rows,
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
                Direction::Rows,
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
                Direction::Rows,
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
                    Direction::Rows,
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
                    Direction::Rows,
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
        let rectangular = ChartData::new(
            vec!["Revenue".to_owned(), "Cost".to_owned()],
            vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
            vec![
                vec![Some(1.0), Some(2.0), Some(3.0)],
                vec![Some(4.0), Some(5.0), Some(6.0)],
            ],
        )
        .unwrap();
        assert_eq!(
            chart_series_count(ChartKind::Line2d, Direction::Rows, &rectangular, "test", 1)
                .unwrap(),
            2
        );
        assert_eq!(
            chart_series_count(
                ChartKind::Line2d,
                Direction::Columns,
                &rectangular,
                "test",
                1,
            )
            .unwrap(),
            3
        );

        let data = ChartData::new(
            vec!["row".to_owned()],
            (0..4).map(|index| format!("column {index}")).collect(),
            vec![vec![Some(1.0); 4]],
        )
        .unwrap();
        assert_eq!(
            chart_series_count(
                ChartKind::MultiDataScatter2d,
                Direction::Columns,
                &data,
                "test",
                1
            )
            .unwrap(),
            1
        );
    }

    #[test]
    fn unsupported_direction_is_rejected() {
        let error = chart_series_count(
            ChartKind::Line2d,
            Direction::from_native(9_001),
            &data_with_rows(2),
            "test",
            1,
        )
        .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("unsupported series direction 9001")
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
