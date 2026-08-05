//! Host bridge for generating embedded Excel chart workbooks.
//!
//! Workbook OPC generation is owned by
//! `litchi_pptx::presentation::embedded::excel`. The legacy `ChartData` input
//! remains a host bridge because chart-part writers still construct that
//! value; conversion is limited to the typed data model and contains no XML
//! or OPC logic.

use crate::error::Result;
use crate::pptx::media_parts::map_pptx_error;
use crate::pptx::parts::chart::{ChartData, ChartType};
use litchi_pptx::presentation::embedded::excel::{Chart, Kind, Series, Workbook};

/// Generate a minimal embedded XLSX workbook from the host chart input.
pub fn generate_chart_excel_data(chart: &ChartData) -> Result<Vec<u8>> {
    let owner_chart = Chart {
        kind: chart_kind(chart.chart_type),
        title: chart.title.clone(),
        series: chart
            .series
            .iter()
            .map(|series| Series {
                name: series.name.clone(),
                values: series.values.clone(),
                categories: series.categories.clone(),
                x_values: series.x_values.clone(),
                bubble_sizes: series.bubble_sizes.clone(),
            })
            .collect(),
        show_legend: chart.show_legend,
        x: chart.x,
        y: chart.y,
        width: chart.width,
        height: chart.height,
    };

    Workbook::from_chart(&owner_chart)
        .map(Workbook::into_bytes)
        .map_err(map_pptx_error)
}

fn chart_kind(value: ChartType) -> Kind {
    match value {
        ChartType::Bar => Kind::Bar,
        ChartType::Column => Kind::Column,
        ChartType::Line => Kind::Line,
        ChartType::Pie => Kind::Pie,
        ChartType::Area => Kind::Area,
        ChartType::Scatter => Kind::Scatter,
        ChartType::Bubble => Kind::Bubble,
        ChartType::Doughnut => Kind::Doughnut,
        ChartType::Radar => Kind::Radar,
        ChartType::Surface => Kind::Surface,
        ChartType::Stock => Kind::Stock,
        ChartType::Unknown => Kind::Unknown,
    }
}
