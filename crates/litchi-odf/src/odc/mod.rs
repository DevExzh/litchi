//! OpenDocument standalone chart (`.odc` and `.otc`) support.

mod document;
mod semantic;

pub use document::{ChartAttribute, ChartDocument, ChartElement, ChartElementKind};
pub use semantic::{
    ChartAxis, ChartAxisDimension, ChartDataPoint, ChartDataSourceLabels, ChartGrid,
    ChartGridClass, ChartLegend, ChartLegendPosition, ChartPlotArea, ChartSeries,
};
