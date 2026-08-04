//! OpenDocument standalone chart (`.odc` and `.otc`) support.

mod authoring;
mod document;
mod mutation;
mod semantic;

pub use authoring::{
    ChartAxisSpec, ChartCachedCell, ChartCachedRow, ChartCachedTable, ChartCachedValue,
    ChartDataLabelSpec, ChartDataPointSpec, ChartDefinition, ChartDomainSpec, ChartEquationSpec,
    ChartExtensionAttribute, ChartExtensionElement, ChartExtensions, ChartGridSpec,
    ChartLegendSpec, ChartPlotAreaSpec, ChartRegressionSpec, ChartSeriesSpec, ChartStyleElement,
    ChartText, serialize_chart_content,
};
pub use document::{ChartAttribute, ChartDocument, ChartElement, ChartElementKind};
pub use mutation::{ChartAxisUpdate, ChartSeriesUpdate};
pub use semantic::{
    ChartAxis, ChartAxisDimension, ChartDataPoint, ChartDataSourceLabels, ChartGrid,
    ChartGridClass, ChartLegend, ChartLegendPosition, ChartPlotArea, ChartSeries,
};
