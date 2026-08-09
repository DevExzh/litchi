//! `OpenDocument` Chart support with semantic responsibility layers.
#![forbid(unsafe_code)]

pub mod authoring;
mod codec;
mod facade;
pub mod flat;
mod package;

pub use authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, ChartClassKind,
    DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, ExtensionAttribute,
    ExtensionElement, Extensions, GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec,
    StyleElement, Text, serialize_axis_fragment, serialize_content, serialize_series_fragment,
};
pub use facade::{Builder, Chart, Commit, Edit, Patch};
pub use flat::{AxisChange, AxisUpdate, FlatChart, FlatChartCommit, FlatChartEdit, FlatChartPatch};
pub use litchi_odf_common::chart;
