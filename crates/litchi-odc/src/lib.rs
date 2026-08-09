//! `OpenDocument` Chart support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::missing_errors_doc,
    clippy::needless_pass_by_value,
    clippy::option_option,
    clippy::struct_field_names,
    reason = "the chart facade uses specification-shaped values and centralizes its typed failure contract"
)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::shadow_reuse,
    reason = "chart codecs follow ODF document order and reuse short-lived XML event bindings"
)]

pub mod authoring;
mod codec;
mod facade;
pub mod flat;
mod limits;
mod package;
mod resource;
mod transaction;
mod validation;

pub use authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, ChartClassKind,
    DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, ExtensionAttribute,
    ExtensionElement, Extensions, GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec,
    StyleElement, Text, serialize_axis_fragment, serialize_content, serialize_content_with_limits,
    serialize_series_fragment,
};
pub use facade::{Builder, Chart, Commit, Edit, Patch, ResourceChange, StylesChange};
pub use flat::{AxisChange, AxisUpdate, FlatChart, FlatChartCommit, FlatChartEdit, FlatChartPatch};
pub use limits::Limits;
pub use litchi_odf_common::chart;
pub use resource::Resource;
pub use transaction::{
    DefinitionChange, DefinitionCommit, DefinitionEdit, DefinitionHistory, DefinitionPatch,
    DefinitionSnapshot, StyleTarget,
};
pub use validation::{validate_formula, validate_range_list};
