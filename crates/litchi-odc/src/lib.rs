//! `OpenDocument` Chart support with semantic responsibility layers.
#![forbid(unsafe_code)]

pub mod authoring;
mod codec;
mod facade;
pub mod flat;
mod limits;
mod merge;
mod package;
mod project;
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
pub use facade::{
    Builder, Chart, Commit, Edit, History, PackageMerge, Patch, ResourceChange, StylesChange,
};
pub use flat::{AxisChange, AxisUpdate, FlatChart, FlatChartCommit, FlatChartEdit, FlatChartPatch};
pub use limits::Limits;
pub use litchi_odf_common::chart;
pub use merge::{Conflict, DefinitionMerge};
pub use resource::Resource;
pub use transaction::{
    DefinitionChange, DefinitionCommit, DefinitionEdit, DefinitionHistory, DefinitionPatch,
    DefinitionSnapshot, StyleTarget,
};
pub use validation::{validate_formula, validate_range_list};
