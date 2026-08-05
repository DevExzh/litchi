//! OpenDocument standalone chart (`.odc` and `.otc`) support.

mod authoring;
mod document;
mod mutation;

/// Shared chart vocabulary and borrowed views used by standalone and embedded
/// ODF chart owners.
pub use litchi_odf_common::chart;

pub use authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, DataLabelSpec, DataPointSpec,
    Definition, DomainSpec, EquationSpec, ExtensionAttribute, ExtensionElement, Extensions,
    GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec, StyleElement, Text,
    serialize_chart_content,
};
pub use chart::{
    Attribute, Axis, Class, DataPoint, Dimension, Element, Grid, Kind, Labels, Legend, PlotArea,
    Position, Series,
};
pub use document::Document;
pub use mutation::{AxisUpdate, SeriesUpdate};
