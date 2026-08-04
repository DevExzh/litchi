//! OpenDocument standalone chart (`.odc` and `.otc`) support.

mod authoring;
mod document;
mod mutation;
mod semantic;

pub use authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, DataLabelSpec, DataPointSpec,
    Definition, DomainSpec, EquationSpec, ExtensionAttribute, ExtensionElement, Extensions,
    GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec, StyleElement, Text,
    serialize_chart_content,
};
pub use document::{Attribute, Document, Element, ElementKind};
pub use mutation::{AxisUpdate, SeriesUpdate};
pub use semantic::{
    Axis, DataPoint, DataSourceLabels, Dimension, Grid, GridClass, Legend, LegendPosition,
    PlotArea, Series,
};
