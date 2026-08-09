//! `OpenDocument` standalone chart (`.odc` and `.otc`) support.

mod document;
mod mutation;

pub use document::Document;
pub use litchi_odf_common::chart::authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, ChartClassKind,
    DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, ExtensionAttribute,
    ExtensionElement, Extensions, GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec,
    StyleElement, Text, serialize_axis_fragment, serialize_content, serialize_series_fragment,
};
pub use litchi_odf_common::chart::{
    Attribute, Axis, Class, DataPoint, Dimension, Element, Grid, Kind, Labels, Legend, PlotArea,
    Position, Series,
};
pub use mutation::{AxisUpdate, SeriesUpdate};
