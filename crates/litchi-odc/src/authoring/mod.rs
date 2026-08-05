//! Layered typed authoring for standalone and embedded ODF charts.

pub mod data;
pub mod extensions;
pub mod model;
pub mod writer;

mod builder;

pub use builder::Builder;
pub use data::{CachedCell, CachedRow, CachedTable, CachedValue};
pub use extensions::{ExtensionAttribute, ExtensionElement, Extensions};
pub use model::{
    AxisSpec, DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, GridSpec,
    LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec, StyleElement, Text,
};
pub use writer::{serialize_axis_fragment, serialize_content, serialize_series_fragment};
