//! Layered typed authoring for standalone and embedded ODF charts.
//!
//! The content model and XML writer are owned by the common chart owner. This
//! facade keeps standalone-package construction local to ODC without
//! duplicating or re-owning chart semantics.

mod builder;

pub use builder::Builder;
pub use litchi_odf_common::chart::authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, DataLabelSpec, DataPointSpec,
    Definition, DomainSpec, EquationSpec, ExtensionAttribute, ExtensionElement, Extensions,
    GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec, StyleElement, Text, data,
    extensions, model, writer,
};
pub use litchi_odf_common::chart::authoring::{
    serialize_axis_fragment, serialize_content, serialize_series_fragment,
};
