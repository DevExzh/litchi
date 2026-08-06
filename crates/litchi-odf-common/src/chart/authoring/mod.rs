//! Layered typed authoring for ODF chart content shared by standalone and
//! embedded chart hosts.
//!
//! Package ownership remains with the concrete family crate. This owner only
//! validates and serializes the chart content grammar, so ODT can author an
//! embedded chart without depending on the peer ODC family crate.

pub mod data;
pub mod extensions;
pub mod model;
pub mod writer;

pub use data::{CachedCell, CachedRow, CachedTable, CachedValue};
pub use extensions::{ExtensionAttribute, ExtensionElement, Extensions};
pub use model::{
    AxisSpec, DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, GridSpec,
    LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec, StyleElement, Text,
};
pub use writer::{serialize_axis_fragment, serialize_content, serialize_series_fragment};
