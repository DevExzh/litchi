//! `OpenDocument` Chart support with semantic responsibility layers.
//!
//! [`Package`] is the immutable owner of one packaged chart (`.odc`/`.otc`)
//! and owns every state verb: `new`, `open*`, `from_bytes*`, `from_reader*`,
//! `from_prepared_package`, `edit`, `apply`, `save*`, `to_bytes` and
//! `write_to`. [`Chart`] is its read-only projection ([`Package::chart`]).
//! The primary transaction family is [`Edit`], [`Patch`], [`Commit`],
//! [`Diagnostics`], [`Error`] and [`LimitKind`]; [`Builder::build`] returns a
//! [`Package`].
#![forbid(unsafe_code)]

pub mod authoring;
mod codec;
pub mod definition;
mod facade;
pub mod flat;
mod limits;
mod merge;
pub mod metadata;
mod package;
mod project;
mod resource;
mod style_format;
mod validation;

pub use authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, ChartClassKind,
    DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, ExtensionAttribute,
    ExtensionElement, Extensions, GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec,
    StyleElement, Text, serialize_axis_fragment, serialize_content, serialize_series_fragment,
};
pub use definition::StyleTarget;
pub use facade::{
    Builder, Chart, Commit, Diagnostics, Edit, History, Package, PackageMerge, Patch,
    ResourceChange, StylesChange,
};
pub use flat::{AxisChange, AxisUpdate, ExactAttribute, ExactChange, ExactTarget};
pub(crate) use limits::Limits;
pub use litchi_odf_common::chart;
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
};
pub use merge::Conflict;
pub use package::ChartPackageKind;
pub use resource::Resource;
pub use style_format::{
    ChartErrorCategory, ChartRegressionMovingType, ChartRegressionType, ChartStyleProperty,
    ChartStylePropertyChange, ChartStyleValue, ChartSymbolName, ChartSymbolType,
};
pub use validation::{validate_formula, validate_range_list};
