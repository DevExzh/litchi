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
pub use litchi_odf_common::chart::authoring::{serialize_axis_fragment, serialize_series_fragment};
pub use litchi_odf_common::chart::{ChartClass, ChartClassKind};

use litchi_core::Result;

/// Serialize a definition using the default ODC limits.
///
/// # Errors
///
/// Returns an error when validation or XML serialization fails.
pub fn serialize_content(definition: &Definition) -> Result<String> {
    serialize_content_with_limits(definition, crate::Limits::default())
}

/// Serialize a definition after ODF range/formula and caller-limit checks.
///
/// # Errors
///
/// Returns an error when validation, serialization, or compactness checks fail.
pub fn serialize_content_with_limits(
    definition: &Definition,
    limits: crate::Limits,
) -> Result<String> {
    crate::validation::validate_definition(definition, limits)?;
    let content = litchi_odf_common::chart::authoring::serialize_content(definition)?.replacen(
        "office:version=\"1.2\"",
        "office:version=\"1.4\"",
        1,
    );
    let compact_limits =
        litchi_odf_common::compact_xml::Limits::new(limits.max_content_bytes(), limits.max_depth())
            .map_err(litchi_core::Error::from)?;
    litchi_odf_common::compact_xml::validate_with_limits(content.as_bytes(), compact_limits)
        .map_err(litchi_core::Error::from)?;
    Ok(content)
}
