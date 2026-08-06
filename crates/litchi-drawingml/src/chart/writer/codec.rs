//! XML/wire codec façade for DrawingML charts.
//!
//! Record-family serialization lives under [`super::semantic`]. This module
//! keeps the public writer contract small and stable while package owners can
//! still provide relationship identifiers through [`write_with_rels`].

use super::semantic::write_chart_space;
use crate::chart::model::Chart;
use std::io::Write;

/// Write a chart to canonical DrawingML chart XML.
pub fn write<W: Write>(writer: &mut W, chart: &Chart) -> std::io::Result<()> {
    write_with_rels(writer, chart, None, None)
}

/// Write chart XML using relationship identifiers assigned by its package owner.
///
/// The chart model owns relationship metadata, while DOCX, PPTX, XLSX, and
/// XLSB package writers remain responsible for allocating and validating OPC
/// relationships. This entry point only serializes the resulting wire form.
pub fn write_with_rels<W: Write>(
    writer: &mut W,
    chart: &Chart,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> std::io::Result<()> {
    write_chart_space(
        writer,
        chart,
        external_data_relationship_id,
        user_shapes_relationship_id,
    )
}
