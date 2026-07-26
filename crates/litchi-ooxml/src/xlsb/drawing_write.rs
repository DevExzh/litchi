//! Typed authoring helpers for XLSB worksheet Drawings and Chart parts.
//!
//! XLSB reuses the ordinary SpreadsheetDrawing and DrawingML Chart XML
//! grammars. Only the worksheet link is binary (`BrtDrawing`).

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsx::WorksheetChart;

pub(crate) const MAX_CHARTS_PER_SHEET: usize = 4_096;
const MAX_DRAWING_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHART_XML_BYTES: usize = 16 * 1024 * 1024;

fn unsupported(detail: impl Into<String>) -> XlsbError {
    XlsbError::UnsupportedFeature(detail.into())
}

pub(crate) fn validate_chart(chart: &WorksheetChart) -> XlsbResult<()> {
    crate::xlsx::chart::validate_chart_anchor(&chart.anchor)?;
    if chart.chart.pivot_source.is_some() {
        return Err(unsupported(
            "XLSB worksheet chart authoring does not yet wire PivotTable views",
        ));
    }
    if chart.chart.external_data.is_some() || chart.external_data_part.is_some() {
        return Err(unsupported(
            "XLSB worksheet chart external-data relationships are not yet authored",
        ));
    }
    if chart.chart.user_shapes.is_some() || chart.user_shapes_part.is_some() {
        return Err(unsupported(
            "XLSB worksheet chart user-shapes relationships are not yet authored",
        ));
    }
    if !chart.additional_relationships.is_empty()
        || !crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)?.is_empty()
    {
        return Err(unsupported(
            "XLSB worksheet chart relationship-bearing extension fragments are not yet authored",
        ));
    }
    Ok(())
}

pub(crate) fn serialize_chart(chart: &WorksheetChart) -> XlsbResult<Vec<u8>> {
    validate_chart(chart)?;
    let xml = crate::xlsx::chart::generate_chart_xml(&chart.chart)?;
    if xml.len() > MAX_CHART_XML_BYTES {
        return Err(XlsbError::InvalidLength {
            expected: MAX_CHART_XML_BYTES,
            found: xml.len(),
        });
    }
    // Treat the shared chart reader as a post-serialization grammar oracle.
    crate::charts::reader::parse_chart(xml.as_slice())?;
    Ok(xml)
}

pub(crate) fn serialize_drawing(charts: &[WorksheetChart]) -> XlsbResult<Vec<u8>> {
    if charts.is_empty() || charts.len() > MAX_CHARTS_PER_SHEET {
        return Err(XlsbError::InvalidFormula(format!(
            "worksheet drawing requires 1 to {MAX_CHARTS_PER_SHEET} charts"
        )));
    }
    let mut xml = String::with_capacity(512 + charts.len() * 512);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
    );
    crate::xlsx::chart::write_worksheet_chart_anchors(&mut xml, charts, 0, 0)?;
    xml.push_str("</xdr:wsDr>");
    if xml.len() > MAX_DRAWING_XML_BYTES {
        return Err(XlsbError::InvalidLength {
            expected: MAX_DRAWING_XML_BYTES,
            found: xml.len(),
        });
    }
    let bytes = xml.into_bytes();
    // The XLSB drawing inventory reader is the package-load oracle.
    crate::xlsb::drawing::parse_drawing_part(&bytes)?;
    Ok(bytes)
}
