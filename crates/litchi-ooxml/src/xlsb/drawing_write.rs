//! Typed authoring helpers for XLSB worksheet Drawings, Image, and Chart parts.
//!
//! XLSB reuses the ordinary SpreadsheetDrawing and DrawingML Chart XML
//! grammars. Only the worksheet link is binary (`BrtDrawing`).

use crate::xlsb::XlsbWorksheetImage;
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsx::WorksheetChart;
use crate::xlsx::writer::shape::{
    ShapeEmitter, XlsxConnectionShapeSpec, XlsxGroupSpec, XlsxShapeSpec,
};
use litchi_core::xml::escape_xml;
use std::fmt::Write as _;

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

pub(crate) fn serialize_drawing(
    images: &[XlsbWorksheetImage],
    charts: &[WorksheetChart],
    shapes: &[XlsxShapeSpec],
    groups: &[XlsxGroupSpec],
    connections: &[XlsxConnectionShapeSpec],
) -> XlsbResult<Vec<u8>> {
    if images.is_empty()
        && charts.is_empty()
        && shapes.is_empty()
        && groups.is_empty()
        && connections.is_empty()
    {
        return Err(XlsbError::InvalidFormula(
            "worksheet drawing requires at least one drawing object".to_string(),
        ));
    }
    if images.len() > crate::xlsb::drawing_image::MAX_XLSB_WORKSHEET_IMAGES
        || charts.len() > MAX_CHARTS_PER_SHEET
    {
        return Err(XlsbError::InvalidFormula(format!(
            "worksheet drawing exceeds the {} image or {MAX_CHARTS_PER_SHEET} chart safety limit",
            crate::xlsb::drawing_image::MAX_XLSB_WORKSHEET_IMAGES
        )));
    }
    let mut total_image_bytes = 0usize;
    for image in images {
        image.validate()?;
        total_image_bytes =
            total_image_bytes
                .checked_add(image.data().len())
                .ok_or(XlsbError::InvalidLength {
                    expected: crate::xlsb::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                    found: usize::MAX,
                })?;
        if total_image_bytes > crate::xlsb::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES {
            return Err(XlsbError::InvalidLength {
                expected: crate::xlsb::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                found: total_image_bytes,
            });
        }
    }

    let mut xml = String::with_capacity(512 + (images.len() + charts.len()) * 512);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
    );
    for (index, image) in images.iter().enumerate() {
        write_image_anchor(&mut xml, image, index)?;
        ensure_drawing_size(xml.len())?;
    }
    crate::xlsx::chart::write_worksheet_chart_anchors(
        &mut xml,
        charts,
        images.len(),
        images.len(),
    )?;
    let mut object_count = 0usize;
    for shape in shapes {
        shape
            .validate(object_count)
            .map_err(XlsbError::InvalidFormula)?;
        object_count += 1;
    }
    for group in groups {
        group
            .validate(object_count)
            .map_err(XlsbError::InvalidFormula)?;
        object_count += 1;
    }
    for connection in connections {
        connection
            .validate(object_count)
            .map_err(XlsbError::InvalidFormula)?;
        object_count += 1;
    }
    let first_shape_id = images
        .len()
        .checked_add(charts.len())
        .and_then(|count| count.checked_add(1))
        .and_then(|id| u32::try_from(id).ok())
        .ok_or_else(|| {
            XlsbError::InvalidFormula("worksheet drawing object ID overflow".to_string())
        })?;
    let mut emitter = ShapeEmitter::for_objects(first_shape_id, shapes, groups, connections)
        .map_err(XlsbError::InvalidFormula)?;
    for shape in shapes {
        emitter
            .write_anchored_shape(&mut xml, shape)
            .map_err(XlsbError::InvalidFormula)?;
        ensure_drawing_size(xml.len())?;
    }
    for group in groups {
        emitter
            .write_anchored_group(&mut xml, group)
            .map_err(XlsbError::InvalidFormula)?;
        ensure_drawing_size(xml.len())?;
    }
    for connection in connections {
        emitter
            .write_anchored_connection(&mut xml, connection)
            .map_err(XlsbError::InvalidFormula)?;
        ensure_drawing_size(xml.len())?;
    }
    xml.push_str("</xdr:wsDr>");
    ensure_drawing_size(xml.len())?;
    // The detailed shared reader verifies the complete shape/group/connector
    // grammar in addition to the lightweight XLSB drawing inventory below.
    crate::xlsx::shapes::parse_drawing_shapes(&xml)?.ok_or_else(|| {
        XlsbError::Encoding("authored drawing lacks an xdr:wsDr root".to_string())
    })?;
    let bytes = xml.into_bytes();
    // The XLSB drawing inventory reader is the package-load oracle.
    crate::xlsb::drawing::parse_drawing_part(&bytes)?;
    Ok(bytes)
}

fn ensure_drawing_size(bytes: usize) -> XlsbResult<()> {
    if bytes > MAX_DRAWING_XML_BYTES {
        return Err(XlsbError::InvalidLength {
            expected: MAX_DRAWING_XML_BYTES,
            found: bytes,
        });
    }
    Ok(())
}

fn write_image_anchor(
    xml: &mut String,
    image: &XlsbWorksheetImage,
    index: usize,
) -> XlsbResult<()> {
    let anchor = image.anchor();
    let object_id = index.checked_add(1).ok_or_else(|| {
        XlsbError::InvalidFormula("worksheet picture object ID overflow".to_string())
    })?;
    xml.push_str("<xdr:twoCellAnchor>");
    write!(
        xml,
        "<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>",
        anchor.from_col,
        anchor.from_col_offset,
        anchor.from_row,
        anchor.from_row_offset
    )
    .map_err(|error| XlsbError::Encoding(error.to_string()))?;
    write!(
        xml,
        "<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>",
        anchor.to_col,
        anchor.to_col_offset,
        anchor.to_row,
        anchor.to_row_offset
    )
    .map_err(|error| XlsbError::Encoding(error.to_string()))?;
    write!(
        xml,
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{object_id}" name="Picture {object_id}""#
    )
    .map_err(|error| XlsbError::Encoding(error.to_string()))?;
    if let Some(description) = image.description() {
        write!(xml, r#" descr="{}""#, escape_xml(description))
            .map_err(|error| XlsbError::Encoding(error.to_string()))?;
    }
    write!(
        xml,
        r#"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId{object_id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#
    )
    .map_err(|error| XlsbError::Encoding(error.to_string()))?;
    Ok(())
}
