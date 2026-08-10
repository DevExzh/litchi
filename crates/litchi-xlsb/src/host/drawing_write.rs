//! Typed authoring helpers for XLSB worksheet Drawings, Image, and Chart parts.
//!
//! XLSB reuses the ordinary SpreadsheetDrawing and DrawingML Chart XML
//! grammars. Only the worksheet link is binary (`BrtDrawing`).

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "DrawingML authoring uses checked object IDs, offsets, sizes, and relationship identities"
)]

use crate::chart::Chart;
use crate::package::error::{Error, Result};
use crate::writer::Image;
use crate::writer::shape::{ConnectionShapeSpec, Emitter, GroupSpec, ShapeSpec};
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use std::fmt::Write as _;

pub(crate) const MAX_CHARTS_PER_SHEET: usize = 4_096;
const MAX_DRAWING_XML_BYTES: usize = 16 * 1024 * 1024;
const CHART_SHEET_EXTENT_X: u64 = 8_582_025;
const CHART_SHEET_EXTENT_Y: u64 = 5_838_825;

pub(crate) fn validate_chart(chart: &Chart) -> Result<()> {
    crate::chart::anchor::validate(&chart.anchor)?;
    crate::package::chart_resources::validate_chart_resources(chart)
}

/// Serialize the single absolute-anchored chart frame used by a chart sheet.
pub(crate) fn serialize_chart_sheet_drawing(title: &str) -> Result<Vec<u8>> {
    if title.encode_utf16().count() > 32_767 || title.contains('\0') {
        return Err(Error::InvalidFormula(
            "chart-sheet drawing title is too long or contains NUL".to_string(),
        ));
    }
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="{CHART_SHEET_EXTENT_X}" cy="{CHART_SHEET_EXTENT_Y}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="{}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#,
        escape_xml(title)
    );
    if xml.len() > MAX_DRAWING_XML_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_DRAWING_XML_BYTES,
            found: xml.len(),
        });
    }
    crate::package::drawing::parse_drawing_part(xml.as_bytes())?;
    Ok(xml.into_bytes())
}

pub(crate) fn serialize_drawing(
    images: &[Image],
    charts: &[Chart],
    shapes: &[ShapeSpec],
    groups: &[GroupSpec],
    connections: &[ConnectionShapeSpec],
) -> Result<Vec<u8>> {
    if images.is_empty()
        && charts.is_empty()
        && shapes.is_empty()
        && groups.is_empty()
        && connections.is_empty()
    {
        return Err(Error::InvalidFormula(
            "worksheet drawing requires at least one drawing object".to_string(),
        ));
    }
    if images.len() > crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGES
        || charts.len() > MAX_CHARTS_PER_SHEET
    {
        return Err(Error::InvalidFormula(format!(
            "worksheet drawing exceeds the {} image or {MAX_CHARTS_PER_SHEET} chart safety limit",
            crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGES
        )));
    }
    let mut total_image_bytes = 0usize;
    for image in images {
        image.validate()?;
        total_image_bytes =
            total_image_bytes
                .checked_add(image.data().len())
                .ok_or(Error::InvalidLength {
                    expected: crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                    found: usize::MAX,
                })?;
        if total_image_bytes > crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES {
            return Err(Error::InvalidLength {
                expected: crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
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
    crate::chart::anchor::write_all(&mut xml, charts, images.len(), images.len())?;
    let first_shape_id = images
        .len()
        .checked_add(charts.len())
        .and_then(|count| count.checked_add(1))
        .and_then(|id| u32::try_from(id).ok())
        .ok_or_else(|| Error::InvalidFormula("worksheet drawing object ID overflow".to_string()))?;
    let mut emitter = Emitter::for_objects(first_shape_id, shapes, groups, connections)?;
    for shape in shapes {
        emitter.write_anchored_shape(&mut xml, shape)?;
        ensure_drawing_size(xml.len())?;
    }
    for group in groups {
        emitter.write_anchored_group(&mut xml, group)?;
        ensure_drawing_size(xml.len())?;
    }
    for connection in connections {
        emitter.write_anchored_connection(&mut xml, connection)?;
        ensure_drawing_size(xml.len())?;
    }
    xml.push_str("</xdr:wsDr>");
    ensure_drawing_size(xml.len())?;
    // The detailed shared reader verifies the complete shape/group/connector
    // grammar in addition to the lightweight XLSB drawing inventory below.
    crate::shapes::read(&xml)?
        .ok_or_else(|| Error::Encoding("authored drawing lacks an xdr:wsDr root".to_string()))?;
    let bytes = xml.into_bytes();
    // The XLSB drawing inventory reader is the package-load oracle.
    crate::package::drawing::parse_drawing_part(&bytes)?;
    Ok(bytes)
}

fn ensure_drawing_size(bytes: usize) -> Result<()> {
    if bytes > MAX_DRAWING_XML_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_DRAWING_XML_BYTES,
            found: bytes,
        });
    }
    Ok(())
}

fn write_image_anchor(xml: &mut String, image: &Image, index: usize) -> Result<()> {
    let object_id = index
        .checked_add(1)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| Error::InvalidFormula("worksheet picture object ID overflow".to_string()))?;
    write_image_anchor_with_ids(xml, image, object_id, &format!("rId{object_id}"), false)
}

pub(crate) fn append_image_anchor(
    source: &[u8],
    image: &Image,
    object_id: u32,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    image.validate()?;
    let _inventory = crate::package::drawing::parse_drawing_part(source)?;
    let root_end = drawing_root_end(source)?;
    let mut anchor = String::new();
    write_image_anchor_with_ids(&mut anchor, image, object_id, relationship_id, true)?;
    let output_len = source
        .len()
        .checked_add(anchor.len())
        .ok_or(Error::InvalidLength {
            expected: MAX_DRAWING_XML_BYTES,
            found: usize::MAX,
        })?;
    ensure_drawing_size(output_len)?;
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&source[..root_end]);
    output.extend_from_slice(anchor.as_bytes());
    output.extend_from_slice(&source[root_end..]);
    let _inventory = crate::package::drawing::parse_drawing_part(&output)?;
    let xml = std::str::from_utf8(&output)
        .map_err(|error| Error::Encoding(format!("drawing XML is not UTF-8: {error}")))?;
    crate::shapes::read(xml)?
        .ok_or_else(|| Error::Encoding("appended drawing lacks an xdr:wsDr root".to_string()))?;
    Ok(output)
}

pub(crate) fn next_drawing_object_id(source: &[u8]) -> Result<u32> {
    let _inventory = crate::package::drawing::parse_drawing_part(source)?;
    let mut reader = quick_xml::Reader::from_reader(source);
    let mut maximum = 0u32;
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Encoding(format!("invalid drawing XML: {error}")))?
        {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                if let Some(value) = unqualified_attribute_value(&element, b"id", reader.decoder())?
                {
                    let id = value.parse::<u32>().map_err(|error| {
                        Error::Encoding(format!("drawing object ID is not u32: {error}"))
                    })?;
                    maximum = maximum.max(id);
                }
            },
            quick_xml::events::Event::Eof => break,
            quick_xml::events::Event::Start(_)
            | quick_xml::events::Event::Empty(_)
            | quick_xml::events::Event::End(_)
            | quick_xml::events::Event::Text(_)
            | quick_xml::events::Event::CData(_)
            | quick_xml::events::Event::Comment(_)
            | quick_xml::events::Event::Decl(_)
            | quick_xml::events::Event::PI(_)
            | quick_xml::events::Event::DocType(_)
            | quick_xml::events::Event::GeneralRef(_) => {},
        }
    }
    maximum
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormula("drawing object ID overflow".to_string()))
}

fn drawing_root_end(source: &[u8]) -> Result<usize> {
    let mut reader = quick_xml::Reader::from_reader(source);
    let mut depth = 0usize;
    loop {
        let position = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Encoding("drawing XML position exceeds usize".to_string()))?;
        match reader
            .read_event()
            .map_err(|error| Error::Encoding(format!("invalid drawing XML: {error}")))?
        {
            quick_xml::events::Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::Encoding("drawing XML nesting depth overflow".to_string())
                })?;
            },
            quick_xml::events::Event::End(element) => {
                if depth == 1 && element.local_name().as_ref() == b"wsDr" {
                    return Ok(position);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Encoding("drawing XML has an unmatched end element".to_string())
                })?;
            },
            quick_xml::events::Event::Eof => {
                return Err(Error::Encoding(
                    "drawing XML has no closing wsDr element".to_string(),
                ));
            },
            quick_xml::events::Event::Empty(_)
            | quick_xml::events::Event::Text(_)
            | quick_xml::events::Event::CData(_)
            | quick_xml::events::Event::Comment(_)
            | quick_xml::events::Event::Decl(_)
            | quick_xml::events::Event::PI(_)
            | quick_xml::events::Event::DocType(_)
            | quick_xml::events::Event::GeneralRef(_) => {},
        }
    }
}

fn write_image_anchor_with_ids(
    xml: &mut String,
    image: &Image,
    object_id: u32,
    relationship_id: &str,
    declare_namespaces: bool,
) -> Result<()> {
    let anchor = image.anchor();
    if declare_namespaces {
        xml.push_str(
            r#"<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
        );
    } else {
        xml.push_str("<xdr:twoCellAnchor>");
    }
    write!(
        xml,
        "<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>",
        anchor.from_col,
        anchor.from_col_offset,
        anchor.from_row,
        anchor.from_row_offset
    )
    .map_err(|error| Error::Encoding(error.to_string()))?;
    write!(
        xml,
        "<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>",
        anchor.to_col,
        anchor.to_col_offset,
        anchor.to_row,
        anchor.to_row_offset
    )
    .map_err(|error| Error::Encoding(error.to_string()))?;
    write!(
        xml,
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{object_id}" name="Picture {object_id}""#
    )
    .map_err(|error| Error::Encoding(error.to_string()))?;
    if let Some(description) = image.description() {
        write!(xml, r#" descr="{}""#, escape_xml(description))
            .map_err(|error| Error::Encoding(error.to_string()))?;
    }
    write!(
        xml,
        r#"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#,
        escape_xml(relationship_id)
    )
    .map_err(|error| Error::Encoding(error.to_string()))?;
    Ok(())
}
