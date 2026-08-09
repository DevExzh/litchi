//! Chart presentation and layout record families.

use super::super::validation::invalid_chart_input;
use super::super::xml::{write_bool_element, write_fragment};
use crate::chart::data::{Layout, TitleText};
use crate::chart::model::{
    ExtensionList, PictureOptions, ShapeProperties, TextProperties, View3D, WallFloor,
};
use crate::chart::series::Marker;
use crate::chart::types::MarkerStyle;
use litchi_core::xml::escape_xml;
use std::io::Write;

pub(super) fn write_title<W: Write>(
    writer: &mut W,
    title: &TitleText,
    layout: Option<&Layout>,
    overlay: bool,
    shape_properties: Option<&ShapeProperties>,
    text_properties: Option<&TextProperties>,
    extension_list: Option<&ExtensionList>,
) -> std::io::Result<()> {
    write!(writer, "<c:title>")?;

    write_title_text(writer, title)?;

    if let Some(layout) = layout {
        write_layout(writer, Some(layout))?;
    }

    write!(
        writer,
        r#"<c:overlay val="{}"/>"#,
        if overlay { "1" } else { "0" }
    )?;
    if let Some(shape_properties) = shape_properties {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = text_properties {
        write_fragment(writer, text_properties.as_xml())?;
    }
    if let Some(extension_list) = extension_list {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:title>")?;

    Ok(())
}

pub(super) fn write_marker<W: Write>(
    writer: &mut W,
    marker: &Marker,
    description: &str,
) -> std::io::Result<()> {
    write_marker_parts(
        writer,
        marker.symbol,
        marker.size,
        marker.shape_properties.as_ref(),
        marker.extension_list.as_ref(),
        description,
    )
}

pub(super) fn write_marker_parts<W: Write>(
    writer: &mut W,
    symbol: Option<MarkerStyle>,
    size: Option<u32>,
    shape_properties: Option<&ShapeProperties>,
    extension_list: Option<&ExtensionList>,
    description: &str,
) -> std::io::Result<()> {
    if size.is_some_and(|size| !(2..=72).contains(&size)) {
        return Err(invalid_chart_input(format!(
            "{description} marker size must be 2-72"
        )));
    }
    write!(writer, "<c:marker>")?;
    if let Some(symbol) = symbol {
        write!(writer, r#"<c:symbol val="{}"/>"#, symbol.xml_value())?;
    }
    if let Some(size) = size {
        write!(writer, r#"<c:size val="{size}"/>"#)?;
    }
    if let Some(shape_properties) = shape_properties {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(extension_list) = extension_list {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:marker>")?;
    Ok(())
}

pub(super) fn write_title_text<W: Write>(writer: &mut W, title: &TitleText) -> std::io::Result<()> {
    match title {
        TitleText::Literal(rich_text) => {
            write!(writer, "<c:tx><c:rich>")?;
            write!(writer, "<a:bodyPr/><a:lstStyle/>")?;
            write!(writer, "<a:p><a:pPr><a:defRPr/></a:pPr>")?;
            write!(
                writer,
                r#"<a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r>"#,
                escape_xml(&rich_text.text)
            )?;
            write!(writer, "</a:p></c:rich></c:tx>")?;
        },
        TitleText::Reference(source_ref) => {
            write!(writer, "<c:tx><c:strRef>")?;
            write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;
            write!(writer, "</c:strRef></c:tx>")?;
        },
    }
    Ok(())
}

pub(super) fn write_view_3d<W: Write>(writer: &mut W, view: &View3D) -> std::io::Result<()> {
    write!(writer, "<c:view3D>")?;

    if let Some(rot_x) = view.rot_x {
        write!(writer, r#"<c:rotX val="{rot_x}"/>"#)?;
    }
    if let Some(rot_y) = view.rot_y {
        write!(writer, r#"<c:rotY val="{rot_y}"/>"#)?;
    }

    write!(
        writer,
        r#"<c:rAngAx val="{}"/>"#,
        if view.right_angle_axes { "1" } else { "0" }
    )?;

    if let Some(perspective) = view.perspective {
        write!(writer, r#"<c:perspective val="{perspective}"/>"#)?;
    }
    if let Some(height) = view.height_percent {
        write!(writer, r#"<c:hPercent val="{height}"/>"#)?;
    }
    if let Some(depth) = view.depth_percent {
        write!(writer, r#"<c:depthPercent val="{depth}"/>"#)?;
    }

    write!(writer, "</c:view3D>")?;

    Ok(())
}

pub(super) fn write_wall_floor<W: Write>(
    writer: &mut W,
    wall_floor: &WallFloor,
) -> std::io::Result<()> {
    if let Some(thickness) = wall_floor.thickness {
        write!(writer, r#"<c:thickness val="{thickness}"/>"#)?;
    }
    if let Some(shape_properties) = wall_floor.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(options) = wall_floor.picture_options.as_ref() {
        write_picture_options(writer, options)?;
    }
    if let Some(extension_list) = wall_floor.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    Ok(())
}

pub(super) fn write_picture_options<W: Write>(
    writer: &mut W,
    options: &PictureOptions,
) -> std::io::Result<()> {
    write!(writer, "<c:pictureOptions>")?;
    for (name, value) in [
        ("applyToFront", options.apply_to_front),
        ("applyToSides", options.apply_to_sides),
        ("applyToEnd", options.apply_to_end),
    ] {
        if let Some(value) = value {
            write_bool_element(writer, name, value)?;
        }
    }
    if let Some(format) = options.picture_format {
        write!(writer, r#"<c:pictureFormat val="{}"/>"#, format.xml_value())?;
    }
    if let Some(unit) = options.picture_stack_unit {
        if !unit.is_finite() || unit <= 0.0 {
            return Err(invalid_chart_input(
                "chart picture stack unit must be finite and positive",
            ));
        }
        write!(writer, r#"<c:pictureStackUnit val="{unit}"/>"#)?;
    }
    write!(writer, "</c:pictureOptions>")?;
    Ok(())
}

pub(super) fn write_layout<W: Write>(
    writer: &mut W,
    layout: Option<&Layout>,
) -> std::io::Result<()> {
    let Some(layout) = layout else {
        return write!(writer, "<c:layout/>");
    };
    for (name, value) in [
        ("x", layout.x),
        ("y", layout.y),
        ("width", layout.width),
        ("height", layout.height),
    ] {
        if value.is_some_and(|value| !value.is_finite()) {
            return Err(invalid_chart_input(format!(
                "chart layout {name} must be finite"
            )));
        }
    }
    write!(writer, "<c:layout><c:manualLayout>")?;
    if let Some(target) = layout.target {
        write!(writer, r#"<c:layoutTarget val="{}"/>"#, target.xml_value())?;
    }
    for (name, mode) in [
        ("xMode", layout.x_mode),
        ("yMode", layout.y_mode),
        ("wMode", layout.width_mode),
        ("hMode", layout.height_mode),
    ] {
        if let Some(mode) = mode {
            write!(writer, r#"<c:{name} val="{}"/>"#, mode.xml_value())?;
        }
    }
    for (name, value) in [
        ("x", layout.x),
        ("y", layout.y),
        ("w", layout.width),
        ("h", layout.height),
    ] {
        if let Some(value) = value {
            write!(writer, r#"<c:{name} val="{value}"/>"#)?;
        }
    }
    write!(writer, "</c:manualLayout></c:layout>")?;
    Ok(())
}
