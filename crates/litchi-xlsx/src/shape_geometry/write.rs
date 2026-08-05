//! Serializer for [`CustomGeometry`] into `a:custGeom` markup.
//!
//! Output is canonical: optional child lists are omitted when empty,
//! attributes at their ECMA-376 defaults are omitted, and everything written
//! here parses back through [`super::shapes::parse_drawing_shapes`]
//! into an identical [`CustomGeometry`].

use std::fmt::Write as _;

use litchi_core::xml::escape::escape_xml;

use super::{
    AdjustHandle, AdjustValue, CustomGeometry, Guide, Path, PathCommand, PathFillMode, Point,
};

/// Serialize one custom geometry as its `a:custGeom` element.
pub(crate) fn write_custom_geometry(xml: &mut String, geometry: &CustomGeometry) {
    xml.push_str("<a:custGeom>");
    write_guide_list(xml, "avLst", &geometry.adjust_values);
    write_guide_list(xml, "gdLst", &geometry.guides);
    if !geometry.adjust_handles.is_empty() {
        xml.push_str("<a:ahLst>");
        for handle in &geometry.adjust_handles {
            write_adjust_handle(xml, handle);
        }
        xml.push_str("</a:ahLst>");
    }
    if !geometry.connection_sites.is_empty() {
        xml.push_str("<a:cxnLst>");
        for site in &geometry.connection_sites {
            let _ = write!(xml, r#"<a:cxn ang="{}">"#, value_attribute(&site.angle));
            write_position(xml, &site.position);
            xml.push_str("</a:cxn>");
        }
        xml.push_str("</a:cxnLst>");
    }
    if let Some(rectangle) = &geometry.text_rectangle {
        let _ = write!(
            xml,
            r#"<a:rect l="{}" t="{}" r="{}" b="{}"/>"#,
            value_attribute(&rectangle.left),
            value_attribute(&rectangle.top),
            value_attribute(&rectangle.right),
            value_attribute(&rectangle.bottom)
        );
    }
    xml.push_str("<a:pathLst>");
    for path in &geometry.paths {
        write_path(xml, path);
    }
    xml.push_str("</a:pathLst></a:custGeom>");
}

fn write_guide_list(xml: &mut String, tag: &str, guides: &[Guide]) {
    if guides.is_empty() {
        return;
    }
    let _ = write!(xml, "<a:{tag}>");
    for guide in guides {
        let _ = write!(
            xml,
            r#"<a:gd name="{}" fmla="{}"/>"#,
            escape_xml(&guide.name),
            escape_xml(&guide.formula.to_string())
        );
    }
    let _ = write!(xml, "</a:{tag}>");
}

fn write_adjust_handle(xml: &mut String, handle: &AdjustHandle) {
    match handle {
        AdjustHandle::Xy(handle) => {
            xml.push_str("<a:ahXY");
            write_guide_reference(xml, "gdRefX", &handle.horizontal_guide);
            write_optional_value(xml, "minX", &handle.minimum_x);
            write_optional_value(xml, "maxX", &handle.maximum_x);
            write_guide_reference(xml, "gdRefY", &handle.vertical_guide);
            write_optional_value(xml, "minY", &handle.minimum_y);
            write_optional_value(xml, "maxY", &handle.maximum_y);
            xml.push('>');
            write_position(xml, &handle.position);
            xml.push_str("</a:ahXY>");
        },
        AdjustHandle::Polar(handle) => {
            xml.push_str("<a:ahPolar");
            write_guide_reference(xml, "gdRefR", &handle.radius_guide);
            write_optional_value(xml, "minR", &handle.minimum_radius);
            write_optional_value(xml, "maxR", &handle.maximum_radius);
            write_guide_reference(xml, "gdRefAng", &handle.angle_guide);
            write_optional_value(xml, "minAng", &handle.minimum_angle);
            write_optional_value(xml, "maxAng", &handle.maximum_angle);
            xml.push('>');
            write_position(xml, &handle.position);
            xml.push_str("</a:ahPolar>");
        },
    }
}

fn write_path(xml: &mut String, path: &Path) {
    let defaults = Path::default();
    xml.push_str("<a:path");
    if path.width != defaults.width {
        let _ = write!(xml, r#" w="{}""#, path.width);
    }
    if path.height != defaults.height {
        let _ = write!(xml, r#" h="{}""#, path.height);
    }
    if path.fill_mode != PathFillMode::Normal {
        let _ = write!(xml, r#" fill="{}""#, path.fill_mode.as_str());
    }
    if path.stroked != defaults.stroked {
        xml.push_str(r#" stroke="0""#);
    }
    if path.extrusion_allowed != defaults.extrusion_allowed {
        xml.push_str(r#" extrusionOk="0""#);
    }
    xml.push('>');
    for command in &path.commands {
        write_command(xml, command);
    }
    xml.push_str("</a:path>");
}

fn write_command(xml: &mut String, command: &PathCommand) {
    match command {
        PathCommand::MoveTo(point) => {
            xml.push_str("<a:moveTo>");
            write_point(xml, point);
            xml.push_str("</a:moveTo>");
        },
        PathCommand::LineTo(point) => {
            xml.push_str("<a:lnTo>");
            write_point(xml, point);
            xml.push_str("</a:lnTo>");
        },
        PathCommand::ArcTo {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        } => {
            let _ = write!(
                xml,
                r#"<a:arcTo wR="{}" hR="{}" stAng="{}" swAng="{}"/>"#,
                value_attribute(width_radius),
                value_attribute(height_radius),
                value_attribute(start_angle),
                value_attribute(swing_angle)
            );
        },
        PathCommand::QuadraticBezierTo { control, end } => {
            xml.push_str("<a:quadBezTo>");
            write_point(xml, control);
            write_point(xml, end);
            xml.push_str("</a:quadBezTo>");
        },
        PathCommand::CubicBezierTo {
            control1,
            control2,
            end,
        } => {
            xml.push_str("<a:cubicBezTo>");
            write_point(xml, control1);
            write_point(xml, control2);
            write_point(xml, end);
            xml.push_str("</a:cubicBezTo>");
        },
        PathCommand::Close => xml.push_str("<a:close/>"),
    }
}

fn write_point(xml: &mut String, point: &Point) {
    let _ = write!(
        xml,
        r#"<a:pt x="{}" y="{}"/>"#,
        value_attribute(&point.x),
        value_attribute(&point.y)
    );
}

fn write_position(xml: &mut String, position: &Point) {
    let _ = write!(
        xml,
        r#"<a:pos x="{}" y="{}"/>"#,
        value_attribute(&position.x),
        value_attribute(&position.y)
    );
}

fn write_guide_reference(xml: &mut String, name: &str, reference: &Option<String>) {
    if let Some(reference) = reference {
        let _ = write!(xml, r#" {name}="{}""#, escape_xml(reference));
    }
}

fn write_optional_value(xml: &mut String, name: &str, value: &Option<AdjustValue>) {
    if let Some(value) = value {
        let _ = write!(xml, r#" {name}="{}""#, value_attribute(value));
    }
}

fn value_attribute(value: &AdjustValue) -> String {
    match value {
        AdjustValue::Value(value) => value.to_string(),
        AdjustValue::Guide(name) => escape_xml(name),
    }
}
