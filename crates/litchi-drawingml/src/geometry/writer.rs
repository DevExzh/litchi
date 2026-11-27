//! Serializer for [`CustomGeometry`] into `a:custGeom` markup.
//!
//! Output is canonical: optional child lists are omitted when empty and
//! attributes at their ECMA-376 defaults are omitted. The resulting subtree
//! can be consumed by any format-owned shape parser into an identical
//! [`CustomGeometry`].

use std::fmt;

use litchi_core::xml::escape::escape_xml;

use super::Limits;
use super::{
    AdjustHandle, AdjustValue, CustomGeometry, Guide, Path, PathCommand, PathFillMode, Point,
};

/// Serialize one custom geometry after checking the complete output budget.
pub fn write_custom_geometry(
    xml: &mut String,
    geometry: &CustomGeometry,
    limits: &Limits,
) -> crate::Result<()> {
    let limits = limits.validate()?;
    let output_len = measure_custom_geometry_with_limits(geometry, limits)?;

    xml.try_reserve(output_len)
        .map_err(|_| crate::Error::Invalid("custom geometry output allocation failed".into()))?;
    write_custom_geometry_into(xml, geometry).map_err(format_error)?;
    Ok(())
}

/// Measure the canonical custom-geometry XML without retaining an output
/// buffer. The selected structural and output limits are applied exactly as
/// they are by [`write_custom_geometry`], allowing an owning writer to admit a
/// geometry before it reserves IDs or starts its surrounding markup.
pub fn measure_custom_geometry(geometry: &CustomGeometry, limits: &Limits) -> crate::Result<usize> {
    measure_custom_geometry_with_limits(geometry, limits.validate()?)
}

fn measure_custom_geometry_with_limits(
    geometry: &CustomGeometry,
    limits: Limits,
) -> crate::Result<usize> {
    validate_geometry_limits(geometry, limits)?;

    // Count the exact canonical output without building a second complete XML
    // fragment merely to discover that it exceeds the selected budget.
    let mut preflight = LengthWriter::default();
    write_custom_geometry_into(&mut preflight, geometry).map_err(format_error)?;
    if preflight.overflow {
        return Err(crate::Error::Invalid(
            "custom geometry output length overflows".into(),
        ));
    }
    if preflight.len > limits.max_output_bytes {
        return Err(crate::Error::limit(
            crate::LimitKind::OutputBytes,
            preflight.len,
            limits.max_output_bytes,
        ));
    }
    Ok(preflight.len)
}

/// Serialize one custom geometry as its `a:custGeom` element.
fn write_custom_geometry_into<W: fmt::Write>(
    xml: &mut W,
    geometry: &CustomGeometry,
) -> fmt::Result {
    xml.write_str("<a:custGeom>")?;
    write_guide_list(xml, "avLst", &geometry.adjust_values)?;
    write_guide_list(xml, "gdLst", &geometry.guides)?;
    if !geometry.adjust_handles.is_empty() {
        xml.write_str("<a:ahLst>")?;
        for handle in &geometry.adjust_handles {
            write_adjust_handle(xml, handle)?;
        }
        xml.write_str("</a:ahLst>")?;
    }
    if !geometry.connection_sites.is_empty() {
        xml.write_str("<a:cxnLst>")?;
        for site in &geometry.connection_sites {
            write!(xml, r#"<a:cxn ang="{}">"#, value_attribute(&site.angle))?;
            write_position(xml, &site.position)?;
            xml.write_str("</a:cxn>")?;
        }
        xml.write_str("</a:cxnLst>")?;
    }
    if let Some(rectangle) = &geometry.text_rectangle {
        write!(
            xml,
            r#"<a:rect l="{}" t="{}" r="{}" b="{}"/>"#,
            value_attribute(&rectangle.left),
            value_attribute(&rectangle.top),
            value_attribute(&rectangle.right),
            value_attribute(&rectangle.bottom)
        )?;
    }
    xml.write_str("<a:pathLst>")?;
    for path in &geometry.paths {
        write_path(xml, path)?;
    }
    xml.write_str("</a:pathLst></a:custGeom>")?;
    Ok(())
}

fn write_guide_list<W: fmt::Write>(xml: &mut W, tag: &str, guides: &[Guide]) -> fmt::Result {
    if guides.is_empty() {
        return Ok(());
    }
    write!(xml, "<a:{tag}>")?;
    for guide in guides {
        write!(
            xml,
            r#"<a:gd name="{}" fmla="{}"/>"#,
            escape_xml(&guide.name),
            escape_xml(&guide.formula.to_string())
        )?;
    }
    write!(xml, "</a:{tag}>")?;
    Ok(())
}

fn write_adjust_handle<W: fmt::Write>(xml: &mut W, handle: &AdjustHandle) -> fmt::Result {
    match handle {
        AdjustHandle::Xy(handle) => {
            xml.write_str("<a:ahXY")?;
            write_guide_reference(xml, "gdRefX", &handle.horizontal_guide)?;
            write_optional_value(xml, "minX", &handle.minimum_x)?;
            write_optional_value(xml, "maxX", &handle.maximum_x)?;
            write_guide_reference(xml, "gdRefY", &handle.vertical_guide)?;
            write_optional_value(xml, "minY", &handle.minimum_y)?;
            write_optional_value(xml, "maxY", &handle.maximum_y)?;
            xml.write_char('>')?;
            write_position(xml, &handle.position)?;
            xml.write_str("</a:ahXY>")?;
        },
        AdjustHandle::Polar(handle) => {
            xml.write_str("<a:ahPolar")?;
            write_guide_reference(xml, "gdRefR", &handle.radius_guide)?;
            write_optional_value(xml, "minR", &handle.minimum_radius)?;
            write_optional_value(xml, "maxR", &handle.maximum_radius)?;
            write_guide_reference(xml, "gdRefAng", &handle.angle_guide)?;
            write_optional_value(xml, "minAng", &handle.minimum_angle)?;
            write_optional_value(xml, "maxAng", &handle.maximum_angle)?;
            xml.write_char('>')?;
            write_position(xml, &handle.position)?;
            xml.write_str("</a:ahPolar>")?;
        },
    }
    Ok(())
}

fn write_path<W: fmt::Write>(xml: &mut W, path: &Path) -> fmt::Result {
    let defaults = Path::default();
    xml.write_str("<a:path")?;
    if path.width != defaults.width {
        write!(xml, r#" w="{}""#, path.width)?;
    }
    if path.height != defaults.height {
        write!(xml, r#" h="{}""#, path.height)?;
    }
    if path.fill_mode != PathFillMode::Normal {
        write!(xml, r#" fill="{}""#, path.fill_mode.as_str())?;
    }
    if path.stroked != defaults.stroked {
        xml.write_str(r#" stroke="0""#)?;
    }
    if path.extrusion_allowed != defaults.extrusion_allowed {
        xml.write_str(r#" extrusionOk="0""#)?;
    }
    xml.write_char('>')?;
    for command in &path.commands {
        write_command(xml, command)?;
    }
    xml.write_str("</a:path>")?;
    Ok(())
}

fn write_command<W: fmt::Write>(xml: &mut W, command: &PathCommand) -> fmt::Result {
    match command {
        PathCommand::MoveTo(point) => {
            xml.write_str("<a:moveTo>")?;
            write_point(xml, point)?;
            xml.write_str("</a:moveTo>")?;
        },
        PathCommand::LineTo(point) => {
            xml.write_str("<a:lnTo>")?;
            write_point(xml, point)?;
            xml.write_str("</a:lnTo>")?;
        },
        PathCommand::ArcTo {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        } => {
            write!(
                xml,
                r#"<a:arcTo wR="{}" hR="{}" stAng="{}" swAng="{}"/>"#,
                value_attribute(width_radius),
                value_attribute(height_radius),
                value_attribute(start_angle),
                value_attribute(swing_angle)
            )?;
        },
        PathCommand::QuadraticBezierTo { control, end } => {
            xml.write_str("<a:quadBezTo>")?;
            write_point(xml, control)?;
            write_point(xml, end)?;
            xml.write_str("</a:quadBezTo>")?;
        },
        PathCommand::CubicBezierTo {
            control1,
            control2,
            end,
        } => {
            xml.write_str("<a:cubicBezTo>")?;
            write_point(xml, control1)?;
            write_point(xml, control2)?;
            write_point(xml, end)?;
            xml.write_str("</a:cubicBezTo>")?;
        },
        PathCommand::Close => xml.write_str("<a:close/>")?,
    }
    Ok(())
}

fn write_point<W: fmt::Write>(xml: &mut W, point: &Point) -> fmt::Result {
    write!(
        xml,
        r#"<a:pt x="{}" y="{}"/>"#,
        value_attribute(&point.x),
        value_attribute(&point.y)
    )
}

fn write_position<W: fmt::Write>(xml: &mut W, position: &Point) -> fmt::Result {
    write!(
        xml,
        r#"<a:pos x="{}" y="{}"/>"#,
        value_attribute(&position.x),
        value_attribute(&position.y)
    )
}

fn write_guide_reference<W: fmt::Write>(
    xml: &mut W,
    name: &str,
    reference: &Option<String>,
) -> fmt::Result {
    if let Some(reference) = reference {
        write!(xml, r#" {name}="{}""#, escape_xml(reference))?;
    }
    Ok(())
}

fn write_optional_value<W: fmt::Write>(
    xml: &mut W,
    name: &str,
    value: &Option<AdjustValue>,
) -> fmt::Result {
    if let Some(value) = value {
        write!(xml, r#" {name}="{}""#, value_attribute(value))?;
    }
    Ok(())
}

fn value_attribute(value: &AdjustValue) -> String {
    match value {
        AdjustValue::Value(value) => value.to_string(),
        AdjustValue::Guide(name) => escape_xml(name),
    }
}

fn validate_geometry_limits(geometry: &CustomGeometry, limits: Limits) -> crate::Result<()> {
    if geometry.adjust_values.len() > limits.max_guides {
        return Err(crate::Error::limit(
            crate::LimitKind::GeometryGuides,
            geometry.adjust_values.len(),
            limits.max_guides,
        ));
    }
    if geometry.guides.len() > limits.max_guides {
        return Err(crate::Error::limit(
            crate::LimitKind::GeometryGuides,
            geometry.guides.len(),
            limits.max_guides,
        ));
    }
    if geometry.adjust_handles.len() > limits.max_adjust_handles {
        return Err(crate::Error::limit(
            crate::LimitKind::GeometrySites,
            geometry.adjust_handles.len(),
            limits.max_adjust_handles,
        ));
    }
    if geometry.connection_sites.len() > limits.max_connection_sites {
        return Err(crate::Error::limit(
            crate::LimitKind::GeometrySites,
            geometry.connection_sites.len(),
            limits.max_connection_sites,
        ));
    }
    if geometry.paths.len() > limits.max_paths {
        return Err(crate::Error::limit(
            crate::LimitKind::GeometryPaths,
            geometry.paths.len(),
            limits.max_paths,
        ));
    }
    for path in &geometry.paths {
        if path.commands.len() > limits.max_path_commands {
            return Err(crate::Error::limit(
                crate::LimitKind::GeometryPathCommands,
                path.commands.len(),
                limits.max_path_commands,
            ));
        }
    }
    Ok(())
}

#[derive(Default)]
struct LengthWriter {
    len: usize,
    overflow: bool,
}

impl fmt::Write for LengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        match self.len.checked_add(value.len()) {
            Some(len) => self.len = len,
            None => self.overflow = true,
        }
        Ok(())
    }
}

fn format_error(_: fmt::Error) -> crate::Error {
    crate::Error::Invalid("custom geometry XML formatting failed".into())
}
