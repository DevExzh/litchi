//! Authoring-time validation of custom geometry against the ECMA-376 value
//! spaces (ST_Coordinate, ST_PositiveCoordinate, ST_Angle, ST_GeomGuideName)
//! and the module limits.
//!
//! Errors are plain strings, matching the shape authoring validators in
//! [`crate::xlsx::writer::shape`].

use super::{
    MAX_ADJUST_HANDLES, MAX_ANGLE, MAX_CONNECTION_SITES, MAX_COORDINATE, MAX_GEOMETRY_GUIDES,
    MAX_GEOMETRY_PATHS, MAX_GUIDE_NAME_BYTES, MAX_PATH_COMMANDS, MAX_POSITIVE_COORDINATE,
    MIN_ANGLE, MIN_COORDINATE, XlsxAdjustHandle, XlsxAdjustValue, XlsxCustomGeometry,
    XlsxGeometryPath, XlsxGeometryPoint, XlsxPathCommand,
};

/// Validate authored custom geometry.
///
/// The schema requires the `a:pathLst` element but allows it to be empty;
/// authored geometry additionally requires at least one path so the shape
/// actually draws something.
pub(crate) fn validate_custom_geometry(geometry: &XlsxCustomGeometry) -> Result<(), String> {
    if geometry.adjust_values.len() > MAX_GEOMETRY_GUIDES
        || geometry.guides.len() > MAX_GEOMETRY_GUIDES
    {
        return Err("custom geometry guide count limit exceeded".to_string());
    }
    for guide in geometry.adjust_values.iter().chain(&geometry.guides) {
        validate_guide_name(&guide.name, "geometry guide name")?;
        for operand in guide.formula.operands() {
            validate_coordinate_value(operand, "geometry formula operand")?;
        }
    }
    if geometry.adjust_handles.len() > MAX_ADJUST_HANDLES {
        return Err("custom geometry adjust handle count limit exceeded".to_string());
    }
    for handle in &geometry.adjust_handles {
        validate_adjust_handle(handle)?;
    }
    if geometry.connection_sites.len() > MAX_CONNECTION_SITES {
        return Err("custom geometry connection site count limit exceeded".to_string());
    }
    for site in &geometry.connection_sites {
        validate_angle_value(&site.angle, "connection site angle")?;
        validate_point(&site.position, "connection site position")?;
    }
    if let Some(rectangle) = &geometry.text_rectangle {
        validate_coordinate_value(&rectangle.left, "text rectangle left edge")?;
        validate_coordinate_value(&rectangle.top, "text rectangle top edge")?;
        validate_coordinate_value(&rectangle.right, "text rectangle right edge")?;
        validate_coordinate_value(&rectangle.bottom, "text rectangle bottom edge")?;
    }
    if geometry.paths.is_empty() {
        return Err("custom geometry requires at least one path".to_string());
    }
    if geometry.paths.len() > MAX_GEOMETRY_PATHS {
        return Err("custom geometry path count limit exceeded".to_string());
    }
    for path in &geometry.paths {
        validate_path(path)?;
    }
    Ok(())
}

fn validate_adjust_handle(handle: &XlsxAdjustHandle) -> Result<(), String> {
    match handle {
        XlsxAdjustHandle::Xy(handle) => {
            validate_guide_reference(&handle.horizontal_guide, "XY handle horizontal guide")?;
            validate_guide_reference(&handle.vertical_guide, "XY handle vertical guide")?;
            for (value, field) in [
                (&handle.minimum_x, "XY handle minimum X"),
                (&handle.maximum_x, "XY handle maximum X"),
                (&handle.minimum_y, "XY handle minimum Y"),
                (&handle.maximum_y, "XY handle maximum Y"),
            ] {
                if let Some(value) = value {
                    validate_coordinate_value(value, field)?;
                }
            }
            validate_point(&handle.position, "XY handle position")
        },
        XlsxAdjustHandle::Polar(handle) => {
            validate_guide_reference(&handle.radius_guide, "polar handle radius guide")?;
            validate_guide_reference(&handle.angle_guide, "polar handle angle guide")?;
            for (value, field) in [
                (&handle.minimum_radius, "polar handle minimum radius"),
                (&handle.maximum_radius, "polar handle maximum radius"),
            ] {
                if let Some(value) = value {
                    validate_coordinate_value(value, field)?;
                }
            }
            for (value, field) in [
                (&handle.minimum_angle, "polar handle minimum angle"),
                (&handle.maximum_angle, "polar handle maximum angle"),
            ] {
                if let Some(value) = value {
                    validate_angle_value(value, field)?;
                }
            }
            validate_point(&handle.position, "polar handle position")
        },
    }
}

fn validate_path(path: &XlsxGeometryPath) -> Result<(), String> {
    if !(0..=MAX_POSITIVE_COORDINATE).contains(&path.width)
        || !(0..=MAX_POSITIVE_COORDINATE).contains(&path.height)
    {
        return Err("geometry path size is outside ST_PositiveCoordinate bounds".to_string());
    }
    if path.commands.len() > MAX_PATH_COMMANDS {
        return Err("geometry path command count limit exceeded".to_string());
    }
    for command in &path.commands {
        validate_command(command)?;
    }
    Ok(())
}

fn validate_command(command: &XlsxPathCommand) -> Result<(), String> {
    match command {
        XlsxPathCommand::MoveTo(point) | XlsxPathCommand::LineTo(point) => {
            validate_point(point, "path command point")
        },
        XlsxPathCommand::ArcTo {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        } => {
            validate_coordinate_value(width_radius, "arc width radius")?;
            validate_coordinate_value(height_radius, "arc height radius")?;
            validate_angle_value(start_angle, "arc start angle")?;
            validate_angle_value(swing_angle, "arc swing angle")
        },
        XlsxPathCommand::QuadraticBezierTo { control, end } => {
            validate_point(control, "path command point")?;
            validate_point(end, "path command point")
        },
        XlsxPathCommand::CubicBezierTo {
            control1,
            control2,
            end,
        } => {
            validate_point(control1, "path command point")?;
            validate_point(control2, "path command point")?;
            validate_point(end, "path command point")
        },
        XlsxPathCommand::Close => Ok(()),
    }
}

fn validate_point(point: &XlsxGeometryPoint, field: &str) -> Result<(), String> {
    validate_coordinate_value(&point.x, field)?;
    validate_coordinate_value(&point.y, field)
}

fn validate_coordinate_value(value: &XlsxAdjustValue, field: &str) -> Result<(), String> {
    match value {
        XlsxAdjustValue::Value(value) if !(MIN_COORDINATE..=MAX_COORDINATE).contains(value) => {
            Err(format!("{field} is outside ST_Coordinate bounds"))
        },
        XlsxAdjustValue::Value(_) => Ok(()),
        XlsxAdjustValue::Guide(name) => validate_guide_name(name, field),
    }
}

fn validate_angle_value(value: &XlsxAdjustValue, field: &str) -> Result<(), String> {
    match value {
        XlsxAdjustValue::Value(value) if !(MIN_ANGLE..=MAX_ANGLE).contains(value) => {
            Err(format!("{field} is outside ST_Angle bounds"))
        },
        XlsxAdjustValue::Value(_) => Ok(()),
        XlsxAdjustValue::Guide(name) => validate_guide_name(name, field),
    }
}

fn validate_guide_reference(name: &Option<String>, field: &str) -> Result<(), String> {
    match name {
        Some(name) => validate_guide_name(name, field),
        None => Ok(()),
    }
}

/// Validate an ST_GeomGuideName token for authoring.
///
/// Names must be non-empty, bounded, XML-safe, free of whitespace (guide
/// references embed in space-delimited formulas), and must not parse as a
/// number (a numeric name would be re-read as a literal).
fn validate_guide_name(name: &str, field: &str) -> Result<(), String> {
    if name.is_empty() {
        return Err(format!("{field} cannot be empty"));
    }
    if name.len() > MAX_GUIDE_NAME_BYTES {
        return Err(format!("{field} is too long"));
    }
    if name.chars().any(char::is_whitespace) {
        return Err(format!("{field} cannot contain whitespace"));
    }
    if name.parse::<i64>().is_ok() {
        return Err(format!("{field} cannot be a number"));
    }
    if name.chars().any(|character| {
        !matches!(
            character as u32,
            0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
    }) {
        return Err(format!("{field} contains an invalid XML character"));
    }
    Ok(())
}
