//! Authoring-time validation of custom geometry against the ECMA-376 value
//! spaces (ST_Coordinate, ST_PositiveCoordinate, ST_Angle, ST_GeomGuideName)
//! and the module limits.
//!
//! Errors are plain strings, matching the shape authoring validators in
//! [`crate::chart_sheet::writer::shape`].

use super::{
    AdjustHandle, AdjustValue, CustomGeometry, MAX_ADJUST_HANDLES, MAX_ANGLE, MAX_CONNECTION_SITES,
    MAX_COORDINATE, MAX_GEOMETRY_GUIDES, MAX_GEOMETRY_PATHS, MAX_GUIDE_NAME_BYTES,
    MAX_PATH_COMMANDS, MAX_POSITIVE_COORDINATE, MIN_ANGLE, MIN_COORDINATE, Path, PathCommand,
    Point,
};

#[derive(Clone, Copy, PartialEq, Eq)]
enum Mode {
    Author,
    Parsed,
}

/// Validate authored custom geometry.
///
/// The schema requires the `a:pathLst` element but allows it to be empty;
/// authored geometry additionally requires at least one path so the shape
/// actually draws something.
pub fn validate_custom_geometry(geometry: &CustomGeometry) -> Result<(), String> {
    validate_geometry(geometry, Mode::Author)
}

/// Validate decoded geometry without imposing the authoring-only requirement
/// that a path list contain a visible path.
pub fn validate_parsed_custom_geometry(geometry: &CustomGeometry) -> Result<(), String> {
    validate_geometry(geometry, Mode::Parsed)
}

fn validate_geometry(geometry: &CustomGeometry, mode: Mode) -> Result<(), String> {
    if geometry.adjust_values.len() > MAX_GEOMETRY_GUIDES
        || geometry.guides.len() > MAX_GEOMETRY_GUIDES
    {
        return Err("custom geometry guide count limit exceeded".to_string());
    }
    for guide in geometry.adjust_values.iter().chain(&geometry.guides) {
        validate_guide_name(&guide.name, "geometry guide name", mode)?;
        for operand in guide.formula.operands() {
            validate_coordinate_value(operand, "geometry formula operand", mode)?;
        }
    }
    if geometry.adjust_handles.len() > MAX_ADJUST_HANDLES {
        return Err("custom geometry adjust handle count limit exceeded".to_string());
    }
    for handle in &geometry.adjust_handles {
        validate_adjust_handle(handle, mode)?;
    }
    if geometry.connection_sites.len() > MAX_CONNECTION_SITES {
        return Err("custom geometry connection site count limit exceeded".to_string());
    }
    for site in &geometry.connection_sites {
        validate_angle_value(&site.angle, "connection site angle", mode)?;
        validate_point(&site.position, "connection site position", mode)?;
    }
    if let Some(rectangle) = &geometry.text_rectangle {
        validate_coordinate_value(&rectangle.left, "text rectangle left edge", mode)?;
        validate_coordinate_value(&rectangle.top, "text rectangle top edge", mode)?;
        validate_coordinate_value(&rectangle.right, "text rectangle right edge", mode)?;
        validate_coordinate_value(&rectangle.bottom, "text rectangle bottom edge", mode)?;
    }
    if mode == Mode::Author && geometry.paths.is_empty() {
        return Err("custom geometry requires at least one path".to_string());
    }
    if geometry.paths.len() > MAX_GEOMETRY_PATHS {
        return Err("custom geometry path count limit exceeded".to_string());
    }
    for path in &geometry.paths {
        validate_path(path, mode)?;
    }
    Ok(())
}

fn validate_adjust_handle(handle: &AdjustHandle, mode: Mode) -> Result<(), String> {
    match handle {
        AdjustHandle::Xy(handle) => {
            validate_guide_reference(&handle.horizontal_guide, "XY handle horizontal guide", mode)?;
            validate_guide_reference(&handle.vertical_guide, "XY handle vertical guide", mode)?;
            for (value, field) in [
                (&handle.minimum_x, "XY handle minimum X"),
                (&handle.maximum_x, "XY handle maximum X"),
                (&handle.minimum_y, "XY handle minimum Y"),
                (&handle.maximum_y, "XY handle maximum Y"),
            ] {
                if let Some(value) = value {
                    validate_coordinate_value(value, field, mode)?;
                }
            }
            validate_point(&handle.position, "XY handle position", mode)
        },
        AdjustHandle::Polar(handle) => {
            validate_guide_reference(&handle.radius_guide, "polar handle radius guide", mode)?;
            validate_guide_reference(&handle.angle_guide, "polar handle angle guide", mode)?;
            for (value, field) in [
                (&handle.minimum_radius, "polar handle minimum radius"),
                (&handle.maximum_radius, "polar handle maximum radius"),
            ] {
                if let Some(value) = value {
                    validate_coordinate_value(value, field, mode)?;
                }
            }
            for (value, field) in [
                (&handle.minimum_angle, "polar handle minimum angle"),
                (&handle.maximum_angle, "polar handle maximum angle"),
            ] {
                if let Some(value) = value {
                    validate_angle_value(value, field, mode)?;
                }
            }
            validate_point(&handle.position, "polar handle position", mode)
        },
    }
}

fn validate_path(path: &Path, mode: Mode) -> Result<(), String> {
    if !(0..=MAX_POSITIVE_COORDINATE).contains(&path.width)
        || !(0..=MAX_POSITIVE_COORDINATE).contains(&path.height)
    {
        return Err("geometry path size is outside ST_PositiveCoordinate bounds".to_string());
    }
    if path.commands.len() > MAX_PATH_COMMANDS {
        return Err("geometry path command count limit exceeded".to_string());
    }
    for command in &path.commands {
        validate_command(command, mode)?;
    }
    Ok(())
}

fn validate_command(command: &PathCommand, mode: Mode) -> Result<(), String> {
    match command {
        PathCommand::MoveTo(point) | PathCommand::LineTo(point) => {
            validate_point(point, "path command point", mode)
        },
        PathCommand::ArcTo {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        } => {
            validate_coordinate_value(width_radius, "arc width radius", mode)?;
            validate_coordinate_value(height_radius, "arc height radius", mode)?;
            validate_angle_value(start_angle, "arc start angle", mode)?;
            validate_angle_value(swing_angle, "arc swing angle", mode)
        },
        PathCommand::QuadraticBezierTo { control, end } => {
            validate_point(control, "path command point", mode)?;
            validate_point(end, "path command point", mode)
        },
        PathCommand::CubicBezierTo {
            control1,
            control2,
            end,
        } => {
            validate_point(control1, "path command point", mode)?;
            validate_point(control2, "path command point", mode)?;
            validate_point(end, "path command point", mode)
        },
        PathCommand::Close => Ok(()),
    }
}

fn validate_point(point: &Point, field: &str, mode: Mode) -> Result<(), String> {
    validate_coordinate_value(&point.x, field, mode)?;
    validate_coordinate_value(&point.y, field, mode)
}

fn validate_coordinate_value(value: &AdjustValue, field: &str, mode: Mode) -> Result<(), String> {
    match value {
        AdjustValue::Value(value) if !(MIN_COORDINATE..=MAX_COORDINATE).contains(value) => {
            Err(format!("{field} is outside ST_Coordinate bounds"))
        },
        AdjustValue::Value(_) => Ok(()),
        AdjustValue::Guide(name) => validate_guide_name(name, field, mode),
    }
}

fn validate_angle_value(value: &AdjustValue, field: &str, mode: Mode) -> Result<(), String> {
    match value {
        AdjustValue::Value(value) if !(MIN_ANGLE..=MAX_ANGLE).contains(value) => {
            Err(format!("{field} is outside ST_Angle bounds"))
        },
        AdjustValue::Value(_) => Ok(()),
        AdjustValue::Guide(name) => validate_guide_name(name, field, mode),
    }
}

fn validate_guide_reference(name: &Option<String>, field: &str, mode: Mode) -> Result<(), String> {
    match name {
        Some(name) => validate_guide_name(name, field, mode),
        None => Ok(()),
    }
}

/// Validate an ST_GeomGuideName token.
///
/// `ST_GeomGuideName` is an `xsd:token`, so internal spaces, empty strings, and
/// numeric spellings remain schema-valid when parsing. Authoring rejects those
/// forms because formulas are space-delimited and numeric references are
/// indistinguishable from literals in the `ST_AdjCoordinate`/`ST_AdjAngle`
/// unions.
fn validate_guide_name(name: &str, field: &str, mode: Mode) -> Result<(), String> {
    if mode == Mode::Author {
        if name.is_empty() {
            return Err(format!("{field} cannot be empty"));
        }
        if name.chars().any(char::is_whitespace) {
            return Err(format!("{field} cannot contain whitespace"));
        }
        if name.parse::<i64>().is_ok() {
            return Err(format!("{field} cannot be a number"));
        }
    }
    if name.len() > MAX_GUIDE_NAME_BYTES {
        return Err(format!("{field} is too long"));
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
