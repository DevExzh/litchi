use super::model::{Geometry, PathKind};
use super::validation::validate;
use crate::prop::{Array, Id, Props, Value};
use crate::{Error, Result};

/// Parses the custom-geometry property family from one shape property table.
///
/// # Errors
///
/// Returns `Error::MalformedGeometry` if `shapePath` is complex or does not
/// contain a scalar value, if a geometry array is malformed, or if the vertex
/// and segment-info arrays fail validation.
pub fn parse<'data>(properties: &Props<'data>) -> Result<Option<Geometry<'data>>> {
    let has_geometry = [Id::ShapePath, Id::Vertices, Id::SegmentInfo]
        .into_iter()
        .any(|id| properties.has(id));
    if !has_geometry {
        return Ok(None);
    }

    let shape_path = shape_path(properties)?;
    let vertices = array_value(properties, Id::Vertices, 8)?;
    let segment_info = array_value(properties, Id::SegmentInfo, 2)?;

    validate(shape_path, vertices, segment_info)?;
    Ok(Some(Geometry::from_parts(
        shape_path,
        vertices,
        segment_info,
    )))
}

fn shape_path(properties: &Props<'_>) -> Result<PathKind> {
    let Some(prop) = properties.prop(Id::ShapePath) else {
        return Ok(PathKind::LinesClosed);
    };
    if prop.is_complex() {
        return Err(Error::MalformedGeometry {
            reason: "shapePath must be a simple property",
        });
    }
    match prop.value() {
        Value::Simple(value) => Ok(PathKind::from_raw((*value).cast_unsigned())),
        Value::Complex(_) | Value::Array(_) => Err(Error::MalformedGeometry {
            reason: "shapePath does not contain a scalar value",
        }),
    }
}

fn array_value<'data>(
    properties: &Props<'data>,
    id: Id,
    element_size: usize,
) -> Result<Option<Array<'data>>> {
    let Some(prop) = properties.prop(id) else {
        return Ok(None);
    };
    if !prop.is_complex() {
        if prop.raw_value() != 0 {
            return Err(Error::MalformedGeometry {
                reason: "non-complex geometry property has a nonzero length",
            });
        }
        return Ok(None);
    }

    let array = match prop.value() {
        Value::Array(array) => *array,
        Value::Simple(_) | Value::Complex(_) => {
            return Err(Error::MalformedGeometry {
                reason: "complex geometry property is not an IMsoArray",
            });
        },
    };
    if array.raw_element_size() as usize != element_size {
        return Err(Error::MalformedGeometry {
            reason: "geometry IMsoArray element size is invalid",
        });
    }
    Ok(Some(array))
}
