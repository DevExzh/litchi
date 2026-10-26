//! Bounded wire codec of shape geometry shared by the Pages, Numbers and
//! Keynote shape engines (record 0375): today the native straight line —
//! a `TSWP.ShapeInfoArchive` whose Bézier path source is one `moveTo`/`lineTo`
//! pair along the x axis.
//!
//! [`line`](crate::shape_style_codec::line) reads the line's drawable
//! geometry and path; [`rewrite_line`](crate::shape_style_codec::rewrite_line)
//! moves it, patching only the geometry components, the natural width and the
//! `lineTo` x coordinate in place (the migration host's byte semantics).
//! Errors, limits and the borrowed candidate are those of
//! [`chart_style_codec`](crate::chart_style_codec).

use std::borrow::Cow;

use crate::chart_style_codec::{
    Budget, Field, allocate, check_bytes, push_varint, read_f32, replace_payload, replace_scalar,
    require_wire_type, singular, visit_fields,
};
pub use crate::chart_style_codec::{Error, LimitKind, Limits, Rewrite};

const WIRE_VARINT: u8 = 0;
const WIRE_LENGTH_DELIMITED: u8 = 2;
const WIRE_FIXED32: u8 = 5;
const SUPER_FIELD: u32 = 1;
const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const SHAPE_PATH_SOURCE_FIELD: u32 = 3;
const GEOMETRY_POSITION_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const GEOMETRY_FLAGS_FIELD: u32 = 3;
const GEOMETRY_ANGLE_FIELD: u32 = 4;
const PATH_HORIZONTAL_FLIP_FIELD: u32 = 1;
const PATH_BEZIER_FIELD: u32 = 5;
const BEZIER_NATURAL_SIZE_FIELD: u32 = 2;
const BEZIER_PATH_FIELD: u32 = 3;
const PATH_ELEMENTS_FIELD: u32 = 1;
const ELEMENT_TYPE_FIELD: u32 = 1;
const ELEMENT_POINTS_FIELD: u32 = 2;
const X_FIELD: u32 = 1;
const Y_FIELD: u32 = 2;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const TOLERANCE: f32 = 0.001;

/// The drawable flags of a native straight line.
pub const LINE_FLAGS: u32 = 3;

/// The drawable geometry of a shape (`TSD.GeometryArchive`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Geometry {
    /// `position` (x, y), if stored.
    pub position: Option<[f32; 2]>,
    /// `size` (width, height), if stored.
    pub size: Option<[f32; 2]>,
    /// `flags`, if stored.
    pub flags: Option<u32>,
    /// `angle` in degrees, if stored.
    pub angle: Option<f32>,
}

/// A native straight line as stored.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Line {
    /// The drawable geometry.
    pub geometry: Geometry,
    /// The Bézier natural width (the line's length).
    pub natural_width: f32,
    /// `PathSourceArchive.horizontalFlip` (the stored direction is reversed).
    pub horizontal_flip: bool,
}

fn message<'a>(
    data: &'a [u8],
    number: u32,
    budget: &mut Budget,
) -> Result<Option<(Field, &'a [u8])>, Error> {
    let Some(field) = singular(data, number, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
    Ok(Some((field, field.payload(data))))
}

fn float(data: &[u8], number: u32, budget: &mut Budget) -> Result<Option<f32>, Error> {
    let Some(field) = singular(data, number, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_FIXED32)?;
    Ok(Some(read_f32(field.payload(data))))
}

fn pair(data: &[u8], budget: &mut Budget) -> Result<[f32; 2], Error> {
    Ok([
        float(data, X_FIELD, budget)?.unwrap_or(0.0),
        float(data, Y_FIELD, budget)?.unwrap_or(0.0),
    ])
}

fn geometry(data: &[u8], budget: &mut Budget) -> Result<Geometry, Error> {
    let position = message(data, GEOMETRY_POSITION_FIELD, budget)?
        .map(|(_, point)| pair(point, budget))
        .transpose()?;
    let size = message(data, GEOMETRY_SIZE_FIELD, budget)?
        .map(|(_, size)| pair(size, budget))
        .transpose()?;
    let flags = match singular(data, GEOMETRY_FLAGS_FIELD, budget)? {
        Some(field) => {
            require_wire_type(&field, WIRE_VARINT)?;
            Some(field.varint as u32)
        },
        None => None,
    };
    let angle = float(data, GEOMETRY_ANGLE_FIELD, budget)?;
    Ok(Geometry {
        position,
        size,
        flags,
        angle,
    })
}

const fn near(left: f32, right: f32) -> bool {
    (left - right).abs() <= TOLERANCE
}

/// The two path elements of a straight line, if `path` is one.
fn straight_elements(
    path: &[u8],
    width: f32,
    budget: &mut Budget,
) -> Result<Option<[Field; 2]>, Error> {
    let mut elements = Vec::new();
    visit_fields(path, budget, |field| {
        if field.number == PATH_ELEMENTS_FIELD {
            require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
            elements.push(field);
        }
        Ok(())
    })?;
    let [move_to, line_to] = elements.as_slice() else {
        return Ok(None);
    };
    for (element, kind, x) in [(move_to, MOVE_TO, 0.0), (line_to, LINE_TO, width)] {
        let element = element.payload(path);
        let element_type = singular(element, ELEMENT_TYPE_FIELD, budget)?;
        let mut points = Vec::new();
        visit_fields(element, budget, |field| {
            if field.number == ELEMENT_POINTS_FIELD {
                require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
                points.push(field);
            }
            Ok(())
        })?;
        let [point] = points.as_slice() else {
            return Ok(None);
        };
        let [px, py] = pair(point.payload(element), budget)?;
        if element_type.map(|field| field.varint) != Some(kind) || !near(px, x) || !near(py, 0.0) {
            return Ok(None);
        }
    }
    Ok(Some([*move_to, *line_to]))
}

/// Read the native straight line of a `TSWP.ShapeInfoArchive` payload, or
/// `None` when its path is not a straight line.
///
/// # Errors
///
/// [`Error::Malformed`], [`Error::DuplicateField`] or [`Error::WireType`] for
/// an invalid payload; [`Error::Limit`] when `limits` is exceeded.
pub fn line(shape_info: &[u8], limits: Limits) -> Result<Option<Line>, Error> {
    check_bytes(shape_info.len(), limits)?;
    let mut budget = Budget::new(limits);
    let Some((_, shape)) = message(shape_info, SUPER_FIELD, &mut budget)? else {
        return Err(Error::Malformed {
            reason: "the shape info has no shape",
        });
    };
    let Some((_, path_source)) = message(shape, SHAPE_PATH_SOURCE_FIELD, &mut budget)? else {
        return Err(Error::Malformed {
            reason: "the shape has no path source",
        });
    };
    let Some((_, bezier)) = message(path_source, PATH_BEZIER_FIELD, &mut budget)? else {
        return Ok(None);
    };
    let Some((_, natural)) = message(bezier, BEZIER_NATURAL_SIZE_FIELD, &mut budget)? else {
        return Ok(None);
    };
    let [width, height] = pair(natural, &mut budget)?;
    let Some((_, path)) = message(bezier, BEZIER_PATH_FIELD, &mut budget)? else {
        return Ok(None);
    };
    if !width.is_finite() || width <= 0.0 || !near(height, 0.0) {
        return Ok(None);
    }
    if straight_elements(path, width, &mut budget)?.is_none() {
        return Ok(None);
    }
    let horizontal_flip = match singular(path_source, PATH_HORIZONTAL_FLIP_FIELD, &mut budget)? {
        Some(field) => {
            require_wire_type(&field, WIRE_VARINT)?;
            field.varint != 0
        },
        None => false,
    };
    let Some((_, drawable)) = message(shape, SUPER_FIELD, &mut budget)? else {
        return Err(Error::Malformed {
            reason: "the shape has no drawable",
        });
    };
    let Some((_, stored)) = message(drawable, DRAWABLE_GEOMETRY_FIELD, &mut budget)? else {
        return Err(Error::Malformed {
            reason: "the drawable has no geometry",
        });
    };
    Ok(Some(Line {
        geometry: geometry(stored, &mut budget)?,
        natural_width: width,
        horizontal_flip,
    }))
}

/// Apply `edit` to the payload of field `number` (which must be present) and
/// splice the result back; `None` when nothing changed.
fn nested(
    data: &[u8],
    number: u32,
    limits: Limits,
    edit: impl FnOnce(&[u8]) -> Result<Option<Vec<u8>>, Error>,
) -> Result<Option<Vec<u8>>, Error> {
    let mut budget = Budget::new(limits);
    let Some((field, payload)) = message(data, number, &mut budget)? else {
        return Err(Error::Malformed {
            reason: "a straight-line field is missing",
        });
    };
    let Some(next) = edit(payload)? else {
        return Ok(None);
    };
    replace_payload(data, &field, &next, limits).map(Some)
}

/// Set a present fixed32 field (absent: error, as the host requires).
fn set_present_float(
    data: &[u8],
    number: u32,
    value: f32,
    limits: Limits,
) -> Result<Option<Vec<u8>>, Error> {
    let mut budget = Budget::new(limits);
    let field = singular(data, number, &mut budget)?.ok_or(Error::Malformed {
        reason: "a straight-line coordinate is missing",
    })?;
    require_wire_type(&field, WIRE_FIXED32)?;
    replace_scalar(data, &field, &value.to_le_bytes(), limits)
}

fn chain(
    current: Option<Vec<u8>>,
    source: &[u8],
    step: impl FnOnce(&[u8]) -> Result<Option<Vec<u8>>, Error>,
) -> Result<Option<Vec<u8>>, Error> {
    let base = current.as_deref().unwrap_or(source);
    Ok(step(base)?.or(current))
}

fn set_pair(data: &[u8], value: [f32; 2], limits: Limits) -> Result<Option<Vec<u8>>, Error> {
    let after_x = set_present_float(data, X_FIELD, value[0], limits)?;
    chain(after_x, data, |base| {
        set_present_float(base, Y_FIELD, value[1], limits)
    })
}

/// Set an optional geometry field: patch the present message or scalar in
/// place, or append the encoded value.
fn set_geometry(
    data: &[u8],
    replacement: Geometry,
    limits: Limits,
) -> Result<Option<Vec<u8>>, Error> {
    let mut current: Option<Vec<u8>> = None;
    for (number, value) in [
        (GEOMETRY_POSITION_FIELD, replacement.position),
        (GEOMETRY_SIZE_FIELD, replacement.size),
    ] {
        let Some(value) = value else {
            return Err(Error::Malformed {
                reason: "a line geometry needs a position and a size",
            });
        };
        current = chain(current, data, |base| {
            let mut budget = Budget::new(limits);
            if singular(base, number, &mut budget)?.is_some() {
                nested(base, number, limits, |pair| set_pair(pair, value, limits))
            } else {
                let mut encoded = Vec::new();
                for (component, coordinate) in [(X_FIELD, value[0]), (Y_FIELD, value[1])] {
                    push_varint(
                        &mut encoded,
                        (u64::from(component) << 3) | u64::from(WIRE_FIXED32),
                    );
                    encoded.extend_from_slice(&coordinate.to_le_bytes());
                }
                let mut output = allocate(base.len() + encoded.len() + 2, limits)?;
                output.extend_from_slice(base);
                push_varint(
                    &mut output,
                    (u64::from(number) << 3) | u64::from(WIRE_LENGTH_DELIMITED),
                );
                push_varint(&mut output, encoded.len() as u64);
                output.extend_from_slice(&encoded);
                Ok(Some(output))
            }
        })?;
    }
    let flags = replacement.flags.map(|flags| {
        let mut bytes = Vec::new();
        push_varint(&mut bytes, u64::from(flags));
        (GEOMETRY_FLAGS_FIELD, WIRE_VARINT, bytes)
    });
    let angle = replacement.angle.map(|angle| {
        (
            GEOMETRY_ANGLE_FIELD,
            WIRE_FIXED32,
            angle.to_le_bytes().to_vec(),
        )
    });
    for (number, wire_type, value) in [flags, angle].into_iter().flatten() {
        current = chain(current, data, |base| {
            let mut budget = Budget::new(limits);
            match singular(base, number, &mut budget)? {
                Some(field) => {
                    require_wire_type(&field, wire_type)?;
                    replace_scalar(base, &field, &value, limits)
                },
                None => {
                    let mut output = allocate(base.len() + value.len() + 1, limits)?;
                    output.extend_from_slice(base);
                    push_varint(&mut output, (u64::from(number) << 3) | u64::from(wire_type));
                    output.extend_from_slice(&value);
                    Ok(Some(output))
                },
            }
        })?;
    }
    Ok(current)
}

fn set_path_width(path: &[u8], width: f32, limits: Limits) -> Result<Option<Vec<u8>>, Error> {
    let mut budget = Budget::new(limits);
    let mut elements = Vec::new();
    visit_fields(path, &mut budget, |field| {
        if field.number == PATH_ELEMENTS_FIELD {
            elements.push(field);
        }
        Ok(())
    })?;
    let [_, line_to] = elements.as_slice() else {
        return Err(Error::Malformed {
            reason: "a straight line needs two path elements",
        });
    };
    let element = line_to.payload(path);
    let next = nested(element, ELEMENT_POINTS_FIELD, limits, |point| {
        set_present_float(point, X_FIELD, width, limits)
    })?;
    let Some(next) = next else {
        return Ok(None);
    };
    replace_payload(path, line_to, &next, limits).map(Some)
}

/// Move a native straight line: replace its drawable geometry (position and
/// size components, flags, angle) and its length (the natural width and the
/// `lineTo` x coordinate), keeping every other byte.
///
/// # Errors
///
/// [`Error::Malformed`] when the payload is not a straight line or misses a
/// coordinate the host requires; the [`line()`]
/// errors; [`Error::Limit`]/[`Error::Allocation`] for the candidate.
pub fn rewrite_line<'source>(
    shape_info: &'source [u8],
    replacement: Geometry,
    width: f32,
    limits: Limits,
) -> Result<Rewrite<'source>, Error> {
    if line(shape_info, limits)?.is_none() {
        return Err(Error::Malformed {
            reason: "the shape is not a native straight line",
        });
    }
    let next = nested(shape_info, SUPER_FIELD, limits, |shape| {
        let after_geometry = nested(shape, SUPER_FIELD, limits, |drawable| {
            nested(drawable, DRAWABLE_GEOMETRY_FIELD, limits, |geometry| {
                set_geometry(geometry, replacement, limits)
            })
        })?;
        chain(after_geometry, shape, |base| {
            nested(base, SHAPE_PATH_SOURCE_FIELD, limits, |path_source| {
                nested(path_source, PATH_BEZIER_FIELD, limits, |bezier| {
                    let after_size = nested(bezier, BEZIER_NATURAL_SIZE_FIELD, limits, |size| {
                        set_present_float(size, X_FIELD, width, limits)
                    })?;
                    chain(after_size, bezier, |base| {
                        nested(base, BEZIER_PATH_FIELD, limits, |path| {
                            set_path_width(path, width, limits)
                        })
                    })
                })
            })
        })
    })?;
    Ok(Rewrite::from_cow(match next {
        Some(bytes) => Cow::Owned(bytes),
        None => Cow::Borrowed(shape_info),
    }))
}

#[cfg(test)]
mod tests;
