//! The preset a shape's path source draws (record 0372; a port of the
//! host's `shapes::path::shape_preset`, read by `slide_shape_preset` and
//! `slide_shapes`): a rectangle or ellipse Bézier path (the leaf's or
//! Keynote's native ellipse), a rounded rectangle or regular polygon scalar
//! path, or an arrow or star point path. Any other path is no preset.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::path::{InnerRadiusRatio, PolygonSides, StarPoints};

use super::super::structure::wire::Fields;
use super::{Error, Preset, lift};

const SHAPE_ARCHIVE: u32 = 1;
const PATH_SOURCE: u32 = 3;
const POINT_SOURCE: u32 = 3;
const SCALAR_SOURCE: u32 = 4;
const BEZIER_SOURCE: u32 = 5;
const NORMALIZED_RECTANGLE_EXTENT: f32 = 100.0;
const ELLIPSE_CONTROL_FACTOR: f32 = 0.552_284_8;
const PATH_COMPARISON_EPSILON: f32 = 0.000_1;
const NATIVE_ELLIPSE_DIAGONAL: f32 = 0.353_553_4;
const NATIVE_ELLIPSE_OUTER_CONTROL: f32 = 0.548_815_55;
const NATIVE_ELLIPSE_INNER_CONTROL: f32 = 0.158_291_24;
const SINGLE_ARROW_HEAD_START_RATIO: f32 = 0.64;
const DOUBLE_ARROW_HEAD_END_RATIO: f32 = 0.40;
const ARROW_SHAFT_RATIO: f32 = 0.34;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const CURVE_TO: u64 = 4;
const CLOSE_SUBPATH: u64 = 5;

type Element = (u64, Vec<(f32, f32)>);

fn message(fields: &Fields, number: u32) -> Result<Option<&[u8]>, Error> {
    match fields.unique(number).map_err(lift)? {
        Some(field) => Ok(Some(field.message().map_err(lift)?)),
        None => Ok(None),
    }
}

fn float(fields: &Fields, number: u32) -> Result<Option<f32>, Error> {
    match fields.unique(number).map_err(lift)? {
        Some(field) => Ok(Some(field.float().map_err(lift)?)),
        None => Ok(None),
    }
}

fn varint(fields: &Fields, number: u32) -> Result<Option<u64>, Error> {
    match fields.unique(number).map_err(lift)? {
        Some(field) => Ok(Some(field.varint().map_err(lift)?)),
        None => Ok(None),
    }
}

/// A `TSP.Point` or `TSP.Size` (two floats, fields 1 and 2).
fn pair(payload: &[u8], wire: WireLimits) -> Result<(f32, f32), Error> {
    let fields = Fields::parse(payload, wire).map_err(lift)?;
    Ok((
        float(&fields, 1)?.ok_or(Error::InvalidSource)?,
        float(&fields, 2)?.ok_or(Error::InvalidSource)?,
    ))
}

fn elements(path: &[u8], wire: WireLimits) -> Result<Vec<Element>, Error> {
    let fields = Fields::parse(path, wire).map_err(lift)?;
    let mut found = Vec::new();
    for element in fields.all(1) {
        let element = Fields::parse(element.message().map_err(lift)?, wire).map_err(lift)?;
        let kind = varint(&element, 1)?.ok_or(Error::InvalidSource)?;
        let mut points = Vec::new();
        for point in element.all(2) {
            points.push(pair(point.message().map_err(lift)?, wire)?);
        }
        found.push((kind, points));
    }
    Ok(found)
}

fn nearly_equal(left: f32, right: f32, extent: f32) -> bool {
    (left - right).abs() <= extent.max(1.0) * PATH_COMPARISON_EPSILON
}

fn matches(actual: &[Element], expected: &[Element], extent: f32) -> bool {
    actual.len() == expected.len()
        && actual.iter().zip(expected).all(|(actual, expected)| {
            actual.0 == expected.0
                && actual.1.len() == expected.1.len()
                && actual
                    .1
                    .iter()
                    .zip(&expected.1)
                    .all(|(a, e)| nearly_equal(a.0, e.0, extent) && nearly_equal(a.1, e.1, extent))
        })
}

fn rectangle(actual: &[Element], width: f32, height: f32) -> bool {
    let expected: [Element; 6] = [
        (MOVE_TO, vec![(0.0, 0.0)]),
        (LINE_TO, vec![(width, 0.0)]),
        (LINE_TO, vec![(width, height)]),
        (LINE_TO, vec![(0.0, height)]),
        (CLOSE_SUBPATH, vec![]),
        (MOVE_TO, vec![(0.0, 0.0)]),
    ];
    // Exact, as the host compares rectangle corners.
    actual == expected.as_slice()
}

fn ellipse(actual: &[Element], width: f32, height: f32) -> bool {
    let (cx, cy) = (width / 2.0, height / 2.0);
    let (kx, ky) = (cx * ELLIPSE_CONTROL_FACTOR, cy * ELLIPSE_CONTROL_FACTOR);
    let authored: [Element; 7] = [
        (MOVE_TO, vec![(cx, 0.0)]),
        (
            CURVE_TO,
            vec![(cx + kx, 0.0), (width, cy - ky), (width, cy)],
        ),
        (
            CURVE_TO,
            vec![(width, cy + ky), (cx + kx, height), (cx, height)],
        ),
        (CURVE_TO, vec![(cx - kx, height), (0.0, cy + ky), (0.0, cy)]),
        (CURVE_TO, vec![(0.0, cy - ky), (cx - kx, 0.0), (cx, 0.0)]),
        (CLOSE_SUBPATH, vec![]),
        (MOVE_TO, vec![(cx, 0.0)]),
    ];
    let (dx, dy) = (
        width * NATIVE_ELLIPSE_DIAGONAL,
        height * NATIVE_ELLIPSE_DIAGONAL,
    );
    let (ox, oy) = (
        width * NATIVE_ELLIPSE_OUTER_CONTROL,
        height * NATIVE_ELLIPSE_OUTER_CONTROL,
    );
    let (ix, iy) = (
        width * NATIVE_ELLIPSE_INNER_CONTROL,
        height * NATIVE_ELLIPSE_INNER_CONTROL,
    );
    let start = (cx + dx, cy - dy);
    let native: [Element; 7] = [
        (MOVE_TO, vec![start]),
        (
            CURVE_TO,
            vec![(cx + ox, cy - iy), (cx + ox, cy + iy), (cx + dx, cy + dy)],
        ),
        (
            CURVE_TO,
            vec![(cx + ix, cy + oy), (cx - ix, cy + oy), (cx - dx, cy + dy)],
        ),
        (
            CURVE_TO,
            vec![(cx - ox, cy + iy), (cx - ox, cy - iy), (cx - dx, cy - dy)],
        ),
        (
            CURVE_TO,
            vec![(cx - ix, cy - oy), (cx + ix, cy - oy), start],
        ),
        (CLOSE_SUBPATH, vec![]),
        (MOVE_TO, vec![start]),
    ];
    let extent = width.max(height);
    matches(actual, &authored, extent) || matches(actual, &native, extent)
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "checked integral and in range first"
)]
fn integral_u8(value: f32) -> Option<u8> {
    (value.is_finite()
        && value.fract() == 0.0
        && value >= f32::from(u8::MIN)
        && value <= f32::from(u8::MAX))
    .then_some(value as u8)
}

fn refused() -> Error {
    Error::InvalidSource
}

/// The preset of one `TSWP.ShapeInfoArchive` payload (`None`: a path no
/// preset draws).
pub(super) fn preset(shape_info: &[u8], wire: WireLimits) -> Result<Option<Preset>, Error> {
    let info = Fields::parse(shape_info, wire).map_err(lift)?;
    let shape = message(&info, SHAPE_ARCHIVE)?.ok_or(Error::InvalidSource)?;
    let shape = Fields::parse(shape, wire).map_err(lift)?;
    let source = message(&shape, PATH_SOURCE)?.ok_or(Error::InvalidSource)?;
    let source = Fields::parse(source, wire).map_err(lift)?;
    if let Some(bezier) = message(&source, BEZIER_SOURCE)? {
        let bezier = Fields::parse(bezier, wire).map_err(lift)?;
        let (Some(size), Some(path)) = (message(&bezier, 2)?, message(&bezier, 3)?) else {
            return Ok(None);
        };
        let (width, height) = pair(size, wire)?;
        let elements = elements(path, wire)?;
        return Ok(
            if rectangle(&elements, width, height)
                || rectangle(
                    &elements,
                    NORMALIZED_RECTANGLE_EXTENT,
                    NORMALIZED_RECTANGLE_EXTENT,
                )
            {
                Some(Preset::Rectangle)
            } else if width > 0.0 && height > 0.0 && ellipse(&elements, width, height) {
                Some(Preset::Ellipse)
            } else {
                None
            },
        );
    }
    if let Some(scalar) = message(&source, SCALAR_SOURCE)? {
        let scalar = Fields::parse(scalar, wire).map_err(lift)?;
        return Ok(match varint(&scalar, 1)? {
            Some(0) => Some(Preset::RoundedRectangle {
                radius: float(&scalar, 2)?.ok_or(refused())?,
            }),
            Some(1) => {
                let sides = integral_u8(float(&scalar, 2)?.ok_or(refused())?).ok_or(refused())?;
                Some(Preset::RegularPolygon {
                    sides: PolygonSides::new(sides).map_err(|_| refused())?,
                })
            },
            _ => None,
        });
    }
    if let Some(point) = message(&source, POINT_SOURCE)? {
        let point = Fields::parse(point, wire).map_err(lift)?;
        let Some(kind) = varint(&point, 1)? else {
            return Ok(None);
        };
        let (x, y) = pair(message(&point, 2)?.ok_or(refused())?, wire)?;
        return Ok(match kind {
            0 | 1 | 10 => {
                let (width, _) = pair(message(&point, 3)?.ok_or(refused())?, wire)?;
                if !width.is_finite() || width <= 0.0 {
                    return Err(refused());
                }
                let expected = if kind == 10 {
                    DOUBLE_ARROW_HEAD_END_RATIO
                } else {
                    SINGLE_ARROW_HEAD_START_RATIO
                };
                if !nearly_equal(x / width, expected, 1.0)
                    || !nearly_equal(y, ARROW_SHAFT_RATIO, 1.0)
                {
                    return Ok(None);
                }
                Some(match kind {
                    0 => Preset::LeftArrow,
                    1 => Preset::RightArrow,
                    _ => Preset::DoubleArrow,
                })
            },
            100 => Some(Preset::Star {
                points: StarPoints::new(integral_u8(x).ok_or(refused())?).map_err(|_| refused())?,
                inner_radius: InnerRadiusRatio::new(y).map_err(|_| refused())?,
            }),
            _ => None,
        });
    }
    Ok(None)
}

impl super::Package {
    /// Read the preset one shape's path draws (host `slide_shape_preset`):
    /// `None` for a path no [`Preset`] draws, such as a line, a custom
    /// Bézier path or a callout.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`super::Refusal::Unsupported`]
    /// for drawables other than shapes and lines.
    pub fn slide_shape_preset<'s>(
        &self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
    ) -> Result<Option<Preset>, Error> {
        let selected = super::select_slide(self, slide.into())?;
        let inventory =
            super::read_inventory(self, selected.identifier, super::read_limits(self)?)?;
        let drawable = drawable.into();
        let item = inventory
            .items
            .get(drawable.as_index())
            .ok_or(Error::Refused(super::Refusal::DrawablePositionNotFound {
                position: drawable.as_position(),
            }))?;
        if !matches!(item.kind, super::Kind::Shape | super::Kind::Line) {
            return Err(Error::Refused(super::Refusal::Unsupported {
                reason: super::Unsupported::Kind { kind: item.kind },
            }));
        }
        let wire = self.wire_limits().map_err(|_| Error::InvalidSource)?;
        let workspace = super::Workspace::new(self).map_err(lift)?;
        let object = workspace.require(item.identifier).map_err(lift)?;
        let payload = &object
            .messages
            .get(item.message_index)
            .ok_or(Error::InvalidSource)?
            .data;
        preset(payload, wire)
    }
}
