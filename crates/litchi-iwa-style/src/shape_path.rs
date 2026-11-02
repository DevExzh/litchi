//! Preset shape paths (record 0413b): the `TSD.PathSourceArchive` a
//! [`Preset`] draws at a frame size, a straight line's path source, and the
//! preset an existing shape's path source draws, and the segment a
//! straight line draws.
//!
//! This is the legacy host's `shapes::path::{shape_path_source,
//! shape_preset}`. The Numbers sheet-drawable author (record 0378), the
//! Keynote slide-drawable author and preset reader (record 0372) and the
//! Pages body shapes all build and read their presets here.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::path::{
    CornerRadius, InnerRadiusRatio, PolygonSides, Preset, StarPoints,
};
use litchi_iwa_common::wire::{WireFieldView, WireView};

/// The host's native arrow proportions (`TSD.PointPathSourceArchive`).
const SINGLE_ARROW_HEAD_START_RATIO: f32 = 0.64;
const DOUBLE_ARROW_HEAD_END_RATIO: f32 = 0.40;
const ARROW_SHAFT_RATIO: f32 = 0.34;
const ELLIPSE_CONTROL_FACTOR: f32 = 0.552_284_8;
const NORMALIZED_RECTANGLE_EXTENT: f32 = 100.0;
const PATH_COMPARISON_EPSILON: f32 = 0.000_1;
const NATIVE_ELLIPSE_DIAGONAL: f32 = 0.353_553_4;
const NATIVE_ELLIPSE_OUTER_CONTROL: f32 = 0.548_815_55;
const NATIVE_ELLIPSE_INNER_CONTROL: f32 = 0.158_291_24;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const CURVE_TO: u64 = 4;
const CLOSE_SUBPATH: u64 = 5;
/// `TSWP.ShapeInfoArchive.super` and `TSD.ShapeArchive.pathsource`.
const SHAPE_ARCHIVE: u32 = 1;
const PATH_SOURCE: u32 = 3;
/// The path families of `TSD.PathSourceArchive`.
const POINT_SOURCE: u32 = 3;
const SCALAR_SOURCE: u32 = 4;
const BEZIER_SOURCE: u32 = 5;

/// A malformed path source or shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A payload is malformed or a stored value cannot be represented.
    #[error("invalid shape path: {reason}")]
    Invalid {
        /// What is wrong.
        reason: &'static str,
    },
}

type Result<T> = std::result::Result<T, Error>;

const fn invalid(reason: &'static str) -> Error {
    Error::Invalid { reason }
}

/// A minimal canonical protobuf writer.
#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn key(&mut self, number: u32, wire_type: u8) {
        self.raw(u64::from(number) << 3 | u64::from(wire_type));
    }

    fn raw(&mut self, mut value: u64) {
        loop {
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the value is masked to seven bits"
            )]
            let byte = (value & 0x7f) as u8;
            value >>= 7;
            if value == 0 {
                self.0.push(byte);
                return;
            }
            self.0.push(byte | 0x80);
        }
    }

    fn varint(mut self, number: u32, value: u64) -> Self {
        self.key(number, 0);
        self.raw(value);
        self
    }

    fn boolean(self, number: u32, value: bool) -> Self {
        self.varint(number, u64::from(value))
    }

    fn float(mut self, number: u32, value: f32) -> Self {
        self.key(number, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn message(mut self, number: u32, inner: &Self) -> Self {
        self.key(number, 2);
        self.raw(inner.0.len() as u64);
        self.0.extend_from_slice(&inner.0);
        self
    }
}

fn point(x: f32, y: f32) -> Writer {
    Writer::default().float(1, x).float(2, y)
}

fn element(kind: u64, points: &[(f32, f32)]) -> Writer {
    let mut element = Writer::default().varint(1, kind);
    for (x, y) in points {
        element = element.message(2, &point(*x, *y));
    }
    element
}

fn path(elements: &[Writer]) -> Writer {
    let mut path = Writer::default();
    for item in elements {
        path = path.message(1, item);
    }
    path
}

fn arrow(kind: u64, head_ratio: f32, width: f32, size: &Writer) -> Writer {
    Writer::default()
        .varint(1, kind)
        .message(2, &point(width * head_ratio, ARROW_SHAFT_RATIO))
        .message(3, size)
}

/// The `TSD.PathSourceArchive` `preset` draws in a `width` × `height`
/// frame (the host's `shape_path_source`); `None` for a rounded-rectangle
/// corner radius above half the shorter side.
#[must_use]
pub fn path_source(preset: Preset, width: f32, height: f32) -> Option<Vec<u8>> {
    let size = point(width, height);
    let source = Writer::default().boolean(1, false).boolean(2, false);
    let source = match preset {
        Preset::Rectangle => source.message(
            BEZIER_SOURCE,
            &Writer::default().message(2, &size).message(
                3,
                &path(&[
                    element(MOVE_TO, &[(0.0, 0.0)]),
                    element(LINE_TO, &[(width, 0.0)]),
                    element(LINE_TO, &[(width, height)]),
                    element(LINE_TO, &[(0.0, height)]),
                    element(CLOSE_SUBPATH, &[]),
                    element(MOVE_TO, &[(0.0, 0.0)]),
                ]),
            ),
        ),
        Preset::RoundedRectangle { corner_radius } => {
            if corner_radius.points() > width.min(height) / 2.0 {
                return None;
            }
            source.message(
                SCALAR_SOURCE,
                &Writer::default()
                    .varint(1, 0)
                    .float(2, corner_radius.points())
                    .message(3, &size)
                    .boolean(4, true),
            )
        },
        Preset::Ellipse => {
            let (cx, cy) = (width / 2.0, height / 2.0);
            let (kx, ky) = (cx * ELLIPSE_CONTROL_FACTOR, cy * ELLIPSE_CONTROL_FACTOR);
            source.message(
                BEZIER_SOURCE,
                &Writer::default().message(2, &size).message(
                    3,
                    &path(&[
                        element(MOVE_TO, &[(cx, 0.0)]),
                        element(CURVE_TO, &[(cx + kx, 0.0), (width, cy - ky), (width, cy)]),
                        element(
                            CURVE_TO,
                            &[(width, cy + ky), (cx + kx, height), (cx, height)],
                        ),
                        element(CURVE_TO, &[(cx - kx, height), (0.0, cy + ky), (0.0, cy)]),
                        element(CURVE_TO, &[(0.0, cy - ky), (cx - kx, 0.0), (cx, 0.0)]),
                        element(CLOSE_SUBPATH, &[]),
                        element(MOVE_TO, &[(cx, 0.0)]),
                    ]),
                ),
            )
        },
        Preset::LeftArrow => source.message(
            POINT_SOURCE,
            &arrow(0, SINGLE_ARROW_HEAD_START_RATIO, width, &size),
        ),
        Preset::RightArrow => source.message(
            POINT_SOURCE,
            &arrow(1, SINGLE_ARROW_HEAD_START_RATIO, width, &size),
        ),
        Preset::DoubleArrow => source.message(
            POINT_SOURCE,
            &arrow(10, DOUBLE_ARROW_HEAD_END_RATIO, width, &size),
        ),
        Preset::RegularPolygon { sides } => source.message(
            SCALAR_SOURCE,
            &Writer::default()
                .varint(1, 1)
                .float(2, f32::from(sides.count()))
                .message(3, &size)
                .boolean(4, false),
        ),
        Preset::Star {
            points,
            inner_radius,
        } => source.message(
            POINT_SOURCE,
            &Writer::default()
                .varint(1, 100)
                .message(2, &point(f32::from(points.count()), inner_radius.ratio()))
                .message(3, &size),
        ),
    };
    Some(source.0)
}

/// The `TSD.PathSourceArchive` of a native straight line of `length`
/// points (a two-element Bézier path).
#[must_use]
pub fn line_path_source(length: f32) -> Vec<u8> {
    Writer::default()
        .boolean(1, false)
        .boolean(2, false)
        .message(
            BEZIER_SOURCE,
            &Writer::default().message(2, &point(length, 0.0)).message(
                3,
                &path(&[
                    element(MOVE_TO, &[(0.0, 0.0)]),
                    element(LINE_TO, &[(length, 0.0)]),
                ]),
            ),
        )
        .0
}

type Element = (u64, Vec<(f32, f32)>);

fn unique<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> Result<Option<WireFieldView<'a>>> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type || found.replace(field).is_some() {
            return Err(invalid("a singular field is repeated or mistyped"));
        }
    }
    Ok(found)
}

fn parse(payload: &[u8], limits: WireLimits) -> Result<WireView<'_>> {
    WireView::parse(payload, limits).map_err(|_| invalid("malformed message"))
}

fn message<'a>(view: &WireView<'a>, number: u32) -> Result<Option<&'a [u8]>> {
    Ok(unique(view, number, 2)?.map(WireFieldView::payload))
}

fn float(view: &WireView<'_>, number: u32) -> Result<Option<f32>> {
    unique(view, number, 5)?
        .map(|field| {
            field
                .payload()
                .try_into()
                .map(f32::from_le_bytes)
                .map_err(|_| invalid("a float is not four bytes"))
        })
        .transpose()
}

fn varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>> {
    unique(view, number, 0)?
        .map(|field| {
            let (value, length) =
                litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                    .map_err(|_| invalid("malformed varint"))?;
            if length == field.payload().len() {
                Ok(value)
            } else {
                Err(invalid("non-canonical varint"))
            }
        })
        .transpose()
}

/// A `TSP.Point` or `TSP.Size` (two floats, fields 1 and 2).
fn pair(payload: &[u8], limits: WireLimits) -> Result<(f32, f32)> {
    let view = parse(payload, limits)?;
    Ok((
        float(&view, 1)?.ok_or(invalid("a point lacks x"))?,
        float(&view, 2)?.ok_or(invalid("a point lacks y"))?,
    ))
}

fn elements(path: &[u8], limits: WireLimits) -> Result<Vec<Element>> {
    let view = parse(path, limits)?;
    let mut found = Vec::new();
    for element in view.fields().filter(|field| field.number() == 1) {
        if element.wire_type() != 2 {
            return Err(invalid("a path element is not a message"));
        }
        let element = parse(element.payload(), limits)?;
        let kind = varint(&element, 1)?.ok_or(invalid("a path element lacks its kind"))?;
        let mut points = Vec::new();
        for point in element.fields().filter(|field| field.number() == 2) {
            if point.wire_type() != 2 {
                return Err(invalid("a path point is not a message"));
            }
            points.push(pair(point.payload(), limits)?);
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

/// The preset the path of one `TSWP.ShapeInfoArchive` payload draws (the
/// host's `shape_preset`): a rectangle or ellipse Bézier path (the authored
/// or the native ellipse), a rounded rectangle or regular polygon scalar
/// path, or an arrow or star point path. `None` for any other path (a
/// line, a custom Bézier path, a callout).
///
/// # Errors
///
/// [`Error::Invalid`] for a malformed shape, path source or preset value.
pub fn preset(shape_info: &[u8], limits: WireLimits) -> Result<Option<Preset>> {
    let refused = || invalid("the path source holds an unrepresentable preset value");
    let info = parse(shape_info, limits)?;
    let shape = message(&info, SHAPE_ARCHIVE)?.ok_or(invalid("the shape info has no shape"))?;
    let shape = parse(shape, limits)?;
    let source = message(&shape, PATH_SOURCE)?.ok_or(invalid("the shape has no path source"))?;
    let source = parse(source, limits)?;
    if let Some(bezier) = message(&source, BEZIER_SOURCE)? {
        let bezier = parse(bezier, limits)?;
        let (Some(size), Some(path)) = (message(&bezier, 2)?, message(&bezier, 3)?) else {
            return Ok(None);
        };
        let (width, height) = pair(size, limits)?;
        let elements = elements(path, limits)?;
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
        let scalar = parse(scalar, limits)?;
        return Ok(match varint(&scalar, 1)? {
            Some(0) => Some(Preset::RoundedRectangle {
                corner_radius: CornerRadius::new(float(&scalar, 2)?.ok_or_else(refused)?)
                    .map_err(|_| refused())?,
            }),
            Some(1) => {
                let sides =
                    integral_u8(float(&scalar, 2)?.ok_or_else(refused)?).ok_or_else(refused)?;
                Some(Preset::RegularPolygon {
                    sides: PolygonSides::new(sides).map_err(|_| refused())?,
                })
            },
            _ => None,
        });
    }
    if let Some(point) = message(&source, POINT_SOURCE)? {
        let point = parse(point, limits)?;
        let Some(kind) = varint(&point, 1)? else {
            return Ok(None);
        };
        let (x, y) = pair(message(&point, 2)?.ok_or_else(refused)?, limits)?;
        return Ok(match kind {
            0 | 1 | 10 => {
                let (width, _) = pair(message(&point, 3)?.ok_or_else(refused)?, limits)?;
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
                points: StarPoints::new(integral_u8(x).ok_or_else(refused)?)
                    .map_err(|_| refused())?,
                inner_radius: InnerRadiusRatio::new(y).map_err(|_| refused())?,
            }),
            _ => None,
        });
    }
    Ok(None)
}

/// The segment a native straight line draws and whether its stored path
/// runs backwards (the host's `shape_line_segment`); `None` for a shape
/// that is not a straight line.
///
/// # Errors
///
/// [`Error::Invalid`] for a malformed shape or a line without a complete
/// frame.
pub fn line_segment(
    shape_info: &[u8],
) -> Result<Option<(litchi_iwa_common::shape::line::Segment, bool)>> {
    use litchi_iwa_common::shape::geometry::{Point, Size};
    use litchi_iwa_protos::shape_style_codec as line_codec;
    let Some(stored) = line_codec::line(shape_info, line_codec::Limits::DEFAULT)
        .map_err(|_error| invalid("the line path is malformed"))?
    else {
        return Ok(None);
    };
    let [x, y] = stored
        .geometry
        .position
        .ok_or(invalid("a line has no position"))?;
    let [width, height] = stored.geometry.size.ok_or(invalid("a line has no size"))?;
    let segment = litchi_iwa_common::shape::line::Segment::from_geometry(
        Point { x, y },
        Size { width, height },
        stored
            .geometry
            .angle
            .ok_or(invalid("a line has no angle"))?,
        stored.natural_width,
    )
    .map_err(|_error| invalid("a line's frame is not a segment"))?;
    Ok(Some(if stored.horizontal_flip {
        (segment.reversed(), true)
    } else {
        (segment, false)
    }))
}

/// Whether two line segments agree up to the `f32` rounding of a frame
/// round trip (how a written segment reads back).
#[must_use]
pub fn segments_match(
    left: litchi_iwa_common::shape::line::Segment,
    right: litchi_iwa_common::shape::line::Segment,
) -> bool {
    let close = |a: f32, b: f32| (a - b).abs() <= 1e-3 * a.abs().max(b.abs()).max(1.0);
    close(left.start().x, right.start().x)
        && close(left.start().y, right.start().y)
        && close(left.end().x, right.end().x)
        && close(left.end().y, right.end().y)
}

/// Replace the path family of a `TSWP.ShapeInfoArchive` payload's path
/// source with `preset`'s at `width` × `height`, keeping every other field
/// (the host's `set_shape_preset`). `None` when the preset does not fit
/// the frame.
///
/// # Errors
///
/// [`Error::Invalid`] for a malformed shape or a path source with several
/// path families.
pub fn with_preset(
    shape_info: &[u8],
    preset: Preset,
    width: f32,
    height: f32,
    limits: WireLimits,
) -> Result<Option<Vec<u8>>> {
    let Some(replacement) = path_source(preset, width, height) else {
        return Ok(None);
    };
    let replacement_view = parse(&replacement, limits)?;
    let family = replacement_view
        .fields()
        .find(|field| (POINT_SOURCE..=8).contains(&field.number()))
        .ok_or(invalid("the preset path has no family"))?
        .raw()
        .to_vec();
    let rewrite = |data: &[u8], number: u32, edit: &dyn Fn(&[u8]) -> Result<Vec<u8>>| {
        let view = parse(data, limits)?;
        let mut output = Vec::with_capacity(data.len() + family.len());
        let mut seen = false;
        for field in view.fields() {
            if field.number() == number {
                if field.wire_type() != 2 || seen {
                    return Err(invalid("a singular message is repeated or mistyped"));
                }
                seen = true;
                let inner = edit(field.payload())?;
                let mut header = Writer::default();
                header.key(number, 2);
                header.raw(inner.len() as u64);
                output.extend_from_slice(&header.0);
                output.extend_from_slice(&inner);
            } else {
                output.extend_from_slice(field.raw());
            }
        }
        if seen {
            Ok(output)
        } else {
            Err(invalid("a required message is missing"))
        }
    };
    let source_edit = |source: &[u8]| -> Result<Vec<u8>> {
        let view = parse(source, limits)?;
        if view
            .fields()
            .filter(|field| (POINT_SOURCE..=8).contains(&field.number()))
            .count()
            > 1
        {
            return Err(invalid("the path source has several families"));
        }
        let mut output = Vec::with_capacity(source.len() + family.len());
        let mut inserted = false;
        for field in view.fields() {
            if (POINT_SOURCE..=8).contains(&field.number()) {
                if !inserted {
                    output.extend_from_slice(&family);
                    inserted = true;
                }
                continue;
            }
            output.extend_from_slice(field.raw());
        }
        if !inserted {
            output.extend_from_slice(&family);
        }
        Ok(output)
    };
    let shape_edit = |shape: &[u8]| rewrite(shape, PATH_SOURCE, &source_edit);
    rewrite(shape_info, SHAPE_ARCHIVE, &shape_edit).map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shape_info(source: &[u8]) -> Vec<u8> {
        let mut shape = Writer::default();
        shape.key(PATH_SOURCE, 2);
        shape.raw(source.len() as u64);
        shape.0.extend_from_slice(source);
        let mut info = Writer::default();
        info.key(SHAPE_ARCHIVE, 2);
        info.raw(shape.0.len() as u64);
        info.0.extend_from_slice(&shape.0);
        info.0
    }

    #[test]
    fn every_preset_reads_back() {
        let presets = [
            Preset::Rectangle,
            Preset::ROUNDED_RECTANGLE,
            Preset::Ellipse,
            Preset::LeftArrow,
            Preset::RightArrow,
            Preset::DoubleArrow,
            Preset::PENTAGON,
            Preset::Star {
                points: StarPoints::new(5).unwrap(),
                inner_radius: InnerRadiusRatio::new(0.5).unwrap(),
            },
        ];
        for preset in presets {
            let source = path_source(preset, 120.0, 80.0).unwrap();
            assert_eq!(
                super::preset(&shape_info(&source), WireLimits::DEFAULT).unwrap(),
                Some(preset)
            );
        }
        let line = line_path_source(50.0);
        assert_eq!(
            preset(&shape_info(&line), WireLimits::DEFAULT).unwrap(),
            None
        );
    }

    #[test]
    fn a_large_corner_radius_does_not_fit() {
        let preset = Preset::RoundedRectangle {
            corner_radius: CornerRadius::new(30.0).unwrap(),
        };
        assert!(path_source(preset, 50.0, 50.0).is_none());
        assert!(path_source(preset, 60.0, 60.0).is_some());
    }

    #[test]
    fn with_preset_replaces_only_the_family() {
        let info = shape_info(&path_source(Preset::Rectangle, 40.0, 30.0).unwrap());
        let changed = with_preset(&info, Preset::Ellipse, 40.0, 30.0, WireLimits::DEFAULT)
            .unwrap()
            .unwrap();
        assert_eq!(
            preset(&changed, WireLimits::DEFAULT).unwrap(),
            Some(Preset::Ellipse)
        );
    }
}
