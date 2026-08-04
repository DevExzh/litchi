//! Strict typed construction and classification for iWork shape paths.

use prost::Message;

use super::geometry::DrawableSize;
use super::line::is_straight_line_bezier;
use crate::IWorkPackage;
use crate::archive::RawMessage;
use crate::protobuf::{tsd, tsp, tswp};
use crate::wire::{patch_length_delimited_field, transform_length_delimited_field};
use crate::{Error, Result};

const NORMALIZED_RECTANGLE_EXTENT: f32 = 100.0;
const MINIMUM_POLYGON_SIDES: u8 = 3;
const MINIMUM_STAR_POINTS: u8 = 3;
const ELLIPSE_CONTROL_FACTOR: f32 = 0.552_284_8;
const PATH_COMPARISON_EPSILON: f32 = 0.000_1;
const NATIVE_ELLIPSE_DIAGONAL: f32 = 0.353_553_4;
const NATIVE_ELLIPSE_OUTER_CONTROL: f32 = 0.548_815_55;
const NATIVE_ELLIPSE_INNER_CONTROL: f32 = 0.158_291_24;
const NATIVE_SINGLE_ARROW_HEAD_START_RATIO: f32 = 0.64;
const NATIVE_DOUBLE_ARROW_HEAD_END_RATIO: f32 = 0.40;
const NATIVE_ARROW_SHAFT_RATIO: f32 = 0.34;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_INFO_SHAPE_FIELD: u32 = 1;
const SHAPE_ARCHIVE_PATH_SOURCE_FIELD: u32 = 3;

/// Corner radius in the path's natural coordinate system.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeCornerRadius(f32);

impl ShapeCornerRadius {
    /// Native default used by iWork for a 100-point rounded rectangle.
    pub const DEFAULT: Self = Self(15.0);

    /// Validate a corner radius measured in path points.
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ParseError(
                "iWork shape corner radius must be finite and non-negative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the radius in path points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Valid number of sides for an iWork regular polygon.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ShapePolygonSides(u8);

impl ShapePolygonSides {
    /// Three-sided regular polygon.
    pub const TRIANGLE: Self = Self(3);
    /// Five-sided regular polygon used by iWork's Pentagon preset.
    pub const PENTAGON: Self = Self(5);

    /// Validate a regular-polygon side count.
    pub fn new(sides: u8) -> Result<Self> {
        if sides < MINIMUM_POLYGON_SIDES {
            return Err(Error::ParseError(format!(
                "iWork regular polygons require at least {MINIMUM_POLYGON_SIDES} sides"
            )));
        }
        Ok(Self(sides))
    }

    /// Return the side count.
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// Valid number of outer points for an iWork star.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ShapeStarPoints(u8);

impl ShapeStarPoints {
    /// Native five-point star control.
    pub const FIVE: Self = Self(5);

    /// Validate a star point count.
    pub fn new(points: u8) -> Result<Self> {
        if points < MINIMUM_STAR_POINTS {
            return Err(Error::ParseError(format!(
                "iWork stars require at least {MINIMUM_STAR_POINTS} points"
            )));
        }
        Ok(Self(points))
    }

    /// Return the outer point count.
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// Ratio between an iWork star's inner and outer radii.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeStarInnerRatio(f32);

impl ShapeStarInnerRatio {
    /// Native default used by iWork's five-point star preset.
    pub const DEFAULT: Self = Self(0.382);

    /// Validate an inner-radius ratio.
    pub fn new(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() || !(0.0..1.0).contains(&ratio) {
            return Err(Error::ParseError(
                "iWork star inner-radius ratio must be finite and in [0, 1)".to_owned(),
            ));
        }
        Ok(Self(ratio))
    }

    /// Return the inner-radius ratio.
    pub const fn get(self) -> f32 {
        self.0
    }
}

/// Source-buildable iWork shape preset.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ShapePreset {
    /// Four-corner Bézier rectangle.
    Rectangle,
    /// Native rounded rectangle with an explicit corner radius.
    RoundedRectangle { corner_radius: ShapeCornerRadius },
    /// Four-segment cubic Bézier ellipse.
    Ellipse,
    /// Native left-facing single arrow with standard head and shaft proportions.
    LeftArrow,
    /// Native right-facing single arrow with standard head and shaft proportions.
    RightArrow,
    /// Native bidirectional arrow with standard head and shaft proportions.
    DoubleArrow,
    /// Native regular polygon with a configurable side count.
    RegularPolygon { sides: ShapePolygonSides },
    /// Native star with configurable point count and inner radius.
    Star {
        points: ShapeStarPoints,
        inner_radius: ShapeStarInnerRatio,
    },
}

impl ShapePreset {
    /// Native default rounded rectangle.
    pub const ROUNDED_RECTANGLE: Self = Self::RoundedRectangle {
        corner_radius: ShapeCornerRadius::DEFAULT,
    };
    /// Native five-sided Pentagon preset.
    pub const PENTAGON: Self = Self::RegularPolygon {
        sides: ShapePolygonSides::PENTAGON,
    };
    /// Native five-point Star preset.
    pub const STAR: Self = Self::Star {
        points: ShapeStarPoints::FIVE,
        inner_radius: ShapeStarInnerRatio::DEFAULT,
    };
}

/// Structural path family used by an ordinary iWork shape.
///
/// Native preset families are distinguished from arbitrary paths so typed
/// source-built shapes can round-trip without relying on localized names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ShapePathKind {
    /// Native two-point straight line.
    Line,
    Rectangle,
    RoundedRectangle,
    Ellipse,
    RegularPolygon,
    Chevron,
    LeftArrow,
    RightArrow,
    DoubleArrow,
    Star,
    Plus,
    BezierPath,
    PointPath,
    ScalarPath,
    Callout,
    ConnectionLine,
    EditableBezierPath,
}

pub(crate) fn shape_path_kind(shape: &tswp::ShapeInfoArchive) -> Result<ShapePathKind> {
    let path = shape.super_.pathsource.as_ref().ok_or_else(|| {
        Error::InvalidFormat("ordinary iWork shape has no path source".to_owned())
    })?;
    let families = [
        path.point_path_source.is_some(),
        path.scalar_path_source.is_some(),
        path.bezier_path_source.is_some(),
        path.callout_path_source.is_some(),
        path.connection_line_path_source.is_some(),
        path.editable_bezier_path_source.is_some(),
    ]
    .into_iter()
    .filter(|present| *present)
    .count();
    if families != 1 {
        return Err(Error::InvalidFormat(format!(
            "ordinary iWork shape must have exactly one path family, found {families}"
        )));
    }
    if let Some(bezier) = &path.bezier_path_source {
        return Ok(if is_straight_line_bezier(bezier) {
            ShapePathKind::Line
        } else if is_rectangle_path(bezier) {
            ShapePathKind::Rectangle
        } else if is_ellipse_path(bezier) {
            ShapePathKind::Ellipse
        } else {
            ShapePathKind::BezierPath
        });
    }
    if let Some(point) = &path.point_path_source {
        Ok(point_path_kind(point))
    } else if let Some(scalar) = &path.scalar_path_source {
        Ok(scalar_path_kind(scalar))
    } else if path.callout_path_source.is_some() {
        Ok(ShapePathKind::Callout)
    } else if path.connection_line_path_source.is_some() {
        Ok(ShapePathKind::ConnectionLine)
    } else {
        Ok(ShapePathKind::EditableBezierPath)
    }
}

pub(crate) fn shape_preset(shape: &tswp::ShapeInfoArchive) -> Result<Option<ShapePreset>> {
    let path = shape.super_.pathsource.as_ref().ok_or_else(|| {
        Error::InvalidFormat("ordinary iWork shape has no path source".to_owned())
    })?;
    if let Some(bezier) = &path.bezier_path_source {
        return Ok(if is_rectangle_path(bezier) {
            Some(ShapePreset::Rectangle)
        } else if is_ellipse_path(bezier) {
            Some(ShapePreset::Ellipse)
        } else {
            None
        });
    }
    if let Some(scalar) = &path.scalar_path_source {
        use tsd::scalar_path_source_archive::ScalarPathSourceType;
        return match scalar
            .r#type
            .and_then(|value| ScalarPathSourceType::try_from(value).ok())
        {
            Some(ScalarPathSourceType::KTsdRoundedRectangle) => {
                let radius = scalar.scalar.ok_or_else(|| {
                    Error::InvalidFormat("rounded rectangle has no corner radius".to_owned())
                })?;
                Ok(Some(ShapePreset::RoundedRectangle {
                    corner_radius: ShapeCornerRadius::new(radius)?,
                }))
            },
            Some(ScalarPathSourceType::KTsdRegularPolygon) => {
                let sides = scalar.scalar.ok_or_else(|| {
                    Error::InvalidFormat("regular polygon has no side count".to_owned())
                })?;
                let sides = integral_u8(sides).ok_or_else(|| {
                    Error::InvalidFormat(
                        "regular polygon side count is not an unsigned integer".to_owned(),
                    )
                })?;
                Ok(Some(ShapePreset::RegularPolygon {
                    sides: ShapePolygonSides::new(sides)?,
                }))
            },
            _ => Ok(None),
        };
    }
    if let Some(point) = &path.point_path_source {
        return point_shape_preset(point);
    }
    Ok(None)
}

pub(crate) fn shape_path_source(
    preset: ShapePreset,
    natural_size: DrawableSize,
) -> Result<tsd::PathSourceArchive> {
    validate_natural_size(natural_size)?;
    let size = tsp::Size {
        width: natural_size.width,
        height: natural_size.height,
    };
    let mut source = tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        ..Default::default()
    };
    match preset {
        ShapePreset::Rectangle => {
            source.bezier_path_source = Some(tsd::BezierPathSourceArchive {
                natural_size: Some(size),
                path: Some(rectangle_path(natural_size)),
                ..Default::default()
            });
        },
        ShapePreset::RoundedRectangle { corner_radius } => {
            let maximum = natural_size.width.min(natural_size.height) / 2.0;
            if corner_radius.points() > maximum {
                return Err(Error::ParseError(format!(
                    "iWork shape corner radius {} exceeds the maximum {maximum}",
                    corner_radius.points()
                )));
            }
            source.scalar_path_source = Some(tsd::ScalarPathSourceArchive {
                r#type: Some(
                    tsd::scalar_path_source_archive::ScalarPathSourceType::KTsdRoundedRectangle
                        as i32,
                ),
                scalar: Some(corner_radius.points()),
                natural_size: Some(size),
                is_curve_continuous: Some(true),
            });
        },
        ShapePreset::Ellipse => {
            source.bezier_path_source = Some(tsd::BezierPathSourceArchive {
                natural_size: Some(size),
                path: Some(ellipse_path(natural_size)),
                ..Default::default()
            });
        },
        ShapePreset::LeftArrow => {
            source.point_path_source = Some(arrow_path_source(
                tsd::point_path_source_archive::PointPathSourceType::KTsdLeftSingleArrow,
                NATIVE_SINGLE_ARROW_HEAD_START_RATIO,
                natural_size,
                size,
            ));
        },
        ShapePreset::RightArrow => {
            source.point_path_source = Some(arrow_path_source(
                tsd::point_path_source_archive::PointPathSourceType::KTsdRightSingleArrow,
                NATIVE_SINGLE_ARROW_HEAD_START_RATIO,
                natural_size,
                size,
            ));
        },
        ShapePreset::DoubleArrow => {
            source.point_path_source = Some(arrow_path_source(
                tsd::point_path_source_archive::PointPathSourceType::KTsdDoubleArrow,
                NATIVE_DOUBLE_ARROW_HEAD_END_RATIO,
                natural_size,
                size,
            ));
        },
        ShapePreset::RegularPolygon { sides } => {
            source.scalar_path_source = Some(tsd::ScalarPathSourceArchive {
                r#type: Some(
                    tsd::scalar_path_source_archive::ScalarPathSourceType::KTsdRegularPolygon
                        as i32,
                ),
                scalar: Some(f32::from(sides.get())),
                natural_size: Some(size),
                is_curve_continuous: Some(false),
            });
        },
        ShapePreset::Star {
            points,
            inner_radius,
        } => {
            source.point_path_source = Some(tsd::PointPathSourceArchive {
                r#type: Some(tsd::point_path_source_archive::PointPathSourceType::KTsdStar as i32),
                point: Some(tsp::Point {
                    x: f32::from(points.get()),
                    y: inner_radius.get(),
                }),
                natural_size: Some(size),
            });
        },
    }
    Ok(source)
}

pub(crate) fn set_shape_preset(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    preset: ShapePreset,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork drawable object {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let data = patch_shape_preset(original, drawable_id, preset)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn patch_shape_preset(data: &[u8], drawable_id: u64, preset: ShapePreset) -> Result<Vec<u8>> {
    let shape = tswp::ShapeInfoArchive::decode(data)?;
    let geometry = shape
        .super_
        .super_
        .geometry
        .as_ref()
        .and_then(|geometry| geometry.size.as_ref())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} has no natural path size"
            ))
        })?;
    let mut replacement = shape_path_source(
        preset,
        DrawableSize {
            width: geometry.width,
            height: geometry.height,
        },
    )?;
    if let Some(path_source) = &shape.super_.pathsource {
        replacement.horizontal_flip = path_source.horizontal_flip;
        replacement.vertical_flip = path_source.vertical_flip;
    }
    let replacement = replacement.encode_to_vec();
    let data = transform_length_delimited_field(data, SHAPE_INFO_SHAPE_FIELD, |shape_archive| {
        let current = tsd::ShapeArchive::decode(shape_archive)?;
        patch_length_delimited_field(
            shape_archive,
            SHAPE_ARCHIVE_PATH_SOURCE_FIELD,
            current.pathsource.is_some(),
            Some(&replacement),
        )
    })?;
    let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
    if shape_preset(&verified)? != Some(preset) {
        return Err(Error::InvalidFormat(
            "iWork shape preset patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn point_path_kind(point: &tsd::PointPathSourceArchive) -> ShapePathKind {
    use tsd::point_path_source_archive::PointPathSourceType;
    match point
        .r#type
        .and_then(|value| PointPathSourceType::try_from(value).ok())
    {
        Some(PointPathSourceType::KTsdLeftSingleArrow) => ShapePathKind::LeftArrow,
        Some(PointPathSourceType::KTsdRightSingleArrow) => ShapePathKind::RightArrow,
        Some(PointPathSourceType::KTsdDoubleArrow) => ShapePathKind::DoubleArrow,
        Some(PointPathSourceType::KTsdStar) => ShapePathKind::Star,
        Some(PointPathSourceType::KTsdPlus) => ShapePathKind::Plus,
        None => ShapePathKind::PointPath,
    }
}

fn point_shape_preset(point: &tsd::PointPathSourceArchive) -> Result<Option<ShapePreset>> {
    use tsd::point_path_source_archive::PointPathSourceType;

    let Some(kind) = point
        .r#type
        .and_then(|value| PointPathSourceType::try_from(value).ok())
    else {
        return Ok(None);
    };
    let control = point
        .point
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat(format!("{kind:?} path has no point control")))?;
    match kind {
        PointPathSourceType::KTsdLeftSingleArrow
        | PointPathSourceType::KTsdRightSingleArrow
        | PointPathSourceType::KTsdDoubleArrow => {
            let size = point.natural_size.as_ref().ok_or_else(|| {
                Error::InvalidFormat(format!("{kind:?} path has no natural size"))
            })?;
            if !size.width.is_finite() || size.width <= 0.0 {
                return Err(Error::InvalidFormat(format!(
                    "{kind:?} path has invalid natural width {}",
                    size.width
                )));
            }
            let horizontal_ratio = control.x / size.width;
            let expected_horizontal_ratio = match kind {
                PointPathSourceType::KTsdDoubleArrow => NATIVE_DOUBLE_ARROW_HEAD_END_RATIO,
                PointPathSourceType::KTsdLeftSingleArrow
                | PointPathSourceType::KTsdRightSingleArrow => NATIVE_SINGLE_ARROW_HEAD_START_RATIO,
                _ => unreachable!("matched arrow variants"),
            };
            if !nearly_equal(horizontal_ratio, expected_horizontal_ratio, 1.0)
                || !nearly_equal(control.y, NATIVE_ARROW_SHAFT_RATIO, 1.0)
            {
                return Ok(None);
            }
            Ok(Some(match kind {
                PointPathSourceType::KTsdLeftSingleArrow => ShapePreset::LeftArrow,
                PointPathSourceType::KTsdRightSingleArrow => ShapePreset::RightArrow,
                PointPathSourceType::KTsdDoubleArrow => ShapePreset::DoubleArrow,
                _ => unreachable!("matched arrow variants"),
            }))
        },
        PointPathSourceType::KTsdStar => {
            let points = integral_u8(control.x).ok_or_else(|| {
                Error::InvalidFormat("star point count is not an unsigned integer".to_owned())
            })?;
            Ok(Some(ShapePreset::Star {
                points: ShapeStarPoints::new(points)?,
                inner_radius: ShapeStarInnerRatio::new(control.y)?,
            }))
        },
        PointPathSourceType::KTsdPlus => Ok(None),
    }
}

fn arrow_path_source(
    kind: tsd::point_path_source_archive::PointPathSourceType,
    horizontal_ratio: f32,
    natural_size: DrawableSize,
    size: tsp::Size,
) -> tsd::PointPathSourceArchive {
    tsd::PointPathSourceArchive {
        r#type: Some(kind as i32),
        point: Some(tsp::Point {
            x: natural_size.width * horizontal_ratio,
            y: NATIVE_ARROW_SHAFT_RATIO,
        }),
        natural_size: Some(size),
    }
}

fn scalar_path_kind(scalar: &tsd::ScalarPathSourceArchive) -> ShapePathKind {
    use tsd::scalar_path_source_archive::ScalarPathSourceType;
    match scalar
        .r#type
        .and_then(|value| ScalarPathSourceType::try_from(value).ok())
    {
        Some(ScalarPathSourceType::KTsdRoundedRectangle) => ShapePathKind::RoundedRectangle,
        Some(ScalarPathSourceType::KTsdRegularPolygon) => ShapePathKind::RegularPolygon,
        Some(ScalarPathSourceType::KTsdChevron) => ShapePathKind::Chevron,
        None => ShapePathKind::ScalarPath,
    }
}

fn validate_natural_size(size: DrawableSize) -> Result<()> {
    if !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(
            "iWork shape natural size must be finite and positive".to_owned(),
        ));
    }
    Ok(())
}

fn rectangle_path(size: DrawableSize) -> tsp::Path {
    use tsp::path::ElementType;

    tsp::Path {
        elements: vec![
            path_element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
            path_element(ElementType::LineTo, vec![point(size.width, 0.0)]),
            path_element(ElementType::LineTo, vec![point(size.width, size.height)]),
            path_element(ElementType::LineTo, vec![point(0.0, size.height)]),
            path_element(ElementType::CloseSubpath, Vec::new()),
            path_element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
        ],
    }
}

fn ellipse_path(size: DrawableSize) -> tsp::Path {
    use tsp::path::ElementType;

    let center_x = size.width / 2.0;
    let center_y = size.height / 2.0;
    let control_x = center_x * ELLIPSE_CONTROL_FACTOR;
    let control_y = center_y * ELLIPSE_CONTROL_FACTOR;
    tsp::Path {
        elements: vec![
            path_element(ElementType::MoveTo, vec![point(center_x, 0.0)]),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x + control_x, 0.0),
                    point(size.width, center_y - control_y),
                    point(size.width, center_y),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(size.width, center_y + control_y),
                    point(center_x + control_x, size.height),
                    point(center_x, size.height),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x - control_x, size.height),
                    point(0.0, center_y + control_y),
                    point(0.0, center_y),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(0.0, center_y - control_y),
                    point(center_x - control_x, 0.0),
                    point(center_x, 0.0),
                ],
            ),
            path_element(ElementType::CloseSubpath, Vec::new()),
            path_element(ElementType::MoveTo, vec![point(center_x, 0.0)]),
        ],
    }
}

fn point(x: f32, y: f32) -> tsp::Point {
    tsp::Point { x, y }
}

fn path_element(r#type: tsp::path::ElementType, points: Vec<tsp::Point>) -> tsp::path::Element {
    tsp::path::Element {
        r#type: r#type as i32,
        points,
    }
}

fn integral_u8(value: f32) -> Option<u8> {
    if value.is_finite()
        && value.fract() == 0.0
        && value >= f32::from(u8::MIN)
        && value <= f32::from(u8::MAX)
    {
        Some(value as u8)
    } else {
        None
    }
}

fn is_rectangle_path(bezier: &tsd::BezierPathSourceArchive) -> bool {
    let Some(size) = bezier.natural_size.as_ref() else {
        return false;
    };
    let Some(path) = bezier.path.as_ref() else {
        return false;
    };
    rectangle_elements_match(path, size.width, size.height)
        || rectangle_elements_match(
            path,
            NORMALIZED_RECTANGLE_EXTENT,
            NORMALIZED_RECTANGLE_EXTENT,
        )
}

fn is_ellipse_path(bezier: &tsd::BezierPathSourceArchive) -> bool {
    let Some(size) = bezier.natural_size.as_ref() else {
        return false;
    };
    let Some(path) = bezier.path.as_ref() else {
        return false;
    };
    if size.width <= 0.0 || size.height <= 0.0 {
        return false;
    }
    ellipse_elements_match(path, size.width, size.height)
        || native_ellipse_elements_match(path, size.width, size.height)
}

fn ellipse_elements_match(path: &tsp::Path, width: f32, height: f32) -> bool {
    let expected = ellipse_path(DrawableSize { width, height });
    path_matches(path, &expected, width.max(height))
}

fn native_ellipse_elements_match(path: &tsp::Path, width: f32, height: f32) -> bool {
    use tsp::path::ElementType;

    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let diagonal_x = width * NATIVE_ELLIPSE_DIAGONAL;
    let diagonal_y = height * NATIVE_ELLIPSE_DIAGONAL;
    let control_x = width * NATIVE_ELLIPSE_OUTER_CONTROL;
    let control_y = height * NATIVE_ELLIPSE_OUTER_CONTROL;
    let inner_control_x = width * NATIVE_ELLIPSE_INNER_CONTROL;
    let inner_control_y = height * NATIVE_ELLIPSE_INNER_CONTROL;
    let start = point(center_x + diagonal_x, center_y - diagonal_y);
    let expected = tsp::Path {
        elements: vec![
            path_element(ElementType::MoveTo, vec![start]),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x + control_x, center_y - inner_control_y),
                    point(center_x + control_x, center_y + inner_control_y),
                    point(center_x + diagonal_x, center_y + diagonal_y),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x + inner_control_x, center_y + control_y),
                    point(center_x - inner_control_x, center_y + control_y),
                    point(center_x - diagonal_x, center_y + diagonal_y),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x - control_x, center_y + inner_control_y),
                    point(center_x - control_x, center_y - inner_control_y),
                    point(center_x - diagonal_x, center_y - diagonal_y),
                ],
            ),
            path_element(
                ElementType::CurveTo,
                vec![
                    point(center_x - inner_control_x, center_y - control_y),
                    point(center_x + inner_control_x, center_y - control_y),
                    start,
                ],
            ),
            path_element(ElementType::CloseSubpath, Vec::new()),
            path_element(ElementType::MoveTo, vec![start]),
        ],
    };
    path_matches(path, &expected, width.max(height))
}

fn path_matches(actual: &tsp::Path, expected: &tsp::Path, extent: f32) -> bool {
    actual.elements.len() == expected.elements.len()
        && actual
            .elements
            .iter()
            .zip(&expected.elements)
            .all(|(actual, expected)| {
                actual.r#type == expected.r#type
                    && actual.points.len() == expected.points.len()
                    && actual
                        .points
                        .iter()
                        .zip(&expected.points)
                        .all(|(actual, expected)| {
                            nearly_equal(actual.x, expected.x, extent)
                                && nearly_equal(actual.y, expected.y, extent)
                        })
            })
}

fn nearly_equal(left: f32, right: f32, extent: f32) -> bool {
    (left - right).abs() <= extent.max(1.0) * PATH_COMPARISON_EPSILON
}

fn rectangle_elements_match(path: &tsp::Path, width: f32, height: f32) -> bool {
    use tsp::path::ElementType;

    let expected = [
        (ElementType::MoveTo, Some((0.0, 0.0))),
        (ElementType::LineTo, Some((width, 0.0))),
        (ElementType::LineTo, Some((width, height))),
        (ElementType::LineTo, Some((0.0, height))),
        (ElementType::CloseSubpath, None),
        (ElementType::MoveTo, Some((0.0, 0.0))),
    ];
    path.elements.len() == expected.len()
        && path
            .elements
            .iter()
            .zip(expected)
            .all(|(element, expected)| {
                ElementType::try_from(element.r#type).ok() == Some(expected.0)
                    && match expected.1 {
                        Some((x, y)) => {
                            element.points.len() == 1
                                && element.points[0].x == x
                                && element.points[0].y == y
                        },
                        None => element.points.is_empty(),
                    }
            })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn path_element(r#type: tsp::path::ElementType, points: Vec<tsp::Point>) -> tsp::path::Element {
        tsp::path::Element {
            r#type: r#type as i32,
            points,
        }
    }

    #[test]
    fn rejects_missing_and_ambiguous_path_families() {
        let mut shape = tswp::ShapeInfoArchive::default();
        shape.super_.pathsource = Some(tsd::PathSourceArchive::default());
        assert!(shape_path_kind(&shape).is_err());

        let path = shape.super_.pathsource.as_mut().unwrap();
        path.bezier_path_source = Some(tsd::BezierPathSourceArchive::default());
        path.scalar_path_source = Some(tsd::ScalarPathSourceArchive::default());
        assert!(shape_path_kind(&shape).is_err());
    }

    #[test]
    fn classifies_pages_normalized_rectangle_independently_of_natural_size() {
        use tsp::path::ElementType;

        let point = |x, y| tsp::Point { x, y };
        let mut shape = tswp::ShapeInfoArchive::default();
        shape.super_.pathsource = Some(tsd::PathSourceArchive {
            bezier_path_source: Some(tsd::BezierPathSourceArchive {
                natural_size: Some(tsp::Size {
                    width: 300.0,
                    height: 150.0,
                }),
                path: Some(tsp::Path {
                    elements: vec![
                        path_element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                        path_element(ElementType::LineTo, vec![point(100.0, 0.0)]),
                        path_element(ElementType::LineTo, vec![point(100.0, 100.0)]),
                        path_element(ElementType::LineTo, vec![point(0.0, 100.0)]),
                        path_element(ElementType::CloseSubpath, Vec::new()),
                        path_element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                    ],
                }),
                ..Default::default()
            }),
            ..Default::default()
        });

        assert_eq!(shape_path_kind(&shape).unwrap(), ShapePathKind::Rectangle);
    }

    #[test]
    fn typed_presets_build_and_recover_without_magic_controls() {
        let size = DrawableSize {
            width: 240.0,
            height: 120.0,
        };
        for (preset, kind) in [
            (ShapePreset::Rectangle, ShapePathKind::Rectangle),
            (
                ShapePreset::ROUNDED_RECTANGLE,
                ShapePathKind::RoundedRectangle,
            ),
            (ShapePreset::Ellipse, ShapePathKind::Ellipse),
            (ShapePreset::LeftArrow, ShapePathKind::LeftArrow),
            (ShapePreset::RightArrow, ShapePathKind::RightArrow),
            (ShapePreset::DoubleArrow, ShapePathKind::DoubleArrow),
            (ShapePreset::PENTAGON, ShapePathKind::RegularPolygon),
            (ShapePreset::STAR, ShapePathKind::Star),
            (
                ShapePreset::RegularPolygon {
                    sides: ShapePolygonSides::new(8).unwrap(),
                },
                ShapePathKind::RegularPolygon,
            ),
            (
                ShapePreset::Star {
                    points: ShapeStarPoints::new(7).unwrap(),
                    inner_radius: ShapeStarInnerRatio::new(0.45).unwrap(),
                },
                ShapePathKind::Star,
            ),
        ] {
            let mut shape = tswp::ShapeInfoArchive::default();
            shape.super_.pathsource = Some(shape_path_source(preset, size).unwrap());
            assert_eq!(shape_path_kind(&shape).unwrap(), kind);
            assert_eq!(shape_preset(&shape).unwrap(), Some(preset));
        }
    }

    #[test]
    fn recognizes_only_native_default_arrow_controls() {
        use tsd::point_path_source_archive::PointPathSourceType;

        let arrow = |kind, width, control_x| {
            let mut shape = tswp::ShapeInfoArchive::default();
            shape.super_.pathsource = Some(tsd::PathSourceArchive {
                point_path_source: Some(tsd::PointPathSourceArchive {
                    r#type: Some(kind as i32),
                    point: Some(tsp::Point {
                        x: control_x,
                        y: NATIVE_ARROW_SHAFT_RATIO,
                    }),
                    natural_size: Some(tsp::Size {
                        width,
                        height: 100.0,
                    }),
                }),
                ..Default::default()
            });
            shape
        };
        let right = arrow(PointPathSourceType::KTsdRightSingleArrow, 100.0, 64.0);
        assert_eq!(shape_preset(&right).unwrap(), Some(ShapePreset::RightArrow));
        let left = arrow(PointPathSourceType::KTsdLeftSingleArrow, 100.0, 64.0);
        assert_eq!(shape_preset(&left).unwrap(), Some(ShapePreset::LeftArrow));
        let double = arrow(PointPathSourceType::KTsdDoubleArrow, 110.0, 44.0);
        assert_eq!(
            shape_preset(&double).unwrap(),
            Some(ShapePreset::DoubleArrow)
        );

        let custom = arrow(PointPathSourceType::KTsdRightSingleArrow, 100.0, 63.0);
        assert_eq!(shape_path_kind(&custom).unwrap(), ShapePathKind::RightArrow);
        assert_eq!(shape_preset(&custom).unwrap(), None);

        let mut missing_control = right;
        missing_control
            .super_
            .pathsource
            .as_mut()
            .unwrap()
            .point_path_source
            .as_mut()
            .unwrap()
            .point = None;
        assert!(shape_preset(&missing_control).is_err());
    }

    #[test]
    fn classifies_keynote_native_diagonal_ellipse() {
        use tsp::path::ElementType;

        let points =
            |values: &[(f32, f32)]| values.iter().map(|&(x, y)| tsp::Point { x, y }).collect();
        let start = (85.355_34, 14.644_661);
        let native = tsp::Path {
            elements: vec![
                path_element(ElementType::MoveTo, points(&[start])),
                path_element(
                    ElementType::CurveTo,
                    points(&[
                        (104.881_55, 34.170_876),
                        (104.881_55, 65.829_124),
                        (85.355_34, 85.355_34),
                    ]),
                ),
                path_element(
                    ElementType::CurveTo,
                    points(&[
                        (65.829_124, 104.881_55),
                        (34.170_876, 104.881_55),
                        (14.644_661, 85.355_34),
                    ]),
                ),
                path_element(
                    ElementType::CurveTo,
                    points(&[
                        (-4.881_553_6, 65.829_124),
                        (-4.881_553_6, 34.170_876),
                        (14.644_661, 14.644_661),
                    ]),
                ),
                path_element(
                    ElementType::CurveTo,
                    points(&[
                        (34.170_876, -4.881_553_6),
                        (65.829_124, -4.881_553_6),
                        start,
                    ]),
                ),
                path_element(ElementType::CloseSubpath, Vec::new()),
                path_element(ElementType::MoveTo, points(&[start])),
            ],
        };
        let mut shape = tswp::ShapeInfoArchive::default();
        shape.super_.pathsource = Some(tsd::PathSourceArchive {
            bezier_path_source: Some(tsd::BezierPathSourceArchive {
                natural_size: Some(tsp::Size {
                    width: 100.0,
                    height: 100.0,
                }),
                path: Some(native),
                ..Default::default()
            }),
            ..Default::default()
        });

        assert_eq!(shape_path_kind(&shape).unwrap(), ShapePathKind::Ellipse);
        assert_eq!(shape_preset(&shape).unwrap(), Some(ShapePreset::Ellipse));
    }

    #[test]
    fn typed_preset_controls_reject_invalid_values_and_sizes() {
        assert!(ShapeCornerRadius::new(-1.0).is_err());
        assert!(ShapeCornerRadius::new(f32::INFINITY).is_err());
        assert!(ShapePolygonSides::new(2).is_err());
        assert!(ShapeStarPoints::new(2).is_err());
        assert!(ShapeStarInnerRatio::new(-0.1).is_err());
        assert!(ShapeStarInnerRatio::new(1.0).is_err());
        assert!(
            shape_path_source(
                ShapePreset::ROUNDED_RECTANGLE,
                DrawableSize {
                    width: 20.0,
                    height: 20.0,
                },
            )
            .is_err()
        );
        assert!(
            shape_path_source(
                ShapePreset::Ellipse,
                DrawableSize {
                    width: f32::NAN,
                    height: 20.0,
                },
            )
            .is_err()
        );
    }

    #[test]
    fn preset_patch_retains_path_source_reflection_fields() {
        let size = DrawableSize {
            width: 240.0,
            height: 120.0,
        };
        let mut path_source = shape_path_source(ShapePreset::Rectangle, size).unwrap();
        path_source.horizontal_flip = Some(true);
        path_source.vertical_flip = Some(false);
        let shape = tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        size: Some(tsp::Size {
                            width: size.width,
                            height: size.height,
                        }),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                pathsource: Some(path_source),
                ..Default::default()
            },
            ..Default::default()
        };

        let changed = patch_shape_preset(&shape.encode_to_vec(), 42, ShapePreset::Ellipse).unwrap();
        let changed = tswp::ShapeInfoArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(shape_preset(&changed).unwrap(), Some(ShapePreset::Ellipse));
        assert_eq!(
            changed.super_.pathsource.as_ref().unwrap().horizontal_flip,
            Some(true)
        );
        assert_eq!(
            changed.super_.pathsource.as_ref().unwrap().vertical_flip,
            Some(false)
        );
    }

    #[test]
    fn preset_patch_preserves_unknown_ancestor_wire_and_restores_exactly() {
        let size = DrawableSize {
            width: 240.0,
            height: 120.0,
        };
        let shape = tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        size: Some(tsp::Size {
                            width: size.width,
                            height: size.height,
                        }),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                pathsource: Some(shape_path_source(ShapePreset::Rectangle, size).unwrap()),
                ..Default::default()
            },
            ..Default::default()
        };
        let original = transform_length_delimited_field(&shape.encode_to_vec(), 1, |shape| {
            let mut shape = shape.to_vec();
            append_unknown_varint(&mut shape, 98, 980);
            Ok(shape)
        })
        .unwrap();
        let mut original_with_unknowns = original;
        append_unknown_varint(&mut original_with_unknowns, 99, 990);

        let changed = patch_shape_preset(&original_with_unknowns, 42, ShapePreset::STAR).unwrap();
        let restored = patch_shape_preset(&changed, 42, ShapePreset::Rectangle).unwrap();

        assert_eq!(restored, original_with_unknowns);
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }
}
