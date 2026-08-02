//! Typed straight-line geometry, native path construction, and wire updates.

use prost::Message;

use crate::IWorkPackage;
use crate::archive::RawMessage;
use crate::protobuf::{tsd, tsp, tswp};
use crate::wire::{
    patch_fixed32_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, Result};

use super::geometry::{
    DrawableGeometry, DrawablePoint, DrawableSize, geometry_from_drawable, patch_drawable_geometry,
};

const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEGREES_PER_RADIAN: f32 = 180.0 / std::f32::consts::PI;
const FULL_ROTATION_DEGREES: f32 = 360.0;
const LINE_COMPARISON_EPSILON: f32 = 0.001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const ZERO_LINE_HEIGHT: f32 = 0.0;

/// A non-degenerate straight line between two points in document coordinates.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct LineSegment {
    start: DrawablePoint,
    end: DrawablePoint,
}

impl LineSegment {
    /// Validate and construct a straight line from two document-space endpoints.
    pub fn new(start: DrawablePoint, end: DrawablePoint) -> Result<Self> {
        if !start.x.is_finite() || !start.y.is_finite() || !end.x.is_finite() || !end.y.is_finite()
        {
            return Err(Error::ParseError(
                "iWork line endpoints must be finite".to_owned(),
            ));
        }
        let delta_x = end.x - start.x;
        let delta_y = end.y - start.y;
        let length = delta_x.hypot(delta_y);
        if !length.is_finite() || length <= 0.0 {
            return Err(Error::ParseError(
                "iWork line endpoints must have a finite positive distance".to_owned(),
            ));
        }
        Ok(Self { start, end })
    }

    /// First endpoint in document coordinates.
    pub const fn start(self) -> DrawablePoint {
        self.start
    }

    /// Second endpoint in document coordinates.
    pub const fn end(self) -> DrawablePoint {
        self.end
    }

    /// Euclidean line length in document points.
    pub fn length(self) -> f32 {
        let delta_x = self.end.x - self.start.x;
        let delta_y = self.end.y - self.start.y;
        delta_x.hypot(delta_y)
    }
}

pub(crate) fn line_geometry(segment: LineSegment) -> DrawableGeometry {
    let delta_x = segment.end.x - segment.start.x;
    let delta_y = segment.end.y - segment.start.y;
    let length = delta_x.hypot(delta_y);
    let center_x = segment.start.x + delta_x / 2.0;
    let center_y = segment.start.y + delta_y / 2.0;
    let angle = (delta_y.atan2(delta_x) * DEGREES_PER_RADIAN).rem_euclid(FULL_ROTATION_DEGREES);
    DrawableGeometry {
        position: Some(DrawablePoint {
            x: center_x - length / 2.0,
            y: center_y,
        }),
        size: Some(DrawableSize {
            width: length,
            height: ZERO_LINE_HEIGHT,
        }),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(angle),
    }
}

pub(crate) fn line_path_source(segment: LineSegment) -> tsd::PathSourceArchive {
    use tsp::path::ElementType;

    let length = segment.length();
    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size {
                width: length,
                height: ZERO_LINE_HEIGHT,
            }),
            path: Some(tsp::Path {
                elements: vec![
                    path_element(ElementType::MoveTo, 0.0),
                    path_element(ElementType::LineTo, length),
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

pub(crate) fn shape_line_segment(shape: &tswp::ShapeInfoArchive) -> Result<Option<LineSegment>> {
    let Some(path_source) = shape.super_.pathsource.as_ref() else {
        return Err(Error::InvalidFormat(
            "ordinary iWork shape has no path source".to_owned(),
        ));
    };
    let Some(bezier) = path_source.bezier_path_source.as_ref() else {
        return Ok(None);
    };
    if !is_straight_line_bezier(bezier) {
        return Ok(None);
    }
    let effective_width = bezier
        .natural_size
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("native straight line has no natural size".to_owned()))?
        .width;
    let geometry = geometry_from_drawable(&shape.super_.super_)?;
    let mut segment = line_segment_from_geometry(geometry, effective_width)?;
    if path_source.horizontal_flip == Some(true) {
        segment = LineSegment::new(segment.end, segment.start)?;
    }
    Ok(Some(segment))
}

pub(crate) fn is_straight_line_bezier(bezier: &tsd::BezierPathSourceArchive) -> bool {
    use tsp::path::ElementType;

    let Some(size) = bezier.natural_size.as_ref() else {
        return false;
    };
    let Some(path) = bezier.path.as_ref() else {
        return false;
    };
    size.width.is_finite()
        && size.width > 0.0
        && nearly_equal(size.height, ZERO_LINE_HEIGHT)
        && path.elements.len() == 2
        && element_matches(&path.elements[0], ElementType::MoveTo, 0.0)
        && element_matches(&path.elements[1], ElementType::LineTo, size.width)
}

pub(crate) fn line_segments_match(left: LineSegment, right: LineSegment) -> bool {
    points_match(left.start, right.start) && points_match(left.end, right.end)
}

pub(crate) fn set_shape_line_segment(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    replacement: LineSegment,
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
        let data = patch_line_shape(
            object.messages[message_index].data.as_slice(),
            drawable_id,
            replacement,
        )?;
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

fn patch_line_shape(data: &[u8], drawable_id: u64, replacement: LineSegment) -> Result<Vec<u8>> {
    let current = tswp::ShapeInfoArchive::decode(data)?;
    if shape_line_segment(&current)?.is_none() {
        return Err(Error::ParseError(format!(
            "iWork drawable {drawable_id} is not a native straight line"
        )));
    }
    let geometry_segment = if current
        .super_
        .pathsource
        .as_ref()
        .is_some_and(|path| path.horizontal_flip == Some(true))
    {
        LineSegment::new(replacement.end, replacement.start)?
    } else {
        replacement
    };
    let geometry = line_geometry(geometry_segment);
    let width = replacement.length();
    let data = transform_length_delimited_field(data, 1, |shape| {
        let shape = transform_length_delimited_field(shape, 1, |drawable| {
            patch_drawable_geometry(drawable, geometry)
        })?;
        transform_length_delimited_field(&shape, 3, |path_source| {
            patch_line_path_source(path_source, width)
        })
    })?;
    let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
    let actual = shape_line_segment(&verified)?.ok_or_else(|| {
        Error::InvalidFormat("iWork line endpoint patch changed the path family".to_owned())
    })?;
    if !line_segments_match(actual, replacement) {
        return Err(Error::InvalidFormat(
            "iWork line endpoint patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn patch_line_path_source(data: &[u8], width: f32) -> Result<Vec<u8>> {
    let current = tsd::PathSourceArchive::decode(data)?;
    let bezier = current.bezier_path_source.as_ref().ok_or_else(|| {
        Error::InvalidFormat("native straight line has no Bézier path source".to_owned())
    })?;
    if !is_straight_line_bezier(bezier) {
        return Err(Error::InvalidFormat(
            "native straight line has malformed Bézier geometry".to_owned(),
        ));
    }
    transform_length_delimited_field(data, 5, |bezier| patch_line_bezier(bezier, width))
}

fn patch_line_bezier(data: &[u8], width: f32) -> Result<Vec<u8>> {
    let data = transform_length_delimited_field(data, 2, |size| {
        patch_fixed32_field(size, 1, true, Some(width.to_bits()))
    })?;
    transform_length_delimited_field(&data, 3, |path| {
        let elements = repeated_length_delimited_payloads(path, 1)?;
        let [move_to, line_to] = elements.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "native straight line must have two path elements, found {}",
                elements.len()
            )));
        };
        let line_to = transform_length_delimited_field(line_to, 2, |point| {
            patch_fixed32_field(point, 1, true, Some(width.to_bits()))
        })?;
        rewrite_repeated_length_delimited_fields(path, 1, &[move_to.to_vec(), line_to])
    })
}

fn line_segment_from_geometry(
    geometry: DrawableGeometry,
    effective_width: f32,
) -> Result<LineSegment> {
    let position = geometry
        .position
        .ok_or_else(|| Error::InvalidFormat("straight iWork line has no position".to_owned()))?;
    let size = geometry
        .size
        .ok_or_else(|| Error::InvalidFormat("straight iWork line has no size".to_owned()))?;
    let angle = geometry.angle.ok_or_else(|| {
        Error::InvalidFormat("straight iWork line has no rotation angle".to_owned())
    })?;
    if !size.width.is_finite()
        || size.width <= 0.0
        || !effective_width.is_finite()
        || effective_width <= 0.0
        || !nearly_equal(size.height, ZERO_LINE_HEIGHT)
    {
        return Err(Error::InvalidFormat(
            "straight iWork line must have positive width and zero height".to_owned(),
        ));
    }
    let radians = angle / DEGREES_PER_RADIAN;
    let half_x = radians.cos() * effective_width / 2.0;
    let half_y = radians.sin() * effective_width / 2.0;
    let center = DrawablePoint {
        x: position.x + effective_width / 2.0,
        y: position.y,
    };
    LineSegment::new(
        DrawablePoint {
            x: center.x - half_x,
            y: center.y - half_y,
        },
        DrawablePoint {
            x: center.x + half_x,
            y: center.y + half_y,
        },
    )
}

fn path_element(r#type: tsp::path::ElementType, x: f32) -> tsp::path::Element {
    tsp::path::Element {
        r#type: r#type as i32,
        points: vec![tsp::Point { x, y: 0.0 }],
    }
}

fn element_matches(element: &tsp::path::Element, r#type: tsp::path::ElementType, x: f32) -> bool {
    tsp::path::ElementType::try_from(element.r#type).ok() == Some(r#type)
        && element.points.len() == 1
        && nearly_equal(element.points[0].x, x)
        && nearly_equal(element.points[0].y, ZERO_LINE_HEIGHT)
}

fn points_match(left: DrawablePoint, right: DrawablePoint) -> bool {
    nearly_equal(left.x, right.x) && nearly_equal(left.y, right.y)
}

fn nearly_equal(left: f32, right: f32) -> bool {
    (left - right).abs() <= LINE_COMPARISON_EPSILON
}

#[cfg(test)]
mod tests {
    use super::*;

    const START: DrawablePoint = DrawablePoint { x: 20.0, y: 40.0 };
    const END: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };

    #[test]
    fn native_line_geometry_and_path_round_trip_endpoints() {
        let segment = LineSegment::new(START, END).unwrap();
        let geometry = line_geometry(segment);
        assert_eq!(geometry.size.unwrap().height, ZERO_LINE_HEIGHT);
        assert_eq!(geometry.angle, Some(45.0));

        let shape = tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: geometry.position.map(|point| tsp::Point {
                            x: point.x,
                            y: point.y,
                        }),
                        size: geometry.size.map(|size| tsp::Size {
                            width: size.width,
                            height: size.height,
                        }),
                        flags: geometry.flags,
                        angle: geometry.angle,
                    }),
                    ..Default::default()
                },
                pathsource: Some(line_path_source(segment)),
                ..Default::default()
            },
            ..Default::default()
        };
        let decoded = shape_line_segment(&shape).unwrap().unwrap();
        assert!(line_segments_match(decoded, segment));
    }

    #[test]
    fn rejects_degenerate_and_malformed_lines() {
        assert!(LineSegment::new(START, START).is_err());
        assert!(
            LineSegment::new(
                DrawablePoint {
                    x: f32::NAN,
                    y: 0.0
                },
                END
            )
            .is_err()
        );
        assert!(
            LineSegment::new(
                DrawablePoint {
                    x: f32::MAX,
                    y: 0.0,
                },
                DrawablePoint {
                    x: -f32::MAX,
                    y: 0.0,
                },
            )
            .is_err()
        );

        let segment = LineSegment::new(START, END).unwrap();
        let mut path = line_path_source(segment);
        path.bezier_path_source
            .as_mut()
            .unwrap()
            .natural_size
            .as_mut()
            .unwrap()
            .height = 1.0;
        assert!(!is_straight_line_bezier(
            path.bezier_path_source.as_ref().unwrap()
        ));
    }

    #[test]
    fn endpoint_patch_preserves_deep_unknown_wire_and_restores_exactly() {
        let original_segment = LineSegment::new(
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawablePoint { x: 120.0, y: 40.0 },
        )
        .unwrap();
        let geometry = line_geometry(original_segment);
        let shape = tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: geometry.position.map(|point| tsp::Point {
                            x: point.x,
                            y: point.y,
                        }),
                        size: geometry.size.map(|size| tsp::Size {
                            width: size.width,
                            height: size.height,
                        }),
                        flags: geometry.flags,
                        angle: geometry.angle,
                    }),
                    ..Default::default()
                },
                pathsource: Some(line_path_source(original_segment)),
                ..Default::default()
            },
            ..Default::default()
        };
        let mut original = shape.encode_to_vec();
        original = transform_length_delimited_field(&original, 1, |shape| {
            let shape = transform_length_delimited_field(shape, 1, |drawable| {
                transform_length_delimited_field(drawable, 1, |geometry| {
                    let mut geometry = geometry.to_vec();
                    append_unknown_varint(&mut geometry, 96, 960);
                    Ok(geometry)
                })
            })?;
            transform_length_delimited_field(&shape, 3, |path_source| {
                transform_length_delimited_field(path_source, 5, |bezier| {
                    let bezier = transform_length_delimited_field(bezier, 2, |size| {
                        let mut size = size.to_vec();
                        append_unknown_varint(&mut size, 97, 970);
                        Ok(size)
                    })?;
                    transform_length_delimited_field(&bezier, 3, |path| {
                        let elements = repeated_length_delimited_payloads(path, 1)?;
                        let [move_to, line_to] = elements.as_slice() else {
                            unreachable!("constructed line has two elements")
                        };
                        let line_to = transform_length_delimited_field(line_to, 2, |point| {
                            let mut point = point.to_vec();
                            append_unknown_varint(&mut point, 98, 980);
                            Ok(point)
                        })?;
                        rewrite_repeated_length_delimited_fields(
                            path,
                            1,
                            &[move_to.to_vec(), line_to],
                        )
                    })
                })
            })
        })
        .unwrap();
        append_unknown_varint(&mut original, 99, 990);

        let replacement = LineSegment::new(
            DrawablePoint { x: 40.0, y: 80.0 },
            DrawablePoint { x: 190.0, y: 80.0 },
        )
        .unwrap();
        let changed = patch_line_shape(&original, 42, replacement).unwrap();
        let restored = patch_line_shape(&changed, 42, original_segment).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn endpoint_patch_preserves_native_horizontal_flip_semantics() {
        let original = LineSegment::new(
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawablePoint { x: 120.0, y: 40.0 },
        )
        .unwrap();
        let reversed = LineSegment::new(original.end(), original.start()).unwrap();
        let geometry = line_geometry(reversed);
        let mut path_source = line_path_source(original);
        path_source.horizontal_flip = Some(true);
        let shape = tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: geometry.position.map(|point| tsp::Point {
                            x: point.x,
                            y: point.y,
                        }),
                        size: geometry.size.map(|size| tsp::Size {
                            width: size.width,
                            height: size.height,
                        }),
                        flags: geometry.flags,
                        angle: geometry.angle,
                    }),
                    ..Default::default()
                },
                pathsource: Some(path_source),
                ..Default::default()
            },
            ..Default::default()
        };
        assert!(line_segments_match(
            shape_line_segment(&shape).unwrap().unwrap(),
            original
        ));

        let replacement = LineSegment::new(
            DrawablePoint { x: 30.0, y: 80.0 },
            DrawablePoint { x: 230.0, y: 80.0 },
        )
        .unwrap();
        let changed = patch_line_shape(&shape.encode_to_vec(), 42, replacement).unwrap();
        let decoded = tswp::ShapeInfoArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(
            decoded.super_.pathsource.as_ref().unwrap().horizontal_flip,
            Some(true)
        );
        assert!(line_segments_match(
            shape_line_segment(&decoded).unwrap().unwrap(),
            replacement
        ));
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }
}
