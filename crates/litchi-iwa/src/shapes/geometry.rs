//! Typed, wire-preserving geometry for iWork drawables.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::{tsd, tsp, tswp};
use crate::wire::{
    patch_fixed32_field, patch_length_delimited_field, patch_varint_field,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;

/// A drawable position in document points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DrawablePoint {
    pub x: f32,
    pub y: f32,
}

/// A drawable size in document points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DrawableSize {
    pub width: f32,
    pub height: f32,
}

/// Geometry stored on an iWork drawable.
///
/// Optional fields retain the distinction between an absent protobuf field and
/// an explicit zero, allowing an update followed by restoration to be byte-exact.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct DrawableGeometry {
    pub position: Option<DrawablePoint>,
    pub size: Option<DrawableSize>,
    pub flags: Option<u32>,
    /// Clockwise rotation in degrees, using iWork's native geometry convention.
    pub angle: Option<f32>,
}

impl DrawableGeometry {
    pub(crate) fn validate(self) -> Result<Self> {
        if let Some(position) = self.position
            && (!position.x.is_finite() || !position.y.is_finite())
        {
            return Err(Error::ParseError(
                "iWork drawable position must be finite".to_owned(),
            ));
        }
        if let Some(size) = self.size
            && (!size.width.is_finite()
                || !size.height.is_finite()
                || size.width < 0.0
                || size.height < 0.0)
        {
            return Err(Error::ParseError(
                "iWork drawable size must be finite and non-negative".to_owned(),
            ));
        }
        if self.angle.is_some_and(|angle| !angle.is_finite()) {
            return Err(Error::ParseError(
                "iWork drawable angle must be finite".to_owned(),
            ));
        }
        Ok(self)
    }
}

/// Return a copy of geometry moved by the same finite amount on both axes.
pub(crate) fn offset_drawable_geometry(
    geometry: DrawableGeometry,
    offset: f32,
) -> Result<DrawableGeometry> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "iWork drawable offset must be finite".to_owned(),
        ));
    }
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("iWork drawable geometry has no position to offset".to_owned())
    })?;
    let position = DrawablePoint {
        x: position.x + offset,
        y: position.y + offset,
    };
    if !position.x.is_finite() || !position.y.is_finite() {
        return Err(Error::ParseError(
            "iWork drawable offset overflows its position".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        ..geometry
    }
    .validate()
}

pub(crate) fn shape_geometry(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<DrawableGeometry> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(drawable_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork drawable object {drawable_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
        )));
    }
    let shape = tswp::ShapeInfoArchive::decode(messages[0].data.as_slice())?;
    geometry_from_drawable(&shape.super_.super_)
}

pub(crate) fn set_shape_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    replacement: DrawableGeometry,
) -> Result<()> {
    let replacement = replacement.validate()?;
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
        let shape = tswp::ShapeInfoArchive::decode(original)?;
        if shape.super_.super_.geometry.is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} has no geometry payload"
            )));
        }
        let data = transform_length_delimited_field(original, 1, |shape_archive| {
            transform_length_delimited_field(shape_archive, 1, |drawable_archive| {
                patch_drawable_geometry(drawable_archive, replacement)
            })
        })?;
        let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
        if geometry_from_drawable(&verified.super_.super_)? != replacement {
            return Err(Error::InvalidFormat(
                "iWork drawable geometry patch failed validation".to_owned(),
            ));
        }
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

pub(crate) fn geometry_from_drawable(drawable: &tsd::DrawableArchive) -> Result<DrawableGeometry> {
    let geometry = drawable
        .geometry
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("iWork drawable has no geometry payload".to_owned()))?;
    DrawableGeometry {
        position: geometry.position.as_ref().map(|position| DrawablePoint {
            x: position.x,
            y: position.y,
        }),
        size: geometry.size.as_ref().map(|size| DrawableSize {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags,
        angle: geometry.angle,
    }
    .validate()
}

pub(crate) fn patch_drawable_geometry(
    data: &[u8],
    replacement: DrawableGeometry,
) -> Result<Vec<u8>> {
    let replacement = replacement.validate()?;
    let drawable = tsd::DrawableArchive::decode(data)?;
    if drawable.geometry.is_none() {
        return Err(Error::InvalidFormat(
            "iWork drawable has no geometry payload".to_owned(),
        ));
    }
    let data = transform_length_delimited_field(data, 1, |geometry| {
        patch_geometry(geometry, replacement)
    })?;
    let verified = tsd::DrawableArchive::decode(data.as_slice())?;
    if geometry_from_drawable(&verified)? != replacement {
        return Err(Error::InvalidFormat(
            "iWork drawable geometry patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn patch_geometry(data: &[u8], replacement: DrawableGeometry) -> Result<Vec<u8>> {
    let current = tsd::GeometryArchive::decode(data)?;
    let mut data = patch_point(data, current.position.as_ref(), replacement.position)?;
    data = patch_size(&data, current.size.as_ref(), replacement.size)?;
    data = patch_varint_field(
        &data,
        3,
        current.flags.is_some(),
        replacement.flags.map(u64::from),
    )?;
    patch_fixed32_field(
        &data,
        4,
        current.angle.is_some(),
        replacement.angle.map(f32::to_bits),
    )
}

fn patch_point(
    data: &[u8],
    current: Option<&tsp::Point>,
    replacement: Option<DrawablePoint>,
) -> Result<Vec<u8>> {
    match (current, replacement) {
        (Some(_), Some(replacement)) => transform_length_delimited_field(data, 1, |point| {
            let point = patch_fixed32_field(point, 1, true, Some(replacement.x.to_bits()))?;
            patch_fixed32_field(&point, 2, true, Some(replacement.y.to_bits()))
        }),
        (Some(_), None) => patch_length_delimited_field(data, 1, true, None),
        (None, Some(replacement)) => {
            let point = tsp::Point {
                x: replacement.x,
                y: replacement.y,
            }
            .encode_to_vec();
            patch_length_delimited_field(data, 1, false, Some(&point))
        },
        (None, None) => Ok(data.to_vec()),
    }
}

fn patch_size(
    data: &[u8],
    current: Option<&tsp::Size>,
    replacement: Option<DrawableSize>,
) -> Result<Vec<u8>> {
    match (current, replacement) {
        (Some(_), Some(replacement)) => transform_length_delimited_field(data, 2, |size| {
            let size = patch_fixed32_field(size, 1, true, Some(replacement.width.to_bits()))?;
            patch_fixed32_field(&size, 2, true, Some(replacement.height.to_bits()))
        }),
        (Some(_), None) => patch_length_delimited_field(data, 2, true, None),
        (None, Some(replacement)) => {
            let size = tsp::Size {
                width: replacement.width,
                height: replacement.height,
            }
            .encode_to_vec();
            patch_length_delimited_field(data, 2, false, Some(&size))
        },
        (None, None) => Ok(data.to_vec()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn geometry_offset_requires_a_finite_position_and_offset() {
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 12.0, y: 24.0 }),
            size: Some(DrawableSize {
                width: 100.0,
                height: 80.0,
            }),
            flags: Some(3),
            angle: Some(15.0),
        };
        assert_eq!(
            offset_drawable_geometry(geometry, 10.0).unwrap(),
            DrawableGeometry {
                position: Some(DrawablePoint { x: 22.0, y: 34.0 }),
                ..geometry
            }
        );
        assert!(offset_drawable_geometry(geometry, f32::NAN).is_err());
        assert!(
            offset_drawable_geometry(
                DrawableGeometry {
                    position: None,
                    ..geometry
                },
                10.0,
            )
            .is_err()
        );
        assert!(
            offset_drawable_geometry(
                DrawableGeometry {
                    position: Some(DrawablePoint {
                        x: f32::MAX,
                        y: f32::MAX,
                    }),
                    ..geometry
                },
                f32::MAX,
            )
            .is_err()
        );
    }

    #[test]
    fn geometry_patch_preserves_nested_unknowns_and_restores_exactly() {
        let geometry = tsd::GeometryArchive {
            position: Some(tsp::Point { x: 10.0, y: 20.0 }),
            size: Some(tsp::Size {
                width: 300.0,
                height: 80.0,
            }),
            flags: Some(0),
            angle: Some(0.0),
        };
        let mut original = geometry.encode_to_vec();
        original = transform_length_delimited_field(&original, 1, |point| {
            let mut point = point.to_vec();
            append_unknown_varint(&mut point, 98, 980);
            Ok(point)
        })
        .unwrap();
        original = transform_length_delimited_field(&original, 2, |size| {
            let mut size = size.to_vec();
            append_unknown_varint(&mut size, 97, 970);
            Ok(size)
        })
        .unwrap();
        append_unknown_varint(&mut original, 99, 990);

        let baseline = geometry_from_drawable(&tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive::decode(original.as_slice()).unwrap()),
            ..Default::default()
        })
        .unwrap();
        let replacement = DrawableGeometry {
            position: Some(DrawablePoint { x: 44.5, y: 55.5 }),
            size: Some(DrawableSize {
                width: 640.0,
                height: 120.0,
            }),
            flags: Some(3),
            angle: Some(0.75),
        };
        let changed = patch_geometry(&original, replacement).unwrap();
        let restored = patch_geometry(&changed, baseline).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn geometry_patch_rejects_duplicate_and_malformed_scalars() {
        let geometry = tsd::GeometryArchive {
            position: Some(tsp::Point { x: 10.0, y: 20.0 }),
            flags: Some(0),
            ..Default::default()
        };
        let replacement = DrawableGeometry {
            position: Some(DrawablePoint { x: 30.0, y: 40.0 }),
            flags: Some(1),
            ..Default::default()
        };

        let mut duplicate = geometry.encode_to_vec();
        duplicate.extend(crate::varint::encode_varint(3 << 3));
        duplicate.extend(crate::varint::encode_varint(2));
        assert!(patch_geometry(&duplicate, replacement).is_err());

        let mut malformed = geometry.encode_to_vec();
        malformed.extend(crate::varint::encode_varint(4 << 3));
        assert!(patch_geometry(&malformed, replacement).is_err());
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }
}
