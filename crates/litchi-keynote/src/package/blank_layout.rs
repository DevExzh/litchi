//! Title and body placeholder frames of a fresh presentation for its slide
//! size (record 0417; the migration host builder's `slide_size`).
//!
//! The host builder lays out the title and body placeholders of the
//! template and the first slide as fractions of the slide size: the title
//! at (10 %, 24 %) of the slide, 80 % × 20 % of it; the body at
//! (15 %, 52 %), 70 % × 14 %. Each placeholder's rectangular bezier path
//! follows its frame. This rewrites exactly those fields on the wire, so a
//! [`crate::Builder`] with a slide size lays out its placeholders as the
//! host did instead of keeping the blank presentation's 1920×1080 frames.

use litchi_iwa_core::RawMessage;

use super::structure::StructureError;
use super::structure::wire::{Fields, Writer};
use super::structure::workspace::Workspace;
use super::{Package, ReadError, ReadResult};
use crate::show::Size;

const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
/// `KN.PlaceholderArchive.kind`.
const PLACEHOLDER_KIND: u32 = 2;
const KIND_TITLE: u64 = 2;
const KIND_BODY: u64 = 3;
/// The `super` field of every archive on the placeholder's inheritance chain
/// (`KN.PlaceholderArchive` → `TSWP.ShapeInfoArchive` → `TSD.ShapeArchive`
/// → `TSD.DrawableArchive`).
const SUPER: u32 = 1;
/// `TSD.ShapeArchive.pathsource`.
const SHAPE_PATH_SOURCE: u32 = 3;
/// `TSD.DrawableArchive.geometry`.
const DRAWABLE_GEOMETRY: u32 = 1;
/// `TSD.GeometryArchive.{position, size}`.
const GEOMETRY_POSITION: u32 = 1;
const GEOMETRY_SIZE: u32 = 2;
/// `TSD.PathSourceArchive.bezier_path_source`.
const PATH_SOURCE_BEZIER: u32 = 5;
/// `TSD.BezierPathSourceArchive.{natural_size, path}`.
const BEZIER_NATURAL_SIZE: u32 = 2;
const BEZIER_PATH: u32 = 3;
/// `TSP.Path.elements`, `TSP.Path.Element.{type, points}`.
const PATH_ELEMENTS: u32 = 1;
const ELEMENT_TYPE: u32 = 1;
const ELEMENT_POINTS: u32 = 2;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const CLOSE_SUBPATH: u64 = 5;

fn invalid(error: StructureError) -> ReadError {
    ReadError::InvalidFormat(format!(
        "Keynote placeholder layout rewrite failed: {error:?}"
    ))
}

/// A frame in points: position and size.
#[derive(Clone, Copy)]
struct Frame {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
}

impl Frame {
    /// The host builder's frame of the placeholder `kind` on a slide of
    /// `size` (single-precision products, as the host computes them).
    fn for_kind(kind: u64, size: Size) -> Option<Self> {
        let (width, height) = (size.width(), size.height());
        match kind {
            KIND_TITLE => Some(Self {
                x: width * 0.1,
                y: height * 0.24,
                width: width * 0.8,
                height: height * 0.2,
            }),
            KIND_BODY => Some(Self {
                x: width * 0.15,
                y: height * 0.52,
                width: width * 0.7,
                height: height * 0.14,
            }),
            _ => None,
        }
    }
}

fn point(x: f32, y: f32) -> Writer {
    Writer::new().float(1, x).float(2, y)
}

/// The rectangle path the host writes for a frame of `width` × `height`.
fn rectangle_path(width: f32, height: f32) -> Vec<u8> {
    let element = |kind: u64, points: &[(f32, f32)]| {
        points.iter().fold(
            Writer::new().varint(ELEMENT_TYPE, kind),
            |writer, &(x, y)| writer.message(ELEMENT_POINTS, &point(x, y)),
        )
    };
    [
        element(MOVE_TO, &[(0.0, 0.0)]),
        element(LINE_TO, &[(width, 0.0)]),
        element(LINE_TO, &[(width, height)]),
        element(LINE_TO, &[(0.0, height)]),
        element(CLOSE_SUBPATH, &[]),
        element(MOVE_TO, &[(0.0, 0.0)]),
    ]
    .iter()
    .fold(Writer::new(), |writer, element| {
        writer.message(PATH_ELEMENTS, element)
    })
    .finish()
}

/// Rewrite one placeholder payload for `size`; `None` when it is not a
/// title or body placeholder or is already laid out for `size`.
fn placeholder(
    payload: &[u8],
    size: Size,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut placeholder = Fields::parse(payload, wire)?;
    let kind = match placeholder.unique(PLACEHOLDER_KIND)? {
        Some(field) => field.varint()?,
        None => return Ok(None),
    };
    let Some(frame) = Frame::for_kind(kind, size) else {
        return Ok(None);
    };
    let Some(shape_info) = placeholder.unique(SUPER)? else {
        return Ok(None);
    };
    let mut shape_info = Fields::parse(shape_info.message()?, wire)?;
    let Some(shape) = shape_info.unique(SUPER)? else {
        return Ok(None);
    };
    let mut shape = Fields::parse(shape.message()?, wire)?;
    let Some(drawable) = shape.unique(SUPER)? else {
        return Ok(None);
    };
    let mut drawable = Fields::parse(drawable.message()?, wire)?;
    let Some(geometry) = drawable.unique(DRAWABLE_GEOMETRY)? else {
        return Ok(None);
    };
    let mut geometry = Fields::parse(geometry.message()?, wire)?;
    geometry.set_message(GEOMETRY_POSITION, &point(frame.x, frame.y).finish());
    let frame_size = point(frame.width, frame.height).finish();
    geometry.set_message(GEOMETRY_SIZE, &frame_size);
    drawable.set_message(DRAWABLE_GEOMETRY, &geometry.encode());
    shape.set_message(SUPER, &drawable.encode());
    if let Some(path_source) = shape.unique(SHAPE_PATH_SOURCE)? {
        let mut path_source = Fields::parse(path_source.message()?, wire)?;
        if let Some(bezier) = path_source.unique(PATH_SOURCE_BEZIER)? {
            let mut bezier = Fields::parse(bezier.message()?, wire)?;
            bezier.set_message(BEZIER_NATURAL_SIZE, &frame_size);
            bezier.set_message(BEZIER_PATH, &rectangle_path(frame.width, frame.height));
            path_source.set_message(PATH_SOURCE_BEZIER, &bezier.encode());
            shape.set_message(SHAPE_PATH_SOURCE, &path_source.encode());
        }
    }
    shape_info.set_message(SUPER, &shape.encode());
    placeholder.set_message(SUPER, &shape_info.encode());
    let rewritten = placeholder.encode();
    Ok((rewritten != payload).then_some(rewritten))
}

impl Package {
    /// Return a reopened copy whose title and body placeholders (template
    /// and slides) have the host builder's frames for a slide of `size`.
    pub(crate) fn with_placeholder_frames(&self, size: Size) -> ReadResult<Self> {
        let mut workspace = Workspace::new(self).map_err(invalid)?;
        let wire = workspace.wire;
        let mut identifiers = self
            .state
            .object_index
            .iter()
            .map(|locator| locator.identifier)
            .collect::<Vec<_>>();
        identifiers.sort_unstable();
        identifiers.dedup();
        let mut changed = false;
        for identifier in identifiers {
            let object = workspace.require(identifier).map_err(invalid)?;
            let mut rewrites = Vec::new();
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ != PLACEHOLDER_MESSAGE_TYPE {
                    continue;
                }
                if let Some(data) = placeholder(&message.data, size, wire).map_err(invalid)? {
                    rewrites.push((index, data));
                }
            }
            let limits = workspace.archive_limits();
            for (index, data) in rewrites {
                workspace
                    .object_mut(identifier)
                    .map_err(invalid)?
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: PLACEHOLDER_MESSAGE_TYPE,
                            data,
                        },
                        limits,
                    )
                    .map_err(|error| invalid(error.into()))?;
                changed = true;
            }
        }
        if !changed {
            return Ok(self.snapshot());
        }
        let bytes = workspace.publish().map_err(invalid)?;
        self.structure_reopen(bytes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// (kind, x, y, width, height, natural width, natural height) of every
    /// title and body placeholder of `package`.
    fn frames(package: &Package) -> Vec<(u64, [f32; 6])> {
        let wire = Workspace::new(package).expect("workspace").wire;
        let mut found = Vec::new();
        for member in ["Index/TemplateSlide-8.iwa", "Index/Slide-14.iwa"] {
            let archive = package.structure_component(member).expect("member");
            for object in &archive.objects {
                for message in &object.messages {
                    if message.type_ != PLACEHOLDER_MESSAGE_TYPE {
                        continue;
                    }
                    let fields = Fields::parse(&message.data, wire).expect("placeholder");
                    let kind = fields
                        .unique(PLACEHOLDER_KIND)
                        .expect("kind")
                        .map_or(0, |field| field.varint().expect("kind"));
                    if kind != KIND_TITLE && kind != KIND_BODY {
                        continue;
                    }
                    let nested = |payload: &[u8], number: u32| -> Vec<u8> {
                        Fields::parse(payload, wire)
                            .expect("fields")
                            .unique(number)
                            .expect("unique")
                            .expect("present")
                            .message()
                            .expect("message")
                            .to_vec()
                    };
                    let floats = |payload: &[u8]| -> [f32; 2] {
                        let fields = Fields::parse(payload, wire).expect("pair");
                        [1, 2].map(|number| {
                            fields
                                .unique(number)
                                .expect("unique")
                                .expect("present")
                                .float()
                                .expect("float")
                        })
                    };
                    let shape = nested(&nested(&message.data, SUPER), SUPER);
                    let geometry = nested(&nested(&shape, SUPER), DRAWABLE_GEOMETRY);
                    let [x, y] = floats(&nested(&geometry, GEOMETRY_POSITION));
                    let [width, height] = floats(&nested(&geometry, GEOMETRY_SIZE));
                    let bezier = nested(&nested(&shape, SHAPE_PATH_SOURCE), PATH_SOURCE_BEZIER);
                    let [natural_width, natural_height] =
                        floats(&nested(&bezier, BEZIER_NATURAL_SIZE));
                    found.push((kind, [x, y, width, height, natural_width, natural_height]));
                }
            }
        }
        found.sort_by_key(|(kind, _)| *kind);
        found
    }

    #[test]
    fn placeholder_frames_follow_the_slide_size() {
        let size = Size::new(1024.0, 768.0).expect("size");
        let laid_out = Package::new()
            .with_placeholder_frames(size)
            .expect("layout");
        let title = [
            1024.0_f32 * 0.1,
            768.0_f32 * 0.24,
            1024.0_f32 * 0.8,
            768.0_f32 * 0.2,
            1024.0_f32 * 0.8,
            768.0_f32 * 0.2,
        ];
        let body = [
            1024.0_f32 * 0.15,
            768.0_f32 * 0.52,
            1024.0_f32 * 0.7,
            768.0_f32 * 0.14,
            1024.0_f32 * 0.7,
            768.0_f32 * 0.14,
        ];
        assert_eq!(
            frames(&laid_out),
            vec![
                (KIND_TITLE, title),
                (KIND_TITLE, title),
                (KIND_BODY, body),
                (KIND_BODY, body)
            ]
        );
        laid_out.validate().expect("valid");
    }

    #[test]
    fn the_blank_size_is_a_no_op() {
        let blank = Package::new();
        let size = Size::new(1_920.0, 1_080.0).expect("size");
        let same = blank.with_placeholder_frames(size).expect("layout");
        assert_eq!(frames(&same), frames(&blank));
        let mut before = Vec::new();
        let mut after = Vec::new();
        blank.write_to(&mut before).expect("write");
        same.write_to(&mut after).expect("write");
        assert_eq!(before, after);
    }
}
