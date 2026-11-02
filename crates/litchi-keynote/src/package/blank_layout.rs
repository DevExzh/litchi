//! The fresh-presentation values of a [`crate::Builder`] in one publication
//! (record 0417): language and locale (record 0372), the slide size with
//! the title and body placeholder frames laid out for it (the migration
//! host builder's `slide_size`), and the first slide's title, subtitle and
//! notes text.
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
use super::{Package, ReadError};
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

/// The fresh-presentation values a [`crate::Builder`] writes in one
/// publication: the document language and locale (`super::locale`), the
/// show's slide size with the title and body placeholder frames laid out
/// for it, and the first slide's title, subtitle and notes text.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct FreshSettings<'a> {
    pub(crate) language: Option<&'a str>,
    pub(crate) locale: Option<&'a str>,
    pub(crate) size: Option<Size>,
    pub(crate) title: Option<&'a str>,
    pub(crate) subtitle: Option<&'a str>,
    pub(crate) notes: Option<&'a str>,
}

/// A failure of [`Package::with_fresh_settings`], typed by the value it
/// concerns.
#[derive(Debug)]
pub(crate) enum FreshError {
    /// The language, locale or size rewrite, or the publication.
    Read(ReadError),
    /// The title or subtitle text.
    Text(super::slide_text::Error),
    /// The speaker notes text.
    Notes(super::slide_notes::Error),
}

impl From<ReadError> for FreshError {
    fn from(error: ReadError) -> Self {
        Self::Read(error)
    }
}

/// `KN.ShowArchive.size`.
const SHOW_SIZE: u32 = 4;
/// `TSWP.ShapeInfoArchive.owned_storage`.
const SHAPE_INFO_OWNED_STORAGE: u32 = 4;
/// `TSD.DrawableArchive.parent`.
const DRAWABLE_PARENT: u32 = 2;
/// `KN.NoteArchive.contained_storage`.
const NOTE_STORAGE: u32 = 1;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// Rewrite the show's size; `None` when it already has it.
fn show(
    payload: &[u8],
    size: Size,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    fields.set_message(SHOW_SIZE, &point(size.width(), size.height()).finish());
    let rewritten = fields.encode();
    Ok((rewritten != payload).then_some(rewritten))
}

/// The nested reference `number` of `payload`, when present.
fn reference(
    payload: &[u8],
    number: u32,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<u64>, StructureError> {
    match Fields::parse(payload, wire)?.unique(number)? {
        Some(field) => Ok(Some(super::structure::wire::reference_identifier(
            field.message()?,
            wire,
        )?)),
        None => Ok(None),
    }
}

/// The (kind, parent, owned storage) of a title or body placeholder.
fn placeholder_storage(
    payload: &[u8],
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<(u64, Option<u64>, Option<u64>)>, StructureError> {
    let placeholder = Fields::parse(payload, wire)?;
    let kind = match placeholder.unique(PLACEHOLDER_KIND)? {
        Some(field) => field.varint()?,
        None => return Ok(None),
    };
    if kind != KIND_TITLE && kind != KIND_BODY {
        return Ok(None);
    }
    let Some(shape_info) = placeholder.unique(SUPER)? else {
        return Ok(None);
    };
    let shape_info = shape_info.message()?;
    let storage = reference(shape_info, SHAPE_INFO_OWNED_STORAGE, wire)?;
    let Some(shape) = Fields::parse(shape_info, wire)?
        .unique(SUPER)?
        .map(|f| f.message().map(<[u8]>::to_vec))
        .transpose()?
    else {
        return Ok(Some((kind, None, storage)));
    };
    let Some(drawable) = Fields::parse(&shape, wire)?
        .unique(SUPER)?
        .map(|f| f.message().map(<[u8]>::to_vec))
        .transpose()?
    else {
        return Ok(Some((kind, None, storage)));
    };
    Ok(Some((
        kind,
        reference(&drawable, DRAWABLE_PARENT, wire)?,
        storage,
    )))
}

/// Which storage a text value replaces.
#[derive(Clone, Copy)]
enum Role {
    Title,
    Subtitle,
    Notes,
}

impl Package {
    /// Return a reopened copy with `settings` written where the host builder
    /// writes them, in one workspace publication (one reopen); values left
    /// `None` are kept. Called on the blank presentation only. The first
    /// slide's text storages are rewritten through the shared text engine
    /// exactly as the slide-text and speaker-notes transactions rewrite
    /// them (the slide preview is invalidated when the title or subtitle
    /// changes), and the candidate is verified by reading them back.
    pub(crate) fn with_fresh_settings(
        &self,
        settings: FreshSettings<'_>,
    ) -> Result<Self, FreshError> {
        let FreshSettings {
            language,
            locale,
            size,
            title,
            subtitle,
            notes,
        } = settings;
        for (role, text) in [
            (Role::Title, title),
            (Role::Subtitle, subtitle),
            (Role::Notes, notes),
        ] {
            if text.is_some_and(|text| text.contains('\u{000e}') || text.contains('\u{fffc}')) {
                return Err(match role {
                    Role::Notes => FreshError::Notes(super::slide_notes::Error::Refused(
                        super::slide_notes::Refusal::ObjectMarkerReplacement,
                    )),
                    _ => FreshError::Text(super::slide_text::Error::Refused(
                        super::slide_text::Refusal::ObjectMarkerReplacement,
                    )),
                });
            }
        }
        if language.is_none()
            && locale.is_none()
            && size.is_none()
            && title.is_none()
            && subtitle.is_none()
            && notes.is_none()
        {
            return Ok(self.snapshot());
        }
        let record = self
            .slide_record_at(0)?
            .ok_or_else(|| ReadError::InvalidFormat("no first slide".to_owned()))?;
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

        // The first slide's title, subtitle and notes storages.
        let mut texts: Vec<(u64, Role, &str)> = Vec::new();
        if title.is_some() || subtitle.is_some() || notes.is_some() {
            for &identifier in &identifiers {
                let object = workspace.require(identifier).map_err(invalid)?;
                for message in &object.messages {
                    match message.type_ {
                        PLACEHOLDER_MESSAGE_TYPE => {
                            if let Some((kind, Some(parent), Some(storage))) =
                                placeholder_storage(&message.data, wire).map_err(invalid)?
                                && parent == record.slide_identifier
                            {
                                match (kind, title, subtitle) {
                                    (KIND_TITLE, Some(text), _) => {
                                        texts.push((storage, Role::Title, text))
                                    },
                                    (KIND_BODY, _, Some(text)) => {
                                        texts.push((storage, Role::Subtitle, text))
                                    },
                                    _ => {},
                                }
                            }
                        },
                        super::NOTE_MESSAGE_TYPE => {
                            if let (Some(text), Some(storage)) = (
                                notes,
                                reference(&message.data, NOTE_STORAGE, wire).map_err(invalid)?,
                            ) {
                                texts.push((storage, Role::Notes, text));
                            }
                        },
                        _ => {},
                    }
                }
            }
            let expected = usize::from(title.is_some())
                + usize::from(subtitle.is_some())
                + usize::from(notes.is_some());
            if texts.len() != expected {
                return Err(ReadError::InvalidFormat("blank text storages".to_owned()).into());
            }
        }
        let archive_limits = workspace.archive_limits();
        let max_message_bytes = self
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|error| ReadError::InvalidFormat(error.to_string()))?
            .max_message_bytes();
        let mut slide_text_changed = false;
        let mut changed = false;
        for identifier in identifiers {
            let object = workspace.require(identifier).map_err(invalid)?;
            let mut rewrites = Vec::new();
            for (index, message) in object.messages.iter().enumerate() {
                let mut data = match message.type_ {
                    PLACEHOLDER_MESSAGE_TYPE => match size {
                        Some(size) => placeholder(&message.data, size, wire),
                        None => Ok(None),
                    },
                    super::SHOW_MESSAGE_TYPE => match size {
                        Some(size) => show(&message.data, size, wire),
                        None => Ok(None),
                    },
                    other => super::locale::rewrite(other, &message.data, language, locale, wire),
                }
                .map_err(invalid)?;
                if message.type_ == STORAGE_MESSAGE_TYPE
                    && let Some(&(_, role, text)) =
                        texts.iter().find(|(storage, ..)| *storage == identifier)
                {
                    let current = data.as_deref().unwrap_or(&message.data);
                    let rewritten = match role {
                        Role::Notes => {
                            let limits =
                                super::slide_notes::storage_rewrite_limits(self, max_message_bytes)
                                    .map_err(FreshError::Notes)?;
                            rewrite_storage(current, text, limits).map_err(|error| {
                                FreshError::Notes(super::slide_notes::map_text_rewrite_error(error))
                            })?
                        },
                        Role::Title | Role::Subtitle => {
                            let limits =
                                super::slide_text::storage_rewrite_limits(self, max_message_bytes)
                                    .map_err(FreshError::Text)?;
                            rewrite_storage(current, text, limits).map_err(|error| {
                                FreshError::Text(super::slide_text::map_text_rewrite_error(error))
                            })?
                        },
                    };
                    if let Some(rewritten) = rewritten {
                        slide_text_changed |= !matches!(role, Role::Notes);
                        data = Some(rewritten);
                    }
                }
                if let Some(data) = data {
                    rewrites.push((index, message.type_, data));
                }
            }
            for (index, message_type, data) in rewrites {
                workspace
                    .object_mut(identifier)
                    .map_err(invalid)?
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: message_type,
                            data,
                        },
                        archive_limits,
                    )
                    .map_err(|error| invalid(error.into()))?;
                changed = true;
            }
        }
        if slide_text_changed {
            let wire_limits = self
                .wire_limits()
                .map_err(|error| ReadError::InvalidFormat(error.to_string()))?;
            let node = workspace
                .object_mut(record.node_identifier)
                .map_err(invalid)?;
            super::slide_preview::invalidate(node, archive_limits, wire_limits).map_err(
                |error| {
                    ReadError::InvalidFormat(format!(
                        "Keynote slide preview invalidation: {error:?}"
                    ))
                },
            )?;
        }
        if !changed {
            return Ok(self.snapshot());
        }
        let bytes = workspace.publish().map_err(invalid)?;
        let candidate = self.structure_reopen(bytes)?;
        let first = crate::SlideSelector::index(0);
        let verified = title.is_none_or(|text| {
            candidate
                .slide_title(first)
                .ok()
                .flatten()
                .as_deref()
                .unwrap_or("")
                == text
        }) && subtitle.is_none_or(|text| {
            candidate
                .slide_body(first)
                .ok()
                .flatten()
                .as_deref()
                .unwrap_or("")
                == text
        }) && notes.is_none_or(|text| {
            candidate
                .slide_notes(first)
                .ok()
                .flatten()
                .as_deref()
                .unwrap_or("")
                == text
        });
        if !verified {
            return Err(ReadError::InvalidFormat("fresh text verification".to_owned()).into());
        }
        Ok(candidate)
    }
}

/// Replace the whole text of a storage payload through the shared text
/// engine (`PreserveOnEqualText`); `None` when it already has `text`.
fn rewrite_storage(
    payload: &[u8],
    text: &str,
    limits: litchi_iwa_text_wire::RewriteLimits,
) -> Result<Option<Vec<u8>>, litchi_iwa_text_wire::RewriteError> {
    let current = litchi_iwa_text_wire::decode_storage(payload, limits)?
        .into_storage()
        .into_text();
    let rewrite = litchi_iwa_text_wire::rewrite_storage_text_with_behavior(
        payload,
        0..current.encode_utf16().count(),
        text,
        litchi_iwa_text_wire::RewriteBehavior::PreserveOnEqualText,
        limits,
    )?;
    Ok(rewrite.changed().then(|| rewrite.into_bytes()))
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
            .with_fresh_settings(FreshSettings {
                size: Some(size),
                ..FreshSettings::default()
            })
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
        let same = blank
            .with_fresh_settings(FreshSettings {
                size: Some(size),
                ..FreshSettings::default()
            })
            .expect("layout");
        assert_eq!(frames(&same), frames(&blank));
        let mut before = Vec::new();
        let mut after = Vec::new();
        blank.write_to(&mut before).expect("write");
        same.write_to(&mut after).expect("write");
        assert_eq!(before, after);
    }
}
