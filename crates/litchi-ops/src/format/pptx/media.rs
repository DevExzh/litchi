//! Media pictures (record 0321, L-PPTX-R11).
//!
//! A `p:pic` whose non-visual properties name an audio or video file is a
//! media picture: `/slide[N]/video[M]`, `/slide[N]/audio[M]` and
//! `/slide[N]/media[M]` (either kind) address it, each kind with its own
//! counter, and it stays the picture it always was (`picture[K]`, and the
//! slide's view lists it as a picture). It reads the frame keys and the
//! PowerPoint 2010 playback trim (`trimStart`, `trimEnd`, in milliseconds).
//!
//! - `add(/slide[N], video|audio|media, {src, poster, name, x, y, width,
//!   height})` embeds the file with its poster image (`media` takes its
//!   kind from the file); OfficeCLI generates a poster when none is given,
//!   Litchi requires one.
//! - `remove` deletes the picture (the ordinary shape removal).
//!
//! Volume, auto-play and loop live in the slide's timing tree and are gaps.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use litchi_pptx::media_parts::{Kind as MediaKind, Summary};
use litchi_pptx::opened::MediaFile;
use litchi_pptx::shape::owned::Handle;

use super::node::{Item, TreeContainer, TreeItem};
use super::spec::{Kind, MEDIA_WRITE_TABLE, MediaSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options, validate_text};
use super::{Editable, Nav, Pptx, steps};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Result};

/// The default media frame: 1in from the top-left corner, 4in × 3in.
const DEFAULT_BOUNDS: (i64, i64, i64, i64) = (914_400, 914_400, 3_657_600, 2_743_200);

/// The media summary of a scene element, when it is a media picture.
pub(super) fn summary(handle: &Handle) -> Result<Option<Summary>> {
    if steps::tree_kind(handle) != Some(Kind::Picture) {
        return Ok(None);
    }
    Ok(litchi_pptx::media_parts::summary(handle.view())?)
}

/// Whether a media picture of `media` kind is an element of `kind`.
pub(super) const fn matches(kind: Kind, media: MediaKind) -> bool {
    matches!(
        (kind, media),
        (Kind::Media, _) | (Kind::Video, MediaKind::Video) | (Kind::Audio, MediaKind::Audio)
    )
}

/// The media pictures of `kind` among `handles`, with their ordinals.
pub(super) fn select(
    handles: impl Iterator<Item = Handle>,
    kind: Kind,
) -> Result<Vec<(u32, Handle)>> {
    let mut out = Vec::new();
    for handle in handles {
        if summary(&handle)?.is_some_and(|summary| matches(kind, summary.kind)) {
            let ordinal = u32::try_from(out.len()).unwrap_or(u32::MAX);
            out.push((ordinal, handle));
        }
    }
    Ok(out)
}

/// The media picture items of `kind` in a container.
pub(super) fn items(
    container: TreeContainer,
    handles: impl Iterator<Item = Handle>,
    kind: Kind,
) -> Result<Vec<Item>> {
    Ok(select(handles, kind)?
        .into_iter()
        .map(|(ordinal, handle)| {
            Item::Tree(TreeItem::in_container(container, kind, handle, ordinal))
        })
        .collect())
}

/// The trim value of a media picture: `start` or `end`, milliseconds.
pub(super) fn trim(tree: &TreeItem, start: bool) -> Result<Option<String>> {
    Ok(summary(&tree.handle)?
        .and_then(|summary| summary.trim)
        .and_then(|trim| if start { trim.start } else { trim.end })
        .map(|offset| offset.as_str().to_owned()))
}

/// The content type, extension and kind of a media file named `src`.
fn media_type(src: &str) -> Option<(&'static str, &'static str, MediaKind)> {
    let header = src
        .strip_prefix("data:")
        .and_then(|rest| rest.split_once(';'))
        .map(|(header, _)| header.to_ascii_lowercase());
    let extension = src
        .rsplit('.')
        .next()
        .map(str::to_ascii_lowercase)
        .unwrap_or_default();
    let table: [(&str, &str, &str, MediaKind); 11] = [
        ("video/mp4", "mp4", "m4v", MediaKind::Video),
        ("video/quicktime", "mov", "mov", MediaKind::Video),
        ("video/x-ms-wmv", "wmv", "wmv", MediaKind::Video),
        ("video/x-msvideo", "avi", "avi", MediaKind::Video),
        ("video/mpeg", "mpeg", "mpg", MediaKind::Video),
        ("video/webm", "webm", "webm", MediaKind::Video),
        ("audio/mpeg", "mp3", "mp3", MediaKind::Audio),
        ("audio/wav", "wav", "wav", MediaKind::Audio),
        ("audio/mp4", "m4a", "m4a", MediaKind::Audio),
        ("audio/x-ms-wma", "wma", "wma", MediaKind::Audio),
        ("audio/ogg", "ogg", "oga", MediaKind::Audio),
    ];
    table
        .into_iter()
        .find_map(|(content_type, ext, other, kind)| {
            let hit = match &header {
                Some(header) => header == content_type,
                None => extension == ext || extension == other,
            };
            hit.then_some((content_type, ext, kind))
        })
}

impl Pptx {
    /// `add(/slide[N], video|audio|media, {...})`.
    pub(super) fn add_media(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let requested = match node.kind().to_ascii_lowercase().as_str() {
            "video" => Some(MediaKind::Video),
            "audio" => Some(MediaKind::Audio),
            _ => None,
        };
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = resolved.item() else {
            return Err(error(
                Code::UnsupportedType,
                "media requires a slide parent",
            ));
        };
        let slide = slide.position as usize;
        drop(resolved);
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "media insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut src = None;
        let mut poster = None;
        let mut name = None;
        let mut geometry = [None; 4];
        let mut assigned = 0_u32;
        for (key, value) in node.props().iter() {
            let (index, entry) = MEDIA_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a media property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match (setter, scalar) {
                    (MediaSetter::Src, Some(Scalar::Text(value))) => src = Some(value.to_owned()),
                    (MediaSetter::Poster, Some(Scalar::Text(value))) => {
                        poster = Some(value.to_owned());
                    },
                    (MediaSetter::Name, Some(Scalar::Text(value))) => {
                        validate_text(value)?;
                        name = Some(value.to_owned());
                    },
                    (setter, Some(Scalar::Length(value))) => {
                        let slot = match setter {
                            MediaSetter::X => 0,
                            MediaSetter::Y => 1,
                            MediaSetter::Width => 2,
                            _ => 3,
                        };
                        geometry[slot] = Some(value.emu());
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let src = src.ok_or_else(|| {
            error(Code::MissingProperty, "media requires a src").with_property("src")
        })?;
        let poster = poster.ok_or_else(|| {
            error(Code::MissingProperty, "media requires a poster image")
                .with_property("poster")
                .with_suggestion("poster=/path/to/thumbnail.png")
        })?;
        let (content_type, extension, kind) = media_type(&src).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the media type is not recognized from the file's extension or data URI",
            )
            .with_property("src")
        })?;
        if requested.is_some_and(|requested| requested != kind) {
            return Err(error(
                Code::InvalidValue,
                "the media file's type does not match the element kind",
            )
            .with_property("src"));
        }
        let data = super::picture::read_source(&src, limits.max_staged_bytes())?;
        let image = super::picture::read_source(
            &poster,
            limits.max_staged_bytes().saturating_sub(data.len()),
        )
        .map_err(|error| error.with_property("poster"))?;
        let (image_type, image_extension) = super::picture::classify(&image).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the poster is not a PNG, JPEG, GIF, BMP or TIFF image",
            )
            .with_property("poster")
        })?;
        let element = match requested {
            Some(MediaKind::Video) => Kind::Video,
            Some(MediaKind::Audio) => Kind::Audio,
            None => Kind::Media,
        };
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        let ordinal = select(scene.roots(), element)?.len();
        let name = name.unwrap_or_else(|| {
            format!(
                "{} {}",
                if kind == MediaKind::Video {
                    "Video"
                } else {
                    "Audio"
                },
                ordinal + 1
            )
        });
        let bounds = (
            geometry[0].unwrap_or(DEFAULT_BOUNDS.0),
            geometry[1].unwrap_or(DEFAULT_BOUNDS.1),
            geometry[2].unwrap_or(DEFAULT_BOUNDS.2),
            geometry[3].unwrap_or(DEFAULT_BOUNDS.3),
        );
        let source = self.editable_owner()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        native.add_media(
            slide,
            kind,
            &MediaFile {
                data: &data,
                content_type,
                extension,
            },
            &MediaFile {
                data: &image,
                content_type: image_type,
                extension: image_extension,
            },
            &name,
            bounds,
        )?;
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new(element.name_static()).with(Predicate::Index(Position::new(ordinal))),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }
}

impl Kind {
    /// The element name of a media kind.
    const fn name_static(self) -> &'static str {
        match self {
            Self::Video => "video",
            Self::Audio => "audio",
            _ => "media",
        }
    }
}
