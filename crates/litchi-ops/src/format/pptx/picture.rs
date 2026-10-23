//! Picture insertion through the native eager edit (record 0252).
//!
//! `add(/slide[N], picture, {src, x, y, width, height, name})` reads the
//! image from a `data:` URI (base64) or a local file path, admits it against
//! the caller's staging limit, classifies it by its signature (PNG, JPEG,
//! GIF, BMP, TIFF), and stages the media part, the slide relationship and
//! the `p:pic` through `opened::Edit::add_picture`. Source owners are
//! promoted to an eager package for the edit, since the source-backed slide
//! edit does not stage new parts.
use super::node::Item;
use super::spec::{Kind, PICTURE_WRITE_TABLE, PictureSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options, validate_text};
use super::{Editable, Nav, Pptx, steps};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{Error as TxError, LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Result};
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use std::sync::Arc;

/// Default picture extent when the caller supplies none: 2in × 1.5in.
const DEFAULT_EXTENT: (i64, i64) = (1_828_800, 1_371_600);

fn allocation(source: std::collections::TryReserveError) -> crate::Error {
    crate::Error::from(TxError::Allocation { source })
}

/// The image content type and file extension of `data`, by signature.
pub(super) fn classify(data: &[u8]) -> Option<(&'static str, &'static str)> {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some(("image/png", "png"))
    } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some(("image/jpeg", "jpeg"))
    } else if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        Some(("image/gif", "gif"))
    } else if data.starts_with(b"BM") {
        Some(("image/bmp", "bmp"))
    } else if data.starts_with(b"II*\0") || data.starts_with(b"MM\0*") {
        Some(("image/tiff", "tiff"))
    } else {
        None
    }
}

/// Read a picture source: a base64 `data:` URI or a local file path.
pub(super) fn read_source(src: &str, maximum: usize) -> Result<Vec<u8>> {
    if let Some(rest) = src.strip_prefix("data:") {
        let (header, payload) = rest.split_once(',').ok_or_else(|| {
            error(
                Code::InvalidValue,
                "a data URI needs a comma before its payload",
            )
            .with_property("src")
        })?;
        if !header.ends_with(";base64") {
            return Err(error(
                Code::UnsupportedCapability,
                "only base64 data URIs are accepted",
            )
            .with_property("src"));
        }
        // Base64 expands by 4/3; refuse before decoding.
        bound(LimitKind::StagedBytes, payload.len() / 4 * 3, maximum)?;
        use base64::Engine as _;
        return base64::engine::general_purpose::STANDARD
            .decode(payload.trim())
            .map_err(|_| {
                error(Code::InvalidValue, "the data URI payload is not base64").with_property("src")
            });
    }
    let length = std::fs::metadata(src)
        .map_err(|source| crate::Error::from(source).with_property("src"))?
        .len();
    bound(
        LimitKind::StagedBytes,
        usize::try_from(length).unwrap_or(usize::MAX),
        maximum,
    )?;
    use std::io::Read as _;
    let mut data = Vec::new();
    data.try_reserve_exact(usize::try_from(length).unwrap_or(0))
        .map_err(allocation)?;
    std::fs::File::open(src)
        .and_then(|file| file.take(maximum as u64 + 1).read_to_end(&mut data))
        .map_err(|source| crate::Error::from(source).with_property("src"))?;
    bound(LimitKind::StagedBytes, data.len(), maximum)?;
    Ok(data)
}

impl Pptx {
    pub(super) fn add_picture(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let parent = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "pictures require a slide parent",
            ));
        };
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "picture insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut src = None;
        let mut name = None;
        let mut alt = None;
        let mut geometry = [None; 4];
        let mut assigned = 0_u32;
        let mut click = super::links::ClickInput::default();
        for (key, value) in node.props().iter() {
            let (index, entry) = PICTURE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a picture property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match (setter, scalar) {
                    (PictureSetter::Click(key), scalar) => click.stage(key, scalar, limits)?,
                    (PictureSetter::Src, Some(Scalar::Text(value))) => {
                        src = Some(value.to_owned());
                    },
                    (PictureSetter::Name, Some(Scalar::Text(value))) => {
                        validate_text(value)?;
                        name = Some(value.to_owned());
                    },
                    (PictureSetter::Alt, Some(Scalar::Text(value))) => {
                        validate_text(value).map_err(|e| e.with_property("alt"))?;
                        alt = Some(value.to_owned());
                    },
                    (setter, Some(Scalar::Length(value))) => {
                        let index = match setter {
                            PictureSetter::X => 0,
                            PictureSetter::Y => 1,
                            PictureSetter::Width => 2,
                            PictureSetter::Height => 3,
                            PictureSetter::Src
                            | PictureSetter::Name
                            | PictureSetter::Alt
                            | PictureSetter::Click(_) => {
                                unreachable!("text picture setters")
                            },
                        };
                        geometry[index] = Some(value.emu());
                    },
                    _ => unreachable!("picture property coercion"),
                }
                Ok(())
            })?;
        }
        let src = src.ok_or_else(|| {
            error(Code::MissingProperty, "pictures require a src").with_property("src")
        })?;
        let data = read_source(&src, limits.max_staged_bytes())?;
        let (content_type, extension) = classify(&data).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the picture is not a PNG, JPEG, GIF, BMP or TIFF image",
            )
            .with_property("src")
        })?;
        let bounds = (
            geometry[0].unwrap_or(0),
            geometry[1].unwrap_or(0),
            geometry[2].unwrap_or(DEFAULT_EXTENT.0),
            geometry[3].unwrap_or(DEFAULT_EXTENT.1),
        );
        let slide = slide.position as usize;
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        let ordinal = scene
            .roots()
            .filter(|shape| steps::tree_kind(shape) == Some(Kind::Picture))
            .count();
        let name = name.unwrap_or_else(|| format!("Picture {}", ordinal + 1));
        let resource = litchi_pptx::media_parts::Resource::new(
            format!("/ppt/media/image1.{extension}"),
            content_type,
            data,
        );
        let source = self.editable_owner()?;
        let base = match &source {
            Editable::Eager(package) => Arc::clone(package),
            Editable::Source(package) => {
                super::eager_package_from_source(package.to_bytes()?, self.0.ooxml_limits)?
            },
        };
        let link = click.for_add(self.slides_in_catalog()?)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        edit.add_picture(slide, &name, &resource, bounds)?;
        if let Some(alt) = alt.as_deref() {
            // The new picture is the last shape of the slide in pre-order.
            edit.set_shape_description(slide, scene.len(), Some(alt))?;
        }
        if let Some(link) = &link {
            // The new picture is the last shape of the slide in pre-order.
            let _ = edit.set_shape_hyperlink(slide, scene.len(), Some(link))?;
        }
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("picture").with(Predicate::Index(Position::new(ordinal))),
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

impl Pptx {
    /// A `set` of a picture's frame (`x`, `y`, `width`, `height`),
    /// alternative text (`alt`) or image (`src`), in one native edit
    /// (record 0333); `None` for other targets. Click actions are split off
    /// before (`try_set_click`).
    pub(super) fn try_set_picture(
        &self,
        target: &crate::target::TargetArg<'_>,
        props: &crate::Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            ["picture", "image", "img", "pic"]
                .iter()
                .any(|name| step.name().eq_ignore_ascii_case(name))
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let Item::Tree(tree) = node.item() else {
            return Ok(None);
        };
        if tree.kind != Kind::Picture {
            return Ok(None);
        }
        let Some(slide) = tree.container.slide() else {
            return Err(error(
                Code::UnsupportedCapability,
                "master and layout pictures are not editable",
            ));
        };
        let tree = tree.clone();
        drop(node);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut geometry = [None; 4];
        let mut alt: Option<Option<String>> = None;
        let mut src = None;
        let mut assigned = 0_u32;
        for (key, value) in props.iter() {
            let (index, entry) = PICTURE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a picture property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Set, value, |setter, scalar| {
                match (setter, scalar) {
                    (PictureSetter::Src, Some(Scalar::Text(value))) => src = Some(value.to_owned()),
                    (PictureSetter::Alt, Some(Scalar::Text(value))) => {
                        validate_text(value).map_err(|e| e.with_property("alt"))?;
                        alt = Some(Some(value.to_owned()));
                    },
                    (PictureSetter::Alt, None) => alt = Some(None),
                    (setter, Some(Scalar::Length(value))) => {
                        let index = match setter {
                            PictureSetter::X => 0,
                            PictureSetter::Y => 1,
                            PictureSetter::Width => 2,
                            _ => 3,
                        };
                        geometry[index] = Some(value.emu());
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        if geometry.iter().all(Option::is_none) && alt.is_none() && src.is_none() {
            return Err(error(
                Code::InvalidArgument,
                "no picture property was assigned",
            ));
        }
        let bounds = if geometry.iter().any(Option::is_some) {
            let current = tree.handle.view().bounds().ok_or_else(|| {
                error(
                    Code::MissingProperty,
                    "picture geometry requires a complete local transform",
                )
                .with_property("x")
            })?;
            Some((
                geometry[0].unwrap_or(current.x()),
                geometry[1].unwrap_or(current.y()),
                geometry[2].unwrap_or(current.width()),
                geometry[3].unwrap_or(current.height()),
            ))
        } else {
            None
        };
        let resource = match &src {
            Some(src) => {
                let data = read_source(src, limits.max_staged_bytes())?;
                let (content_type, extension) = classify(&data).ok_or_else(|| {
                    error(
                        Code::UnsupportedCapability,
                        "the picture is not a PNG, JPEG, GIF, BMP or TIFF image",
                    )
                    .with_property("src")
                })?;
                Some(litchi_pptx::media_parts::Resource::new(
                    format!("/ppt/media/image1.{extension}"),
                    content_type,
                    data,
                ))
            },
            None => None,
        };
        let slide = slide as usize;
        let shape = tree.handle.position().get();
        let source = self.editable_owner_for_write()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let mut operations = 0;
        if let Some(bounds) = bounds {
            operations += edit.set_shape_properties(
                slide,
                &[
                    litchi_pptx::opened::ShapePropertiesReplacement::new(Position::new(shape))
                        .with_bounds(bounds),
                ],
            )?;
        }
        if let Some(alt) = &alt {
            operations += usize::from(edit.set_shape_description(slide, shape, alt.as_deref())?);
        }
        if let Some(resource) = &resource {
            operations += usize::from(edit.set_picture_image(slide, shape, resource)?);
        }
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            operations,
        )))
    }
}
