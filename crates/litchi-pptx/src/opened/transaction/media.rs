//! Slide media pictures (record 0321).
//!
//! [`Edit::add_media`] embeds one audio or video file with its poster
//! image as a `p:pic`: the media part (`/ppt/media/mediaN.ext`), its
//! `video`/`audio` and PowerPoint 2010 `media` slide relationships, the
//! poster part and its image relationship. Removal is the ordinary
//! [`Edit::remove_shape`].

use litchi_opc::PackURI;

use super::{Edit, invalid, next_shape_id};
use crate::Result;
use crate::media_parts::{Extension, Kind, List, Picture, Poster, Resource, Transform};

/// One embedded file: its bytes, content type and file extension.
#[derive(Debug, Clone)]
pub struct MediaFile<'a> {
    /// The file's bytes.
    pub data: &'a [u8],
    /// Its content type (`video/mp4`, `audio/mpeg`, `image/png`).
    pub content_type: &'a str,
    /// Its file extension, without the dot (`mp4`).
    pub extension: &'a str,
}

fn free_relationship(rels: &litchi_opc::Relationships, taken: &[String]) -> Result<String> {
    for number in 1..=u32::MAX {
        let id = format!("rId{number}");
        if rels.get(&id).is_none() && !taken.contains(&id) {
            return Ok(id);
        }
    }
    Err(invalid("the slide has no free relationship ID"))
}

fn free_part(package: &litchi_opc::OpcPackage, stem: &str, extension: &str) -> Result<PackURI> {
    if extension.is_empty() || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        return Err(invalid("the media file extension is not alphanumeric"));
    }
    crate::shape::theme::next_part_uri(package, stem, &format!(".{extension}"))
}

impl Edit {
    /// Embed an audio or video file with its poster image as a media
    /// picture on a slide, at `bounds` (`(x, y, width, height)` in EMU).
    /// Returns the new picture's shape ID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, a content type that does not
    /// match `kind` (or a poster that is not an image), invalid name text,
    /// non-positive extents, or exhausted shape IDs.
    pub fn add_media<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        kind: Kind,
        media: &MediaFile<'_>,
        poster: &MediaFile<'_>,
        name: &str,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        if bounds.2 <= 0 || bounds.3 <= 0 {
            return Err(invalid("media picture extents must be positive"));
        }
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        crate::package::signature::refuse_signed_change(&candidate)?;
        let slide_part = candidate.get_part(&selected.part_name)?;
        let media_limits = self
            .semantic_limits()
            .media_parts()
            .constrained_by_source(candidate.read_limits());
        let shape_limits = self.semantic_limits().shape();
        let conformance =
            crate::media_parts::document_conformance_with_limits(slide_part.blob(), media_limits)?;
        let scene = crate::shape::Scene::read(slide_part.blob(), shape_limits)?;
        let shape_id = next_shape_id(&scene, "media picture")?;
        drop(scene);
        let rels = slide_part.rels();
        let mut taken = Vec::new();
        for _ in 0..3 {
            let id = free_relationship(rels, &taken)?;
            taken.push(id);
        }
        let media_name = free_part(&candidate, "/ppt/media/media", media.extension)?;
        let poster_name = free_part(&candidate, "/ppt/media/image", poster.extension)?;
        let picture = Picture {
            shape_id,
            name: name.to_owned(),
            kind,
            relationship_id: taken[0].clone(),
            resource: Some(Resource::new(
                media_name.as_str(),
                media.content_type,
                media.data.to_vec(),
            )),
            poster: Some(Poster {
                relationship_id: taken[2].clone(),
                resource: Some(Resource::new(
                    poster_name.as_str(),
                    poster.content_type,
                    poster.data.to_vec(),
                )),
            }),
            transform: Some(Transform::emu(bounds.0, bounds.1, bounds.2, bounds.3)?),
            office_extension: Some(Extension {
                embed_relationship_id: Some(taken[1].clone()),
                link_relationship_id: None,
                trim: None,
                fade: None,
                bookmarks: Vec::new(),
                extensions: None,
            }),
        };
        let list = List {
            pictures: vec![picture],
        };
        crate::media_parts::store_with_limits(
            &mut candidate,
            &selected.part_name,
            &list,
            conformance,
            true,
            media_limits,
        )?;
        let staged = crate::shape::Scene::read(
            candidate.get_part(&selected.part_name)?.blob(),
            shape_limits,
        )?;
        let added = staged
            .iter()
            .find(|shape| shape.id() == Some(shape_id))
            .ok_or_else(|| invalid("the media picture did not round-trip"))?;
        if crate::media_parts::summary(added)?.map(|summary| summary.kind) != Some(kind) {
            return Err(invalid("the media picture did not round-trip"));
        }
        crate::package::signature::settle_unsigned_change(&mut candidate)?;
        self.working = candidate;
        Ok(shape_id)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::MediaFile;
    use crate::media_parts::Kind;
    use litchi_core::Position;

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";

    #[test]
    fn two_media_pictures_on_one_slide() {
        let source = crate::Package::new().unwrap();
        let mut edit = source.edit().unwrap();
        edit.insert_slide(Position::new(0), "Blank").unwrap();
        let video = MediaFile {
            data: b"not really mp4",
            content_type: "video/mp4",
            extension: "mp4",
        };
        let audio = MediaFile {
            data: b"not really mp3",
            content_type: "audio/mpeg",
            extension: "mp3",
        };
        let poster = MediaFile {
            data: PNG,
            content_type: "image/png",
            extension: "png",
        };
        let first = edit
            .add_media(
                0_usize,
                Kind::Video,
                &video,
                &poster,
                "Video 1",
                (0, 0, 100, 100),
            )
            .unwrap();
        let second = edit
            .add_media(
                0_usize,
                Kind::Audio,
                &audio,
                &poster,
                "Audio 1",
                (0, 0, 100, 100),
            )
            .unwrap();
        assert_ne!(first, second);
        assert!(
            edit.add_media(
                0_usize,
                Kind::Audio,
                &video,
                &poster,
                "Bad",
                (0, 0, 100, 100)
            )
            .is_err()
        );
        let package = edit.commit().unwrap().into_package();
        let presentation = package.presentation().unwrap();
        let slide = presentation.slide(0).unwrap().unwrap();
        let scene = slide.scene().unwrap();
        let kinds: Vec<_> = scene
            .iter()
            .filter_map(|shape| crate::media_parts::summary(shape.view()).unwrap())
            .map(|summary| summary.kind)
            .collect();
        assert_eq!(kinds, vec![Kind::Video, Kind::Audio]);
    }
}
