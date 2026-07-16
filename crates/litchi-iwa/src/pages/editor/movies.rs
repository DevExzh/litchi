//! Body-anchored movie CRUD for Pages documents.

use std::time::Duration;

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::{add_component_external_reference, component_identifier_for_entry};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

mod graph;

use graph::*;

/// One ordinary file-backed movie anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesMovieInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub movie_data_identifier: u64,
    pub poster_image_data_identifier: u64,
    pub geometry: DrawableGeometry,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
    pub duration: Duration,
}

/// Typed layout and playback metadata for a newly created Pages movie.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesMovieOptions {
    /// Top-left position on the page, in points.
    pub position: DrawablePoint,
    /// Displayed movie size, in points.
    pub size: DrawableSize,
    /// Untransformed media dimensions reported to Pages, in points.
    pub natural_size: DrawableSize,
    /// Playable duration of the source movie.
    pub duration: Duration,
}

impl PagesMovieOptions {
    /// Create options whose displayed and natural dimensions are identical.
    pub const fn new(position: DrawablePoint, size: DrawableSize, duration: Duration) -> Self {
        Self {
            position,
            size,
            natural_size: size,
            duration,
        }
    }

    /// Set media dimensions independently of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: DrawableSize) -> Self {
        self.natural_size = natural_size;
        self
    }
}

/// Result of removing one body-anchored Pages movie and its private graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesMovie {
    pub movie: PagesMovieInfo,
    /// Assets culled because the removed movie held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl PagesEditor {
    /// List ordinary file-backed movies anchored to the body in text-flow order.
    pub fn body_movies(&self) -> Result<Vec<PagesMovieInfo>> {
        body_movie_infos(self)
    }

    /// Add an independently editable movie anchored at a UTF-16 body position.
    ///
    /// The movie, poster, title/caption stand-ins, body attachment,
    /// object-replacement character, z-order, UUIDs, component data references,
    /// and `Data/*` assets are constructed directly from typed values. No source
    /// drawable or package is copied.
    pub fn add_body_movie(
        &mut self,
        anchor_character_index: usize,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: PagesMovieOptions,
    ) -> Result<PagesMovieInfo> {
        let (geometry, duration_seconds) = movie_creation_values(options)?;
        let root = root_document(self.package())?;
        let style_id = movie_style_id(self.package(), &root)?;
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = MovieObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let movie_asset = media.insert_unreferenced(preferred_movie_filename, movie_data)?;
        if movie_asset.media_type != crate::MediaType::Video {
            return Err(Error::ParseError(format!(
                "Pages body movies require video data, not {}",
                movie_asset.media_type.name()
            )));
        }
        let poster_asset = media.insert_unreferenced(preferred_poster_filename, poster_data)?;
        if poster_asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Pages movie posters require image data, not {}",
                poster_asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        let objects = movie_objects(
            ids,
            self.body_storage_id,
            style_id,
            movie_asset.data_identifier,
            poster_asset.data_identifier,
            geometry,
            options.natural_size,
            duration_seconds,
            root.left_margin.unwrap_or_default(),
        )?;
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;

        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            ids.attachment,
        )?;
        patch_pages_zorder(&mut staged, None, Some(ids.drawable))?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &ids.uuid_objects())?;
        for data_identifier in [movie_asset.data_identifier, poster_asset.data_identifier] {
            add_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                ids.drawable,
            )?;
        }
        let style_archive = find_object_archive(&staged, style_id)?;
        if let Some(stylesheet_component) = component_identifier_for_entry(&staged, &style_archive)?
            && stylesheet_component != DOCUMENT_OBJECT_ID
        {
            add_component_external_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                stylesheet_component,
                style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .body_movies()?
            .into_iter()
            .find(|movie| movie.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages movie creation failed validation".to_owned())
            })?;
        let created_graph = body_movie_graph(&verified, ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.anchor_character_index != expected_anchor
            || created.movie_data_identifier != movie_asset.data_identifier
            || created.poster_image_data_identifier != poster_asset.data_identifier
            || created.geometry != geometry
            || created.original_size != Some(options.natural_size)
            || created.natural_size != Some(options.natural_size)
            || created.duration.as_secs_f32() != duration_seconds
            || created_graph.object_ids != ids.all()
            || verified.extract_media(movie_asset.data_identifier)? != movie_data
            || verified.extract_media(poster_asset.data_identifier)? != poster_data
        {
            return Err(Error::InvalidFormat(
                "Pages movie creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read typed geometry for one body-anchored movie.
    pub fn body_movie_geometry(&self, drawable_object_id: u64) -> Result<DrawableGeometry> {
        Ok(body_movie_graph(self, drawable_object_id)?.info.geometry)
    }

    /// Update movie geometry while preserving unknown movie fields.
    pub fn set_body_movie_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = body_movie_graph(self, drawable_object_id)?;
        let position = geometry.position.ok_or_else(|| {
            Error::ParseError("Pages body movie geometry requires a position".to_owned())
        })?;
        let left_margin = root_document(self.package())?
            .left_margin
            .unwrap_or_default();
        let mut staged = self.package().clone();
        set_movie_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        set_movie_attachment_position(
            &mut staged,
            &source.archive_name,
            source.attachment_id,
            position,
            left_margin,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_movie_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages movie geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the video bytes referenced by one body-anchored movie.
    pub fn replace_body_movie_data(
        &mut self,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = body_movie_graph(self, drawable_object_id)?;
        self.replace_media(source.info.movie_data_identifier, replacement)
    }

    /// Replace the poster image referenced by one body-anchored movie.
    pub fn replace_body_movie_poster(
        &mut self,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = body_movie_graph(self, drawable_object_id)?;
        self.replace_media(source.info.poster_image_data_identifier, replacement)
    }

    /// Remove a body movie, its attachment/private graph, and unshared assets.
    pub fn remove_body_movie(&mut self, drawable_object_id: u64) -> Result<RemovedPagesMovie> {
        let source = body_movie_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = source.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            let object_archive = find_object_archive(&staged, *identifier)?;
            staged.update_archive(&object_archive, |archive| {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages movie object {identifier} is missing from {object_archive}"
                    ))
                })?;
                Ok(())
            })?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages movie object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let mut media = IWorkMediaEditor::from_package(staged)?;
        let mut removed_data_identifiers = Vec::new();
        let data_identifiers = source
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<HashSet<_>>();
        for identifier in data_identifiers {
            if media
                .asset(identifier)
                .is_some_and(|asset| !asset.is_referenced())
            {
                media.remove_unreferenced(identifier)?;
                removed_data_identifiers.push(identifier);
            }
        }
        removed_data_identifiers.sort_unstable();
        let staged = media.into_package();
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let remaining_assets = verified.media_assets()?;
        if verified
            .body_movies()?
            .iter()
            .any(|movie| movie.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Pages movie deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesMovie {
            movie: source.info,
            removed_data_identifiers,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-pages-movie";
    const REPLACEMENT_MOVIE: &[u8] = b"\0\0\0\x18ftypqt  replacement-pages-movie";
    const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-pages-poster";
    const REPLACEMENT_POSTER: &[u8] = b"\x89PNG\r\n\x1a\nreplacement-pages-poster";
    const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 180.0,
    };

    fn options() -> PagesMovieOptions {
        PagesMovieOptions::new(POSITION, SIZE, Duration::from_secs(8))
    }

    #[test]
    fn scratch_document_supports_movie_crud_without_a_source_package() {
        let mut editor = PagesEditor::create_with_text("Quarterly report").unwrap();
        let anchor = "Quarterly report".encode_utf16().count();
        assert!(editor.body_movies().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());

        let created = editor
            .add_body_movie(anchor, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        assert_eq!(created.anchor_character_index, anchor as u32);
        assert_eq!(created.geometry.position, Some(POSITION));
        assert_eq!(created.geometry.size, Some(SIZE));
        assert_eq!(created.original_size, Some(SIZE));
        assert_eq!(created.natural_size, Some(SIZE));
        assert_eq!(created.duration, Duration::from_secs(8));
        assert_eq!(editor.body_text().unwrap(), "Quarterly report\u{fffc}");
        assert_eq!(
            editor.extract_media(created.movie_data_identifier).unwrap(),
            MOVIE
        );
        assert_eq!(
            editor
                .extract_media(created.poster_image_data_identifier)
                .unwrap(),
            POSTER
        );

        let roundtripped = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.body_movies().unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 72.0, y: 216.0 }),
            size: Some(DrawableSize {
                width: 400.0,
                height: 225.0,
            }),
            flags: Some(3),
            angle: Some(7.5),
        };
        editor
            .set_body_movie_geometry(created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .body_movie_geometry(created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );
        assert_eq!(
            editor
                .replace_body_movie_data(created.drawable_object_id, REPLACEMENT_MOVIE)
                .unwrap(),
            MOVIE
        );
        assert_eq!(
            editor
                .replace_body_movie_poster(created.drawable_object_id, REPLACEMENT_POSTER)
                .unwrap(),
            POSTER
        );
        editor
            .set_drawable_comment(created.drawable_object_id, "Remove this movie after review")
            .unwrap();

        let removed = editor
            .remove_body_movie(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.movie.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [
                created.movie_data_identifier,
                created.poster_image_data_identifier,
            ]
        );
        assert_eq!(editor.body_text().unwrap(), "Quarterly report");
        assert!(editor.body_movies().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_movie_creation_and_cross_type_edits_are_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        for result in [
            editor.add_body_movie(4, "poster.png", POSTER, "poster.png", POSTER, options()),
            editor.add_body_movie(4, "movie.mov", MOVIE, "movie.mov", MOVIE, options()),
            editor.add_body_movie(
                4,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                PagesMovieOptions::new(POSITION, SIZE, Duration::ZERO),
            ),
            editor.add_body_movie(5, "movie.mov", MOVIE, "poster.png", POSTER, options()),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        let image = editor
            .add_body_image(4, "poster.png", POSTER, POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_movie_geometry(image.drawable_object_id, DrawableGeometry::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .replace_body_movie_data(image.drawable_object_id, MOVIE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn removing_an_earlier_movie_shifts_and_preserves_later_attachments() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let first = editor
            .add_body_movie(4, "first.mov", MOVIE, "first.png", POSTER, options())
            .unwrap();
        let second = editor
            .add_body_movie(
                5,
                "second.mov",
                REPLACEMENT_MOVIE,
                "second.png",
                REPLACEMENT_POSTER,
                options(),
            )
            .unwrap();

        editor.remove_body_movie(first.drawable_object_id).unwrap();
        let remaining = editor.body_movies().unwrap();
        assert_eq!(remaining.len(), 1);
        assert_eq!(remaining[0].drawable_object_id, second.drawable_object_id);
        assert_eq!(remaining[0].anchor_character_index, 4);
        assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}");
        assert_eq!(
            editor.extract_media(second.movie_data_identifier).unwrap(),
            REPLACEMENT_MOVIE
        );
        assert_eq!(
            editor
                .extract_media(second.poster_image_data_identifier)
                .unwrap(),
            REPLACEMENT_POSTER
        );
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }
}
