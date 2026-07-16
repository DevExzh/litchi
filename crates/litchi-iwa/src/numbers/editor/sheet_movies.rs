//! Standalone movie-object CRUD for Numbers sheets.

use std::time::Duration;

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

pub(super) mod graph;

use graph::*;

/// One ordinary file-backed movie owned directly by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetMovieInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub movie_data_identifier: u64,
    pub poster_image_data_identifier: u64,
    pub geometry: DrawableGeometry,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
    pub duration: Duration,
}

/// Typed layout and playback metadata for a newly created Numbers movie.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct NumbersSheetMovieOptions {
    /// Top-left position on the sheet, in points.
    pub position: DrawablePoint,
    /// Displayed movie size, in points.
    pub size: DrawableSize,
    /// Untransformed media dimensions reported to Numbers, in points.
    pub natural_size: DrawableSize,
    /// Playable duration of the source movie.
    pub duration: Duration,
}

impl NumbersSheetMovieOptions {
    /// Create options whose displayed and natural dimensions are identical.
    pub const fn new(position: DrawablePoint, size: DrawableSize, duration: Duration) -> Self {
        Self {
            position,
            size,
            natural_size: size,
            duration,
        }
    }

    /// Set dimensions independent of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: DrawableSize) -> Self {
        self.natural_size = natural_size;
        self
    }
}

/// Result of removing one sheet-owned movie and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetMovie {
    pub movie: NumbersSheetMovieInfo,
    /// Assets culled because the removed movie held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl NumbersEditor {
    /// List ordinary file-backed movies owned directly by one reachable sheet.
    pub fn sheet_movies(&self, sheet_id: u64) -> Result<Vec<NumbersSheetMovieInfo>> {
        movie_infos(self, sheet_id)
    }

    /// Add an independently editable, file-backed movie to a reachable sheet.
    ///
    /// The movie, poster, title/caption stand-ins, sheet ownership, style link,
    /// UUIDs, component data references, and `Data/*` assets are constructed
    /// directly from typed values. No source drawable or package is copied.
    pub fn add_sheet_movie(
        &mut self,
        sheet_id: u64,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: NumbersSheetMovieOptions,
    ) -> Result<NumbersSheetMovieInfo> {
        let (geometry, duration_seconds) = movie_creation_values(options)?;
        let context = movie_creation_context(self, sheet_id)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(&self.package)?)?;

        let mut media = IWorkMediaEditor::from_package(self.package.clone())?;
        let movie_asset = media.insert_unreferenced(preferred_movie_filename, movie_data)?;
        if movie_asset.media_type != crate::MediaType::Video {
            return Err(Error::ParseError(format!(
                "Numbers sheet movies require video data, not {}",
                movie_asset.media_type.name()
            )));
        }
        let poster_asset = media.insert_unreferenced(preferred_poster_filename, poster_data)?;
        if poster_asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Numbers movie posters require image data, not {}",
                poster_asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        let objects = movie_objects(
            ids,
            sheet_id,
            context.style_id,
            movie_asset.data_identifier,
            poster_asset.data_identifier,
            geometry,
            options.natural_size,
            duration_seconds,
        )?;
        staged.update_archive(&context.archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &context.archive_name,
            sheet_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, context.component_id, &ids.all())?;
        for data_identifier in [movie_asset.data_identifier, poster_asset.data_identifier] {
            add_component_data_reference(
                &mut staged,
                context.component_id,
                data_identifier,
                ids.drawable,
            )?;
        }
        if context.stylesheet_component_id != context.component_id {
            add_component_external_reference(
                &mut staged,
                context.component_id,
                context.stylesheet_component_id,
                context.style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_movies(sheet_id)?
            .into_iter()
            .find(|movie| movie.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers movie creation failed validation".to_owned())
            })?;
        let created_graph = movie_graph(&verified, sheet_id, ids.drawable)?;
        if created.movie_data_identifier != movie_asset.data_identifier
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
                "Numbers movie creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read typed geometry for one ordinary sheet movie.
    pub fn sheet_movie_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(movie_graph(self, sheet_id, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update movie position, size, flags, and rotation while preserving
    /// unknown movie fields.
    pub fn set_sheet_movie_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = movie_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_movie_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_movie_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers movie geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the video bytes referenced by one ordinary sheet movie.
    pub fn replace_sheet_movie_data(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = movie_graph(self, sheet_id, drawable_object_id)?;
        self.replace_media(source.info.movie_data_identifier, replacement)
    }

    /// Replace the poster image referenced by one ordinary sheet movie.
    pub fn replace_sheet_movie_poster(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = movie_graph(self, sheet_id, drawable_object_id)?;
        self.replace_media(source.info.poster_image_data_identifier, replacement)
    }

    /// Remove an ordinary movie, its private graph, and unshared assets.
    pub fn remove_sheet_movie(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetMovie> {
        let source = movie_graph(self, sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            source.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers movie object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers movie object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
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
            .sheet_movies(sheet_id)?
            .iter()
            .any(|movie| movie.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Numbers movie deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetMovie {
            movie: source.info,
            removed_data_identifiers,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-numbers-movie";
    const REPLACEMENT_MOVIE: &[u8] = b"\0\0\0\x18ftypqt  replacement-numbers-movie";
    const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-numbers-poster";
    const REPLACEMENT_POSTER: &[u8] = b"\x89PNG\r\n\x1a\nreplacement-numbers-poster";
    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 180.0,
    };

    fn options() -> NumbersSheetMovieOptions {
        NumbersSheetMovieOptions::new(POSITION, SIZE, Duration::from_secs(8))
    }

    #[test]
    fn scratch_spreadsheet_supports_movie_crud_without_a_source_package() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Movies")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        assert!(editor.sheet_movies(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());

        let created = editor
            .add_sheet_movie(
                sheet_id,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                options(),
            )
            .unwrap();
        assert_eq!(created.sheet_id, sheet_id);
        assert_eq!(created.geometry.position, Some(POSITION));
        assert_eq!(created.geometry.size, Some(SIZE));
        assert_eq!(created.original_size, Some(SIZE));
        assert_eq!(created.natural_size, Some(SIZE));
        assert_eq!(created.duration, Duration::from_secs(8));
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

        let roundtripped = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.sheet_movies(sheet_id).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 64.0, y: 96.0 }),
            size: Some(DrawableSize {
                width: 192.0,
                height: 108.0,
            }),
            flags: Some(3),
            angle: Some(9.0),
        };
        editor
            .set_sheet_movie_geometry(sheet_id, created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .sheet_movie_geometry(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );
        assert_eq!(
            editor
                .replace_sheet_movie_data(sheet_id, created.drawable_object_id, REPLACEMENT_MOVIE,)
                .unwrap(),
            MOVIE
        );
        assert_eq!(
            editor
                .replace_sheet_movie_poster(
                    sheet_id,
                    created.drawable_object_id,
                    REPLACEMENT_POSTER,
                )
                .unwrap(),
            POSTER
        );
        editor
            .set_sheet_drawable_comment(
                sheet_id,
                created.drawable_object_id,
                "Remove this movie after review",
            )
            .unwrap();

        let removed = editor
            .remove_sheet_movie(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.movie.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [
                created.movie_data_identifier,
                created.poster_image_data_identifier,
            ]
        );
        assert!(editor.sheet_movies(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_movie_creation_and_cross_type_edits_are_transactional() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

        for result in [
            editor.add_sheet_movie(
                sheet_id,
                "poster.png",
                POSTER,
                "poster.png",
                POSTER,
                options(),
            ),
            editor.add_sheet_movie(sheet_id, "movie.mov", MOVIE, "movie.mov", MOVIE, options()),
            editor.add_sheet_movie(
                sheet_id,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                NumbersSheetMovieOptions::new(POSITION, SIZE, Duration::ZERO),
            ),
            editor.add_sheet_movie(999, "movie.mov", MOVIE, "poster.png", POSTER, options()),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        let image = editor
            .add_sheet_image(sheet_id, "poster.png", POSTER, POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_movie_geometry(
                    sheet_id,
                    image.drawable_object_id,
                    DrawableGeometry::default(),
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .replace_sheet_movie_data(sheet_id, image.drawable_object_id, MOVIE)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
