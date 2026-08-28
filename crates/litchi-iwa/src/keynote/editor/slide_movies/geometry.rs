//! Wire-preserving movie drawable geometry edits.

use super::*;
use crate::shapes::{
    DrawableFlipAxis, DrawableProperties, drawable_properties, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};
use litchi_core::Position;
use litchi_keynote::slide::media::geometry::{
    MovieFlipAxis as FocusedMovieFlipAxis, MovieGeometry,
};
use litchi_keynote::{MovieSelector, Package as KeynotePackage};

const MOVIE_DRAWABLE_FIELD: u32 = 1;

impl KeynoteEditor {
    /// Read the focused semantic geometry of one file-backed movie.
    ///
    /// The selector is resolved by `litchi_keynote::Package`; no native movie
    /// identifier crosses this compatibility seam.
    pub fn slide_movie_geometry_by_selector(
        &self,
        slide_position: Position,
        movie_selector: MovieSelector,
    ) -> Result<Option<MovieGeometry>> {
        focused_movie_geometry_package(self)?
            .slide_movie_geometry(slide_position, movie_selector)
            .map_err(map_focused_movie_geometry_error)
    }

    /// Replace the focused semantic geometry of one file-backed movie.
    pub fn set_slide_movie_geometry_by_selector(
        &mut self,
        slide_position: Position,
        movie_selector: MovieSelector,
        geometry: MovieGeometry,
    ) -> Result<()> {
        let package = focused_movie_geometry_package(self)?;
        let edit = package
            .edit_slide_movie_geometry(slide_position, movie_selector)
            .map_err(map_focused_movie_geometry_error)?
            .set(geometry)
            .map_err(map_focused_movie_geometry_error)?;
        let commit = edit.commit().map_err(map_focused_movie_geometry_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_movie_geometry_commit(self, commit)
    }

    /// Restore the focused movie's displayed size from its archive-free
    /// original-size metadata while preserving its current position.
    pub fn restore_slide_movie_original_size_by_selector(
        &mut self,
        slide_position: Position,
        movie_selector: MovieSelector,
    ) -> Result<MovieGeometry> {
        let package = focused_movie_geometry_package(self)?;
        let current = package
            .slide_movie_geometry(slide_position, movie_selector)
            .map_err(map_focused_movie_geometry_error)?
            .ok_or_else(|| {
                Error::InvalidFormat("focused Keynote movie geometry is unavailable".to_owned())
            })?;
        let original_size = package
            .slides()
            .map_err(|error| {
                Error::InvalidFormat(format!("focused Keynote slides failed: {error}"))
            })?
            .get(slide_position.get())
            .and_then(|slide| {
                slide
                    .video_movies()
                    .filter(|movie| movie.kind() == MovieKind::File)
                    .nth(movie_selector.as_index())
            })
            .and_then(|movie| movie.original_size())
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "focused Keynote movie has no original-size metadata".to_owned(),
                )
            })?;
        let restored = MovieGeometry::new(current.position(), original_size).map_err(|error| {
            Error::InvalidFormat(format!("invalid focused movie geometry: {error}"))
        })?;
        let edit = package
            .edit_slide_movie_geometry(slide_position, movie_selector)
            .map_err(map_focused_movie_geometry_error)?
            .set(restored)
            .map_err(map_focused_movie_geometry_error)?;
        let commit = edit.commit().map_err(map_focused_movie_geometry_error)?;
        if !commit.patch().is_noop() {
            replace_from_focused_movie_geometry_commit(self, commit)?;
        }
        Ok(restored)
    }

    /// Apply a native Arrange flip through a typed movie selector.
    ///
    /// The focused package owns the transform read and rewrite. The returned
    /// legacy geometry is only a compatibility projection for callers that
    /// still consume the physical movie listing; it is never used to choose a
    /// fallback mutation path.
    pub fn flip_slide_movie_by_selector(
        &mut self,
        slide_position: Position,
        movie_selector: MovieSelector,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let focused_axis = match axis {
            DrawableFlipAxis::Horizontal => FocusedMovieFlipAxis::Horizontal,
            DrawableFlipAxis::Vertical => FocusedMovieFlipAxis::Vertical,
        };
        let package = focused_movie_geometry_package(self)?;
        let edit = package
            .edit_slide_movie_geometry(slide_position, movie_selector)
            .map_err(map_focused_movie_geometry_error)?;
        let commit = edit
            .flip(focused_axis)
            .map_err(map_focused_movie_geometry_error)?
            .commit()
            .map_err(map_focused_movie_geometry_error)?;
        if !commit.patch().is_noop() {
            replace_from_focused_movie_geometry_commit(self, commit)?;
        }
        self.slide_movies(slide_position.get())?
            .into_iter()
            .filter(|movie| movie.kind == MovieKind::File)
            .nth(movie_selector.as_index())
            .map(|movie| movie.geometry)
            .ok_or_else(|| {
                Error::InvalidFormat("focused Keynote movie selector is unavailable".to_owned())
            })
    }
}

fn focused_movie_geometry_package(editor: &KeynoteEditor) -> Result<KeynotePackage> {
    let bytes = editor.to_bytes()?;
    KeynotePackage::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote movie geometry source failed: {error}"
        ))
    })
}

fn replace_from_focused_movie_geometry_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::SlideMovieGeometryCommit,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote movie geometry write failed: {error}"
        ))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_movie_geometry_error(error: litchi_keynote::SlideMovieGeometryError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote movie geometry operation failed: {error}"
    ))
}

pub(in crate::keynote::editor) fn set_movie_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let message_index = *message_index;
        let data = transform_length_delimited_field(
            object.messages[message_index].data.as_slice(),
            MOVIE_DRAWABLE_FIELD,
            |drawable| patch_drawable_geometry(drawable, geometry),
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(in crate::keynote::editor) fn set_movie_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    properties: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::MovieArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::MovieArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Keynote movie properties patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            *message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn offset_movie(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    offset: f32,
) -> Result<()> {
    let geometry = {
        let archive = package.archive(archive_name)?;
        let object = archive.object(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote movie {movie_id} payload is missing"))
            })?;
        let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
        geometry_from_drawable(&movie.super_)?
    };
    let position = geometry
        .position
        .ok_or_else(|| Error::InvalidFormat(format!("Keynote movie {movie_id} has no position")))?;
    set_movie_geometry(
        package,
        archive_name,
        movie_id,
        DrawableGeometry {
            position: Some(crate::shapes::DrawablePoint {
                x: position.x + offset,
                y: position.y + offset,
            }),
            ..geometry
        },
    )
}
