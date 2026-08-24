//! Selector-first native title and caption bridging for Keynote slide movies.

use litchi_core::Position;
use litchi_keynote::MovieSelector;

use super::*;

impl KeynoteEditor {
    /// Read the native title attached to one movie selected by source position.
    pub fn slide_movie_title_by_selector(
        &self,
        slide_position: Position,
        selector: MovieSelector,
    ) -> Result<Option<String>> {
        focused_movie_title_package(self)?
            .slide_movie_title(slide_position, selector)
            .map_err(map_focused_movie_title_error)
    }

    /// Create or replace one ordinary slide movie's native title by selector.
    pub fn set_slide_movie_title_by_selector(
        &mut self,
        slide_position: Position,
        selector: MovieSelector,
        title: &str,
    ) -> Result<()> {
        let package = focused_movie_title_package(self)?;
        let edit = package
            .edit_slide_movie_title(slide_position, selector)
            .map_err(map_focused_movie_title_error)?;
        let edit = edit.set(title).map_err(map_focused_movie_title_error)?;
        let commit = edit.commit().map_err(map_focused_movie_title_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_movie_title_commit(self, commit)
    }

    /// Remove one ordinary slide movie's native title by selector.
    ///
    /// Returns whether a title was present. Native iWork removal preserves the
    /// prior title graph for undo history and attaches a fresh empty stand-in.
    pub fn remove_slide_movie_title_by_selector(
        &mut self,
        slide_position: Position,
        selector: MovieSelector,
    ) -> Result<bool> {
        let package = focused_movie_title_package(self)?;
        let edit = package
            .edit_slide_movie_title(slide_position, selector)
            .map_err(map_focused_movie_title_error)?;
        let had_title = edit.before().is_some();
        let edit = edit.clear().map_err(map_focused_movie_title_error)?;
        let commit = edit.commit().map_err(map_focused_movie_title_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_movie_title_commit(self, commit)?;
        Ok(had_title)
    }

    /// Read one ordinary slide movie's native caption by semantic selector.
    pub fn slide_movie_caption_by_selector(
        &self,
        slide_position: Position,
        selector: MovieSelector,
    ) -> Result<Option<String>> {
        focused_movie_caption_package(self)?
            .slide_movie_caption(slide_position, selector)
            .map_err(map_focused_movie_caption_error)
    }

    /// Create or replace one ordinary slide movie's native caption by selector.
    pub fn set_slide_movie_caption_by_selector(
        &mut self,
        slide_position: Position,
        selector: MovieSelector,
        caption: &str,
    ) -> Result<()> {
        let package = focused_movie_caption_package(self)?;
        let edit = package
            .edit_slide_movie_caption(slide_position, selector)
            .map_err(map_focused_movie_caption_error)?;
        let edit = edit.set(caption).map_err(map_focused_movie_caption_error)?;
        let commit = edit.commit().map_err(map_focused_movie_caption_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_movie_caption_commit(self, commit)
    }

    /// Remove one ordinary slide movie's native caption by selector.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_slide_movie_caption_by_selector(
        &mut self,
        slide_position: Position,
        selector: MovieSelector,
    ) -> Result<bool> {
        let package = focused_movie_caption_package(self)?;
        let edit = package
            .edit_slide_movie_caption(slide_position, selector)
            .map_err(map_focused_movie_caption_error)?;
        let had_caption = edit.before().is_some();
        let edit = edit.clear().map_err(map_focused_movie_caption_error)?;
        let commit = edit.commit().map_err(map_focused_movie_caption_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_movie_caption_commit(self, commit)?;
        Ok(had_caption)
    }
}

fn focused_movie_caption_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote movie caption source failed: {error}"
        ))
    })
}

fn replace_from_focused_movie_caption_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::SlideMovieCaptionCommit,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote movie caption write failed: {error}"
        ))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_movie_caption_error(error: litchi_keynote::SlideMovieCaptionError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote movie caption operation failed: {error}"
    ))
}

fn focused_movie_title_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote movie title source failed: {error}"
        ))
    })
}

fn replace_from_focused_movie_title_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::SlideMovieTitleCommit,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote movie title write failed: {error}"))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_movie_title_error(error: litchi_keynote::SlideMovieTitleError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote movie title operation failed: {error}"
    ))
}
