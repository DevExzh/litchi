//! Selector-first native caption CRUD for Keynote slide charts.

use super::*;

impl KeynoteEditor {
    /// Read one chart caption through semantic slide/chart selection.
    pub fn slide_chart_caption_by_selector<'selector>(
        &self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<Option<String>> {
        let selector = selector.into();
        focused_chart_caption_package(self)?
            .slide_chart_caption(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_caption_error)
    }

    /// Create or replace one chart caption through semantic selection.
    pub fn set_slide_chart_caption_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        caption: impl AsRef<str>,
    ) -> Result<()> {
        let selector = selector.into();
        let package = focused_chart_caption_package(self)?;
        let commit = package
            .edit_slide_chart_caption(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_caption_error)?
            .set(caption)
            .map_err(map_focused_chart_caption_error)?
            .commit()
            .map_err(map_focused_chart_caption_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_chart_caption_commit(self, commit)
    }

    /// Remove one chart caption through semantic selection.
    ///
    /// Returns whether a caption was present. Clearing an absent caption is
    /// an exact package-level no-op.
    pub fn remove_slide_chart_caption_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<bool> {
        let selector = selector.into();
        let package = focused_chart_caption_package(self)?;
        let edit = package
            .edit_slide_chart_caption(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_caption_error)?;
        let had_caption = edit.before().is_some();
        let commit = edit
            .clear()
            .map_err(map_focused_chart_caption_error)?
            .commit()
            .map_err(map_focused_chart_caption_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_chart_caption_commit(self, commit)?;
        Ok(had_caption)
    }
}

fn focused_chart_caption_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote chart caption source failed: {error}"
        ))
    })
}

fn replace_from_focused_chart_caption_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::ChartCaptionCommit,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote chart caption write failed: {error}"
        ))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_chart_caption_error(error: litchi_keynote::ChartCaptionError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart caption operation failed: {error}"
    ))
}
