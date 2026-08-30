//! Selector-first native axis-title CRUD for Keynote slide charts.

use super::*;
use litchi_keynote::{Axis, SlideSelector};

impl KeynoteEditor {
    /// Read the title shown by Keynote for one selected slide-chart axis.
    pub fn slide_chart_axis_title_by_selector<'selector>(
        &self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        axis: Axis,
    ) -> Result<Option<String>> {
        let selector = selector.into();
        focused_chart_axis_title_package(self)?
            .slide_chart_axis_title(SlideSelector::index(slide_index), selector, axis)
            .map_err(map_focused_chart_axis_title_error)
    }

    /// Create or replace the title shown by Keynote for one selected
    /// slide-chart axis through an atomic focused-package transaction.
    pub fn set_slide_chart_axis_title_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        axis: Axis,
        title: impl AsRef<str>,
    ) -> Result<()> {
        let selector = selector.into();
        let package = focused_chart_axis_title_package(self)?;
        let title = title.as_ref();
        let commit = package
            .edit_slide_chart_axis_title(SlideSelector::index(slide_index), selector, axis)
            .map_err(map_focused_chart_axis_title_error)?
            .set(title)
            .map_err(map_focused_chart_axis_title_error)?
            .commit()
            .map_err(map_focused_chart_axis_title_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_chart_axis_title_commit(
            self,
            commit,
            slide_index,
            selector,
            axis,
            Some(title),
        )
    }

    /// Remove the title shown by Keynote for one selected slide-chart axis
    /// through an atomic focused-package transaction.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_slide_chart_axis_title_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        axis: Axis,
    ) -> Result<bool> {
        let selector = selector.into();
        let package = focused_chart_axis_title_package(self)?;
        let edit = package
            .edit_slide_chart_axis_title(SlideSelector::index(slide_index), selector, axis)
            .map_err(map_focused_chart_axis_title_error)?;
        let had_visible_title = edit.before().is_some();
        let commit = edit
            .clear()
            .map_err(map_focused_chart_axis_title_error)?
            .commit()
            .map_err(map_focused_chart_axis_title_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_chart_axis_title_commit(
            self,
            commit,
            slide_index,
            selector,
            axis,
            None,
        )?;
        Ok(had_visible_title)
    }
}

fn focused_chart_axis_title_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(map_focused_chart_axis_title_read_error)
}

fn replace_from_focused_chart_axis_title_commit<'selector>(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::ChartAxisTitleCommit,
    slide_index: usize,
    selector: ChartSelector<'selector>,
    axis: Axis,
    expected: Option<&str>,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote chart axis title write failed: {error}"
        ))
    })?;
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let actual = focused_chart_axis_title_package(&reopened)?
        .slide_chart_axis_title(SlideSelector::index(slide_index), selector, axis)
        .map_err(map_focused_chart_axis_title_error)?;
    if actual.as_deref() != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart axis title update failed semantic verification".to_owned(),
        ));
    }
    *editor = reopened;
    Ok(())
}

fn map_focused_chart_axis_title_error(error: litchi_keynote::ChartAxisTitleError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart axis title operation failed: {error}"
    ))
}

fn map_focused_chart_axis_title_read_error(error: litchi_keynote::ReadError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart axis title source failed: {error}"
    ))
}
