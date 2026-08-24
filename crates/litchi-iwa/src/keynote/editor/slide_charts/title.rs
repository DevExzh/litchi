//! Selector-first native title CRUD for Keynote slide charts.

use super::*;

impl KeynoteEditor {
    /// Read the chart title for a chart selected on one slide.
    ///
    /// The host graph is checked first so malformed ownership or title
    /// stand-ins fail before the focused selector transaction is entered.
    pub fn slide_chart_title_by_selector<'selector>(
        &self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<Option<String>> {
        let selector = selector.into();
        self.resolve_chart_selector(slide_index, selector)?;
        focused_chart_title_package(self)?
            .slide_chart_title(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_title_error)
    }

    /// Set the chart title for a chart selected on one slide through the
    /// focused selector-first transaction.
    pub fn set_slide_chart_title_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        title: impl AsRef<str>,
    ) -> Result<()> {
        let selector = selector.into();
        self.resolve_chart_selector(slide_index, selector)?;
        let package = focused_chart_title_package(self)?;
        let edit = package
            .edit_slide_chart_title(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_title_error)?
            .set(title)
            .map_err(map_focused_chart_title_error)?;
        let commit = edit.commit().map_err(map_focused_chart_title_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_chart_title_commit(self, commit)
    }

    /// Remove the chart title for a chart selected on one slide through the
    /// focused selector-first transaction.
    pub fn remove_slide_chart_title_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<bool> {
        let selector = selector.into();
        self.resolve_chart_selector(slide_index, selector)?;
        let package = focused_chart_title_package(self)?;
        let edit = package
            .edit_slide_chart_title(litchi_core::Position::new(slide_index), selector)
            .map_err(map_focused_chart_title_error)?;
        let had_visible_title = edit.before().is_some();
        let commit = edit
            .clear()
            .map_err(map_focused_chart_title_error)?
            .commit()
            .map_err(map_focused_chart_title_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_chart_title_commit(self, commit)?;
        Ok(had_visible_title)
    }
}

fn focused_chart_title_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(map_focused_chart_title_read_error)
}

pub(super) fn focused_chart_catalog(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<ChartCatalog> {
    focused_chart_title_package(editor)?
        .slide_chart_catalog(litchi_core::Position::new(slide_index))
        .map_err(map_focused_chart_title_error)
}

fn replace_from_focused_chart_title_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::ChartTitleCommit,
) -> Result<()> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote chart title write failed: {error}"))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_chart_title_error(error: litchi_keynote::ChartTitleError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart title operation failed: {error}"
    ))
}

fn map_focused_chart_title_read_error(error: litchi_keynote::ReadError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart title source failed: {error}"
    ))
}
