//! Selector-first chart Arrange compatibility for Keynote slide charts.
//!
//! The focused `litchi_keynote::Package` owns chart Arrange reads and writes.
//! The selector-based `KeynoteEditor` methods below are compatibility bridges
//! for the legacy editor while all physical reads and writes stay focused.

use super::*;
use litchi_keynote::{ChartSelector, Package as FocusedKeynotePackage, SlideSelector};

impl KeynoteEditor {
    /// Read one slide chart's Arrange-panel state by semantic chart selector.
    pub fn slide_chart_arrangement_by_selector<'selector>(
        &self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
    ) -> Result<ChartArrangement> {
        let selector = selector.into();
        focused_chart_arrangement_package(self)?
            .slide_chart_arrangement(SlideSelector::index(slide_index), selector)
            .map_err(map_focused_chart_arrangement_error)
    }

    /// Set one slide chart's Arrange-panel state by semantic chart selector.
    pub fn set_slide_chart_arrangement_by_selector<'selector>(
        &mut self,
        slide_index: usize,
        selector: impl Into<ChartSelector<'selector>>,
        arrangement: ChartArrangement,
    ) -> Result<()> {
        let selector = selector.into();
        let package = focused_chart_arrangement_package(self)?;
        let commit = package
            .edit_slide_chart_arrangement(SlideSelector::index(slide_index), selector)
            .map_err(map_focused_chart_arrangement_error)?
            .set(arrangement)
            .commit()
            .map_err(map_focused_chart_arrangement_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_chart_arrangement_commit(self, commit)
    }
}

/// Populate the public Keynote chart listing with the focused semantic value.
///
/// The focused owner performs one checked chart traversal for the complete
/// slide. This keeps listing proportional to the slide graph instead of
/// selecting and decoding the focused package once per chart.
pub(super) fn fill_focused_chart_arrangements(
    editor: &KeynoteEditor,
    slide_index: usize,
    charts: &mut [KeynoteSlideChartInfo],
    chart_positions: &[usize],
) -> Result<()> {
    if charts.is_empty() && chart_positions.is_empty() {
        return Ok(());
    }
    let arrangements = focused_chart_arrangements(editor, slide_index)?;
    let order_matches = chart_positions.len() == charts.len()
        && chart_positions
            .iter()
            .copied()
            .enumerate()
            .all(|(position, chart_position)| position == chart_position);
    if arrangements.len() != charts.len() || !order_matches {
        return Err(Error::InvalidFormat(format!(
            "focused Keynote chart arrangement listing does not match host order (focused {}, host {}, positions {})",
            arrangements.len(),
            charts.len(),
            chart_positions.len()
        )));
    }
    for (chart, arrangement) in charts.iter_mut().zip(arrangements) {
        chart.arrangement = arrangement;
    }
    Ok(())
}

pub(super) fn focused_chart_arrangement_at(
    editor: &KeynoteEditor,
    slide_index: usize,
    chart_count: usize,
    chart_position: usize,
) -> Result<ChartArrangement> {
    if chart_position >= chart_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart arrangement position {chart_position} is outside chart count {chart_count}"
        )));
    }
    focused_chart_arrangement_package(editor)?
        .slide_chart_arrangement(
            SlideSelector::index(slide_index),
            ChartSelector::index(chart_position),
        )
        .map_err(map_focused_chart_arrangement_error)
}

fn focused_chart_arrangements(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<Box<[ChartArrangement]>> {
    focused_chart_arrangement_package(editor)?
        .slide_chart_arrangements(SlideSelector::index(slide_index))
        .map_err(map_focused_chart_arrangement_error)
}

fn focused_chart_arrangement_package(editor: &KeynoteEditor) -> Result<FocusedKeynotePackage> {
    let bytes = editor.to_bytes()?;
    FocusedKeynotePackage::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote chart arrangement source failed: {error}"
        ))
    })
}

fn replace_from_focused_chart_arrangement_commit(
    editor: &mut KeynoteEditor,
    commit: litchi_keynote::ChartArrangementCommit,
) -> Result<()> {
    let expected = commit.patch().after();
    let slide_position = commit.patch().slide_position();
    let chart_position = commit.patch().chart_position();
    let mut writer = FallibleBytes::default();
    commit.package().write_to(&mut writer).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote chart arrangement write failed: {error}"
        ))
    })?;
    let bytes = writer.into_bytes();
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let actual = focused_chart_arrangement_package(&reopened)?
        .slide_chart_arrangement(
            SlideSelector::position(slide_position),
            ChartSelector::position(chart_position),
        )
        .map_err(map_focused_chart_arrangement_error)?;
    if actual != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart arrangement update failed semantic verification".to_owned(),
        ));
    }
    *editor = reopened;
    Ok(())
}

#[derive(Debug, Default)]
struct FallibleBytes {
    bytes: Vec<u8>,
}

impl FallibleBytes {
    fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

impl std::io::Write for FallibleBytes {
    fn write(&mut self, source: &[u8]) -> std::io::Result<usize> {
        self.bytes
            .try_reserve_exact(source.len())
            .map_err(|_error| {
                std::io::Error::other("Keynote chart arrangement handoff allocation failed")
            })?;
        self.bytes.extend_from_slice(source);
        Ok(source.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn map_focused_chart_arrangement_error(error: litchi_keynote::ChartArrangementError) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote chart arrangement operation failed: {error}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_chart_arrangement_selector_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                Kind::Line2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement_by_selector(0, ChartSelector::index(0))
                .unwrap(),
            ChartArrangement::default()
        );
        assert_eq!(
            editor.slide_charts(0).unwrap()[0].arrangement,
            ChartArrangement::default()
        );
        let focused = litchi_keynote::Package::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused.slide_chart_arrangement(0usize, 0usize).unwrap(),
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        editor
            .set_slide_chart_arrangement_by_selector(0, ChartSelector::index(0), constrained)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement_by_selector(0, ChartSelector::index(0))
                .unwrap(),
            constrained
        );
        assert_eq!(editor.slide_charts(0).unwrap()[0].arrangement, constrained);
        let focused = litchi_keynote::Package::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            focused.slide_chart_arrangement(0usize, 0usize).unwrap(),
            constrained
        );

        let duplicate = editor
            .duplicate_slide_chart(0, chart_selector(&editor, &chart))
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement_by_selector(0, chart_selector(&editor, &duplicate))
                .unwrap(),
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        editor
            .set_slide_chart_arrangement_by_selector(0, chart_selector(&editor, &duplicate), locked)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement_by_selector(0, chart_selector(&editor, &chart))
                .unwrap(),
            constrained
        );
        assert_eq!(
            editor
                .slide_chart_arrangement_by_selector(0, chart_selector(&editor, &duplicate))
                .unwrap(),
            locked
        );

        editor
            .remove_slide_chart(0, chart_selector(&editor, &duplicate))
            .unwrap();
        editor
            .set_slide_chart_arrangement_by_selector(
                0,
                chart_selector(&editor, &chart),
                ChartArrangement::default(),
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned(), "C".to_owned()],
            vec![vec![Some(8.0), Some(20.0), Some(42.0)]],
        )
        .unwrap()
    }
}
