//! Focused chart Arrange projection for Keynote slide charts.
//!
//! The focused `litchi_keynote::Package` owns chart Arrange reads and writes.
//! The helpers below populate the host listing and validate lifecycle reads;
//! physical reads and writes stay focused.

use super::*;
use litchi_keynote::{ChartSelector, Package as FocusedKeynotePackage, SlideSelector};

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
            editor.slide_charts(0).unwrap()[0].arrangement,
            ChartArrangement::default()
        );
        let focused = litchi_keynote::Package::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused.slide_chart_arrangement(0usize, 0usize).unwrap(),
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        set_focused_arrangement(&mut editor, 0, constrained);
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
                .slide_charts(0)
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .arrangement,
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        set_focused_arrangement(&mut editor, 1, locked);
        assert_eq!(
            editor
                .slide_charts(0)
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == chart.drawable_object_id)
                .unwrap()
                .arrangement,
            constrained
        );
        assert_eq!(
            editor
                .slide_charts(0)
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .arrangement,
            locked
        );

        editor
            .remove_slide_chart(0, chart_selector(&editor, &duplicate))
            .unwrap();
        set_focused_arrangement(&mut editor, 0, ChartArrangement::default());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn focused_arrangement_edit_preserves_a_large_chart_listing_and_inverse() {
        const CHART_COUNT: usize = 32;
        const TARGET_POSITION: usize = 17;

        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        editor
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
        for _ in 1..CHART_COUNT {
            editor
                .duplicate_slide_chart(0, ChartSelector::index(0))
                .unwrap();
        }

        let baseline = editor.to_bytes().unwrap();
        let baseline_listing = editor.slide_charts(0).unwrap();
        assert_eq!(baseline_listing.len(), CHART_COUNT);
        assert!(
            baseline_listing
                .iter()
                .all(|chart| chart.arrangement == ChartArrangement::default())
        );

        let focused = litchi_keynote::Package::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused
                .slide_chart_arrangement(0usize, TARGET_POSITION)
                .unwrap(),
            ChartArrangement::default()
        );
        let replacement = ChartArrangement::default()
            .with_locked(true)
            .with_constrain_proportions(true);
        let changed = focused
            .edit_slide_chart_arrangement(0usize, TARGET_POSITION)
            .unwrap()
            .set(replacement)
            .commit()
            .unwrap();
        let mut candidate_bytes = Vec::new();
        changed.package().write_to(&mut candidate_bytes).unwrap();
        let reopened = KeynoteEditor::from_bytes(&candidate_bytes).unwrap();
        let mut expected_listing = baseline_listing.clone();
        expected_listing[TARGET_POSITION].arrangement = replacement;
        assert_eq!(reopened.slide_charts(0).unwrap(), expected_listing);

        let restored = changed
            .package()
            .apply_slide_chart_arrangement(&changed.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, baseline);
    }

    fn set_focused_arrangement(
        editor: &mut KeynoteEditor,
        chart_position: usize,
        arrangement: ChartArrangement,
    ) {
        let focused = litchi_keynote::Package::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = focused
            .edit_slide_chart_arrangement(
                SlideSelector::index(0),
                ChartSelector::index(chart_position),
            )
            .unwrap()
            .set(arrangement)
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        *editor = KeynoteEditor::from_bytes(&bytes).unwrap();
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
