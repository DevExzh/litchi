//! Focused chart Arrange projection for Pages body charts.
//!
//! The focused `litchi_pages::Package` owns chart Arrange reads and writes.
//! These helpers populate the host listing and validate lifecycle reads;
//! physical reads and writes stay focused.

use super::*;
use litchi_pages::{BodyChartSelector, Package as FocusedPagesPackage};

/// Populate the public Pages chart listing with one checked focused traversal.
pub(super) fn fill_focused_chart_arrangements(
    editor: &PagesEditor,
    charts: &mut [PagesBodyChartInfo],
    chart_positions: &[usize],
) -> Result<()> {
    if charts.is_empty() && chart_positions.is_empty() {
        return Ok(());
    }
    let arrangements = focused_chart_arrangement_package(editor)?
        .body_chart_arrangements()
        .map_err(map_focused_chart_arrangement_error)?;
    let order_matches = chart_positions.len() == charts.len()
        && chart_positions
            .iter()
            .copied()
            .enumerate()
            .all(|(position, chart_position)| position == chart_position);
    if arrangements.len() != charts.len() || !order_matches {
        return Err(Error::InvalidFormat(format!(
            "focused Pages chart arrangement listing does not match host order (focused {}, host {}, positions {})",
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
    editor: &PagesEditor,
    chart_count: usize,
    chart_position: usize,
) -> Result<ChartArrangement> {
    if chart_position >= chart_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart arrangement position {chart_position} is outside chart count {chart_count}"
        )));
    }
    focused_chart_arrangement_package(editor)?
        .body_chart_arrangement(BodyChartSelector::index(chart_position))
        .map_err(map_focused_chart_arrangement_error)
}

fn focused_chart_arrangement_package(editor: &PagesEditor) -> Result<FocusedPagesPackage> {
    let bytes = editor.to_bytes()?;
    FocusedPagesPackage::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Pages chart arrangement source failed: {error}"
        ))
    })
}

fn map_focused_chart_arrangement_error(error: litchi_pages::BodyChartArrangementError) -> Error {
    Error::InvalidFormat(format!(
        "focused Pages chart arrangement operation failed: {error}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_focused_chart_arrangement_selector_crud() {
        let body = "Chart arrangement";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
            editor.body_charts().unwrap()[0].arrangement,
            ChartArrangement::default()
        );
        let focused = FocusedPagesPackage::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused.body_chart_arrangement(0usize).unwrap(),
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        set_focused_arrangement(&mut editor, 0, constrained);
        assert_eq!(editor.body_charts().unwrap()[0].arrangement, constrained);
        let focused = FocusedPagesPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(focused.body_chart_arrangement(0usize).unwrap(), constrained);

        let duplicate = editor
            .duplicate_body_chart(
                chart.drawable_object_id,
                editor.body_text().unwrap().encode_utf16().count(),
            )
            .unwrap();
        assert_eq!(
            editor
                .body_charts()
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
                .body_charts()
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == chart.drawable_object_id)
                .unwrap()
                .arrangement,
            constrained
        );
        assert_eq!(
            editor
                .body_charts()
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .arrangement,
            locked
        );

        editor
            .remove_body_chart(duplicate.drawable_object_id)
            .unwrap();
        set_focused_arrangement(&mut editor, 0, ChartArrangement::default());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn focused_chart_arrangement_edit_reopens_and_inverts_exactly() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Chart arrangement")
            .build()
            .unwrap();
        editor
            .add_body_chart(
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
        let focused = FocusedPagesPackage::from_bytes(&baseline).unwrap();
        let replacement = ChartArrangement::default()
            .with_locked(true)
            .with_constrain_proportions(true);
        let changed = focused
            .edit_body_chart_arrangement(BodyChartSelector::index(0))
            .unwrap()
            .set(replacement)
            .commit()
            .unwrap();
        let mut changed_bytes = Vec::new();
        changed.package().write_to(&mut changed_bytes).unwrap();
        let reopened = PagesEditor::from_bytes(&changed_bytes).unwrap();
        assert_eq!(reopened.body_charts().unwrap()[0].arrangement, replacement);

        let restored = changed
            .package()
            .apply_body_chart_arrangement(&changed.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, baseline);
    }

    #[test]
    fn focused_chart_arrangement_edit_preserves_a_large_body_chart_listing_and_inverse() {
        const CHART_COUNT: usize = 32;
        const TARGET_POSITION: usize = 17;

        let mut editor = PagesDocumentBuilder::new()
            .body_text("Chart arrangement")
            .build()
            .unwrap();
        let source = editor
            .add_body_chart(
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
            let anchor = editor.body_text().unwrap().encode_utf16().count();
            editor
                .duplicate_body_chart(source.drawable_object_id, anchor)
                .unwrap();
        }

        let baseline = editor.to_bytes().unwrap();
        let baseline_listing = editor.body_charts().unwrap();
        assert_eq!(baseline_listing.len(), CHART_COUNT);
        assert!(
            baseline_listing
                .iter()
                .all(|chart| chart.arrangement == ChartArrangement::default())
        );

        let focused = FocusedPagesPackage::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused
                .body_chart_arrangement(BodyChartSelector::index(TARGET_POSITION))
                .unwrap(),
            ChartArrangement::default()
        );
        let replacement = ChartArrangement::default()
            .with_locked(true)
            .with_constrain_proportions(true);
        let changed = focused
            .edit_body_chart_arrangement(BodyChartSelector::index(TARGET_POSITION))
            .unwrap()
            .set(replacement)
            .commit()
            .unwrap();
        let mut changed_bytes = Vec::new();
        changed.package().write_to(&mut changed_bytes).unwrap();
        let reopened = PagesEditor::from_bytes(&changed_bytes).unwrap();
        let mut expected_listing = baseline_listing;
        expected_listing[TARGET_POSITION].arrangement = replacement;
        assert_eq!(reopened.body_charts().unwrap(), expected_listing);

        let restored = changed
            .package()
            .apply_body_chart_arrangement(&changed.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, baseline);
    }

    fn set_focused_arrangement(
        editor: &mut PagesEditor,
        chart_position: usize,
        arrangement: ChartArrangement,
    ) {
        let focused = FocusedPagesPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = focused
            .edit_body_chart_arrangement(BodyChartSelector::index(chart_position))
            .unwrap()
            .set(arrangement)
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        *editor = PagesEditor::from_bytes(&bytes).unwrap();
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
