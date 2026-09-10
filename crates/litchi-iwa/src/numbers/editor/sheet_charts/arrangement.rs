//! Focused chart Arrange projection for Numbers sheet charts.
//!
//! The focused `litchi_numbers::Package` owns chart Arrange reads and writes.
//! These private helpers keep the legacy Numbers editor listing and graph
//! lifecycle APIs useful without exposing native drawable identifiers through
//! the focused package boundary.

use super::*;
use litchi_numbers::{ChartSelector, Package as FocusedNumbersPackage, SheetSelector};

/// Populate the legacy host chart listing from one focused batch read.
///
/// Chart selectors are semantic positions within the sheet. The host graph
/// walk supplies the same source order, so the focused batch is read once and
/// then copied into the already validated graph records.
pub(super) fn fill_focused_chart_arrangements(
    editor: &NumbersEditor,
    sheet_id: u64,
    charts: &mut [NumbersSheetChartInfo],
    chart_positions: &[usize],
) -> Result<()> {
    if charts.is_empty() && chart_positions.is_empty() {
        return Ok(());
    }
    let arrangements = focused_chart_arrangements(editor, sheet_id)?;
    let order_matches = chart_positions.len() == charts.len()
        && chart_positions
            .iter()
            .copied()
            .enumerate()
            .all(|(position, chart_position)| position == chart_position);
    if arrangements.len() != charts.len() || !order_matches {
        return Err(Error::InvalidFormat(format!(
            "focused Numbers chart arrangement listing does not match host order (focused {}, host {}, positions {})",
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

/// Read one chart Arrange value through its checked semantic sheet/chart
/// position. This is used by lifecycle methods that return one graph record;
/// ordinary listing uses the batch helper above.
pub(super) fn focused_chart_arrangement_for_drawable(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartArrangement> {
    let sheet_index = sheet_index(editor, sheet_id)?;
    let (chart_position, chart_count) = chart_position(editor, sheet_id, drawable_object_id)?;
    focused_chart_arrangement_at(editor, sheet_index, chart_count, chart_position)
}

fn focused_chart_arrangement_at(
    editor: &NumbersEditor,
    sheet_index: usize,
    chart_count: usize,
    chart_position: usize,
) -> Result<ChartArrangement> {
    if chart_position >= chart_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart arrangement position {chart_position} is outside chart count {chart_count}"
        )));
    }
    focused_chart_arrangement_package(editor)?
        .sheet_chart_arrangement(
            SheetSelector::index(sheet_index),
            ChartSelector::index(chart_position),
        )
        .map_err(map_focused_chart_arrangement_error)
}

fn focused_chart_arrangements(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<Box<[ChartArrangement]>> {
    let sheet_index = sheet_index(editor, sheet_id)?;
    focused_chart_arrangement_package(editor)?
        .sheet_chart_arrangements(SheetSelector::index(sheet_index))
        .map_err(map_focused_chart_arrangement_error)
}

fn focused_chart_arrangement_package(editor: &NumbersEditor) -> Result<FocusedNumbersPackage> {
    let bytes = editor.to_bytes()?;
    FocusedNumbersPackage::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers chart arrangement source failed: {error}"
        ))
    })
}

fn map_focused_chart_arrangement_error(error: litchi_numbers::SheetChartArrangementError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers chart arrangement operation failed: {error}"
    ))
}

fn sheet_index(editor: &NumbersEditor, sheet_id: u64) -> Result<usize> {
    editor
        .sheets()?
        .into_iter()
        .find(|sheet| sheet.object_id == sheet_id)
        .map(|sheet| sheet.index)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))
}

/// Resolve a host chart graph to the focused package's semantic sheet and
/// chart positions. Native identifiers remain private to this adapter.
pub(super) fn focused_chart_data_target(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<(usize, usize)> {
    let sheet_index = sheet_index(editor, sheet_id)?;
    let (chart_position, _) = chart_position(editor, sheet_id, drawable_object_id)?;
    Ok((sheet_index, chart_position))
}

/// Resolve a native drawable to its semantic chart position while preserving
/// the host's source-order ownership checks. This is intentionally private;
/// focused callers only receive `ChartSelector` positions.
fn chart_position(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<(usize, usize)> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut chart_position = 0usize;
    let mut target_position = None;
    for reference in sheet.drawable_infos {
        let Some(archive_name) = locations.get(&reference.identifier) else {
            return Err(Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            )));
        };
        let archive = editor.package().archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        if !object
            .messages
            .iter()
            .any(|message| message.type_ == CHART_MESSAGE_TYPE)
        {
            continue;
        }
        if reference.identifier == drawable_object_id {
            target_position = Some(chart_position);
        }
        chart_position = chart_position
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers sheet chart position overflow".to_owned()))?;
    }
    target_position
        .map(|position| (position, chart_position))
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} does not own chart {drawable_object_id}"
            ))
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_chart_arrangement_selector_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
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
            editor.sheet_charts(sheet_id).unwrap()[0].arrangement,
            ChartArrangement::default()
        );
        let focused = litchi_numbers::Package::from_bytes(&baseline).unwrap();
        assert_eq!(
            focused
                .sheet_chart_arrangement(SheetSelector::index(0), ChartSelector::index(0))
                .unwrap(),
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        let changed = focused
            .edit_sheet_chart_arrangement(SheetSelector::index(0), ChartSelector::index(0))
            .unwrap()
            .set(constrained)
            .commit()
            .unwrap();
        let restored = changed
            .package()
            .apply_sheet_chart_arrangement(&changed.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, baseline);

        set_focused_arrangement(&mut editor, 0, constrained);
        assert_eq!(
            editor.sheet_charts(sheet_id).unwrap()[0].arrangement,
            constrained
        );
        let focused = litchi_numbers::Package::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            focused
                .sheet_chart_arrangement(SheetSelector::index(0), ChartSelector::index(0))
                .unwrap(),
            constrained
        );
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_charts(sheet_id)
                .unwrap()
                .into_iter()
                .find(|candidate| candidate.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .arrangement,
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        set_focused_arrangement(&mut editor, 1, locked);
        let listing = editor.sheet_charts(sheet_id).unwrap();
        assert_eq!(
            listing
                .iter()
                .find(|candidate| candidate.drawable_object_id == chart.drawable_object_id)
                .unwrap()
                .arrangement,
            constrained
        );
        assert_eq!(
            listing
                .iter()
                .find(|candidate| candidate.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .arrangement,
            locked
        );

        editor
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        set_focused_arrangement(&mut editor, 0, ChartArrangement::default());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn focused_batch_preserves_a_large_numbers_chart_listing_and_inverse() {
        const CHART_COUNT: usize = 32;
        const TARGET_POSITION: usize = 17;

        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let first = editor
            .add_sheet_chart(
                sheet_id,
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
                .duplicate_sheet_chart(sheet_id, first.drawable_object_id)
                .unwrap();
        }

        let baseline = editor.to_bytes().unwrap();
        let baseline_listing = editor.sheet_charts(sheet_id).unwrap();
        assert_eq!(baseline_listing.len(), CHART_COUNT);
        assert!(
            baseline_listing
                .iter()
                .all(|chart| chart.arrangement == ChartArrangement::default())
        );

        let focused = litchi_numbers::Package::from_bytes(&baseline).unwrap();
        let focused_listing = focused
            .sheet_chart_arrangements(SheetSelector::index(0))
            .unwrap();
        assert_eq!(focused_listing.len(), CHART_COUNT);
        assert!(
            focused_listing
                .iter()
                .all(|arrangement| *arrangement == ChartArrangement::default())
        );

        let replacement = ChartArrangement::default()
            .with_locked(true)
            .with_constrain_proportions(true);
        let changed = focused
            .edit_sheet_chart_arrangement(
                SheetSelector::index(0),
                ChartSelector::index(TARGET_POSITION),
            )
            .unwrap()
            .set(replacement)
            .commit()
            .unwrap();
        let changed_listing = changed
            .package()
            .sheet_chart_arrangements(SheetSelector::index(0))
            .unwrap();
        assert_eq!(changed_listing.len(), CHART_COUNT);
        assert_eq!(changed_listing[TARGET_POSITION], replacement);
        assert!(
            changed_listing
                .iter()
                .enumerate()
                .all(|(position, arrangement)| {
                    position == TARGET_POSITION || *arrangement == ChartArrangement::default()
                })
        );

        let mut candidate_bytes = Vec::new();
        changed.package().write_to(&mut candidate_bytes).unwrap();
        let reopened = NumbersEditor::from_bytes(&candidate_bytes).unwrap();
        let mut expected_listing = baseline_listing.clone();
        expected_listing[TARGET_POSITION].arrangement = replacement;
        assert_eq!(reopened.sheet_charts(sheet_id).unwrap(), expected_listing);

        let restored = changed
            .package()
            .apply_sheet_chart_arrangement(&changed.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, baseline);
    }

    fn set_focused_arrangement(
        editor: &mut NumbersEditor,
        chart_position: usize,
        arrangement: ChartArrangement,
    ) {
        let focused = litchi_numbers::Package::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = focused
            .edit_sheet_chart_arrangement(
                SheetSelector::index(0),
                ChartSelector::index(chart_position),
            )
            .unwrap()
            .set(arrangement)
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        *editor = NumbersEditor::from_bytes(&bytes).unwrap();
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
