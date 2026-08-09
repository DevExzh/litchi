#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_xlsx::page_breaks::{Break, Collection, PageBreaks};
use litchi_xlsx::{
    Cell, DurablePatch, Error, Formula, History, HistoryLimits, LocalStyle, MergeChoice,
    MergeLimits, Package, Value, Workbook,
};

fn two_sheet_workbook() -> Workbook {
    let mut package = Package::create().unwrap();
    let mut breaks = package.edit_page_breaks("Sheet1").unwrap();
    breaks
        .set_horizontal(
            Collection::horizontal([Break::new(12, 0, 16_383).unwrap().with_manual(true)]).unwrap(),
        )
        .unwrap();
    breaks.commit().unwrap();
    let workbook = package.into_workbook().unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.add("Target").unwrap();
    edit.commit().unwrap().into_workbook()
}

#[test]
fn durable_json_is_deterministic_stale_checked_reversible_and_reopenable() {
    let base = two_sheet_workbook();
    let mut edit = base.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", "durable")
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let first = durable.to_deterministic_json().unwrap();
    let second = durable.to_deterministic_json().unwrap();
    assert_eq!(first, second);

    let parsed = DurablePatch::from_deterministic_json(&first).unwrap();
    let applied = parsed.apply(&base).unwrap();
    let reopened = Workbook::from_bytes(applied.to_plain_bytes().unwrap()).unwrap();
    assert!(matches!(
        reopened
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("A1")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Value(Value::Text(text)) if text.as_str() == "durable"
    ));

    let restored = parsed.inverse().apply(&applied).unwrap();
    assert_eq!(
        restored.to_plain_bytes().unwrap(),
        base.to_plain_bytes().unwrap()
    );

    let mut divergent_edit = base.edit().unwrap();
    divergent_edit
        .sheet("Target")
        .unwrap()
        .unwrap()
        .set("B2", "stale")
        .unwrap();
    let divergent_book = divergent_edit.commit().unwrap().into_workbook();
    assert!(matches!(
        parsed.apply(&divergent_book),
        Err(Error::PatchConflict { .. })
    ));

    let sealed_json = durable.seal().to_deterministic_json().unwrap();
    let sealed = litchi_xlsx::SealedPatch::from_deterministic_json(&sealed_json).unwrap();
    assert_eq!(
        sealed.apply(&base).unwrap().to_plain_bytes().unwrap(),
        applied.to_plain_bytes().unwrap()
    );
}

#[test]
fn commits_record_exact_weighted_history_and_refuse_small_budgets() {
    let base = two_sheet_workbook();
    let mut edit = base.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("C3", "history")
        .unwrap();
    let commit = edit.commit().unwrap();

    let mut history = History::new(base.clone(), HistoryLimits::new(2, 256 * 1024 * 1024));
    assert!(commit.record(&mut history).unwrap().is_empty());
    assert!(history.can_undo());
    assert!(history.retained_weight() > 0);
    assert!(history.undo());
    assert_eq!(
        history.current().to_plain_bytes().unwrap(),
        base.to_plain_bytes().unwrap()
    );
    assert!(history.redo());

    let mut too_small = History::new(base.clone(), HistoryLimits::new(2, 1));
    assert!(matches!(
        commit.record(&mut too_small),
        Err(Error::DurablePatch(
            litchi_core::PatchError::HistoryWeight { .. }
        ))
    ));
    assert_eq!(
        too_small.current().to_plain_bytes().unwrap(),
        base.to_plain_bytes().unwrap()
    );
}

#[test]
fn three_way_planning_keeps_disjoint_work_and_requires_conflict_resolution() {
    let base = two_sheet_workbook();
    let limits = MergeLimits::new(2, 64, 128, 64);
    let mut disjoint_left = base.edit().unwrap();
    disjoint_left
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", "left")
        .unwrap();
    let mut disjoint_right = base.edit().unwrap();
    disjoint_right
        .sheet("Target")
        .unwrap()
        .unwrap()
        .set("B1", "right")
        .unwrap();
    let plan = disjoint_left
        .plan_three_way(disjoint_right, limits)
        .unwrap();
    assert_eq!(plan.automatic_len(), 2);
    assert!(plan.conflicts().is_empty());
    let merged = plan.finish().unwrap().commit().unwrap().into_workbook();
    assert!(
        merged
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("A1")
            .unwrap()
            .stored()
            .is_some()
    );
    assert!(
        merged
            .sheet("Target")
            .unwrap()
            .unwrap()
            .cell("B1")
            .unwrap()
            .stored()
            .is_some()
    );

    let mut conflict_left = base.edit().unwrap();
    conflict_left
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("D4", "left")
        .unwrap();
    let mut conflict_right = base.edit().unwrap();
    conflict_right
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("D4", "right")
        .unwrap();
    let mut conflict_plan = conflict_left
        .plan_three_way(conflict_right, limits)
        .unwrap();
    assert!(!conflict_plan.conflicts().is_empty());
    conflict_plan.resolve(MergeChoice::Left);
    let chosen = conflict_plan
        .finish()
        .unwrap()
        .commit()
        .unwrap()
        .into_workbook();
    assert!(matches!(
        chosen
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("D4")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Value(Value::Text(text)) if text.as_str() == "left"
    ));
}

#[test]
fn page_break_copy_and_move_are_atomic_dependency_checked_transfers() {
    let base = two_sheet_workbook();
    let source = base
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .page_breaks()
        .unwrap();
    let mut copy = base.edit().unwrap();
    assert!(copy.copy_page_breaks("Sheet1", "Target").unwrap().is_some());
    let copied = copy.commit().unwrap().into_workbook();
    assert_eq!(
        copied
            .sheet("Target")
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap(),
        source
    );

    let mut move_edit = copied.edit().unwrap();
    assert!(
        move_edit
            .move_page_breaks("Target", "Sheet1")
            .unwrap()
            .is_some()
    );
    let moved_book = move_edit.commit().unwrap().into_workbook();
    assert_eq!(
        moved_book
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap(),
        source
    );
    assert_eq!(
        moved_book
            .sheet("Target")
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap(),
        PageBreaks::new()
    );

    let reopened = Workbook::from_bytes(moved_book.to_plain_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .sheet("Target")
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap(),
        PageBreaks::new()
    );
}

#[test]
fn cell_copy_and_move_close_formula_text_and_style_dependencies() {
    let base = two_sheet_workbook();
    let style = base.styles().unwrap().base().unwrap();
    let mut prepare = base.edit().unwrap();
    let mut source = prepare.sheet("Sheet1").unwrap().unwrap();
    source
        .set("A1", Formula::new("B1*2").unwrap())
        .unwrap()
        .set("A2", "shared text")
        .unwrap()
        .style("A1", &style)
        .unwrap()
        .style("A2", &style)
        .unwrap();
    let prepared = prepare.commit().unwrap().into_workbook();

    let mut copy = prepared.edit().unwrap();
    assert!(
        copy.copy_cells("Sheet1", "A1:A2", "Target", "B2")
            .unwrap()
            .is_some()
    );
    let commit = copy.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let durable_json = durable.to_deterministic_json().unwrap();
    assert_eq!(durable_json, durable.to_deterministic_json().unwrap());
    let parsed = DurablePatch::from_deterministic_json(&durable_json).unwrap();
    let copied = parsed.apply(&prepared).unwrap();
    let target = copied.sheet("Target").unwrap().unwrap();
    assert!(matches!(
        target.cell("B2").unwrap().stored().unwrap(),
        Cell::Formula(formula)
            if formula.text() == "C2*2" && formula.cached().is_none()
    ));
    assert!(matches!(
        target.cell("B3").unwrap().stored().unwrap(),
        Cell::Value(Value::Text(text)) if text.as_str() == "shared text"
    ));
    assert!(matches!(
        target.local_style("B2").unwrap(),
        Some(LocalStyle::Shared(_))
    ));

    let reopened = Workbook::from_bytes(copied.to_plain_bytes().unwrap()).unwrap();
    assert!(matches!(
        reopened
            .sheet("Target")
            .unwrap()
            .unwrap()
            .cell("B2")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Formula(formula) if formula.text() == "C2*2"
    ));
    assert_eq!(
        parsed
            .inverse()
            .apply(&copied)
            .unwrap()
            .to_plain_bytes()
            .unwrap(),
        prepared.to_plain_bytes().unwrap()
    );

    let mut history = History::new(prepared.clone(), HistoryLimits::new(2, 256 * 1024 * 1024));
    assert!(commit.record(&mut history).unwrap().is_empty());
    assert!(history.undo());
    assert!(history.redo());

    let mut move_edit = copied.edit().unwrap();
    assert!(
        move_edit
            .move_cells("Target", "B2:B3", "Sheet1", "D4")
            .unwrap()
            .is_some()
    );
    let moved = move_edit.commit().unwrap().into_workbook();
    assert!(
        moved
            .sheet("Target")
            .unwrap()
            .unwrap()
            .cell("B2")
            .unwrap()
            .stored()
            .is_none()
    );
    assert!(matches!(
        moved
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("D4")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Formula(formula) if formula.text() == "E4*2"
    ));
    let reopened_move = Workbook::from_bytes(moved.to_plain_bytes().unwrap()).unwrap();
    assert!(matches!(
        reopened_move
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("D4")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Formula(formula) if formula.text() == "E4*2"
    ));

    let mut overlap = prepared.edit().unwrap();
    overlap
        .move_cells("Sheet1", "A1:A2", "Sheet1", "A2")
        .unwrap();
    let overlapped = overlap.commit().unwrap().into_workbook();
    let sheet = overlapped.sheet("Sheet1").unwrap().unwrap();
    assert!(sheet.cell("A1").unwrap().stored().is_none());
    assert!(matches!(
        sheet.cell("A2").unwrap().stored().unwrap(),
        Cell::Formula(formula) if formula.text() == "B2*2"
    ));
    assert!(matches!(
        sheet.cell("A3").unwrap().stored().unwrap(),
        Cell::Value(Value::Text(text)) if text.as_str() == "shared text"
    ));
}

#[test]
fn cell_transfer_participates_in_non_applying_three_way_plans() {
    let base = two_sheet_workbook();
    let mut prepare = base.edit().unwrap();
    prepare
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", Formula::new("B1+1").unwrap())
        .unwrap();
    let prepared = prepare.commit().unwrap().into_workbook();
    let limits = MergeLimits::new(2, 64, 128, 64);

    let mut left = prepared.edit().unwrap();
    left.copy_cells("Sheet1", "A1", "Target", "B2").unwrap();
    let mut right = prepared.edit().unwrap();
    right
        .sheet("Target")
        .unwrap()
        .unwrap()
        .set("H8", "right")
        .unwrap();
    let plan = left.plan_three_way(right, limits).unwrap();
    assert!(plan.conflicts().is_empty());
    let merged = plan.finish().unwrap().commit().unwrap().into_workbook();
    assert!(matches!(
        merged
            .sheet("Target")
            .unwrap()
            .unwrap()
            .cell("B2")
            .unwrap()
            .stored()
            .unwrap(),
        Cell::Formula(formula) if formula.text() == "C2+1"
    ));

    let mut conflict_left = prepared.edit().unwrap();
    conflict_left
        .copy_cells("Sheet1", "A1", "Target", "B2")
        .unwrap();
    let mut conflict_right = prepared.edit().unwrap();
    conflict_right
        .sheet("Target")
        .unwrap()
        .unwrap()
        .set("B2", "conflict")
        .unwrap();
    assert!(
        !conflict_left
            .plan_three_way(conflict_right, limits)
            .unwrap()
            .conflicts()
            .is_empty()
    );
}

#[test]
fn cell_transfer_is_bounded_and_malformed_package_bytes_stay_test_only() {
    let base = two_sheet_workbook();
    let mut edit = base.edit().unwrap();
    assert!(
        edit.copy_cells("Sheet1", "A1:XFD1048576", "Target", "A1")
            .is_err()
    );

    let truncated_zip = b"PK\x03\x04\0\0\0".to_vec();
    assert!(Workbook::from_bytes(truncated_zip).is_err());

    let array_book = Workbook::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsx/MatrixFormulaEvalTestData.xlsx"
    ))
    .unwrap();
    let mut array_edit = array_book.edit().unwrap();
    assert!(array_edit.copy_cells(0usize, "I3", 0usize, "A100").is_err());
}

#[test]
fn real_poi_fixture_round_trips_and_inverse_preserves_the_exact_package() {
    let base = Workbook::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xlsx"
    ))
    .unwrap();
    let before = base.to_plain_bytes().unwrap();
    let names = base
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    let source = names
        .iter()
        .find(|name| {
            let value = base
                .sheet(name.as_str())
                .unwrap()
                .unwrap()
                .page_breaks()
                .unwrap();
            value.horizontal().is_some() || value.vertical().is_some()
        })
        .unwrap();
    let target = names.iter().find(|name| *name != source).unwrap();

    let mut edit = base.edit().unwrap();
    edit.copy_page_breaks(source.as_str(), target.as_str())
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let changed = durable.apply(&base).unwrap();
    let reopened = Workbook::from_bytes(changed.to_plain_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .sheet(target.as_str())
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap(),
        base.sheet(source.as_str())
            .unwrap()
            .unwrap()
            .page_breaks()
            .unwrap()
    );
    let restored = durable.inverse().apply(&changed).unwrap();
    assert_eq!(restored.to_plain_bytes().unwrap(), before);
}

#[test]
fn real_poi_and_libreoffice_formulas_copy_and_fully_reopen() {
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../..");
    let cases = [
        (
            format!("{root}/test-data/poi/test-data/spreadsheet/shared_formulas.xlsx"),
            "A41",
            "AA1000",
            Some(("A1", "Z999", "Currently Using")),
        ),
        (
            format!(
                "{root}/test-data/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/basic.xlsx"
            ),
            "B19",
            "Y1000*10",
            None,
        ),
    ];
    for (path, source, expected, shared_text) in cases {
        let base = Workbook::open(path).unwrap();
        let before = base.to_plain_bytes().unwrap();
        let sheet_name = base.sheet(0usize).unwrap().unwrap().name().to_owned();
        let mut edit = base.edit().unwrap();
        edit.copy_cells(sheet_name.as_str(), source, sheet_name.as_str(), "Z1000")
            .unwrap();
        if let Some((source, target, _)) = shared_text {
            edit.copy_cells(sheet_name.as_str(), source, sheet_name.as_str(), target)
                .unwrap();
        }
        let durable = edit.commit().unwrap().patch().durable().unwrap();
        let changed = durable.apply(&base).unwrap();
        let reopened = Workbook::from_bytes(changed.to_plain_bytes().unwrap()).unwrap();
        assert!(matches!(
            reopened
                .sheet(sheet_name.as_str())
                .unwrap()
                .unwrap()
                .cell("Z1000")
                .unwrap()
                .stored()
                .unwrap(),
            Cell::Formula(formula)
                if formula.text() == expected && formula.cached().is_none()
        ));
        if let Some((_, target, expected)) = shared_text {
            let sheet = reopened.sheet(sheet_name.as_str()).unwrap().unwrap();
            assert!(matches!(
                sheet.cell(target).unwrap().stored().unwrap(),
                Cell::Value(Value::Text(text)) if text.as_str() == expected
            ));
            assert!(matches!(
                sheet.local_style(target).unwrap(),
                Some(LocalStyle::Shared(_))
            ));
        }
        assert_eq!(
            durable
                .inverse()
                .apply(&changed)
                .unwrap()
                .to_plain_bytes()
                .unwrap(),
            before
        );
    }
}
