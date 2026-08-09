#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_xlsx::page_breaks::{Break, Collection, PageBreaks};
use litchi_xlsx::{
    Cell, DurablePatch, Error, History, HistoryLimits, MergeChoice, MergeLimits, Package, Value,
    Workbook,
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
