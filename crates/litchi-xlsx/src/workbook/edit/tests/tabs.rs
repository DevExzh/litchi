//! Tab selection, visibility, rename, and reorder edit tests.

use super::super::*;
use super::support::{
    active_second_sheet_workbook, part_text, rename_reference_workbook, three_sheet_workbook,
    two_sheet_workbook,
};

use crate::cell::Value;
use crate::formula::Formula;
use litchi_opc::{BlobPart, TargetMode};

#[test]
fn tab_visibility_is_selector_first_reversible_and_active_safe() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let source_bytes = source.to_bytes().expect("source bytes");
    assert_eq!(
        source.active_sheet().map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".to_owned())
    );

    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet1")
        .expect("name lookup")
        .expect("tab")
        .hide();
    let hidden = edit.commit().expect("hide active tab");
    assert_eq!(hidden.patch().len(), 2);
    let (before, after) = hidden.patch().changes()[0]
        .active()
        .expect("implicit active relocation");
    assert_eq!((before.name(), before.position()), ("Sheet1", 0));
    assert_eq!((after.name(), after.position()), ("Sheet2", 1));
    assert!(matches!(
        hidden.patch().changes()[1],
        Change::Visibility {
            position: 0,
            before: Visibility::Visible,
            after: Visibility::Hidden,
            ..
        }
    ));
    assert_eq!(
        hidden
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
    assert_eq!(
        hidden
            .workbook()
            .active_sheet()
            .map(|sheet| sheet.name().to_owned()),
        Some("Sheet2".to_owned())
    );
    assert!(
        hidden
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .is_active()
    );

    let restored = hidden
        .workbook()
        .apply(&hidden.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);

    let mut last = hidden.workbook().edit().expect("last visible edit");
    last.tab(1usize)
        .expect("position lookup")
        .expect("tab")
        .very_hide();
    assert!(matches!(
        last.commit(),
        Err(Error::TabEditBlocked {
            sheet,
            position: 1,
            reason: TabEditBlock::LastVisibleTab,
        }) if sheet == "Sheet2"
    ));

    let mut swap = hidden.workbook().edit().expect("swap visibility");
    swap.tab("Sheet1").expect("lookup").expect("tab").show();
    swap.tab(1usize).expect("lookup").expect("tab").very_hide();
    let swapped = swap.commit().expect("swap commit");
    assert_eq!(
        swapped
            .workbook()
            .sheet(1usize)
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::VeryHidden
    );
    assert_eq!(
        swapped
            .workbook()
            .active_sheet()
            .map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".to_owned())
    );

    let mut no_op = source.edit().expect("no-op edit");
    no_op.tab(0usize).expect("lookup").expect("tab").show();
    assert!(no_op.commit().expect("no-op commit").patch().is_empty());
}
#[test]
fn active_tab_is_selector_first_reversible_and_composable() {
    for kind in [WorksheetKind::Worksheet, WorksheetKind::Chart] {
        let source = two_sheet_workbook(kind);
        let source_bytes = source.to_bytes().expect("source bytes");
        assert!(
            source
                .sheet("Sheet1")
                .expect("lookup")
                .expect("first tab")
                .is_active()
        );
        assert!(
            !source
                .sheet(1usize)
                .expect("lookup")
                .expect("second tab")
                .is_active()
        );

        let mut edit = source.edit().expect("edit");
        edit.tab(1usize)
            .expect("position lookup")
            .expect("tab")
            .activate();
        assert_eq!(edit.len(), 1);
        let committed = edit.commit().expect("activate");
        assert_eq!(committed.patch().len(), 1);
        let (before, after) = committed.patch().changes()[0]
            .active()
            .expect("active change");
        assert_eq!((before.name(), before.position()), ("Sheet1", 0));
        assert_eq!((after.name(), after.position()), ("Sheet2", 1));
        let active = committed.workbook().active_sheet().expect("active sheet");
        assert_eq!(active.name(), "Sheet2");
        assert_eq!(active.kind(), kind);
        assert!(active.is_active());
        assert_eq!(
            committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );

        let mut no_op = committed.workbook().edit().expect("no-op edit");
        no_op
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        assert!(no_op.commit().expect("no-op commit").patch().is_empty());
    }

    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut cell = source.edit().expect("cell edit");
    cell.sheet("Sheet2")
        .expect("lookup")
        .expect("worksheet")
        .set("A1", "active payload")
        .expect("cell");
    let mut active = source.edit().expect("active edit");
    active
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    cell.join(active).expect("orthogonal join");
    let committed = cell.commit().expect("joined commit");
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.sheets[1].part_uri)
            .count(),
        1
    );
    let active = committed.workbook().active_sheet().expect("active sheet");
    assert_eq!(active.name(), "Sheet2");
    assert!(matches!(
        active.cell("A1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "active payload"
    ));

    let mut replaced = source.edit().expect("replacement edit");
    replaced
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    replaced
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .activate();
    assert_eq!(replaced.len(), 1);
    assert!(
        replaced
            .commit()
            .expect("last activation wins")
            .patch()
            .is_empty()
    );
}

#[test]
fn active_tab_requires_final_visibility_and_conflicts_globally() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut hide = source.edit().expect("hide edit");
    hide.tab("Sheet2").expect("lookup").expect("tab").hide();
    let hidden = hide.commit().expect("hidden source");

    let mut blocked = hidden.workbook().edit().expect("blocked edit");
    blocked
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    assert!(matches!(
        blocked.commit(),
        Err(Error::TabEditBlocked {
            sheet,
            position: 1,
            reason: TabEditBlock::NotVisible,
        }) if sheet == "Sheet2"
    ));

    let mut repaired = hidden.workbook().edit().expect("repair edit");
    repaired
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .show()
        .activate();
    let repaired = repaired.commit().expect("show and activate");
    let active = repaired.workbook().active_sheet().expect("active sheet");
    assert_eq!(active.name(), "Sheet2");
    assert!(active.visibility().is_visible());
    assert_eq!(repaired.patch().len(), 2);

    let mut contradictory = source.edit().expect("contradictory edit");
    contradictory
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate()
        .very_hide();
    assert!(matches!(
        contradictory.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::NotVisible,
            ..
        })
    ));

    let mut left = source.edit().expect("left");
    left.tab("Sheet2").expect("lookup").expect("tab").activate();
    let mut right = source.edit().expect("right");
    right
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .activate();
    let error = left.join(right).expect_err("active-tab intents are global");
    let conflicts = error.conflicts().expect("active conflict");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_active());
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");

    let mut active = source.edit().expect("active");
    active
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let mut visibility = source.edit().expect("visibility");
    visibility
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .hide();
    active.join(visibility).expect("orthogonal facets");
    let committed = active.commit().expect("joined commit");
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Sheet2"
    );
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
}

#[test]
fn tab_rename_is_typed_dependency_aware_move_first_and_reversible() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_part = source.inner.sheets[0].part_uri.clone();
    let source_id = source.inner.sheets[0].native_id;
    let source_relationship = source.inner.sheets[0].relationship_id.clone();

    let mut edit = source.edit().expect("edit");
    edit.tab("data")
        .expect("caseless lookup")
        .expect("tab")
        .rename(String::from("Input 2026"))
        .expect("checked rename");
    edit.sheet("Calc")
        .expect("Calc lookup")
        .expect("Calc sheet")
        .set("D1", "composed")
        .expect("same-part cell edit");
    let committed = edit.commit().expect("rename commit");
    assert!(
        committed
            .workbook()
            .sheet("Data")
            .expect("lookup")
            .is_none()
    );
    let renamed = committed
        .workbook()
        .sheet("INPUT 2026")
        .expect("Unicode caseless lookup")
        .expect("renamed sheet");
    assert_eq!(renamed.name(), "Input 2026");
    assert_eq!(renamed.data.part_uri, source_part);
    assert_eq!(renamed.data.native_id, source_id);
    assert_eq!(renamed.data.relationship_id, source_relationship);
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed.patch().changes()[0].renamed(),
        Some((0, "Data", "Input 2026"))
    );

    let calc = committed
        .workbook()
        .sheet("Calc")
        .expect("lookup")
        .expect("Calc");
    assert!(matches!(
        calc.cell("A1").expect("formula cell").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "'Input 2026'!A1"
    ));
    assert!(matches!(
        calc.cell("D1").expect("composed cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "composed"
    ));
    for uri in [
        "/xl/workbook.xml",
        "/xl/worksheets/sheet2.xml",
        "/xl/tables/table1.xml",
        "/xl/charts/chart1.xml",
        "/xl/pivotCache/pivotCacheDefinition1.xml",
        "/docProps/app.xml",
    ] {
        let text = part_text(committed.workbook(), uri);
        assert!(
            text.contains("Input 2026"),
            "expected renamed dependency in {uri}: {text}"
        );
    }
    assert!(
        part_text(committed.workbook(), "/xl/externalLinks/externalLink1.xml")
            .contains("[1]Data!A1")
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse rename");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
}

#[test]
fn tab_rename_validates_names_collisions_swaps_and_join_facets() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut invalid = source.edit().expect("invalid edit");
    let error = invalid
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("bad/name")
        .expect_err("invalid name");
    assert!(matches!(error, Error::SheetName(_)));

    let mut collision = source.edit().expect("collision edit");
    collision
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("sheet2")
        .expect("valid spelling");
    assert!(matches!(
        collision.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut swap = source.edit().expect("swap edit");
    swap.tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Sheet2")
        .expect("first swap");
    swap.tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .rename("Sheet1")
        .expect("second swap");
    let swapped = swap.commit().expect("simultaneous swap");
    assert_eq!(
        swapped
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet2", "Sheet1"]
    );

    let mut left = source.edit().expect("left");
    left.tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("First")
        .expect("left rename");
    let mut same = source.edit().expect("same");
    same.tab(0usize)
        .expect("lookup")
        .expect("tab")
        .rename("Other")
        .expect("same rename");
    let error = left.join(same).expect_err("same name facet conflicts");
    assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());

    let mut right = source.edit().expect("right");
    right
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .rename(Name::new("Second").expect("prevalidated"))
        .expect("moved typed name");
    left.join(right).expect("disjoint names join");
    let joined = left.commit().expect("joined renames");
    assert!(joined.workbook().sheet("first").expect("lookup").is_some());
    assert!(joined.workbook().sheet("second").expect("lookup").is_some());
}

#[test]
fn tab_reorder_is_selector_first_dependency_aware_and_reversible() {
    let source = three_sheet_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    assert!(
        edit.move_before("Sheet3", "Sheet1")
            .expect("move lookup")
            .is_some()
    );
    assert_eq!(edit.len(), 1);
    let committed = edit.commit().expect("reorder");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet3", "Sheet1", "Sheet2"]
    );
    assert_eq!(
        committed
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .name(),
        "Sheet2"
    );
    assert_eq!(
        committed
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .position(),
        2
    );
    assert!(matches!(
        committed
            .workbook()
            .sheet("Sheet3")
            .expect("lookup")
            .expect("Sheet3")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "three"
    ));
    assert_eq!(
        committed
            .workbook()
            .defined_names()
            .iter()
            .map(|name| (name.name.as_str(), name.local_sheet_id))
            .collect::<Vec<_>>(),
        [
            ("FirstLocal", Some(1)),
            ("ThirdLocal", Some(0)),
            ("Global", None),
        ]
    );
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(committed.patch().changes()[0].sheet(), "Sheet3");
    assert_eq!(committed.patch().changes()[0].moved(), Some((2, 0)));
    let (before, after) = committed.patch().changes()[1]
        .active()
        .expect("active position remap");
    assert_eq!((before.name(), before.position()), ("Sheet2", 1));
    assert_eq!((after.name(), after.position()), ("Sheet2", 2));

    let inverse = committed.patch().inverse();
    assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
    let restored = committed
        .workbook()
        .apply(&inverse)
        .expect("inverse reorder");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);

    let chart_source = two_sheet_workbook(WorksheetKind::Chart);
    let chart_bytes = chart_source.to_bytes().expect("chart source bytes");
    let mut chart_edit = chart_source.edit().expect("chart edit");
    chart_edit
        .move_before("Sheet2", "Sheet1")
        .expect("chart lookup")
        .expect("chart tabs");
    let chart = chart_edit.commit().expect("chart reorder");
    let first = chart
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("tab");
    assert_eq!(
        (first.name(), first.kind()),
        ("Sheet2", WorksheetKind::Chart)
    );
    assert_eq!(
        chart
            .workbook()
            .apply(&chart.patch().inverse())
            .expect("chart inverse")
            .workbook()
            .to_bytes()
            .expect("chart bytes"),
        chart_bytes
    );
}

#[test]
fn tab_reorder_composes_with_other_facets_and_conflicts_globally() {
    let source = three_sheet_workbook();
    let mut order = source.edit().expect("order");
    order
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    let mut active = source.edit().expect("active");
    active
        .tab("Sheet3")
        .expect("lookup")
        .expect("tab")
        .activate();
    let mut visibility = source.edit().expect("visibility");
    visibility
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .hide();
    let mut cell = source.edit().expect("cell");
    cell.sheet("Sheet3")
        .expect("lookup")
        .expect("sheet")
        .set("B1", "moved payload")
        .expect("cell");
    order.join(active).expect("order and active");
    order.join(visibility).expect("order and visibility");
    order.join(cell).expect("order and cell");
    let committed = order.commit().expect("composed reorder");
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("Sheet3", 0));
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
    assert!(matches!(
        active.cell("B1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "moved payload"
    ));
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.workbook_uri)
            .count(),
        1
    );
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.sheets[2].part_uri)
            .count(),
        1
    );

    let mut left = source.edit().expect("left order");
    left.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    let mut right = source.edit().expect("right order");
    right
        .move_after("Sheet1", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let error = left.join(right).expect_err("order is one global facet");
    let conflicts = error.conflicts().expect("order conflict");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_order());

    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut same_position = source.edit().expect("same-position edit");
    same_position
        .move_before("Sheet2", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    same_position
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let committed = same_position
        .commit()
        .expect("active identity changes at the same position");
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("Sheet2", 0));
    let active_change = committed
        .patch()
        .changes()
        .iter()
        .find_map(Change::active)
        .expect("semantic active change");
    assert_eq!(active_change.0.position(), 0);
    assert_eq!(active_change.1.position(), 0);
    assert_eq!(active_change.0.name(), "Sheet1");
    assert_eq!(active_change.1.name(), "Sheet2");
}

#[test]
fn numeric_tab_moves_are_checked_and_all_positions_round_trip() {
    for from in 0..3usize {
        for to in 0..3usize {
            let source = three_sheet_workbook();
            let source_bytes = source.to_bytes().expect("source bytes");
            let mut expected = vec!["Sheet1", "Sheet2", "Sheet3"];
            let moved = expected.remove(from);
            expected.insert(to, moved);
            let mut edit = source.edit().expect("edit");
            assert!(edit.move_to(from, to).expect("lookup").is_some());
            let committed = edit.commit().expect("move");
            assert_eq!(
                committed
                    .workbook()
                    .sheets()
                    .map(|sheet| sheet.name().to_owned())
                    .collect::<Vec<_>>(),
                expected.into_iter().map(str::to_owned).collect::<Vec<_>>()
            );
            let restored = committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse");
            assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
        }
    }

    let source = three_sheet_workbook();
    let mut missing = source.edit().expect("missing edit");
    assert!(
        missing
            .move_before("Absent", "Sheet1")
            .expect("lookup")
            .is_none()
    );
    assert!(missing.move_to("Sheet1", 3).expect("bounds").is_none());
    assert!(missing.commit().expect("no-op").patch().is_empty());

    let mut cancelled = source.edit().expect("cancelled edit");
    cancelled
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .move_after("Sheet3", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    assert!(cancelled.is_empty());
    assert_eq!(cancelled.len(), 0);
    assert!(cancelled.commit().expect("cancelled").patch().is_empty());

    let mut cancelled = source.edit().expect("cancelled join edit");
    cancelled
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .move_after("Sheet3", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let mut effective = source.edit().expect("effective join edit");
    effective
        .move_before("Sheet2", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .join(effective)
        .expect("cancelled order has no conflict");
    let joined = cancelled.commit().expect("joined order");
    assert_eq!(
        joined
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet2", "Sheet1", "Sheet3"]
    );

    let source_bytes = source.to_bytes().expect("source bytes");
    let mut sequence = source.edit().expect("sequence edit");
    sequence
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    sequence
        .move_after("Sheet1", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let sequence = sequence.commit().expect("move sequence");
    assert_eq!(
        sequence
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet3", "Sheet2", "Sheet1"]
    );
    assert_eq!(sequence.patch().len(), 2);
    assert_eq!(sequence.patch().changes()[0].moved(), Some((2, 0)));
    assert_eq!(sequence.patch().changes()[1].moved(), Some((1, 2)));
    let inverse = sequence.patch().inverse();
    assert_eq!(inverse.changes()[0].moved(), Some((2, 1)));
    assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
    assert_eq!(
        sequence
            .workbook()
            .apply(&inverse)
            .expect("sequence inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn tab_reorder_blocks_protection_and_revision_tracking() {
    let source = three_sheet_workbook();
    let mut package = source.inner.package.clone();
    let workbook = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook");
    let xml = std::str::from_utf8(workbook.blob())
        .expect("UTF-8")
        .replace(
            "<bookViews>",
            "<workbookProtection lockStructure=\"1\"/><bookViews>",
        );
    workbook.set_blob(xml.into_bytes());
    let protected = Workbook::from_package(package).expect("protected workbook");
    let mut edit = protected.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));
    let mut rename = protected.edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("checked name");
    assert!(matches!(
        rename.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));

    let mut package = source.inner.package.clone();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/revisions/revisionHeaders1.xml").expect("revision URI"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml"
                .to_owned(),
            br#"<headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
                .to_vec(),
        )))
        .expect("revision part");
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"
                .to_owned(),
            "revisions/revisionHeaders1.xml".to_owned(),
            "rIdRevisionHeaders".to_owned(),
            TargetMode::Internal,
        )
        .expect("revision relationship");
    let tracked = Workbook::from_package(package).expect("tracked workbook");
    let mut edit = tracked.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::TrackedWorkbook,
            ..
        })
    ));
    let mut rename = tracked.edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("checked name");
    assert!(matches!(
        rename.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::TrackedWorkbook,
            ..
        })
    ));
}

#[test]
fn tab_visibility_composes_with_cells_and_conflicts_by_facet() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut cell = source.edit().expect("cell edit");
    cell.sheet("Sheet2")
        .expect("lookup")
        .expect("worksheet")
        .set("A1", "preserved while hidden")
        .expect("cell");
    let mut tab = source.edit().expect("tab edit");
    tab.tab(1usize).expect("lookup").expect("tab").hide();
    cell.join(tab).expect("orthogonal join");
    let committed = cell.commit().expect("joined commit");
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.workbook_uri)
            .count(),
        1
    );
    let sheet = committed
        .workbook()
        .sheet("Sheet2")
        .expect("lookup")
        .expect("sheet");
    assert_eq!(sheet.visibility(), &Visibility::Hidden);
    assert!(matches!(
        sheet.cell("A1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "preserved while hidden"
    ));

    let mut left = source.edit().expect("left");
    left.tab("Sheet2").expect("lookup").expect("tab").hide();
    let mut right = source.edit().expect("right");
    right.tab(1usize).expect("lookup").expect("tab").very_hide();
    let error = left.join(right).expect_err("same tab facet must conflict");
    let conflicts = error.conflicts().expect("tab conflicts");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_tab());
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");
}

#[test]
fn tab_visibility_applies_to_non_worksheet_sheet_kinds() {
    let source = two_sheet_workbook(WorksheetKind::Chart);
    assert_eq!(
        source
            .sheet("Sheet2")
            .expect("lookup")
            .expect("chart sheet")
            .kind(),
        WorksheetKind::Chart
    );
    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .very_hide();
    let committed = edit.commit().expect("chart tab commit");
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("chart sheet")
            .visibility(),
        &Visibility::VeryHidden
    );
}

#[test]
fn active_relocation_synchronizes_worksheet_and_chart_view_selection() {
    for kind in [WorksheetKind::Worksheet, WorksheetKind::Chart] {
        let source = active_second_sheet_workbook(kind);
        let source_bytes = source.to_bytes().expect("source bytes");
        assert_eq!(
            source.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet2".to_owned())
        );
        let mut edit = source.edit().expect("edit");
        edit.tab("Sheet2").expect("lookup").expect("tab").hide();
        let committed = edit.commit().expect("active hide");
        assert_eq!(
            committed
                .workbook()
                .active_sheet()
                .map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".to_owned())
        );
        let new_active = committed
            .workbook()
            .inner
            .package
            .get_part(&committed.workbook().inner.sheets[0].part_uri)
            .expect("new active part")
            .blob();
        assert!(
            std::str::from_utf8(new_active)
                .expect("new active XML")
                .contains(r#"tabSelected="1""#)
        );
        let old_active = committed
            .workbook()
            .inner
            .package
            .get_part(&committed.workbook().inner.sheets[1].part_uri)
            .expect("old active part")
            .blob();
        assert!(
            !std::str::from_utf8(old_active)
                .expect("old active XML")
                .contains("tabSelected")
        );
        assert_eq!(
            committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );
    }
}

#[test]
fn tab_visibility_blocks_protected_workbook_structure() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let protected = Workbook::from_package(package).expect("protected workbook");
    let mut edit = protected.edit().expect("edit");
    edit.tab("Sheet2").expect("lookup").expect("tab").hide();
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));

    let mut activation = protected.edit().expect("activation edit");
    activation
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let activated = activation
        .commit()
        .expect("structure protection permits selection");
    assert_eq!(
        activated
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .name(),
        "Sheet2"
    );
}

#[test]
fn showing_an_unknown_producer_state_repairs_it_explicitly() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" state="show" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let source = Workbook::from_package(package).expect("producer workbook");
    assert!(matches!(
        source
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        Visibility::Unknown(value) if value.as_ref() == "show"
    ));
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet2").expect("lookup").expect("tab").show();
    let committed = edit.commit().expect("repair commit");
    assert!(
        committed
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .visibility()
            .is_visible()
    );
    assert_eq!(
        committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}
