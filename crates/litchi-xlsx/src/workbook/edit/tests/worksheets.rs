//! Worksheet insertion, addition, removal, and dependency tests.

use super::super::*;
use super::support::{
    part_text, rename_reference_workbook, three_sheet_workbook, two_sheet_workbook,
};

use crate::cell::Value;
use crate::formula::Formula;
use litchi_opc::{BlobPart, TargetMode};

#[test]
fn worksheet_add_is_atomic_populatable_active_and_reversible() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.add(String::from("Summary")).expect("new sheet");
        assert_eq!(sheet.name(), "Summary");
        assert_eq!(sheet.position(), 1);
        sheet
            .set("A1", "created in one transaction")
            .and_then(|sheet| sheet.set("B2", Formula::new("1+1").expect("checked formula")))
            .expect("new cells");
        sheet.row(3u32).expect("row").hide();
        sheet.column(2u32).expect("column").hide();
        sheet.activate();
    }
    let committed = edit.commit().expect("create commit");
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Summary"]
    );
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Summary"
    );
    let summary = committed
        .workbook()
        .sheet("summary")
        .expect("caseless lookup")
        .expect("created sheet");
    assert_eq!(summary.data.native_id, 2);
    assert!(matches!(
        summary.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "created in one transaction"
    ));
    assert!(matches!(
        summary.cell("B2").expect("B2").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "1+1"
    ));
    assert!(summary.row(3u32).expect("row").hidden());
    assert!(summary.column(2u32).expect("column").hidden());
    assert!(
        committed
            .patch()
            .changes()
            .iter()
            .any(|change| { change.created() == Some((1, "Summary")) })
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse create");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    let replayed = source.apply(committed.patch()).expect("forward replay");
    assert!(
        replayed
            .workbook()
            .sheet("Summary")
            .expect("lookup")
            .is_some()
    );
}
#[test]
fn worksheet_insert_is_selector_first_order_aware_and_reversible() {
    let source = three_sheet_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("move lookup")
        .expect("both base tabs");
    assert!(
        edit.add_before("Never Added", "Absent")
            .expect("missing lookup")
            .is_none()
    );
    {
        let mut sheet = edit
            .add_before("Before A", "Sheet1")
            .expect("lookup")
            .expect("anchor");
        assert_eq!(sheet.position(), 1);
        sheet.set("A1", "before-a").expect("payload");
    }
    edit.add_before("Before B", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "before-b")
        .expect("payload");
    edit.add_after("After A", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "after-a")
        .expect("payload");
    edit.add_after("After B", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "after-b")
        .expect("payload")
        .activate();
    {
        let sheet = edit
            .add_before("Before Third", 2usize)
            .expect("numeric lookup")
            .expect("numeric anchor");
        assert_eq!(sheet.position(), 0);
    }
    edit.add("Tail").expect("tail");

    let committed = edit.commit().expect("insert commit");
    let names = committed
        .workbook()
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(
        names,
        [
            "Before Third",
            "Sheet3",
            "Before A",
            "Before B",
            "Sheet1",
            "After A",
            "After B",
            "Sheet2",
            "Tail",
        ]
    );
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("After B", 6));
    assert!(matches!(
        committed
            .workbook()
            .sheet("Before A")
            .expect("lookup")
            .expect("created")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "before-a"
    ));
    assert_eq!(
        committed
            .workbook()
            .defined_names()
            .iter()
            .map(|name| (name.name.as_str(), name.local_sheet_id))
            .collect::<Vec<_>>(),
        [
            ("FirstLocal", Some(4)),
            ("ThirdLocal", Some(1)),
            ("Global", None),
        ]
    );
    assert_eq!(
        committed
            .patch()
            .changes()
            .iter()
            .filter_map(Change::created)
            .collect::<Vec<_>>(),
        [
            (2, "Before A"),
            (3, "Before B"),
            (5, "After A"),
            (6, "After B"),
            (0, "Before Third"),
            (8, "Tail"),
        ]
    );

    let committed_bytes = committed.workbook().to_bytes().expect("committed bytes");
    assert_eq!(
        source
            .apply(committed.patch())
            .expect("forward replay")
            .workbook()
            .to_bytes()
            .expect("replayed bytes"),
        committed_bytes
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

#[test]
fn worksheet_insert_join_preserves_explicit_edit_order() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut left = source.edit().expect("left");
    left.add_before("Left", "Sheet2")
        .expect("lookup")
        .expect("anchor");
    let mut right = source.edit().expect("right");
    right
        .add_before("Right", 1usize)
        .expect("lookup")
        .expect("anchor");
    left.join(right).expect("disjoint insertion join");
    assert_eq!(
        left.commit()
            .expect("joined insertion")
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Left", "Right", "Sheet2"]
    );
}

#[test]
fn worksheet_add_validates_names_visibility_and_parallel_joins() {
    let source = Workbook::new().expect("source workbook");

    let mut duplicate = source.edit().expect("duplicate edit");
    duplicate.add("sheet1").expect("checked spelling");
    assert!(matches!(
        duplicate.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut duplicate = source.edit().expect("anchored duplicate edit");
    duplicate
        .add_before("sheet1", "Sheet1")
        .expect("lookup")
        .expect("anchor");
    assert!(matches!(
        duplicate.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut hidden = source.edit().expect("hidden edit");
    {
        let mut sheet = hidden.add("Hidden").expect("new sheet");
        sheet.hide().activate();
    }
    assert!(matches!(
        hidden.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::NotVisible,
            ..
        })
    ));

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .tab("Sheet1")
        .expect("lookup")
        .expect("existing tab")
        .hide();
    replacement.add("Replacement").expect("visible replacement");
    let replaced = replacement.commit().expect("replacement commit");
    assert_eq!(
        replaced.workbook().active_sheet().expect("active").name(),
        "Replacement"
    );
    assert!(matches!(
        replaced
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("old tab")
            .visibility(),
        Visibility::Hidden
    ));

    let mut left = source.edit().expect("left");
    left.add("North")
        .expect("North")
        .set("A1", 1_i32)
        .expect("cell");
    let mut right = source.edit().expect("right");
    right
        .add("South")
        .expect("South")
        .set("A1", 2_i32)
        .expect("cell");
    left.join(right).expect("disjoint appends join");
    let joined = left.commit().expect("joined create");
    assert_eq!(
        joined
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "North", "South"]
    );

    let mut one = source.edit().expect("one");
    one.add("Résumé").expect("first name");
    let mut two = source.edit().expect("two");
    two.add("RE\u{301}SUME\u{301}").expect("equivalent name");
    let error = one.join(two).expect_err("equivalent creations conflict");
    assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());
}

#[test]
fn worksheet_add_closes_rename_formula_and_extended_property_dependencies() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.tab("Data")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("rename");
    edit.add("New & Sheet")
        .expect("new sheet")
        .set(
            "A1",
            Formula::new("Data!A1").expect("source-snapshot formula"),
        )
        .expect("formula");
    let committed = edit.commit().expect("composed structural commit");
    let created = committed
        .workbook()
        .sheet("New & Sheet")
        .expect("lookup")
        .expect("created sheet");
    assert!(matches!(
        created.cell("A1").expect("formula").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "Input!A1"
    ));
    let properties = part_text(committed.workbook(), "/docProps/app.xml");
    assert!(properties.contains("size=\"4\""));
    assert!(properties.contains(
            "<vt:lpstr>Input</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>New &amp; Sheet</vt:lpstr><vt:lpstr>Input!Print_Area</vt:lpstr>"
        ));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
}

#[test]
fn worksheet_add_synchronizes_optional_properties_during_a_simultaneous_reorder() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.move_before("Calc", "Data")
        .expect("move")
        .expect("both tabs");
    edit.add_before("Middle", "Data")
        .expect("lookup")
        .expect("anchor");
    edit.add("Tail").expect("new sheet");
    let committed = edit.commit().expect("composed commit");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Calc", "Middle", "Data", "Tail"]
    );
    let properties = part_text(committed.workbook(), "/docProps/app.xml");
    assert!(properties.contains("size=\"5\""));
    assert!(properties.contains(concat!(
        "<vt:lpstr>Calc</vt:lpstr>",
        "<vt:lpstr>Middle</vt:lpstr>",
        "<vt:lpstr>Data</vt:lpstr>",
        "<vt:lpstr>Tail</vt:lpstr>",
        "<vt:lpstr>Data!Print_Area</vt:lpstr>"
    )));
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

#[test]
fn worksheet_add_allocates_strict_graph_identity_without_exposing_it() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let main = package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part");
    main.set_blob(
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Strict1" sheetId="7" r:id="tab"/></sheets></workbook>"#.to_vec(),
        );
    main.rels_mut().remove("rId1").expect("old worksheet rel");
    main.rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::STRICT_WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "tab".to_owned(),
            TargetMode::Internal,
        )
        .expect("strict worksheet rel");
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("URI"))
            .expect("worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/></worksheet>"#.to_vec(),
        );
    let source = Workbook::from_package(package).expect("strict source");
    let mut edit = source.edit().expect("edit");
    edit.add_before("Strict2", "Strict1")
        .expect("lookup")
        .expect("strict anchor");
    let committed = edit.commit().expect("strict create");
    let sheet = committed
        .workbook()
        .sheet("Strict2")
        .expect("lookup")
        .expect("created");
    assert_eq!(sheet.position(), 0);
    assert_eq!(sheet.data.native_id, 1);
    assert_eq!(sheet.data.relationship_id, "rId1");
    assert_eq!(sheet.data.part_uri.as_str(), "/xl/worksheets/sheet2.xml");
    let main = committed
        .workbook()
        .inner
        .package
        .get_part(&committed.workbook().inner.workbook_uri)
        .expect("workbook part");
    assert_eq!(
        main.rels()
            .get(&sheet.data.relationship_id)
            .expect("new relationship")
            .reltype(),
        litchi_opc::constants::relationship_type::STRICT_WORKSHEET
    );
    assert!(
        part_text(committed.workbook(), sheet.data.part_uri.as_str())
            .contains("http://purl.oclc.org/ooxml/spreadsheetml/main")
    );
}

#[test]
fn worksheet_add_blocks_protected_and_compatibility_owned_catalogs() {
    for xml in [
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.as_slice(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><mc:AlternateContent><mc:Choice Requires="x"><x:payload/></mc:Choice><mc:Fallback/></mc:AlternateContent></sheets></workbook>"#.as_slice(),
        ] {
            let baseline = Workbook::new().expect("baseline");
            let mut package = baseline.inner.package.clone();
            package
                .get_part_mut(&baseline.inner.workbook_uri)
                .expect("workbook")
                .set_blob(xml.to_vec());
            let source = Workbook::from_package(package).expect("source");
            let mut edit = source.edit().expect("edit");
            edit.add_before("Blocked", "Sheet1")
                .expect("lookup")
                .expect("anchor");
            assert!(matches!(
                edit.commit(),
                Err(Error::TabEditBlocked {
                    reason: TabEditBlock::ProtectedWorkbook
                        | TabEditBlock::MarkupCompatibility,
                    ..
                })
            ));
        }
}

#[test]
fn worksheet_remove_is_selector_first_active_relocating_and_reversible() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create
        .add("Delete")
        .expect("Delete")
        .set("A1", "removed payload")
        .expect("payload")
        .activate();
    create
        .add("Keep")
        .expect("Keep")
        .set("A1", 42_i32)
        .expect("retained payload");
    let source = create.commit().expect("create source").into_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    assert_eq!(source.active_sheet().expect("active").name(), "Delete");

    let mut edit = source.edit().expect("remove edit");
    assert!(edit.remove("missing").expect("missing selector").is_none());
    edit.remove("delete")
        .expect("selector")
        .expect("Delete worksheet");
    assert_eq!(edit.len(), 1);
    let committed = edit.commit().expect("remove commit");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Keep"]
    );
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Keep"
    );
    assert!(
        committed
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );
    assert!(matches!(
        committed.patch().changes().first(),
        Some(Change::Remove {
            sheet,
            position: 1,
            ..
        }) if sheet.as_ref() == "Delete"
    ));
    assert!(
            committed
                .patch()
                .changes()
                .iter()
                .any(|change| matches!(change, Change::Active { after, .. } if after.name() == "Keep" && after.position() == 1))
        );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse remove");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    let replayed = source.apply(committed.patch()).expect("forward replay");
    assert!(
        replayed
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );

    let mut activate_last = source.edit().expect("activate last");
    activate_last
        .tab("Keep")
        .expect("lookup")
        .expect("Keep")
        .activate();
    let active_last = activate_last
        .commit()
        .expect("active-last source")
        .into_workbook();
    let mut remove_last = active_last.edit().expect("remove last");
    remove_last.remove("Keep").expect("lookup").expect("Keep");
    assert_eq!(
        remove_last
            .commit()
            .expect("remove last active")
            .workbook()
            .active_sheet()
            .expect("replacement active")
            .name(),
        "Delete"
    );
}

#[test]
fn worksheet_remove_blocks_live_formulas_last_sheet_and_mixed_edits() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    create
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set("A1", Formula::new("Delete!A1").expect("formula"))
        .expect("formula cell");
    let source = create.commit().expect("source").into_workbook();
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            sheet,
            reason: RemoveBlock::IncomingReference,
            ..
        }) if sheet == "Delete"
    ));

    let single = Workbook::new().expect("single");
    let mut last = single.edit().expect("last edit");
    last.remove(0usize).expect("lookup").expect("only sheet");
    assert!(matches!(
        last.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::LastSheet,
            ..
        })
    ));

    let baseline = Workbook::new().expect("visibility baseline");
    let mut create = baseline.edit().expect("visibility create");
    create
        .tab("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .hide();
    create.add("Delete").expect("Delete").activate();
    let visibility = create.commit().expect("visibility source").into_workbook();
    let mut last_visible = visibility.edit().expect("last-visible edit");
    last_visible
        .remove("Delete")
        .expect("lookup")
        .expect("Delete");
    assert!(matches!(
        last_visible.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::LastVisibleTab,
            ..
        })
    ));

    let mut mixed = source.edit().expect("mixed edit");
    mixed
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set("B1", 1_i32)
        .expect("cell edit");
    assert!(matches!(
        mixed.remove("Delete"),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::MixedEdit,
            ..
        })
    ));

    let baseline = Workbook::new().expect("dynamic baseline");
    let mut create = baseline.edit().expect("dynamic create");
    create.add("Delete").expect("Delete");
    create
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set(
            "A1",
            Formula::new(r#"INDIRECT("Delete!A1")"#).expect("dynamic formula"),
        )
        .expect("dynamic cell");
    let dynamic = create.commit().expect("dynamic source").into_workbook();
    let mut remove = dynamic.edit().expect("dynamic removal");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::UnmodeledReference,
            ..
        })
    ));
}

#[test]
fn worksheet_remove_joins_disjoint_plans_and_blocks_unknown_dependencies() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("North").expect("North");
    create.add("South").expect("South");
    let source = create.commit().expect("source").into_workbook();

    let mut north = source.edit().expect("north edit");
    north.remove("North").expect("lookup").expect("North");
    let mut south = source.edit().expect("south edit");
    south.remove(2usize).expect("lookup").expect("South");
    north.join(south).expect("disjoint removals join");
    let committed = north.commit().expect("joined removal");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1"]
    );

    let mut package = source.inner.package.clone();
    let custom_uri = PackURI::new("/customXml/item1.xml").expect("custom URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            custom_uri,
            "application/xml".to_owned(),
            b"<root><futureFormulaCache>North!A1</futureFormulaCache></root>".to_vec(),
        )))
        .expect("custom part");
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-custom".to_owned(),
            "../customXml/item1.xml".to_owned(),
            "customRef".to_owned(),
            TargetMode::Internal,
        )
        .expect("custom relationship");
    let unknown = Workbook::from_package(package).expect("unknown producer workbook");
    let mut edit = unknown.edit().expect("remove edit");
    edit.remove("North").expect("lookup").expect("North");
    assert!(matches!(
        edit.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::UnmodeledReference,
            part,
            ..
        }) if part == "/customXml/item1.xml"
    ));
}

#[test]
fn worksheet_remove_blocks_macro_projects_and_extra_incoming_relationships() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();

    let mut macro_package = source.inner.package.clone();
    macro_package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/vbaProject.bin").expect("VBA URI"),
            litchi_opc::constants::content_type::OFC_VBA_PROJECT.to_owned(),
            vec![0, 1, 2, 3],
        )))
        .expect("VBA part");
    macro_package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::VBA_PROJECT.to_owned(),
            "vbaProject.bin".to_owned(),
            "vbaProject".to_owned(),
            TargetMode::Internal,
        )
        .expect("VBA relationship");
    let macro_book = Workbook::from_package(macro_package).expect("macro workbook");
    let mut remove = macro_book.edit().expect("macro remove");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::MacroProject,
            ..
        })
    ));

    let target = source
        .inner
        .sheets
        .get(1)
        .expect("Delete sheet")
        .part_uri
        .clone();
    let mut incoming_package = source.inner.package.clone();
    let mut referrer = BlobPart::new(
        PackURI::new("/xl/custom.xml").expect("custom URI"),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    referrer
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-incoming".to_owned(),
            target.relative_ref("/xl"),
            "sheetRef".to_owned(),
            TargetMode::Internal,
        )
        .expect("incoming relationship");
    incoming_package
        .try_add_part(Box::new(referrer))
        .expect("referrer part");
    let incoming = Workbook::from_package(incoming_package).expect("incoming workbook");
    let mut remove = incoming.edit().expect("incoming remove");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::IncomingRelationship,
            part,
            ..
        }) if part == "/xl/custom.xml"
    ));
}

#[test]
fn worksheet_remove_blocks_custom_workbook_view_identity() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();
    let native_id = source.inner.sheets[1].native_id;
    let mut package = source.inner.package.clone();
    let workbook = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook");
    let xml = std::str::from_utf8(workbook.blob())
            .expect("workbook UTF-8")
            .replace(
                "</workbook>",
                &format!(
                    r#"<customWorkbookViews><customWorkbookView name="Delete view" guid="{{00000000-0000-0000-0000-000000000001}}" activeSheetId="{native_id}"/></customWorkbookViews></workbook>"#
                ),
            );
    assert!(xml.contains("customWorkbookView"));
    workbook.set_blob(xml.into_bytes());
    let source = Workbook::from_package(package).expect("custom-view workbook");
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::IncomingReference,
            part,
            ..
        }) if part == "/xl/workbook.xml"
    ));
}

#[test]
fn worksheet_remove_accepts_case_equivalent_opc_targets() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();
    let relationship_id = source.inner.sheets[1].relationship_id.clone();
    let mut package = source.inner.package.clone();
    let relationships = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut();
    let relationship = relationships
        .remove(&relationship_id)
        .expect("worksheet relationship");
    relationships
        .try_add_relationship(
            relationship.reltype().to_owned(),
            "worksheets/SHEET2.XML".to_owned(),
            relationship_id,
            TargetMode::Internal,
        )
        .expect("case-equivalent relationship");
    let source = Workbook::from_package(package).expect("case-equivalent workbook");
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(
        remove
            .commit()
            .expect("case-equivalent removal")
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );
}
