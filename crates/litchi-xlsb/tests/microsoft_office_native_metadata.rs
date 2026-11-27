//! Native Excel metadata, grouped formulas and source-preserving edits.

use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xlsb::{Package, Workbook};
use std::io::Cursor;

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

#[test]
fn native_form_control_drawing_mirrors_survive_inert_cell_edits() {
    let bytes = native("excel-form-controls.xlsb");
    let package = Package::from_bytes(&bytes).unwrap();
    let elements = package.workbook().unwrap().drawing_elements(0).unwrap();
    assert_eq!(elements.len(), 11);
    // Excel's XLSB derivative carries MCE compatibility graphic frames,
    // not the XLSX xdr:sp controls. The generic drawing owner retains each
    // wrapper as opaque rather than inventing a selected shape identity.
    assert!(
        elements
            .iter()
            .all(|e| e.kind == litchi_xlsb::drawing::edit::ElementKind::Other
                && e.id.is_none()
                && e.name.is_none())
    );
    let parts = members(&bytes);
    use sha2::{Digest, Sha256};
    let digest = |data: &[u8]| {
        Sha256::digest(data)
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>()
    };
    assert_eq!(
        digest(&parts["xl/drawings/drawing1.xml"]),
        "031fc3826b614f6df90b7b20f665ae974589c163f832af370bb6ad233a22d49c"
    );
    assert_eq!(
        digest(&parts["xl/drawings/vmlDrawing1.vml"]),
        "622a97e86fcb0abed36496f4ed67f46ca361194f50c032d12003f318a7032215"
    );
    // Native XLSB keeps drawing/VML mirrors but does not carry x14 ctrlProps.
    assert!(
        members(&bytes)
            .keys()
            .all(|name| !name.starts_with("xl/ctrlProps/") && !name.starts_with("xl/activeX/"))
    );
    let mut edit = package.edit_cell_values("Sheet1").unwrap();
    edit.set_number(litchi_xlsb::cell_values::Reference::new(4, 1).unwrap(), 9.0)
        .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    assert_preserved_except(&bytes, &output, &["xl/worksheets/sheet1.bin"]);
    let reopened = Package::from_bytes(&output).unwrap();
    assert!(commit.patch().before() == bytes);
    assert!(commit.patch().after() == output);
    let after = reopened.workbook().unwrap().drawing_elements(0).unwrap();
    assert_eq!(after, elements);
    let workbook = Workbook::new(Cursor::new(output)).unwrap();
    assert_eq!(
        cached_number(workbook.worksheet(0).unwrap().get_cell(4, 1).unwrap()),
        9.0
    );
    assert!(
        reopened
            .apply_cell_values(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(package.to_bytes().unwrap() == bytes);
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn assert_preserved_except(before: &[u8], after: &[u8], changed: &[&str]) {
    let before = members(before);
    let after = members(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, bytes) in &before {
        if !changed.contains(&name.as_str()) {
            assert!(after[name] == *bytes, "unexpected change to {name}");
        }
    }
}

fn outside_collection(
    bytes: &[u8],
    begin: litchi_xlsb::raw::Kind,
    end: litchi_xlsb::raw::Kind,
) -> Vec<(u16, Vec<u8>)> {
    let mut inside = false;
    let mut other = Vec::new();
    for record in litchi_xlsb::raw::Records::new(bytes, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = record.unwrap();
        if record.kind() == begin {
            assert!(!inside);
            inside = true;
        } else if record.kind() == end {
            assert!(inside);
            inside = false;
        } else if !inside {
            other.push((record.kind().get(), record.payload().to_vec()));
        }
    }
    assert!(!inside);
    other
}

fn cached_number(cell: &litchi_xlsb::package::Cell) -> f64 {
    match cell.value() {
        CellValue::Float(value) => *value,
        CellValue::Int(value) => *value as f64,
        CellValue::Formula {
            cached_value: Some(value),
            ..
        } => match value.as_ref() {
            CellValue::Float(value) => *value,
            CellValue::Int(value) => *value as f64,
            value => panic!("expected a numeric cache, got {value:?}"),
        },
        value => panic!("expected stored numeric data, got {value:?}"),
    }
}

#[test]
fn native_slicer_caches_keep_identity_bindings_and_selected_items() {
    use litchi_xlsb::slicer::{self, CrossFilter, SortOrder, Source};
    let bytes = native("excel-pivot-controls.xlsb");
    let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
    let caches = workbook.slicer_caches().unwrap();
    assert_eq!(caches.len(), 2);
    let parts = members(&bytes);
    for (index, (cache, (name, hierarchy, selected))) in caches
        .iter()
        .zip([
            ("Slicer_Region", "Region", [(0, true), (1, false)]),
            ("Slicer_Product", "Product", [(1, true), (0, true)]),
        ])
        .enumerate()
    {
        assert_eq!(cache.name, name);
        assert_eq!(cache.hierarchy, hierarchy);
        assert_eq!(cache.pivot_tables.len(), 1);
        assert_eq!(cache.pivot_tables[0].sheet_id, 2);
        assert_eq!(cache.pivot_tables[0].name, "PivotTable1");
        let Source::Native(source) = &cache.source else {
            panic!("expected native non-OLAP cache");
        };
        assert_eq!(source.cache_id, 2018148052);
        assert_eq!(source.sort_order, SortOrder::Ascending);
        assert_eq!(source.cross_filter, CrossFilter::ShowItemsWithoutData);
        assert!(source.sort_using_custom_lists);
        assert!(source.show_all_items);
        assert_eq!(source.items.len(), 2);
        for (item, (id, selected)) in source.items.iter().zip(selected) {
            assert_eq!(item.cache_index, id);
            assert_eq!(item.selected, selected);
            assert!(!item.no_data);
        }
        let encoded = slicer::write_cache(cache, litchi_xlsb::raw::Limits::DEFAULT).unwrap();
        assert_eq!(
            encoded,
            parts[&format!("xl/slicerCaches/slicerCache{}.bin", index + 1)]
        );
    }
}

#[test]
fn native_slicer_caption_edit_preserves_caches_identities_and_other_members() {
    let bytes = native("excel-pivot-controls.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit_slicers().unwrap();
    let mut views = edit.views(1).unwrap();
    assert_eq!(
        views
            .items
            .iter()
            .map(|v| v.name.as_str())
            .collect::<Vec<_>>(),
        ["Region", "Product"]
    );
    assert!(!edit.set_views(1, views.clone()).unwrap());
    views.items[0].caption = Some("Regional sales reviewed".into());
    assert!(edit.set_views(1, views.clone()).unwrap());
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(reopened.edit_slicers().unwrap().views(1).unwrap(), views);
    assert_preserved_except(&before, &after, &["xl/slicers/slicer1.bin"]);
    let undone = reopened
        .apply_slicers(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&undone.package().to_bytes().unwrap()),
        members(&before)
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn native_pivot_cache_records_tables_and_view_bindings_keep_saved_values() {
    use litchi_xlsb::package::pivot::{self, PivotCacheItemValue as Value, PivotCacheSourceType};
    for (file, cache_id, owner, total_row, total_column, total) in [
        ("excel-visual-analytics.xlsb", 5, 0, 6, 3, 175.5),
        ("excel-pivot-controls.xlsb", 1, 1, 5, 2, 9.5),
    ] {
        let bytes = native(file);
        let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
        assert_eq!(workbook.pivot_views().len(), 1);
        let view = &workbook.pivot_views()[0];
        assert_eq!(
            (view.name(), view.cache_id(), view.sheet_index()),
            ("PivotTable1", cache_id, owner)
        );
        assert_eq!(workbook.pivot_cache_definitions().len(), 1);
        let cache = workbook.pivot_cache_definition(cache_id).unwrap();
        assert_eq!(cache.record_count, 8);
        assert!(cache.save_data);
        assert!(!cache.invalid);
        let source = cache.source.as_ref().unwrap();
        assert_eq!(source.source_type, PivotCacheSourceType::Worksheet);
        assert_eq!(
            source.worksheet.as_ref().unwrap().named_range.as_deref(),
            Some("Table1")
        );
        assert_eq!(
            cache
                .fields
                .iter()
                .map(|f| f.name.as_str())
                .collect::<Vec<_>>(),
            ["Date", "Region", "Product", "Revenue", "Units", "Status"]
        );
        let parts = members(&bytes);
        let records = pivot::parse_pivot_cache_records(
            &parts["xl/pivotCache/pivotCacheRecords1.bin"],
            cache,
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(records.record_count, 8);
        assert_eq!(records.records.len(), 8);
        for (record, (month, day, region, product, revenue, units, status)) in
            records.records.iter().zip([
                (1, 15, "North", "Widget", 12.5, 3.0, "Open"),
                (1, 20, "South", "Widget", 25.0, 5.0, "Closed"),
                (2, 10, "North", "Gadget", 9.5, 2.0, "Open"),
                (2, 18, "South", "Gadget", 48.0, 4.0, "Hold"),
                (3, 3, "North", "Widget", 16.5, 3.0, "Open"),
                (3, 21, "South", "Gadget", 20.0, 4.0, "Closed"),
                (4, 2, "North", "Gadget", 14.0, 2.0, "Hold"),
                (4, 18, "South", "Widget", 30.0, 6.0, "Open"),
            ])
        {
            assert_eq!(record.values.len(), 6);
            let values = record
                .values
                .iter()
                .enumerate()
                .map(|(field, value)| match value {
                    Value::Index(index) => {
                        &cache.fields[field].shared_items.items[*index as usize].value
                    },
                    value => value,
                })
                .collect::<Vec<_>>();
            let Value::DateTime(date) = values[0] else {
                panic!("expected stored date");
            };
            assert_eq!(
                (
                    date.year,
                    date.month,
                    date.day,
                    date.hour,
                    date.minute,
                    date.second
                ),
                (2026, month, day, 0, 0, 0)
            );
            assert_eq!(values[1], &Value::String(region.into()));
            assert_eq!(values[2], &Value::String(product.into()));
            assert_eq!(values[3], &Value::Number(revenue));
            assert_eq!(values[4], &Value::Number(units));
            assert_eq!(values[5], &Value::String(status.into()));
        }
        assert_eq!(workbook.structured_tables().len(), 1);
        let (table_owner, table) = &workbook.structured_tables()[0];
        assert_eq!(*table_owner, 1);
        assert_eq!(table.display_name.as_deref(), Some("Table1"));
        assert_eq!(
            (
                table.range.first_row,
                table.range.last_row,
                table.range.first_column,
                table.range.last_column
            ),
            (0, 8, 0, 5)
        );
        assert_eq!(table.header_row_count, 1);
        assert!(table.columns.iter().all(|c| c.name.is_none()));
        assert_eq!(
            table
                .columns
                .iter()
                .map(|c| c.caption.as_deref().unwrap())
                .collect::<Vec<_>>(),
            ["Date", "Region", "Product", "Revenue", "Units", "Status"]
        );
        let pivot_sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            cached_number(pivot_sheet.get_cell(total_row, total_column).unwrap()),
            total
        );
    }
}

#[test]
fn native_notes_read_and_edit_keep_other_note_vml_and_package_payloads() {
    let bytes = native("excel-interaction-print.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let before = source.to_bytes().unwrap();
    let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
    let notes = workbook.sheet_comments(0).unwrap();
    assert_eq!(notes.len(), 2);
    for (note, (row, text)) in notes.iter().zip([
        (1, "Ryker Zhu:\nLEGACY_NOTE_002: native Excel cell note."),
        (
            2,
            "Ryker Zhu:\nCOMMENT_COMMAND_002: comment inserted through Review ribbon.",
        ),
    ]) {
        assert_eq!((note.row, note.col), (row, 0));
        assert_eq!(note.author, "Ryker Zhu");
        assert_eq!(note.text, text);
        assert_eq!(
            note.runs
                .iter()
                .map(|r| (r.character_index, r.font_id))
                .collect::<Vec<_>>(),
            [(0, 3), (10, 2), (11, 2)]
        );
    }
    assert_eq!(workbook.worksheet(0).unwrap().comments(), notes);
    let lazy = litchi_xlsb::SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
    assert_eq!(
        lazy.sheet(0_usize).unwrap().unwrap().comments().unwrap(),
        notes
    );
    let mut edit = source.edit_comments().unwrap();
    assert!(!edit.set(0, 1, 0, &notes[0].author, &notes[0].text).unwrap());
    assert!(
        edit.set(0, 1, 0, &notes[0].author, "Reviewed native note")
            .unwrap()
    );
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let updated = reopened.edit_comments().unwrap().comments(0).unwrap();
    let mut expected = notes[0].clone();
    expected.text = "Reviewed native note".into();
    // CommentText requires a first formatting run; plain replacement uses
    // the default font and drops the source's mixed formatting.
    expected.runs = vec![litchi_xlsb::comments::Run {
        character_index: 0,
        font_id: 0,
    }];
    assert_eq!(updated, [expected, notes[1].clone()]);
    assert_preserved_except(&before, &after, &["xl/comments1.bin"]);
    let undone = reopened
        .apply_comments(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&undone.package().to_bytes().unwrap()),
        members(&before)
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn native_1904_layout_and_page_break_edit_preserve_serials_and_other_parts() {
    use litchi_xlsb::page_breaks::Break;
    let bytes = native("excel-date1904-breaks.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
    assert!(workbook.is_1904_date_system());
    let sheet = workbook.worksheet(0).unwrap();
    for (row, value) in [
        (2, 0.0),
        (3, 1.0),
        (4, 43889.0),
        (5, 44848.5515625),
        (6, -1.0),
        (7, 0.2708333333333333),
    ] {
        assert_eq!(cached_number(sheet.get_cell(row, 1).unwrap()), value);
    }
    assert_eq!(
        sheet
            .row_infos()
            .iter()
            .filter(|r| r.hidden)
            .map(|r| r.row)
            .collect::<Vec<_>>(),
        [13, 14]
    );
    let hidden = sheet
        .column_infos()
        .iter()
        .filter(|c| c.hidden)
        .collect::<Vec<_>>();
    assert_eq!(hidden.len(), 1);
    assert_eq!(
        (
            hidden[0].first_column,
            hidden[0].last_column,
            hidden[0].width
        ),
        (9, 9, 0.0)
    );
    let wide = sheet
        .column_infos()
        .iter()
        .find(|c| c.first_column == 0)
        .unwrap();
    assert_eq!((wide.last_column, wide.width), (1, 32.83203125));
    let mut edit = source.edit_page_breaks().unwrap();
    let original = edit.page_breaks(0).unwrap();
    assert_eq!(original.rows, [Break::row(19)]);
    assert_eq!(original.columns, [Break::column(6)]);
    assert!(!edit.set_page_breaks(0, &original).unwrap());
    let mut expected = original.clone();
    expected.rows.push(Break::row(30));
    assert!(edit.set_page_breaks(0, &expected).unwrap());
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened.edit_page_breaks().unwrap().page_breaks(0).unwrap(),
        expected
    );
    assert_preserved_except(&bytes, &after, &["xl/worksheets/sheet1.bin"]);
    let original_parts = members(&bytes);
    let updated_parts = members(&after);
    assert_eq!(
        outside_collection(
            &original_parts["xl/worksheets/sheet1.bin"],
            litchi_xlsb::raw::kind::BEGIN_RW_BRK,
            litchi_xlsb::raw::kind::END_RW_BRK
        ),
        outside_collection(
            &updated_parts["xl/worksheets/sheet1.bin"],
            litchi_xlsb::raw::kind::BEGIN_RW_BRK,
            litchi_xlsb::raw::kind::END_RW_BRK
        ),
    );
    let undone = reopened
        .apply_page_breaks(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&undone.package().to_bytes().unwrap()),
        members(&bytes)
    );
}

#[test]
fn native_array_shared_and_dynamic_groups_keep_ranges_tokens_and_all_caches() {
    use litchi_xlsb::formula::{Group, GroupKind};
    use litchi_xlsb::raw::{Limits, Records, kind};
    let bytes = native("excel-formula-owners.xlsb");
    let parts = members(&bytes);
    let mut definitions = std::collections::BTreeMap::new();
    for (part, expected) in [
        (
            "xl/worksheets/sheet1.bin",
            vec![
                (GroupKind::Array, "C4:C7"),
                (GroupKind::Shared, "D5:D68"),
                (GroupKind::Shared, "D69:D103"),
            ],
        ),
        (
            "xl/worksheets/sheet2.bin",
            vec![(GroupKind::Array, "A4:B6"), (GroupKind::Array, "D4:E5")],
        ),
    ] {
        let mut groups = Vec::new();
        for record in Records::new(&parts[part], Limits::DEFAULT) {
            let record = record.unwrap();
            let group = if record.kind() == kind::ARR_FMLA {
                Some(Group::parse_array(record.payload(), Limits::DEFAULT).unwrap())
            } else if record.kind() == kind::SHR_FMLA {
                Some(Group::parse_shared(record.payload(), Limits::DEFAULT).unwrap())
            } else {
                None
            };
            if let Some(group) = group {
                assert!(!group.formula.rgce.is_empty());
                assert_eq!(
                    group.to_record_data(Limits::DEFAULT).unwrap(),
                    record.payload()
                );
                definitions.insert(group.range.to_a1(), group.formula.clone());
                groups.push((group.kind, group.range.to_a1()));
            }
        }
        assert_eq!(
            groups
                .iter()
                .map(|(kind, range)| (*kind, range.as_str()))
                .collect::<Vec<_>>(),
            expected
        );
    }
    let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
    let first = workbook.worksheet(0).unwrap();
    for (row, expected) in (3..7).zip([20.0, 60.0, 120.0, 200.0]) {
        let cell = first.get_cell(row, 2).unwrap();
        assert_eq!(cached_number(cell), expected);
        let CellValue::Formula {
            formula,
            is_array,
            array_range,
            ..
        } = cell.value()
        else {
            panic!("expected CSE formula");
        };
        assert_eq!(formula, "(A4:A7*B4:B7)");
        assert!(*is_array);
        assert_eq!(array_range.as_deref(), Some("C4:C7"));
    }
    for row in 3..103 {
        let cell = first.get_cell(row, 3).unwrap();
        assert_eq!(cached_number(cell), f64::from((row + 1) * 2));
        let CellValue::Formula { formula, .. } = cell.value() else {
            panic!("expected copied/shared formula");
        };
        assert_eq!(formula, "(ROW()*2)");
    }
    let dynamic = workbook.worksheet(1).unwrap();
    for (row, col, expected) in [
        (3, 0, 10.0),
        (3, 1, 12.0),
        (4, 0, 14.0),
        (4, 1, 16.0),
        (5, 0, 18.0),
        (5, 1, 20.0),
        (3, 3, 4.0),
        (3, 4, 30.0),
        (4, 3, 5.0),
        (4, 4, 40.0),
        (3, 5, 50.0),
    ] {
        let cell = dynamic.get_cell(row, col).unwrap();
        assert!(cell.is_formula());
        assert_eq!(
            cell.formula_resolution_status(),
            Some(litchi_xlsb::FormulaResolutionStatus::Opaque(
                litchi_xlsb::FormulaOpacityReason::Unsupported
            ))
        );
        assert_eq!(cached_number(cell), expected);
        assert_eq!(cell.cached_value(), Some(&CellValue::Float(expected)));
        assert!(cell.formula_bytes().is_some());
        if col == 5 {
            assert_eq!(cell.formula_range(), None);
        } else {
            let range = if col < 2 { "A4:B6" } else { "D4:E5" };
            assert_eq!(cell.formula_range().as_deref(), Some(range));
            assert_eq!(
                cell.formula_definition_bytes(),
                Some(definitions[range].rgce.as_slice())
            );
            assert_eq!(
                cell.formula_definition_extra_bytes(),
                Some(definitions[range].rgcb.as_slice())
            );
        }
    }
    // What-if results are cached scalars; reading them never applies inputs.
    for (row, expected) in (3..7).zip([7.0, 9.0, 11.0, 13.0]) {
        assert_eq!(cached_number(first.get_cell(row, 9).unwrap()), expected);
    }
    for (row, values) in (3..7).zip([
        [15.0, 25.0, 35.0, 45.0],
        [25.0, 45.0, 65.0, 85.0],
        [35.0, 65.0, 95.0, 125.0],
        [45.0, 85.0, 125.0, 165.0],
    ]) {
        for (column, value) in (12..16).zip(values) {
            assert_eq!(cached_number(first.get_cell(row, column).unwrap()), value);
        }
    }
}

#[test]
fn native_scenarios_retain_inputs_and_metadata_edits_do_not_apply_them() {
    let bytes = native("excel-formula-owners.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(&bytes)).unwrap();
    let manager = workbook.scenarios(0).unwrap().unwrap();
    assert_eq!(manager.current(), Some(0));
    assert_eq!(manager.shown(), None);
    assert_eq!(manager.scenarios().len(), 2);
    for (scenario, (name, comment, values)) in manager.scenarios().iter().zip([
        ("BASELINE_009", "Native scenario: X=10, Y=2.", ["10", "2"]),
        ("ALTERNATE_009", "Native scenario: X=25, Y=4.", ["25", "4"]),
    ]) {
        assert_eq!(scenario.name(), name);
        assert_eq!(scenario.comment(), comment);
        assert!(scenario.locked());
        assert!(!scenario.hidden());
        assert_eq!(scenario.changed_cells().len(), 2);
        for (cell, (row, value)) in scenario
            .changed_cells()
            .iter()
            .zip([3, 4].into_iter().zip(values))
        {
            assert_eq!((cell.row(), cell.column(), cell.value()), (row, 7, value));
        }
    }
    let mut expected = manager.clone();
    let mut second = manager.scenarios()[1].clone();
    second.set_comment("Reviewed alternate scenario").unwrap();
    expected.set_scenario(1, second).unwrap();
    let mut edit = source.edit_scenarios().unwrap();
    assert!(!edit.set_scenarios(0, &manager).unwrap());
    assert!(edit.set_scenarios(0, &expected).unwrap());
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened.edit_scenarios().unwrap().scenarios(0).unwrap(),
        Some(expected)
    );
    let readback = Workbook::new(Cursor::new(&after)).unwrap();
    let sheet = readback.worksheet(0).unwrap();
    for (row, value) in [(3, 10.0), (4, 2.0), (5, 25.0)] {
        assert_eq!(cached_number(sheet.get_cell(row, 7).unwrap()), value);
    }
    assert_preserved_except(&bytes, &after, &["xl/worksheets/sheet1.bin"]);
    let original_parts = members(&bytes);
    let updated_parts = members(&after);
    assert_eq!(
        outside_collection(
            &original_parts["xl/worksheets/sheet1.bin"],
            litchi_xlsb::raw::kind::BEGIN_SCEN_MAN,
            litchi_xlsb::raw::kind::END_SCEN_MAN
        ),
        outside_collection(
            &updated_parts["xl/worksheets/sheet1.bin"],
            litchi_xlsb::raw::kind::BEGIN_SCEN_MAN,
            litchi_xlsb::raw::kind::END_SCEN_MAN
        ),
    );
    let undone = reopened
        .apply_scenarios(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&undone.package().to_bytes().unwrap()),
        members(&bytes)
    );
}

#[test]
fn native_slicer_cache_identity_refuses_malformed_envelopes_and_limits() {
    use litchi_xlsb::{
        raw::{Encoder, Kind, Limits, Records, kind},
        slicer,
    };
    let bytes = native("excel-pivot-controls.xlsb");
    let parts = members(&bytes);
    let original = &parts["xl/slicerCaches/slicerCache1.bin"];
    let records: Vec<_> = Records::new(original, Limits::DEFAULT)
        .map(|r| {
            let r = r.unwrap();
            (r.kind(), r.payload().to_vec())
        })
        .collect();
    let encode = |records: &[(Kind, Vec<u8>)]| {
        let mut bytes = Vec::new();
        let mut encoder = Encoder::new(&mut bytes, Limits::DEFAULT);
        for (kind, payload) in records {
            encoder.write_record(*kind, payload).unwrap();
        }
        bytes
    };
    for case in 0..7 {
        let mut damaged = records.clone();
        match case {
            0 => {
                damaged[1].1[0] = 2;
            }, // Product count disagrees with payload.
            1 => {
                damaged[2].1.pop();
            }, // Truncated UID.
            2 => {
                damaged.insert(3, damaged[2].clone());
            }, // Duplicate UID.
            3 => {
                damaged.remove(3);
            }, // Missing AC end.
            4 => {
                damaged[2].0 = kind::BEGIN_SLICER_CACHE_DEF;
            },
            5 => {
                damaged[3].1.push(0);
            }, // Nonempty AC end.
            6 => {
                damaged.splice(4..4, records[1..4].iter().cloned());
            },
            _ => unreachable!(),
        }
        assert!(
            slicer::parse_cache(&encode(&damaged), Limits::DEFAULT).is_err(),
            "case {case}"
        );
    }
    let mut parsed = slicer::parse_cache(original, Limits::DEFAULT).unwrap();
    let prefix = parsed.alternate_content.clone();
    parsed.alternate_content.extend_from_slice(&prefix);
    assert!(slicer::write_cache(&parsed, Limits::DEFAULT).is_err());
    assert!(slicer::parse_cache(original, Limits::new(15, 32767)).is_err());
}

#[test]
fn native_slicer_cache_metadata_edit_keeps_parts_relationships_and_inverse() {
    let bytes = native("excel-pivot-controls.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit_slicers().unwrap();
    let mut expected = edit.caches().unwrap();
    assert!(!edit.set_caches(expected.clone()).unwrap());
    let litchi_xlsb::slicer::Source::Native(state) = &mut expected[0].source else {
        panic!("expected native cache");
    };
    state.sort_order = litchi_xlsb::slicer::SortOrder::Descending;
    assert!(edit.set_caches(expected.clone()).unwrap());
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(reopened.edit_slicers().unwrap().caches().unwrap(), expected);
    assert_preserved_except(&bytes, &after, &["xl/slicerCaches/slicerCache1.bin"]);
    let undone = reopened
        .apply_slicers(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&undone.package().to_bytes().unwrap()),
        members(&bytes)
    );
}

#[test]
fn native_slicer_cache_reorder_and_add_remove_handle_sparse_relationship_ids() {
    let bytes = native("excel-pivot-controls.xlsb");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit_slicers().unwrap();
    let original = edit.caches().unwrap();
    let mut reversed = original.clone();
    reversed.reverse();
    assert!(edit.set_caches(reversed.clone()).unwrap());
    let committed = edit.commit().unwrap();
    let reordered = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&reordered).unwrap();
    assert_eq!(reopened.edit_slicers().unwrap().caches().unwrap(), reversed);
    assert_preserved_except(&bytes, &reordered, &["xl/workbook.bin"]);
    let restored = reopened
        .apply_slicers(&committed.patch().inverse())
        .unwrap();
    assert_eq!(
        members(&restored.package().to_bytes().unwrap()),
        members(&bytes)
    );

    let mut edit = source.edit_slicers().unwrap();
    let mut extra = original[0].clone();
    extra.name = "Additional_Region".into();
    extra.alternate_content.clear(); // A new cache must not reuse a native UID.
    let mut expected = original.clone();
    expected.push(extra);
    assert!(edit.set_caches(expected.clone()).unwrap());
    let added = edit.commit().unwrap();
    let reopened = Package::from_bytes(&added.package().to_bytes().unwrap()).unwrap();
    let mut edit = reopened.edit_slicers().unwrap();
    assert_eq!(edit.caches().unwrap(), expected);
    assert_eq!(
        edit.remove_cache("Additional_Region").unwrap(),
        expected.pop()
    );
    let removed = edit.commit().unwrap();
    assert_eq!(
        removed.package().edit_slicers().unwrap().caches().unwrap(),
        original
    );
    let inverse = reopened.apply_slicers(&added.patch().inverse()).unwrap();
    assert_eq!(
        members(&inverse.package().to_bytes().unwrap()),
        members(&bytes)
    );
}

#[test]
fn native_binary_index_binds_metadata_prefix_to_actual_formula_cell_only() {
    use litchi_xlsb::binary_index::{BinaryIndex, Limits};
    let bytes = native("excel-formula-owners.xlsb");
    let parts = members(&bytes);
    let index = &parts["xl/worksheets/binaryIndex2.bin"];
    let worksheet = &parts["xl/worksheets/sheet2.bin"];
    let parsed = BinaryIndex::parse(index, Limits::default()).unwrap();
    let source_anchor = parsed.blocks()[0]
        .rows
        .columns
        .iter()
        .find(|c| c.row == 3)
        .unwrap();
    assert_eq!(source_anchor.offset, 358); // Native BrtCellMeta, six bytes before A4.
    let bound = parsed.bind(worksheet, Limits::default()).unwrap();
    let anchor = bound.start_offset(3, 0).unwrap().unwrap();
    assert_eq!(
        (anchor.row, anchor.column, anchor.offset, anchor.record_kind),
        (3, 0, 364, litchi_xlsb::raw::kind::FMLA_NUM.get())
    );
    let lazy = litchi_xlsb::SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
    let values = lazy
        .sheet(1_usize)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap();
    for (row, column, number) in [
        (3, 0, 10.0),
        (3, 1, 12.0),
        (3, 3, 4.0),
        (3, 5, 50.0),
        (5, 1, 20.0),
    ] {
        assert_eq!(
            values.cached_value(row, column).unwrap(),
            Some(CellValue::Float(number))
        );
    }
    // The third native row's sub-offset is 199 + the 159-byte block base.
    assert_eq!(&index[58..62], &199_u32.to_le_bytes());
    for wrong in [198_u32, 200, 201, 206, 291] {
        let mut damaged = index.clone();
        damaged[58..62].copy_from_slice(&wrong.to_le_bytes());
        assert!(
            BinaryIndex::parse(&damaged, Limits::default())
                .unwrap()
                .bind(worksheet, Limits::default())
                .is_err(),
            "accepted sub-offset {wrong}"
        );
    }
    let mut wrong_prefix = worksheet.clone();
    assert_eq!(wrong_prefix[358], 0x31);
    wrong_prefix[358] = 0x33; // A different record cannot stand in for cell metadata.
    assert!(parsed.bind(&wrong_prefix, Limits::default()).is_err());
}

#[test]
fn native_drawing_picture_and_rectangle_keep_resolved_payload_and_anchors() {
    use litchi_xlsb::package::drawing::AnchorKind;
    let bytes = native("excel-conditional-drawings.xlsb");
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.drawings_parsed());
    let drawing = workbook.sheet_drawing(0).unwrap();
    assert_eq!(drawing.drawing.anchors.len(), 2);
    assert_eq!(drawing.images.len(), 1);
    assert!(drawing.charts.is_empty());
    let picture = &drawing.images[0];
    assert_eq!(picture.picture_name, "Picture 2");
    assert_eq!(picture.rel_id, "rId1");
    assert_eq!(
        picture.description.as_deref(),
        Some("CF_PICTURE_005: synthetic blue and yellow stripes anchored over worksheet cells.")
    );
    use sha2::{Digest, Sha256};
    assert_eq!(
        Sha256::digest(picture.data.as_ref())
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>(),
        "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8"
    );
    for (anchor, expected) in drawing
        .drawing
        .anchors
        .iter()
        .zip([(0, 11, 1, 16), (5, 15, 7, 19)])
    {
        let AnchorKind::TwoCell { from, to, .. } = &anchor.anchor else {
            panic!("native two-cell anchor")
        };
        assert_eq!((from.column, from.row, to.column, to.row), expected);
    }
    assert_eq!(drawing.shapes.len(), 1);
    let litchi_xlsb::shapes::Object::Shape(shape) = &drawing.shapes[0].object else {
        panic!("rounded rectangle")
    };
    assert_eq!(
        shape.non_visual.name.as_deref(),
        Some("Rounded Rectangle 3")
    );
    assert_eq!(shape.preset(), Some(litchi_xlsb::shapes::Preset::RoundRect));
    assert_eq!(shape.text_body.as_ref().unwrap().text(), "CF_SHAPE_005");
}
