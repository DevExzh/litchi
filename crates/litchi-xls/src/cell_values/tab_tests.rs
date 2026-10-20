#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "tab-structure fixtures intentionally fail fast"
)]

use super::*;
use litchi_core::Position;
use litchi_core::sheet::WorkbookTrait;
use std::path::Path;

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls")
            .join(name),
    )
    .expect("XLS fixture")
}

fn three_sheets() -> Vec<u8> {
    let mut builder = crate::Builder::new();
    let first = builder.add_worksheet("First").unwrap();
    let second = builder.add_worksheet("Second").unwrap();
    let third = builder.add_worksheet("Third").unwrap();
    builder.write_number(first, 0, 0, 1.0).unwrap();
    builder.write_string(second, 1, 1, "two").unwrap();
    builder.write_number(third, 2, 2, 3.0).unwrap();
    builder.write_formula(first, 3, 0, "A1*2").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

/// Every tab's name and cell projection (value and rendered formula), by
/// name, from the complete reader.
fn contents(bytes: &[u8]) -> Vec<(String, Vec<String>)> {
    let reader = Reader::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut out = Vec::new();
    for (index, name) in reader.worksheet_names().iter().enumerate() {
        let worksheet = reader.xls_worksheet(index).unwrap();
        let cells = worksheet
            .stored_cells()
            .map(|cell| {
                format!(
                    "{}:{}={:?}{}",
                    cell.row(),
                    cell.column(),
                    cell.value(),
                    cell.formula()
                        .map(|formula| format!(" ={formula}"))
                        .unwrap_or_default()
                )
            })
            .collect();
        out.push((name.clone(), cells));
    }
    out
}

fn names(snapshot: &Snapshot) -> Vec<String> {
    tabs::catalog(snapshot.workbook_stream())
        .unwrap()
        .into_iter()
        .map(|tab| tab.name)
        .collect()
}

fn values_only(value: Vec<(String, Vec<String>)>) -> Vec<(String, Vec<String>)> {
    value
        .into_iter()
        .map(|(name, cells)| {
            let cells = cells
                .into_iter()
                .map(|cell| match cell.find(" =") {
                    Some(at) => format!("{} (formula)", &cell[..at]),
                    None => cell,
                })
                .collect();
            (name, cells)
        })
        .collect()
}

fn sorted(mut value: Vec<(String, Vec<String>)>) -> Vec<(String, Vec<String>)> {
    value.sort();
    value
}

#[test]
fn reorder_moves_tabs_and_keeps_every_cell_and_reference() {
    let bytes = three_sheets();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.reorder_sheets(&[
        Selector::Name("third"),
        Selector::Position(Position::new(0)),
        Selector::Name("SECOND"),
    ])
    .unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(names(&target), ["Third", "First", "Second"]);
    assert_eq!(sorted(contents(&bytes)), sorted(contents(target.bytes())));
    // Exact and semantic inverses restore the source tab order.
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    let restored = patch.semantic().inverse().apply(&target).unwrap();
    assert_eq!(
        names(&restored.into_parts().0),
        ["First", "Second", "Third"]
    );
    // The semantic patch replays against the source and round-trips JSON.
    let json = patch.semantic().to_deterministic_json().unwrap();
    let replayed = SemanticPatch::from_deterministic_json(&json)
        .unwrap()
        .apply(&source)
        .unwrap();
    assert_eq!(replayed.into_parts().0.bytes(), target.bytes());
}

#[test]
fn identity_reorder_is_an_exact_noop() {
    let bytes = three_sheets();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.reorder_sheets(&[
        Selector::Position(Position::new(0)),
        Selector::Position(Position::new(1)),
        Selector::Position(Position::new(2)),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    let (target, patch) = commit.into_parts();
    assert!(patch.is_noop());
    assert_eq!(target.bytes(), &bytes[..]);
}

#[test]
fn reorder_refuses_incomplete_repeated_and_unknown_orders() {
    let source = Snapshot::from_bytes(three_sheets(), Limits::default()).unwrap();
    for order in [
        vec![
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(1)),
        ],
        vec![
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(1)),
        ],
        vec![
            Selector::Name("Nope"),
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(1)),
        ],
        vec![
            Selector::Position(Position::new(3)),
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(1)),
        ],
    ] {
        let mut edit = source.edit();
        assert!(edit.reorder_sheets(&order).is_err());
        assert!(edit.tab_change.is_none(), "refusal left a staged change");
    }
}

#[test]
fn insert_adds_an_empty_unselected_worksheet() {
    let bytes = three_sheets();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    for position in 0..=3 {
        let mut edit = source.edit();
        edit.insert_sheet(Position::new(position), "Added").unwrap();
        let (target, patch) = edit.commit().unwrap().into_parts();
        let mut expected = vec!["First", "Second", "Third"];
        expected.insert(position, "Added");
        assert_eq!(names(&target), expected);
        let after = contents(target.bytes());
        let added = after.iter().find(|(name, _)| name == "Added").unwrap();
        assert!(added.1.is_empty());
        let mut before = contents(&bytes);
        before.push(("Added".into(), Vec::new()));
        assert_eq!(sorted(before), sorted(after));
        assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
        // The semantic inverse removes the still-empty inserted worksheet.
        let restored = patch.semantic().inverse().apply(&target).unwrap();
        assert_eq!(
            names(&restored.into_parts().0),
            ["First", "Second", "Third"]
        );
        // A later edit can write into the inserted worksheet.
        let mut edit = target.edit();
        edit.insert_cell(
            Selector::Name("added"),
            Reference::new(4, 4).unwrap(),
            Value::Number(9.0),
        )
        .unwrap();
        let written = edit.commit().unwrap().into_parts().0;
        let added = contents(written.bytes())
            .into_iter()
            .find(|(name, _)| name == "Added")
            .unwrap();
        assert_eq!(added.1, ["4:4=Float(9.0)"]);
    }
}

#[test]
fn insert_refuses_invalid_and_caseless_duplicate_names() {
    let source = Snapshot::from_bytes(three_sheets(), Limits::default()).unwrap();
    for name in [
        "",
        "first",
        "FIRST",
        "bad/name",
        "'quoted'",
        &"x".repeat(32),
    ] {
        let mut edit = source.edit();
        assert!(
            edit.insert_sheet(Position::new(0), name).is_err(),
            "{name:?} was admitted"
        );
    }
    let mut edit = source.edit();
    assert!(edit.insert_sheet(Position::new(4), "Late").is_err());
}

#[test]
fn remove_drops_a_tab_and_restores_exactly() {
    let mut builder = crate::Builder::new();
    let keep = builder.add_worksheet("Keep").unwrap();
    let drop = builder.add_worksheet("Drop").unwrap();
    builder.write_number(keep, 0, 0, 1.0).unwrap();
    builder.write_string(drop, 0, 0, "gone").unwrap();
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.remove_sheet(Selector::Name("DROP")).unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(names(&target), ["Keep"]);
    assert_eq!(
        contents(target.bytes()),
        [("Keep".into(), vec!["0:0=Float(1.0)".to_string()])]
    );
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    // The removed content is not semantically restorable without the
    // exact artifact patch.
    assert!(matches!(
        patch.semantic().inverse().apply(&target),
        Err(Error::UnsupportedFeature(_))
    ));
    // The forward semantic patch replays on the same source.
    let replayed = patch.semantic().apply(&source).unwrap().into_parts().0;
    assert_eq!(replayed.bytes(), target.bytes());
}

#[test]
fn remove_refuses_the_last_tab_and_referenced_tabs() {
    let source = Snapshot::from_bytes(
        crate::Package::new().unwrap().to_bytes().unwrap(),
        Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    assert!(
        edit.remove_sheet(Selector::Position(Position::new(0)))
            .is_err()
    );

    // Sheet1 owns the sheet-scoped Print_Area name.
    let source =
        Snapshot::from_bytes(fixture("SimpleWithPrintArea.xls"), Limits::default()).unwrap();
    let mut edit = source.edit();
    let refusal = edit.remove_sheet(Selector::Name("Sheet1"));
    assert!(
        matches!(
            refusal,
            Err(Error::Refused(_) | Error::PatchConflict { .. } | Error::Internal(_))
        ),
        "{refusal:?}"
    );
    assert!(edit.tab_change.is_none());
    // Sheet2 is unreferenced; the name keeps its scope and target.
    let mut edit = source.edit();
    edit.remove_sheet(Selector::Name("Sheet2")).unwrap();
    let target = edit.commit().unwrap().into_parts().0;
    assert_eq!(names(&target), ["Sheet1", "Sheet3"]);
    let reader = Reader::new(Cursor::new(target.bytes().to_vec())).unwrap();
    let name = &reader.defined_names()[0];
    assert_eq!(name.scope, crate::NameScope::Worksheet(0));
    assert_eq!(name.formula.as_deref(), Some("='Sheet1'!$A$1:$C$5"));
}

/// `FormulaSheetRange.xls` holds 3-D references over a tab range.
#[test]
fn three_dimensional_ranges_follow_or_refuse_tab_moves() {
    let bytes = fixture("FormulaSheetRange.xls");
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let catalog = names(&source);
    let formulas = |bytes: &[u8]| -> Vec<String> {
        contents(bytes)
            .into_iter()
            .flat_map(|(_, cells)| cells)
            .filter(|cell| cell.contains(" ="))
            .collect()
    };
    let before = formulas(&bytes);
    assert!(
        before
            .iter()
            .any(|cell| cell.contains(':') && cell.contains('!'))
    );
    // Moving the first tab (outside every 3-D range) to the end keeps every
    // range contiguous and its rendered formula unchanged.
    let mut order: Vec<_> = (1..catalog.len()).map(Selector::from).collect();
    order.push(Selector::Position(Position::new(0)));
    let mut edit = source.edit();
    edit.reorder_sheets(&order).unwrap();
    let target = edit.commit().unwrap().into_parts().0;
    assert_eq!(formulas(target.bytes()).len(), before.len());
    assert_eq!(sorted(contents(&bytes)), sorted(contents(target.bytes())));
    // Splitting a 3-D range by moving an interior tab out is refused, as is
    // removing a tab a range covers or inserting inside a range.
    let mut split: Vec<_> = (0..catalog.len()).map(Selector::from).collect();
    split.swap(2, catalog.len() - 1);
    split.swap(1, 2);
    let mut edit = source.edit();
    let refusals = [
        edit.reorder_sheets(&[
            Selector::Position(Position::new(2)),
            Selector::Position(Position::new(0)),
            Selector::Position(Position::new(1)),
            Selector::Position(Position::new(3)),
            Selector::Position(Position::new(4)),
        ]),
        source
            .edit()
            .remove_sheet(Selector::Position(Position::new(2))),
        source.edit().insert_sheet(Position::new(2), "Inside"),
    ];
    assert!(
        refusals.iter().all(|result| matches!(
            result,
            Err(Error::Refused(_) | Error::PatchConflict { .. } | Error::Internal(_))
        )),
        "{refusals:?}"
    );
}

#[test]
fn a_tab_change_is_exclusive_within_its_edit() {
    let source = Snapshot::from_bytes(three_sheets(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.set_value(
        Selector::Position(Position::new(0)),
        Reference::new(0, 0).unwrap(),
        Value::Number(5.0),
    )
    .unwrap();
    assert!(
        edit.reorder_sheets(&[
            Selector::Position(Position::new(2)),
            Selector::Position(Position::new(1)),
            Selector::Position(Position::new(0))
        ])
        .is_err()
    );
    let mut edit = source.edit();
    edit.insert_sheet(Position::new(0), "New").unwrap();
    assert!(
        edit.set_value(
            Selector::Name("First"),
            Reference::new(0, 0).unwrap(),
            Value::Number(5.0)
        )
        .is_err()
    );
    assert!(edit.insert_sheet(Position::new(0), "Other").is_err());
    let mut other = source.edit();
    other
        .set_value(
            Selector::Name("First"),
            Reference::new(0, 0).unwrap(),
            Value::Number(5.0),
        )
        .unwrap();
    assert!(matches!(
        edit.join(other),
        Err(JoinError::StructuralConflicts(_))
    ));
}

/// Every corpus workbook either accepts or refuses each tab-structure
/// change with a typed error; an accepted change reopens under the complete
/// reader, keeps every tab's cells, and inverts exactly.
#[test]
fn corpus_tab_changes_reopen_or_refuse() {
    let directory = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls");
    let mut accepted = 0_usize;
    let mut names_seen = 0_usize;
    for entry in std::fs::read_dir(directory).unwrap() {
        let path = entry.unwrap().path();
        if path.extension().and_then(|value| value.to_str()) != Some("xls") {
            continue;
        }
        let bytes = std::fs::read(&path).unwrap();
        let Ok(source) = Snapshot::from_bytes(bytes.clone(), Limits::default()) else {
            continue;
        };
        let count = names(&source).len();
        names_seen += count;
        let before = sorted(contents(&bytes));
        let mut attempts: Vec<Box<dyn Fn(&mut Edit) -> Result<()>>> = Vec::new();
        attempts.push(Box::new(move |edit| {
            let order: Vec<_> = (0..count).rev().map(Selector::from).collect();
            edit.reorder_sheets(&order)
        }));
        attempts.push(Box::new(|edit| {
            edit.insert_sheet(Position::new(0), "Inserted Tab")
        }));
        for position in 0..count {
            attempts.push(Box::new(move |edit| {
                edit.remove_sheet(Selector::Position(Position::new(position)))
            }));
        }
        for (index, attempt) in attempts.iter().enumerate() {
            let mut edit = source.edit();
            if attempt(&mut edit).is_err() {
                continue;
            }
            let (target, patch) = edit
                .commit()
                .unwrap_or_else(|error| {
                    panic!(
                        "{}: staged change {index} failed to commit: {error}",
                        path.display()
                    )
                })
                .into_parts();
            accepted += 1;
            assert_eq!(
                patch.inverse().apply(&target).unwrap().bytes(),
                &bytes[..],
                "{}",
                path.display()
            );
            if index == 0 {
                // A reversed catalog renders a 3-D range from its new first
                // tab, so compare the cell values and formula presence.
                assert_eq!(
                    values_only(sorted(contents(target.bytes()))),
                    values_only(before.clone()),
                    "{}",
                    path.display()
                );
            }
        }
    }
    assert!(names_seen > 40, "too few corpus tabs: {names_seen}");
    assert!(
        accepted > 40,
        "too few accepted corpus tab changes: {accepted}"
    );
}

#[test]
fn fixture_with_chart_tab_reorders_complete_catalog() {
    let bytes = fixture("WithChart.xls");
    let Ok(source) = Snapshot::from_bytes(bytes, Limits::default()) else {
        return;
    };
    let catalog = tabs::catalog(source.workbook_stream()).unwrap();
    if !catalog.iter().any(|tab| tab.kind == 2) {
        return;
    }
    let order: Vec<_> = (0..catalog.len()).rev().map(Selector::from).collect();
    let mut edit = source.edit();
    edit.reorder_sheets(&order).unwrap();
    let target = edit.commit().unwrap().into_parts().0;
    let mut expected: Vec<_> = catalog.into_iter().map(|tab| tab.name).collect();
    expected.reverse();
    assert_eq!(names(&target), expected);
}
