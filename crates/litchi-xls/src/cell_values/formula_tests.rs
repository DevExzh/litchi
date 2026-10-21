#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "formula fixtures intentionally fail fast"
)]

use super::*;

fn fixture() -> Vec<u8> {
    let mut builder = crate::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 2.0).unwrap();
    builder.write_number(data, 0, 1, 7.0).unwrap();
    builder.write_string(data, 1, 0, "label").unwrap();
    builder.write_number(data, 2, 2, 5.0).unwrap();
    builder.write_formula(data, 3, 0, "A1+1").unwrap();
    builder.add_worksheet("Other").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

fn reference(row: u32, column: u32) -> Reference {
    Reference::new(row, column).unwrap()
}

/// Rendered formula text and storage at one coordinate, from the complete
/// reader and the edit inventory.
fn formula_at(snapshot: &Snapshot, row: u32, column: u32) -> (Option<String>, Option<Storage>) {
    let reader = Reader::new(Cursor::new(snapshot.bytes().to_vec())).unwrap();
    let rendered = reader
        .xls_worksheet(0)
        .unwrap()
        .get_cell(row, column)
        .and_then(|cell| cell.formula().map(str::to_owned));
    let storage = snapshot
        .worksheet(Selector::Position(Position::new(0)))
        .unwrap()
        .unwrap()
        .cell(reference(row, column))
        .unwrap()
        .map(Cell::storage);
    (rendered, storage)
}

#[test]
fn formulas_are_authored_into_absent_and_scalar_cells() {
    let bytes = fixture();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let original_style = source
        .worksheet(Selector::Position(Position::new(0)))
        .unwrap()
        .unwrap()
        .cell(reference(0, 1))
        .unwrap()
        .unwrap()
        .style();
    let mut edit = source.edit();
    // Absent, standalone number, shared string, and a cell on the same
    // sheet's other row compose in one edit.
    edit.set_formula(Selector::Name("data"), reference(5, 5), "A1*2")
        .unwrap();
    edit.set_formula(
        Selector::Position(Position::new(0)),
        reference(0, 1),
        "A1*3",
    )
    .unwrap();
    edit.set_formula(
        Selector::Position(Position::new(0)),
        reference(1, 0),
        "SUM(A1:B1)",
    )
    .unwrap();
    edit.insert_cell(
        Selector::Position(Position::new(0)),
        reference(6, 0),
        Value::Number(1.5),
    )
    .unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(
        formula_at(&target, 5, 5),
        (Some("=(A1*2)".into()), Some(Storage::Formula))
    );
    assert_eq!(
        formula_at(&target, 0, 1),
        (Some("=(A1*3)".into()), Some(Storage::Formula))
    );
    assert_eq!(
        formula_at(&target, 1, 0),
        (Some("=SUM((A1:B1))".into()), Some(Storage::Formula))
    );
    let replaced = target
        .worksheet(Selector::Position(Position::new(0)))
        .unwrap()
        .unwrap()
        .cell(reference(0, 1))
        .unwrap()
        .unwrap()
        .clone();
    assert_eq!(replaced.style(), original_style);
    // Exact inverse restores bytes; the semantic patch replays and inverts.
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    let replayed = patch.semantic().apply(&source).unwrap().into_parts().0;
    assert_eq!(replayed.bytes(), target.bytes());
    let restored = patch.semantic().inverse().apply(&target).unwrap();
    let restored = restored.into_parts().0;
    assert_eq!(formula_at(&restored, 0, 1), (None, Some(Storage::Number)));
    assert_eq!(formula_at(&restored, 1, 0), (None, Some(Storage::LabelSst)));
    assert_eq!(formula_at(&restored, 5, 5), (None, None));
    let json = patch.semantic().to_deterministic_json().unwrap();
    assert!(SemanticPatch::from_deterministic_json(&json).is_ok());
}

#[test]
fn existing_formula_cells_and_bad_expressions_are_refused_atomically() {
    let source = Snapshot::from_bytes(fixture(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.set_formula(Selector::Position(Position::new(0)), reference(0, 0), "B1")
        .unwrap();
    for (row, column, expression) in [(3, 0, "A1+2"), (2, 2, ""), (2, 2, "Other!A1")] {
        let before = edit.clone();
        assert!(
            edit.set_formula(
                Selector::Position(Position::new(0)),
                reference(row, column),
                expression
            )
            .is_err(),
            "{expression:?}"
        );
        assert_eq!(
            edit.structural_changes.len(),
            before.structural_changes.len()
        );
        assert_eq!(edit.resource_changes.len(), before.resource_changes.len());
    }
    // A second operation on a replaced coordinate is refused.
    assert!(
        edit.set_value(
            Selector::Position(Position::new(0)),
            reference(0, 0),
            Value::Number(1.0)
        )
        .is_err()
    );
    assert!(
        edit.insert_rows(Selector::Position(Position::new(0)), 10, 1)
            .is_err(),
        "an axis shift would move the authored Formula"
    );
    let target = edit.commit().unwrap().into_parts().0;
    assert_eq!(
        formula_at(&target, 0, 0),
        (Some("=B1".into()), Some(Storage::Formula))
    );
}
