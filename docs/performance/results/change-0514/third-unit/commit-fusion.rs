//! Public transaction coverage for the conditional worksheet fusion path.

use std::sync::{Arc, Barrier};

use litchi_opc::PackURI;
use litchi_sheet::Cell as Address;

use super::super::*;
use super::support::{part_text, styled_workbook};
use crate::cell::{Cell, Value};

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn workbook_with_sheet(source: impl Into<Vec<u8>>) -> Workbook {
    let baseline = Workbook::new().expect("baseline workbook");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(source.into());
    Workbook::from_package(package).expect("workbook fixture")
}

fn unknown_attribute_sheet(value: &str) -> Vec<u8> {
    format!(
        r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" q:opaque="{value}"><v>1</v></c></row></sheetData></worksheet>"#
    )
    .into_bytes()
}

#[test]
fn changed_cell_commit_preserves_unknown_namespaces_and_has_exact_inverse() {
    let source = workbook_with_sheet(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="urn:litchi:future"><dimension ref="A1:B1"/><sheetData q:container="keep"><row r="1" q:row="keep"><c r="A1" q:opaque="yes"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
    );
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 9_i32)
        .expect("cell edit");
    let committed = edit.commit().expect("commit");

    assert_eq!(committed.patch().len(), 1);
    assert_eq!(
        part_text(committed.workbook(), "/xl/worksheets/sheet1.xml"),
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="urn:litchi:future"><dimension ref="A1:B1"/><sheetData q:container="keep"><row r="1" q:row="keep"><c r="A1" q:opaque="yes"><v>9</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
    );
    assert_eq!(
        source
            .apply(committed.patch())
            .expect("forward patch")
            .workbook()
            .to_bytes()
            .expect("patched bytes"),
        committed.workbook().to_bytes().expect("committed bytes")
    );
    assert_eq!(
        committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse patch")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn cold_same_value_edit_discards_snapshot_only_failure_as_an_exact_no_op() {
    let source = workbook_with_sheet(unknown_attribute_sheet("&missing;"));
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 1_i32)
        .expect("same-value edit");
    let committed = edit.commit().expect("semantic no-op must succeed");

    assert!(committed.patch().is_empty());
    assert_eq!(
        committed.workbook().to_bytes().expect("no-op bytes"),
        source_bytes
    );
}

#[test]
fn prepopulated_store_same_value_edit_keeps_the_hot_no_op_fallback() {
    let source = workbook_with_sheet(unknown_attribute_sheet("&missing;"));
    assert!(
        source
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .cell("A1")
            .expect("cell")
            .stored()
            .is_some()
    );
    assert!(source.inner.sheets[0].cells.get().is_some());
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 1_i32)
        .expect("same-value edit");
    let committed = edit.commit().expect("hot semantic no-op must succeed");

    assert!(committed.patch().is_empty());
    assert_eq!(
        committed.workbook().to_bytes().expect("no-op bytes"),
        source_bytes
    );
}

#[test]
fn effective_edit_exposes_deferred_snapshot_error_at_rewrite_boundary() {
    let source = workbook_with_sheet(unknown_attribute_sheet("&missing;"));
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 2_i32)
        .expect("cell edit");
    let error = edit
        .commit()
        .expect_err("snapshot tag must reject the edit");
    assert!(
        error.to_string().contains("missing"),
        "unexpected error: {error}"
    );
}

#[test]
fn semantic_cell_error_precedes_deferred_snapshot_error() {
    let source = workbook_with_sheet(format!(
        r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" s="not-a-style" q:opaque="&missing;"/></row></sheetData></worksheet>"#
    ));
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 2_i32)
        .expect("cell edit");
    let error = edit.commit().expect_err("semantic style syntax must win");
    assert!(
        error.to_string().contains("invalid worksheet cell style"),
        "unexpected error: {error}"
    );
    assert!(!error.to_string().contains("missing"));
}

#[test]
fn workbook_style_validation_precedes_deferred_snapshot_error() {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(
            format!(
                r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" s="2" q:opaque="&missing;"><v>1</v></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
        );
    let source = Workbook::from_package(package).expect("style fixture");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 2_i32)
        .expect("cell edit");
    let error = edit
        .commit()
        .expect_err("style catalog validation must win");
    assert!(
        error.to_string().contains("references shared style 2"),
        "unexpected error: {error}"
    );
    assert!(!error.to_string().contains("missing"));
}

#[test]
fn unknown_cell_projection_precedes_deferred_snapshot_error() {
    let source = workbook_with_sheet(format!(
        r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" t="future" q:opaque="&missing;"><v>payload</v></c></row></sheetData></worksheet>"#
    ));
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 2_i32)
        .expect("cell edit");
    assert!(matches!(
        edit.commit(),
        Err(Error::EditBlocked {
            reason: EditBlock::UnknownCell,
            ..
        })
    ));
}

#[test]
fn prepopulated_store_effective_edit_keeps_exact_source_rewrite_and_inverse() {
    let source = workbook_with_sheet(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" q:opaque="yes"><v>1</v></c></row></sheetData></worksheet>"#,
    );
    assert!(
        source
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .cell("A1")
            .expect("cell")
            .stored()
            .is_some()
    );
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", 3_i32)
        .expect("cell edit");
    let committed = edit.commit().expect("warm edit");
    assert_eq!(committed.patch().len(), 1);
    assert!(
        part_text(committed.workbook(), "/xl/worksheets/sheet1.xml")
            .contains(r#"<c r="A1" q:opaque="yes"><v>3</v></c>"#)
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
fn concurrent_cold_read_and_commit_publish_only_canonical_stores() {
    let source = Arc::new(workbook_with_sheet(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
    ));
    let barrier = Arc::new(Barrier::new(9));
    let (committed, observations) = std::thread::scope(|scope| {
        let commit_source = Arc::clone(&source);
        let commit_barrier = Arc::clone(&barrier);
        let commit = scope.spawn(move || {
            commit_barrier.wait();
            let mut edit = commit_source.edit().expect("edit");
            edit.sheet("Sheet1")
                .expect("sheet lookup")
                .expect("worksheet")
                .set("A1", 9_i32)
                .expect("cell edit");
            edit.commit().expect("commit").into_workbook()
        });

        let mut readers = Vec::new();
        for _ in 0..8 {
            let read_source = Arc::clone(&source);
            let read_barrier = Arc::clone(&barrier);
            readers.push(scope.spawn(move || {
                read_barrier.wait();
                read_source
                    .sheet("Sheet1")
                    .expect("sheet lookup")
                    .expect("worksheet")
                    .cell("A1")
                    .expect("cell")
                    .stored()
                    .cloned()
            }));
        }

        let committed = commit.join().expect("commit worker");
        let observations = readers
            .into_iter()
            .map(|reader| reader.join().expect("read worker"))
            .collect::<Vec<_>>();
        (committed, observations)
    });

    assert!(observations.iter().all(|observed| matches!(
        observed,
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "1"
    )));
    let source_store = source.inner.sheets[0]
        .cells
        .get()
        .expect("cold readers must publish the source store");
    assert!(matches!(
        source_store.get(Address::from_a1("A1").expect("address")),
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "1"
    ));
    assert!(matches!(
        committed
            .inner
            .sheets[0]
            .cells
            .get()
            .expect("validated commit store")
            .get(Address::from_a1("A1").expect("address")),
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "9"
    ));
}
