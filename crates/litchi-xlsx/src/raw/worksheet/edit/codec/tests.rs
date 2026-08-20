//! Focused regression tests for the worksheet codec seams.

use super::{Attribute, Tag, scan, sibling_name, write_tag};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[test]
fn sibling_names_retain_the_source_prefix() {
    assert_eq!(sibling_name("x:row", "c"), "x:c");
    assert_eq!(sibling_name("row", "c"), "c");
}

#[test]
fn wire_tags_escape_and_replace_attributes_deterministically() {
    let tag = Tag {
        name: "x:c".into(),
        attributes: vec![Attribute {
            name: "r".into(),
            value: "A1".into(),
        }]
        .into_boxed_slice(),
    };
    let mut output = Vec::new();
    write_tag(&mut output, &tag, true, &["r"], &[("r", "A&B".to_owned())]);
    assert_eq!(output, br#"<x:c r="A&amp;B"/>"#);
}

#[test]
fn snapshot_scan_builds_sorted_edit_slots() {
    let source = format!(
        r#"<x:worksheet xmlns:x="{MAIN}"><x:dimension ref="A1:B2"/><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row><x:row r="2"><x:c r="B2"/></x:row></x:sheetData></x:worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    assert_eq!(layout.sheet_data.rows.len(), 2);
    assert_eq!(layout.sheet_data.rows[0].cells.len(), 1);
    assert_eq!(layout.sheet_data.rows[1].cells[0].address.a1(), "B2");
}

#[test]
fn snapshot_scan_restores_namespace_scope_after_a_rebinding() {
    let source = format!(
        r#"<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1"/></x:row><x:row xmlns:x="urn:future" r="2"><x:c r="B2"/></x:row><x:row r="3"><x:c r="C3"/></x:row></x:sheetData></x:worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    assert_eq!(
        layout
            .sheet_data
            .rows
            .iter()
            .map(|row| row.number)
            .collect::<Vec<_>>(),
        [1, 3]
    );
}
