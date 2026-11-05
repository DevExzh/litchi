//! Record 0326: a note's `get` picks the selected note without building
//! every note (XLSX and XLSB). Each step form must still give the node the
//! listing gives at that position, in row-major order of the cells.
#![cfg(any(feature = "xlsx", feature = "xlsb"))]

use litchi_core::FileFormat;
use litchi_ops::{File, NewNode, OpenOptions, SourceMode};

/// Cells in neither row-major nor column-major order.
const CELLS: [&str; 9] = ["C3", "A1", "B10", "A10", "Z2", "B2", "A2", "AA1", "C1"];

fn with_notes(format: FileFormat) -> File {
    let mut file = litchi_ops::create(format).unwrap();
    for (at, cell) in CELLS.iter().enumerate() {
        file = file
            .add(
                "/Sheet1",
                NewNode::new("comment")
                    .prop("ref", *cell)
                    .prop("text", format!("note {at}"))
                    .prop("author", format!("author {at}")),
            )
            .unwrap()
            .file;
    }
    file
}

fn describe(node: &litchi_ops::node::Node<'_>) -> String {
    let view = node.view();
    let format: Vec<String> = view
        .format()
        .unwrap()
        .map(|prop| format!("{}={}", prop.key(), prop.value()))
        .collect();
    format!(
        "{} {:?} {format:?}",
        node.path(),
        view.text().unwrap().map(|text| text.into_owned())
    )
}

fn check(format: FileFormat) {
    let built = with_notes(format);
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            built.to_bytes().unwrap(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let listed: Vec<String> = file
            .query("comment")
            .unwrap()
            .map(|node| describe(&node.unwrap()))
            .collect();
        assert_eq!(listed.len(), CELLS.len());
        for (at, expected) in listed.iter().enumerate() {
            let path = format!("/Sheet1/comment[{}]", at + 1);
            assert_eq!(
                &describe(&file.get(path.as_str()).unwrap()),
                expected,
                "{path}"
            );
        }
        assert_eq!(
            &describe(&file.get("/Sheet1/comment[last()]").unwrap()),
            listed.last().unwrap()
        );
        assert_eq!(&describe(&file.get("/Sheet1/comment").unwrap()), &listed[0]);
        let error = file.get("/Sheet1/comment[10]").unwrap_err();
        assert_eq!(error.code().as_str(), "path_not_found", "{error}");
    }
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_note_lookup_matches_the_listing() {
    check(FileFormat::Xlsx);
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_note_lookup_matches_the_listing() {
    check(FileFormat::Xlsb);
}
