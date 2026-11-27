//! PPTX and XLSX handles promote a source-backed owner to an owning package
//! for some writes (notes, comments, charts). The promoted package must carry
//! the edits the handle already committed, so the saved package equals the
//! one the eager owner saves for the same edits.
#![cfg(all(feature = "pptx", feature = "xlsx"))]

use std::collections::BTreeMap;
use std::io::{Cursor, Read as _};

use litchi_core::FileFormat;
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).unwrap();
            let mut data = Vec::new();
            member.read_to_end(&mut data).unwrap();
            (member.name().unwrap().into_owned(), data)
        })
        .collect()
}

/// Apply `edit` to a source-backed and to an eager open of `input`, and
/// require equal saved packages that both contain every `expected` text.
fn check(input: &[u8], edit: impl Fn(File) -> File, expected: &[&str]) {
    let source = litchi_ops::open_bytes(input.to_vec()).unwrap();
    if let File::Pptx(pptx) = &source {
        assert!(pptx.is_source_backed(), "the default open is source-backed");
    }
    let eager = litchi_ops::open_bytes_with(
        input.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .unwrap();
    let saved = members(&edit(source).to_bytes().unwrap());
    let oracle = members(&edit(eager).to_bytes().unwrap());
    let all: String = saved
        .values()
        .map(|data| String::from_utf8_lossy(data).into_owned())
        .collect();
    for needle in expected {
        assert!(all.contains(needle), "the saved package lost {needle:?}");
    }
    assert_eq!(
        saved, oracle,
        "the saved package differs from the eager one"
    );
}

fn deck() -> Vec<u8> {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
        .add(
            "/slide[1]",
            NewNode::new("shape")
                .prop("width", "2in")
                .prop("height", "1in")
                .prop("text", "Shape"),
        )
        .unwrap()
        .file
        .to_bytes()
        .unwrap()
}

fn workbook() -> Vec<u8> {
    let mut file = litchi_ops::create(FileFormat::Xlsx).unwrap();
    for (cell, value) in [("A1", "Month"), ("B1", "East"), ("A2", "Jan"), ("B2", "10")] {
        file = file
            .set(
                format!("/Sheet1/{cell}").as_str(),
                props! { "value" => value },
            )
            .unwrap();
    }
    file.to_bytes().unwrap()
}

#[test]
fn pptx_notes_and_comment_writes_keep_earlier_shape_edits() {
    check(
        &deck(),
        |file| {
            file.set("/slide[1]/shape[1]", props! { "text" => "SHAPE_CHANGED" })
                .unwrap()
                .add(
                    "/slide[1]",
                    NewNode::new("notes").prop("text", "NOTES_ADDED"),
                )
                .unwrap()
                .file
                .set("/slide[1]/shape[1]", props! { "text" => "SHAPE_AGAIN" })
                .unwrap()
                .add(
                    "/slide[1]",
                    NewNode::new("comment")
                        .prop("author", "Ann")
                        .prop("text", "COMMENT_ADDED"),
                )
                .unwrap()
                .file
        },
        &["SHAPE_AGAIN", "NOTES_ADDED", "COMMENT_ADDED"],
    );
}

#[test]
fn xlsx_comment_and_chart_writes_keep_earlier_cell_edits() {
    check(
        &workbook(),
        |file| {
            file.set("/Sheet1/A2", props! { "value" => "CELL_CHANGED" })
                .unwrap()
                .add(
                    "/Sheet1",
                    NewNode::new("comment")
                        .prop("ref", "B2")
                        .prop("text", "COMMENT_ADDED")
                        .prop("author", "Ann"),
                )
                .unwrap()
                .file
                .set("/Sheet1/A1", props! { "value" => "CELL_AGAIN" })
                .unwrap()
                .add(
                    "/Sheet1",
                    NewNode::new("chart")
                        .prop("chartType", "bar")
                        .prop("dataRange", "Sheet1!A1:B2"),
                )
                .unwrap()
                .file
        },
        &["CELL_CHANGED", "CELL_AGAIN", "COMMENT_ADDED"],
    );
}
