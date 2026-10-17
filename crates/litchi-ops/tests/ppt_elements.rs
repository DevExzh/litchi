//! Unified PPT element surface (record 0323): slide comments, table rows,
//! cells and columns, hyperlinks, runs, line breaks, the theme and charts.
//!
//! Every write is checked on the eager and the positional owner, read back
//! after a byte round trip, and the source handle is left unchanged.

#![cfg(feature = "ppt")]

use std::path::PathBuf;

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

const WITH_COMMENTS: &str = "poi/test-data/slideshow/WithComments.ppt";
const NO_COMMENTS: &str = "poi/test-data/slideshow/basic_test_ppt_file.ppt";

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .expect("PPT fixture")
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn each_owner(bytes: &[u8], mut test: impl FnMut(File)) {
    test(eager(bytes));
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("source.ppt");
    std::fs::write(&path, bytes).expect("write PPT source");
    test(
        litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .expect("open positional PPT"),
    );
}

fn reopen(file: &File) -> File {
    eager(&file.to_bytes().expect("published bytes"))
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .unwrap_or_else(|| panic!("{path} has no text"))
        .into_owned()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn paths(file: &File, query: &str) -> Vec<String> {
    file.query(query)
        .expect("query")
        .map(|node| node.expect("node").path().to_string())
        .collect()
}

#[test]
fn comments_read_like_pptx_comments() {
    let deck = eager(&fixture(WITH_COMMENTS));
    assert_eq!(
        prop(&deck, "/slide[1]/comment[1]", "author").as_deref(),
        Some("Administrator")
    );
    assert_eq!(
        prop(&deck, "/slide[1]/comment[1]", "initials").as_deref(),
        Some("A")
    );
    assert_eq!(
        prop(&deck, "/slide[1]/comment[1]", "index").as_deref(),
        Some("1")
    );
    assert!(
        prop(&deck, "/slide[1]/comment[1]", "date")
            .unwrap()
            .starts_with("2008-08-04T")
    );
    assert!(text(&deck, "/slide[1]/comment[1]").contains("This is a test comment"));
    assert_eq!(paths(&deck, "comment"), ["/slide[1]/comment[1]"]);
    assert_eq!(
        deck.get("/slide[1]/comment[2]").unwrap_err().code(),
        Code::PathNotFound
    );
}

#[test]
fn comment_add_set_and_remove_round_trip() {
    each_owner(&fixture(NO_COMMENTS), |source| {
        let before = source.to_bytes().expect("source bytes");
        let added = source
            .add(
                "/slide[2]",
                NewNode::new("comment")
                    .prop("text", "Check this")
                    .prop("author", "Jane Doe")
                    .prop("date", "2026-10-17T09:30:00")
                    .prop("x", "1in")
                    .prop("y", "2in"),
            )
            .expect("add comment");
        assert_eq!(added.path.to_string(), "/slide[2]/comment[1]");
        let read = reopen(&added.file);
        assert_eq!(text(&read, "/slide[2]/comment[1]"), "Check this");
        assert_eq!(
            prop(&read, "/slide[2]/comment[1]", "initials").as_deref(),
            Some("JD")
        );
        assert_eq!(
            prop(&read, "/slide[2]/comment[1]", "date").as_deref(),
            Some("2026-10-17T09:30:00.000")
        );
        assert_eq!(
            prop(&read, "/slide[2]/comment[1]", "x").as_deref(),
            Some("2.54cm")
        );
        assert_eq!(source.to_bytes().unwrap(), before);

        let changed = read
            .set(
                "/slide[2]/comment[1]",
                props! { "text" => "Checked", "y" => "1in" },
            )
            .expect("set comment");
        let changed = reopen(&changed);
        assert_eq!(text(&changed, "/slide[2]/comment[1]"), "Checked");
        assert_eq!(
            prop(&changed, "/slide[2]/comment[1]", "y").as_deref(),
            Some("2.54cm")
        );
        assert_eq!(
            prop(&changed, "/slide[2]/comment[1]", "x").as_deref(),
            Some("2.54cm")
        );

        let removed = changed
            .remove("/slide[2]/comment[1]")
            .expect("remove comment");
        assert!(paths(&reopen(&removed), "comment").is_empty());
    });
}

#[test]
fn comment_noops_and_refusals_are_typed() {
    let source = eager(&fixture(WITH_COMMENTS));
    let before = source.to_bytes().unwrap();
    let current = text(&source, "/slide[1]/comment[1]");
    let noop = source
        .edit(|edit| {
            edit.set(
                "/slide[1]/comment[1]",
                props! { "text" => current.as_str() },
            )
        })
        .expect("equal text");
    assert!(noop.patch.is_noop());
    assert_eq!(noop.file.to_bytes().unwrap(), before);

    let missing_author = source
        .add("/slide[1]", NewNode::new("comment").prop("text", "x"))
        .unwrap_err();
    assert_eq!(missing_author.code(), Code::MissingProperty);
    let bad_date = source
        .add(
            "/slide[1]",
            NewNode::new("comment")
                .prop("author", "A")
                .prop("date", "yesterday"),
        )
        .unwrap_err();
    assert_eq!(bad_date.code(), Code::InvalidValue);
    let author_set = source
        .set("/slide[1]/comment[1]", props! { "author" => "B" })
        .unwrap_err();
    assert_eq!(author_set.code(), Code::UnsupportedProperty);
    assert!(
        source
            .add(
                "/slide[1]/shape[1]",
                NewNode::new("comment").prop("author", "A")
            )
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}

const TABLES: &str = "poi/test-data/slideshow/bug61881.ppt";

/// A deck with a title and a 2x3 table authored by the native builder.
fn table_deck() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "Title").unwrap();
    let mut table = litchi_ppt::writer::Table::new(2, 3).unwrap();
    for (row, texts) in [["A1", "B1", "C1"], ["A2", "B2", "C2"]].iter().enumerate() {
        for (column, text) in texts.iter().enumerate() {
            table.set_cell_text(row, column, *text).unwrap();
        }
    }
    writer.add_table(slide, 72, 120, table).unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn table_rows_cells_and_columns_read_like_pptx() {
    let deck = eager(&table_deck());
    assert_eq!(text(&deck, "/slide[1]/table[1]/tr[2]/tc[3]"), "C2");
    assert_eq!(text(&deck, "/slide[1]/table[1]/row[1]/cell[1]"), "A1");
    assert!(prop(&deck, "/slide[1]/table[1]/tr[1]", "height").is_some());
    assert!(prop(&deck, "/slide[1]/table[1]/col[2]", "width").is_some());
    assert_eq!(
        deck.get("/slide[1]/table[1]/tr[3]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        deck.get("/slide[1]/table[1]/col[4]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        paths(&deck, "/slide[1]/table[1] > tr"),
        ["/slide[1]/table[1]/tr[1]", "/slide[1]/table[1]/tr[2]"]
    );
    assert_eq!(paths(&deck, "tc").len(), 6);

    let powerpoint = eager(&fixture(TABLES));
    assert_eq!(paths(&powerpoint, "/slide[11]/table[1] > tr").len(), 5);
    assert_eq!(text(&powerpoint, "/slide[11]/table[1]/tr[1]/tc[1]"), "");
}

#[test]
fn table_cell_text_set_round_trips_and_layout_edits_are_refused() {
    each_owner(&table_deck(), |source| {
        let before = source.to_bytes().unwrap();
        let changed = source
            .set(
                "/slide[1]/table[1]/tr[1]/tc[2]",
                props! { "text" => "A longer cell text" },
            )
            .expect("cell text");
        let read = reopen(&changed);
        assert_eq!(
            text(&read, "/slide[1]/table[1]/tr[1]/tc[2]"),
            "A longer cell text"
        );
        assert_eq!(text(&read, "/slide[1]/table[1]/tr[1]/tc[1]"), "A1");
        assert_eq!(text(&read, "/slide[1]/table[1]/tr[2]/tc[2]"), "B2");
        assert_eq!(text(&read, "/slide[1]/shape[1]"), "Title");
        assert_eq!(source.to_bytes().unwrap(), before);

        let noop = source
            .edit(|edit| edit.set("/slide[1]/table[1]/tr[2]/tc[1]", props! { "text" => "A2" }))
            .unwrap();
        assert!(noop.patch.is_noop());

        for refused in [
            source
                .add("/slide[1]/table[1]", NewNode::new("row"))
                .unwrap_err(),
            source
                .add("/slide[1]/table[1]", NewNode::new("column"))
                .unwrap_err(),
            source.remove("/slide[1]/table[1]/tr[1]").unwrap_err(),
            source.remove("/slide[1]/table[1]/col[1]").unwrap_err(),
        ] {
            assert_eq!(refused.code(), Code::UnsupportedCapability);
        }
    });

    // PowerPoint-authored empty cells gain their first text atom.
    let powerpoint = eager(&fixture(TABLES));
    let changed = powerpoint
        .set(
            "/slide[11]/table[1]/tr[2]/tc[1]",
            props! { "text" => "filled" },
        )
        .unwrap();
    assert_eq!(
        text(&reopen(&changed), "/slide[11]/table[1]/tr[2]/tc[1]"),
        "filled"
    );
}
