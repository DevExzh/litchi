//! The last DOCX P0 writes of record 0345 through the unified API, after
//! OfficeCLI's word schemas: the hyperlink screen tip, cell insertion and
//! cell text on `add`, the table `col` element, and comment replies,
//! resolution and anchors (`commentsExtended`). Each write is read back
//! after a byte round trip on both owners.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _};

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

fn ok<T>(result: litchi_ops::Result<T>) -> T {
    result.unwrap_or_else(|error| panic!("{error} ({:?})", error.property()))
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
}

fn children(file: &File, path: &str) -> Option<u32> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .child_count()
        .unwrap()
}

fn part(file: &File, name: &str) -> Option<String> {
    let bytes = file.to_bytes().expect("publish");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name(name).ok()?;
    let mut contents = String::new();
    member.read_to_string(&mut contents).expect("read part");
    Some(contents)
}

fn document_xml(file: &File) -> String {
    part(file, "word/document.xml").expect("main document")
}

fn reopen(file: &File, eager: bool) -> File {
    litchi_ops::open_bytes_with(
        file.to_bytes().unwrap(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn with_paragraph(text: &str) -> File {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    ok(file.add("/body", NewNode::new("paragraph").prop("text", text))).file
}

fn with_table(rows: &str, cols: &str) -> File {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    ok(file.add(
        "/body",
        NewNode::new("table").prop("rows", rows).prop("cols", cols),
    ))
    .file
}

// ---------------------------------------------------------------------------
// Hyperlink tooltip
// ---------------------------------------------------------------------------

/// `hyperlink.json`: `tooltip` seeds the new link's `w:tooltip`; Litchi
/// also sets and clears it.
#[test]
fn hyperlink_tooltip_is_added_set_and_cleared() {
    for eager in [true, false] {
        let file = reopen(&with_paragraph("See "), eager);
        let added = ok(file.add(
            "/body/p[1]",
            NewNode::new("hyperlink")
                .prop("url", "https://example.test/")
                .prop("text", "the report")
                .prop("tooltip", "Opens the report"),
        ));
        assert_eq!(added.path.to_string(), "/body/p[1]/hyperlink[1]");
        let file = reopen(&added.file, eager);
        assert_eq!(
            prop(&file, "/body/p[1]/hyperlink[1]", "tooltip").as_deref(),
            Some("Opens the report")
        );
        assert!(document_xml(&file).contains(r#"w:tooltip="Opens the report""#));

        let file = reopen(
            &ok(file.set(
                "/body/p[1]/hyperlink[1]",
                props! { "tooltip" => "Q3 & Q4 <draft>" },
            )),
            eager,
        );
        assert_eq!(
            prop(&file, "/body/p[1]/hyperlink[1]", "tooltip").as_deref(),
            Some("Q3 & Q4 <draft>")
        );
        assert!(document_xml(&file).contains(r#"w:tooltip="Q3 &amp; Q4 &lt;draft&gt;""#));

        // An identical value publishes the input bytes.
        let before = file.to_bytes().unwrap();
        let same = ok(file.set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => "Q3 & Q4 <draft>" },
        ));
        assert_eq!(same.to_bytes().unwrap(), before);

        let cleared = ok(file.set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => litchi_ops::PropValue::Null },
        ));
        let cleared = reopen(&cleared, eager);
        assert_eq!(prop(&cleared, "/body/p[1]/hyperlink[1]", "tooltip"), None);
        assert!(!document_xml(&cleared).contains("tooltip"));
    }
}

/// A hyperlink without a tooltip is written exactly as before.
#[test]
fn hyperlink_without_tooltip_has_no_tooltip_attribute() {
    let file = with_paragraph("See ");
    let added = ok(file.add(
        "/body/p[1]",
        NewNode::new("hyperlink")
            .prop("anchor", "Intro")
            .prop("text", "intro"),
    ));
    assert!(!document_xml(&added.file).contains("tooltip"));
    let error = added
        .file
        .set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => "bad\u{1}tip" },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
}

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

/// `table-cell.json`: `add cell` under a row with `text` and `width`; the
/// grid gains a column and the other rows a padding cell (OfficeCLI's
/// `AddCell`).
#[test]
fn a_cell_is_added_to_a_row_and_the_grid_follows() {
    for eager in [true, false] {
        let file = reopen(&with_table("2", "2"), eager);
        let added = ok(file.add(
            "/body/tbl[1]/tr[1]",
            NewNode::new("cell")
                .prop("text", "Total")
                .prop("width", "1440"),
        ));
        assert_eq!(added.path.to_string(), "/body/tbl[1]/tr[1]/tc[3]");
        let file = reopen(&added.file, eager);
        assert_eq!(
            text(&file, "/body/tbl[1]/tr[1]/tc[3]").as_deref(),
            Some("Total")
        );
        assert_eq!(
            prop(&file, "/body/tbl[1]/tr[1]/tc[3]", "width").as_deref(),
            Some("1440dxa")
        );
        // The other row is padded and the grid has three columns.
        assert_eq!(children(&file, "/body/tbl[1]/tr[2]"), Some(3));
        let xml = document_xml(&file);
        assert_eq!(xml.matches("<w:gridCol ").count(), 3, "{xml}");

        // A placement before the first cell, with a percentage width.
        let first = ok(file.add(
            "/body/tbl[1]/tr[2]",
            NewNode::new("tc")
                .prop("text", "Lead")
                .prop("width", "25%")
                .at(litchi_ops::Placement::Start),
        ));
        assert_eq!(first.path.to_string(), "/body/tbl[1]/tr[2]/tc[1]");
        let file = reopen(&first.file, eager);
        assert_eq!(
            prop(&file, "/body/tbl[1]/tr[2]/tc[1]", "width").as_deref(),
            Some("25%")
        );
        assert_eq!(
            text(&file, "/body/tbl[1]/tr[2]/tc[1]").as_deref(),
            Some("Lead")
        );
    }
}

/// Cell `text` is now written by `add`, and still by `set`.
#[test]
fn cell_text_is_set_and_refused_values_leave_the_input() {
    let file = with_table("1", "1");
    let file = ok(file.set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "Changed" }));
    assert_eq!(
        text(&file, "/body/tbl[1]/tr[1]/tc[1]").as_deref(),
        Some("Changed")
    );
    let before = file.to_bytes().unwrap();
    let error = file
        .add(
            "/body/tbl[1]/tr[1]",
            NewNode::new("cell").prop("width", "wide"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    let error = file
        .add(
            "/body/tbl[1]/tr[1]",
            NewNode::new("cell").prop("valign", "top"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    assert_eq!(file.to_bytes().unwrap(), before);
}

// ---------------------------------------------------------------------------
// Columns
// ---------------------------------------------------------------------------

/// `table-column.json`: `add column` with `width` and `text`; Litchi also
/// reads a column's width and sets it, and removes columns.
#[test]
fn table_columns_are_added_read_set_and_removed() {
    for eager in [true, false] {
        let file = reopen(&with_table("2", "2"), eager);
        assert_eq!(
            prop(&file, "/body/tbl[1]/col[1]", "width").as_deref(),
            Some("4680")
        );
        let added = ok(file.add(
            "/body/tbl[1]",
            NewNode::new("column")
                .prop("width", "2400")
                .prop("text", "Q1"),
        ));
        assert_eq!(added.path.to_string(), "/body/tbl[1]/col[3]");
        let file = reopen(&added.file, eager);
        assert_eq!(
            prop(&file, "/body/tbl[1]/col[3]", "width").as_deref(),
            Some("2400")
        );
        for row in 1..=2 {
            assert_eq!(
                text(&file, &format!("/body/tbl[1]/tr[{row}]/tc[3]")).as_deref(),
                Some("Q1")
            );
        }
        // Before the first column; the width is the grid's average.
        let first = ok(file.add(
            "/body/tbl[1]",
            NewNode::new("col").at(litchi_ops::Placement::Start),
        ));
        assert_eq!(first.path.to_string(), "/body/tbl[1]/col[1]");
        let file = reopen(&first.file, eager);
        assert_eq!(
            prop(&file, "/body/tbl[1]/col[1]", "width").as_deref(),
            Some("3920")
        );
        assert_eq!(children(&file, "/body/tbl[1]/tr[1]"), Some(4));

        // A width set rewrites the grid column and the cells' widths.
        let set = ok(file.set("/body/tbl[1]/col[2]", props! { "width" => "2cm" }));
        let set = reopen(&set, eager);
        assert_eq!(
            prop(&set, "/body/tbl[1]/col[2]", "width").as_deref(),
            Some("1134")
        );
        assert_eq!(
            prop(&set, "/body/tbl[1]/tr[1]/tc[2]", "width").as_deref(),
            Some("1134dxa")
        );
        let before = set.to_bytes().unwrap();
        let same = ok(set.set("/body/tbl[1]/col[2]", props! { "width" => "1134" }));
        assert_eq!(same.to_bytes().unwrap(), before);

        let removed = ok(set.remove("/body/tbl[1]/col[1]"));
        let removed = reopen(&removed, eager);
        assert_eq!(children(&removed, "/body/tbl[1]/tr[2]"), Some(3));
        assert_eq!(
            prop(&removed, "/body/tbl[1]/col[1]", "width").as_deref(),
            Some("1134")
        );
        assert_eq!(document_xml(&removed).matches("<w:gridCol ").count(), 3);
    }
}

/// The last column of a table is refused, as are a missing column and a
/// zero width (merged slots are refused by the leaf, `grid.rs`).
#[test]
fn the_last_column_and_bad_widths_are_refused() {
    let file = with_table("1", "1");
    let error = file.remove("/body/tbl[1]/col[1]").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
    let error = file.get("/body/tbl[1]/col[2]").unwrap_err();
    assert_eq!(error.code(), Code::PathNotFound, "{error}");
    let error = file
        .add("/body/tbl[1]", NewNode::new("col").prop("width", "0"))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

/// `comment.json`: a reply (`parentId`), resolution (`done`, alias
/// `resolved`) and the anchor's path (`anchoredTo`), through
/// `commentsExtended`.
#[test]
fn comment_replies_resolution_and_anchors() {
    for eager in [true, false] {
        let file = reopen(&with_paragraph("Reviewed text"), eager);
        let root = ok(file.add(
            "/body/p[1]",
            NewNode::new("comment")
                .prop("author", "Ada")
                .prop("text", "Check"),
        ));
        assert_eq!(root.path.to_string(), "/comments/comment[1]");
        let file = reopen(&root.file, eager);
        let root_id = prop(&file, "/comments/comment[1]", "id").unwrap();
        assert_eq!(
            prop(&file, "/comments/comment[1]", "done").as_deref(),
            Some("false")
        );
        assert_eq!(prop(&file, "/comments/comment[1]", "parentId"), None);
        assert_eq!(
            prop(&file, "/comments/comment[1]", "anchoredTo").as_deref(),
            Some("/body/p[1]")
        );
        // A plain comment writes no modern-comment part, as before.
        assert!(part(&file, "word/commentsExtended.xml").is_none());

        let reply = ok(file.add(
            "/body/p[1]",
            NewNode::new("comment")
                .prop("author", "Grace")
                .prop("text", "Fixed")
                .prop("parentId", root_id.as_str())
                .prop("done", "true"),
        ));
        assert_eq!(reply.path.to_string(), "/comments/comment[2]");
        let file = reopen(&reply.file, eager);
        assert_eq!(
            prop(&file, "/comments/comment[2]", "parentId").as_deref(),
            Some(root_id.as_str())
        );
        assert_eq!(
            prop(&file, "/comments/comment[2]", "done").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(&file, "/comments/comment[1]", "done").as_deref(),
            Some("false")
        );
        let extended = part(&file, "word/commentsExtended.xml").expect("commentsExtended");
        assert_eq!(extended.matches("<w15:commentEx ").count(), 2, "{extended}");
        assert!(extended.contains("w15:paraIdParent="), "{extended}");

        // Resolve the root; `resolved` is OfficeCLI's alias.
        let file = reopen(
            &ok(file.set("/comments/comment[1]", props! { "resolved" => true })),
            eager,
        );
        assert_eq!(
            prop(&file, "/comments/comment[1]", "done").as_deref(),
            Some("true")
        );
        let file = reopen(
            &ok(file.set("/comments/comment[1]", props! { "done" => false })),
            eager,
        );
        assert_eq!(
            prop(&file, "/comments/comment[1]", "done").as_deref(),
            Some("false")
        );
        // A parent is set on add only.
        let error = file
            .set("/comments/comment[1]", props! { "parentId" => 0 })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    }
}

/// `runStart` places the range after a run; `range=false` writes a point
/// comment; a reply to a missing comment is refused.
#[test]
fn comment_anchor_controls_and_refusals() {
    let file = with_paragraph("one ");
    let file = ok(file.add("/body/p[1]", NewNode::new("run").prop("text", "two"))).file;
    let ranged = ok(file.add(
        "/body/p[1]",
        NewNode::new("comment")
            .prop("author", "Ada")
            .prop("text", "after one")
            .prop("runStart", "1"),
    ));
    let xml = document_xml(&ranged.file);
    let start = xml.find("<w:commentRangeStart").expect("range start");
    assert!(
        xml.find("one ").unwrap() < start && start < xml.find(">two<").unwrap(),
        "{xml}"
    );

    let point = ok(ranged.file.add(
        "/body/p[1]",
        NewNode::new("comment")
            .prop("author", "Ada")
            .prop("text", "point")
            .prop("range", "false"),
    ));
    let xml = document_xml(&point.file);
    assert_eq!(xml.matches("<w:commentRangeStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:commentReference").count(), 2, "{xml}");
    assert_eq!(
        prop(&point.file, "/comments/comment[2]", "anchoredTo").as_deref(),
        Some("/body/p[1]")
    );

    let before = point.file.to_bytes().unwrap();
    let error = point
        .file
        .add(
            "/body/p[1]",
            NewNode::new("comment")
                .prop("author", "Ada")
                .prop("parentId", "99"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    assert_eq!(point.file.to_bytes().unwrap(), before);
}

/// A comment anchored in a table cell reports the cell paragraph's path.
#[test]
fn a_cell_comment_is_anchored_in_its_cell() {
    let file = with_table("2", "2");
    let file = ok(file.set("/body/tbl[1]/tr[2]/tc[1]", props! { "text" => "cell" }));
    let added = ok(file.add(
        "/body/tbl[1]/tr[2]/tc[1]/p[1]",
        NewNode::new("comment")
            .prop("author", "Ada")
            .prop("text", "here"),
    ));
    assert_eq!(
        prop(&added.file, "/comments/comment[1]", "anchoredTo").as_deref(),
        Some("/body/tbl[1]/tr[2]/tc[1]/p[1]")
    );
}
