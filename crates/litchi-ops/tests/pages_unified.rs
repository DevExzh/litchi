//! Unified Pages reads and writes (record 0299): paragraphs (set, add,
//! remove, move, swap), anchored objects, body tables, headers and footers,
//! footnotes, section names, document properties, creation, package
//! directories and coalesced batches, over the leaf's focused transactions.

#![cfg(feature = "pages")]

use std::collections::BTreeMap;
use std::io::{Cursor, Read};

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, File, NewNode, Placement, create, open, open_bytes, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(path)
}

fn read(path: &str) -> File {
    open_bytes(std::fs::read(fixture(path)).expect("read fixture")).expect("open")
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("path")
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).expect("path");
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn zip_members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    (0..archive.len())
        .map(|index| {
            let mut entry = archive.by_index(index).expect("entry");
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).expect("entry bytes");
            (entry.name().to_owned(), bytes)
        })
        .collect()
}

fn pages_comment_input() -> File {
    // The tracked Pages table fixture has the required table topology but no
    // root comment.  Author one through the leaf's semantic transaction so
    // the unified test exercises a populated, reply-free source; host-vs-leaf
    // differential coverage remains in the iWork parity cases.
    let source = std::fs::read(fixture("pages/body-table-cells-native.pages"))
        .expect("read Pages comment source");
    let package = litchi_pages::Package::from_bytes(&source).expect("open Pages source");
    let commit = package
        .edit_body_table_cell_comment(
            litchi_pages::BodyTableSelector::index(0),
            litchi_pages::table::format::At::new(1, 1),
        )
        .expect("select Pages comment cell")
        .set("D7 native Pages root")
        .expect("author Pages root")
        .commit()
        .expect("commit Pages root");
    let mut bytes = Vec::new();
    commit
        .package()
        .write_to(&mut bytes)
        .expect("serialize Pages comment source");
    open_bytes(bytes).expect("open authored Pages source")
}

fn paragraphs(file: &File) -> Vec<String> {
    file.query("p")
        .expect("query")
        .map(|node| {
            node.expect("node")
                .view()
                .text()
                .expect("text")
                .map(std::borrow::Cow::into_owned)
                .unwrap_or_default()
        })
        .collect()
}

#[test]
fn paragraphs_add_remove_move_and_swap() {
    let file = read("pages/basic.pages");
    assert_eq!(
        paragraphs(&file),
        [
            "Litchi native Pages fixture",
            "Buffa lazy-view migration verification",
            "2026-08-07"
        ]
    );
    let litchi_ops::Added {
        file: added, path, ..
    } = file
        .add(
            "/body",
            NewNode::new("p")
                .prop("text", "Inserted second")
                .at(Placement::Position(litchi_core::Position::new(1))),
        )
        .expect("add");
    assert_eq!(path.to_string(), "/body/p[2]");
    assert_eq!(
        text(&reopen(&added), "/body/p[2]").as_deref(),
        Some("Inserted second")
    );
    let litchi_ops::Added {
        file: appended,
        path,
        ..
    } = added
        .add("/body", NewNode::new("paragraph").prop("text", "Last"))
        .expect("append");
    assert_eq!(path.to_string(), "/body/p[5]");

    let removed = appended.remove("/body/p[1]").expect("remove");
    assert_eq!(
        paragraphs(&reopen(&removed)),
        [
            "Inserted second",
            "Buffa lazy-view migration verification",
            "2026-08-07",
            "Last"
        ]
    );
    let moved = removed
        .move_to("/body/p[4]", Placement::Start)
        .expect("move");
    assert_eq!(paragraphs(&moved)[0], "Last");
    let swapped = moved.swap("/body/p[1]", "/body/p[4]").expect("swap");
    let reopened = reopen(&swapped);
    assert_eq!(paragraphs(&reopened)[0], "2026-08-07");
    assert_eq!(paragraphs(&reopened)[3], "Last");
    // The source handle is untouched.
    assert_eq!(paragraphs(&file).len(), 3);
}

#[test]
fn paragraphs_move_across_sections() {
    let file = read("pages/body-sections-unicode.pages");
    let moved = file
        .move_to("/body/p[1]", Placement::End)
        .expect("move to the second section");
    let reopened = reopen(&moved);
    assert_eq!(
        paragraphs(&reopened),
        [
            "First section marker.",
            "Second section marker — end.",
            "Pages borrowed body 😀"
        ]
    );
    assert_eq!(
        prop(&reopened, "/body/p[3]", "section").as_deref(),
        Some("2")
    );
}

#[test]
fn anchored_objects_survive_paragraph_writes_and_go_with_their_paragraph() {
    let file = read("pages/body-table-read-native.pages");
    let anchored = file
        .query("p[objects>0]")
        .expect("query")
        .next()
        .expect("an anchored paragraph")
        .expect("node")
        .path()
        .to_string();
    let edited = file
        .set(&anchored, props! { "text" => "Caption" })
        .expect("set keeps the table");
    let reopened = reopen(&edited);
    assert_eq!(text(&reopened, &anchored).as_deref(), Some("Caption"));
    assert_eq!(prop(&reopened, &anchored, "objects").as_deref(), Some("1"));
    assert_eq!(
        prop(&reopened, "/", "tableCount"),
        prop(&file, "/", "tableCount")
    );
    // A paragraph whose anchors are body tables goes with its tables
    // (record 0301).
    let tables: i64 = prop(&file, "/", "tableCount")
        .expect("table count")
        .parse()
        .expect("count");
    let paragraphs: i64 = prop(&file, "/", "paragraphCount")
        .expect("paragraph count")
        .parse()
        .expect("count");
    let removed = file.remove(&anchored).expect("remove with its table");
    let reopened = reopen(&removed);
    assert_eq!(
        prop(&reopened, "/", "tableCount"),
        Some((tables - 1).to_string())
    );
    assert_eq!(
        prop(&reopened, "/", "paragraphCount"),
        Some((paragraphs - 1).to_string())
    );
    // Every other paragraph of the 16 corpus files that 0296 refused now
    // takes a set (see `corpus_fidelity`); a plain one here.
    let first = file
        .set("/body/p[1]", props! { "text" => "Retitled" })
        .expect("set");
    assert_eq!(
        text(&reopen(&first), "/body/p[1]").as_deref(),
        Some("Retitled")
    );
}

#[test]
fn tables_read_cells_and_remove() {
    let file = read("pages/body-table-cells-native.pages");
    let rows = prop(&file, "/body/tbl[1]", "rows").expect("rows");
    assert!(rows.parse::<usize>().expect("count") > 0);
    let cells: Vec<_> = file
        .query("tc")
        .expect("query")
        .map(|node| node.expect("node").kind())
        .collect();
    assert!(!cells.is_empty());
    assert!(
        file.query("tc[type=text]").expect("query").next().is_some(),
        "the fixture holds text cells"
    );
    // The table takes part in a formula graph; a value write refreshes the
    // formulas that read it (record 0376).
    let written = file
        .set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "x" })
        .expect("a value write into a formula table");
    assert_eq!(
        prop(&reopen(&written), "/body/tbl[1]/tr[1]/tc[1]", "text").as_deref(),
        Some("x")
    );
    let tables = prop(&file, "/", "tableCount").expect("tables");
    let removed = file.remove("/body/tbl[1]").expect("remove the table");
    let reopened = reopen(&removed);
    let left: usize = prop(&reopened, "/", "tableCount")
        .expect("tables")
        .parse()
        .expect("count");
    assert_eq!(left + 1, tables.parse::<usize>().expect("count"));
}

#[test]
fn body_table_cell_comments_round_trip_through_the_unified_route() {
    let file = pages_comment_input();
    let cell = file
        .query("tc")
        .expect("query")
        .filter_map(|node| node.ok())
        .find(|node| {
            let path = node.path().to_string();
            prop(&file, &path, "comment").is_some()
        })
        .expect("a native cell comment")
        .path()
        .to_string();
    let original = file.to_bytes().expect("source bytes");
    let existing = prop(&file, &cell, "comment").expect("comment text");

    // A repeated root-comment write is an exact no-op, including the ZIP.
    let same = file
        .set(&cell, props! { "comment" => existing.as_str() })
        .expect("same comment");
    assert_eq!(same.to_bytes().expect("same bytes"), original);

    let commit = file
        .edit(|tx| tx.set(&cell, props! { "comment" => "D7 unified Pages root" }))
        .expect("set comment");
    let edited = reopen(commit.file());
    assert_eq!(
        prop(&edited, &cell, "comment").as_deref(),
        Some("D7 unified Pages root")
    );

    // Null is the canonical clear operation.  The inverse is source-bound
    // and restores the exact package, including untouched comment metadata.
    let cleared = edited
        .set(&cell, props! { "comment" => litchi_ops::PropValue::Null })
        .expect("clear comment");
    assert_eq!(prop(&reopen(&cleared), &cell, "comment"), None);
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().expect("inverse bytes"), original);

    // A selector outside the table is a terminal focused-route refusal.
    let error = file
        .set(
            "/body/tbl[999]/tr[1]/tc[1]",
            props! { "comment" => "must refuse" },
        )
        .expect_err("missing table");
    assert_eq!(error.code(), Code::PathNotFound);

    // The source handle and an unrelated cell remain unchanged.
    assert_eq!(file.to_bytes().expect("source unchanged"), original);
    let other = file
        .query("tc")
        .expect("query")
        .filter_map(|node| node.ok())
        .map(|node| node.path().to_string())
        .find(|path| path != &cell)
        .expect("another cell");
    assert_eq!(
        prop(&edited, &other, "text"),
        prop(&file, &other, "text"),
        "comment write stays local"
    );
}

#[test]
fn pages_cell_comment_schema_is_clearable() {
    let property = litchi_ops::schema(FileFormat::Pages)
        .element("cell")
        .expect("cell schema")
        .property("comment")
        .expect("comment schema");
    assert!(property.get());
    assert!(property.set());
    assert!(property.clearable());
    assert!(!property.add());
}

#[test]
fn body_table_cell_comment_creates_a_root_from_a_rootless_file_cell() {
    let file = read("pages/body-table-cells-native.pages");
    let cell = "/body/tbl[1]/tr[2]/tc[2]";
    assert_eq!(prop(&file, cell, "comment"), None);
    let original = file.to_bytes().expect("source bytes");
    let added = file
        .set(cell, props! { "comment" => "D7 Pages created root" })
        .expect("create root comment");
    assert_eq!(
        prop(&reopen(&added), cell, "comment").as_deref(),
        Some("D7 Pages created root")
    );
    let commit = file
        .edit(|tx| tx.set(cell, props! { "comment" => "D7 Pages inverse root" }))
        .expect("create root in a transaction");
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse root creation");
    assert_eq!(restored.to_bytes().expect("inverse bytes"), original);
}

#[test]
fn body_table_cell_comment_stale_recipient_and_batch_proofs_are_terminal_and_local() {
    let file = read("pages/body-table-comment-hidden-native-saved.pages");
    let populated = "/body/tbl[1]/tr[2]/tc[2]";
    let first = "/body/tbl[1]/tr[1]/tc[1]";
    let second = "/body/tbl[1]/tr[1]/tc[2]";
    assert!(prop(&file, populated, "comment").is_some());
    assert!(prop(&file, first, "comment").is_none());
    assert!(prop(&file, second, "comment").is_none());
    let original = file.to_bytes().expect("source bytes");

    let left = file
        .edit(|tx| tx.set(first, props! { "comment" => "D7 stale left" }))
        .expect("left edit");
    let right = file
        .edit(|tx| tx.set(first, props! { "comment" => "D7 stale right" }))
        .expect("right edit");
    let recipient = right.file().to_bytes().expect("recipient bytes");
    let stale = left
        .patch()
        .apply(right.file())
        .expect_err("stale patch recipient");
    assert_eq!(stale.code(), Code::InvalidInput);
    assert_eq!(
        right.file().to_bytes().expect("recipient unchanged"),
        recipient
    );

    let invalid = file
        .edit(|tx| {
            tx.set(first, props! { "comment" => "must not publish" })?;
            tx.set(
                "/body/tbl[999]/tr[1]/tc[1]",
                props! { "comment" => "must refuse" },
            )
        })
        .expect_err("batch failure is atomic");
    assert_eq!(invalid.code(), Code::PathNotFound);
    assert_eq!(file.to_bytes().expect("source unchanged"), original);

    let before_comment = litchi_pages::Package::from_bytes(&original)
        .expect("leaf source")
        .body_table_cell_comment(
            litchi_pages::BodyTableSelector::index(0),
            litchi_pages::table::format::At::new(1, 1),
        )
        .expect("untouched comment read")
        .expect("untouched comment");
    let batch = file
        .edit(|tx| {
            tx.set(first, props! { "comment" => "D7 batch one" })?;
            tx.set(second, props! { "comment" => "D7 batch two" })
        })
        .expect("multi-cell batch");
    let reopened = reopen(batch.file());
    assert_eq!(
        prop(&reopened, first, "comment").as_deref(),
        Some("D7 batch one")
    );
    assert_eq!(
        prop(&reopened, second, "comment").as_deref(),
        Some("D7 batch two")
    );
    let after_bytes = reopened.to_bytes().expect("batch bytes");
    let after_comment = litchi_pages::Package::from_bytes(&after_bytes)
        .expect("reopen leaf")
        .body_table_cell_comment(
            litchi_pages::BodyTableSelector::index(0),
            litchi_pages::table::format::At::new(1, 1),
        )
        .expect("untouched comment reread")
        .expect("untouched comment after batch");
    assert_eq!(after_comment, before_comment, "untouched metadata/replies");
    let before_members = zip_members(&original);
    let after_members = zip_members(&after_bytes);
    assert_eq!(
        before_members.get("preview.jpg"),
        after_members.get("preview.jpg"),
        "untouched package member"
    );
}

#[test]
fn section_names_and_text() {
    let file = read("pages/body-sections-unicode.pages");
    let named = file
        .set("/section[1]", props! { "name" => "Opening" })
        .expect("name");
    let reopened = reopen(&named);
    assert_eq!(
        prop(&reopened, "/section[1]", "name").as_deref(),
        Some("Opening")
    );
    let error = file
        .set("/section[1]", props! { "paragraphCount" => 3 })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn headers_and_footers_are_read_and_written() {
    let file = read("pages/basic.pages");
    // One header and one footer per template (first, even, odd) and slot.
    assert_eq!(file.query("header").expect("query").count(), 9);
    assert_eq!(file.query("footer").expect("query").count(), 9);
    assert_eq!(
        prop(&file, "/section[1]/header[1]", "template").as_deref(),
        Some("first")
    );
    let edited = file
        .set("/section[1]/footer[3]", props! { "text" => "Page footer" })
        .expect("footer set");
    let reopened = reopen(&edited);
    assert_eq!(
        text(&reopened, "/section[1]/footer[3]").as_deref(),
        Some("Page footer")
    );
    assert_eq!(paragraphs(&reopened), paragraphs(&file));
}

#[test]
fn footnotes_are_added_written_and_removed() {
    let file = read("pages/basic.pages");
    assert_eq!(prop(&file, "/footnotes", "count").as_deref(), Some("0"));
    let litchi_ops::Added {
        file: noted, path, ..
    } = file
        .add(
            "/body/p[2]",
            NewNode::new("footnote").prop("text", "A note"),
        )
        .expect("add footnote");
    assert_eq!(path.to_string(), "/footnotes/footnote[1]");
    let reopened = reopen(&noted);
    assert_eq!(
        text(&reopened, "/footnotes/footnote[1]").as_deref(),
        Some("A note")
    );
    // The anchor sits at the end of the paragraph, which keeps its text.
    assert_eq!(
        text(&reopened, "/body/p[2]").as_deref(),
        Some("Buffa lazy-view migration verification")
    );
    assert_eq!(
        prop(&reopened, "/body/p[2]", "objects").as_deref(),
        Some("1")
    );
    let rewritten = noted
        .set("/footnotes/footnote[1]", props! { "text" => "Revised" })
        .expect("set footnote");
    assert_eq!(
        text(&reopen(&rewritten), "/footnotes/footnote[1]").as_deref(),
        Some("Revised")
    );
    let removed = rewritten.remove("/footnotes/footnote[1]").expect("remove");
    let reopened = reopen(&removed);
    assert_eq!(prop(&reopened, "/footnotes", "count").as_deref(), Some("0"));
    assert_eq!(paragraphs(&reopened), paragraphs(&file));
}

#[test]
fn metadata_writes_through_the_metadata_path() {
    let file = read("pages/basic.pages");
    let edited = file
        .set(
            "/metadata",
            props! { "title" => "Quarterly", "author" => "Ryker", "keywords" => "q3" },
        )
        .expect("metadata");
    let metadata = reopen(&edited).metadata().expect("read").expect("present");
    assert_eq!(metadata.title.as_deref(), Some("Quarterly"));
    assert_eq!(metadata.author.as_deref(), Some("Ryker"));
    assert_eq!(metadata.keywords, ["q3"]);
    let cleared = edited
        .set(
            "/metadata",
            props! { "title" => litchi_ops::PropValue::Null },
        )
        .expect("clear");
    assert_eq!(cleared.metadata().unwrap().unwrap().title, None);
    let error = file
        .set("/metadata", props! { "subject" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .set("/metadata", props! { "bogus" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn create_builds_the_minimal_document() {
    let file = create(FileFormat::Pages).expect("create");
    assert!(matches!(file, File::Pages(_)));
    assert_eq!(paragraphs(&file), [""]);
    let edited = file
        .set("/body/p[1]", props! { "text" => "Hello" })
        .expect("set");
    let edited = edited
        .add("/body", NewNode::new("p").prop("text", "World"))
        .expect("add")
        .file;
    assert_eq!(paragraphs(&reopen(&edited)), ["Hello", "World"]);
    assert_eq!(
        create(FileFormat::Pages).unwrap().to_bytes().unwrap(),
        file.to_bytes().unwrap(),
        "creation is deterministic"
    );
    let capabilities = Capabilities::of(FileFormat::Pages);
    for verb in [
        Verb::Create,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Set,
    ] {
        assert!(capabilities.verb(verb).is_supported(), "{verb:?}");
    }
}

#[test]
fn package_directories_open_save_in_place_and_save_as_single_files() {
    let file = open(fixture("directory/pages/basic.pages")).expect("open directory");
    assert!(matches!(file, File::Pages(_)));
    assert_eq!(
        paragraphs(&file),
        paragraphs(&read("pages/basic.pages")),
        "the directory reads like its single-file form"
    );
    let edited = file
        .set("/body/p[1]", props! { "text" => "From a directory" })
        .expect("set");
    let directory = tempfile::tempdir().expect("tempdir");
    let path = directory.path().join("converted.pages");
    edited.save_as(&path).expect("save as");
    let reopened = open(&path).expect("reopen");
    assert_eq!(
        text(&reopened, "/body/p[1]").as_deref(),
        Some("From a directory")
    );
    assert!(path.is_file());

    // `save()` writes a directory-opened package back as a directory
    // (record 0301), and the bundle reopens to the same bytes.
    let bundle = directory.path().join("bundle.pages");
    copy_tree(&fixture("directory/pages/basic.pages"), &bundle);
    let file = open(&bundle).expect("open copy");
    let edited = file
        .set("/body/p[1]", props! { "text" => "Saved in place" })
        .expect("set");
    edited.save().expect("save in place");
    assert!(bundle.join("Index.zip").is_file());
    let reopened = open(&bundle).expect("reopen bundle");
    assert_eq!(
        text(&reopened, "/body/p[1]").as_deref(),
        Some("Saved in place")
    );
    assert_eq!(reopened.to_bytes().unwrap(), edited.to_bytes().unwrap());
}

fn copy_tree(from: &std::path::Path, to: &std::path::Path) {
    std::fs::create_dir_all(to).expect("mkdir");
    for entry in std::fs::read_dir(from).expect("read dir") {
        let entry = entry.expect("entry");
        let target = to.join(entry.file_name());
        if entry.file_type().expect("type").is_dir() {
            copy_tree(&entry.path(), &target);
        } else {
            std::fs::copy(entry.path(), target).expect("copy");
        }
    }
}

#[test]
fn a_batch_of_paragraph_sets_is_one_leaf_transaction() {
    let file = read("pages/basic.pages");
    let commit = file
        .edit(|tx| {
            for round in 0..100 {
                for index in 1..=3 {
                    tx.set(
                        &format!("/body/p[{index}]"),
                        props! { "text" => format!("round {round} paragraph {index}") },
                    )?;
                }
            }
            Ok(())
        })
        .expect("edit");
    assert_eq!(commit.diagnostics().operations(), 300);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(
        paragraphs(&reopen(commit.file())),
        [
            "round 99 paragraph 1",
            "round 99 paragraph 2",
            "round 99 paragraph 3"
        ]
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn a_batch_mixing_structure_and_sets_keeps_its_order() {
    let file = read("pages/basic.pages");
    let commit = file
        .edit(|tx| {
            tx.set("/body/p[1]", props! { "text" => "A" })?;
            tx.add("/body", NewNode::new("p").prop("text", "B"))?;
            tx.set("/body/p[4]", props! { "text" => "C" })?;
            tx.remove("/body/p[2]")?;
            Ok(())
        })
        .expect("edit");
    assert_eq!(paragraphs(&reopen(commit.file())), ["A", "2026-08-07", "C"]);
}

#[test]
fn paragraph_fonts_are_written() {
    let file = read("pages/basic.pages");
    let styled = file
        .set(
            "/body/p[1]",
            props! { "font.bold" => true, "font.size" => "20pt", "font.color" => "0000FF" },
        )
        .expect("style a paragraph");
    let bytes = styled.to_bytes().expect("publish");
    let package = litchi_pages::Package::from_bytes(&bytes).expect("leaf");
    let runs = package.section_text_runs(0).expect("runs");
    assert_eq!(runs[0].style().bold, Some(true));
    assert_eq!(runs[0].style().size, Some(20.0));
    assert_eq!(runs[0].style().color, Some([0, 0, 255]));
    assert_eq!(
        text(&reopen(&styled), "/body/p[1]"),
        text(&file, "/body/p[1]")
    );
    let commit = file
        .edit(|tx| tx.set("/body/p[1]", props! { "font.italic" => true }))
        .expect("italic");
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    let error = file
        .set("/body/p[1]", props! { "font.size" => "-3pt" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn body_table_cells_take_values() {
    let file = read("pages/body-table-catalog-native.pages");
    let written = file
        .edit(|tx| {
            tx.set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "Header" })?;
            tx.set("/body/tbl[1]/tr[2]/tc[2]", props! { "value" => 12.5 })?;
            Ok(())
        })
        .expect("cell values");
    let reopened = reopen(written.file());
    assert_eq!(
        text(&reopened, "/body/tbl[1]/tr[1]/tc[1]").as_deref(),
        Some("Header")
    );
    assert_eq!(
        prop(&reopened, "/body/tbl[1]/tr[2]/tc[2]", "type").as_deref(),
        Some("number")
    );
    let restored = written
        .patch()
        .inverse()
        .apply(written.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    let cleared = written
        .file()
        .set(
            "/body/tbl[1]/tr[1]/tc[1]",
            props! { "value" => litchi_ops::PropValue::Null },
        )
        .expect("clear");
    assert_eq!(
        prop(&reopen(&cleared), "/body/tbl[1]/tr[1]/tc[1]", "type").as_deref(),
        Some("empty")
    );
}

#[test]
fn sections_are_added_and_removed() {
    // Record 0312 (request 11): a section break and a copy of the page
    // masters, headers and footers of the section it follows.
    let file = read("pages/basic.pages");
    let litchi_ops::Added {
        file: added, path, ..
    } = file
        .add("/", NewNode::new("section").prop("name", "Appendix"))
        .expect("add section");
    assert_eq!(path.to_string(), "/section[2]");
    let reopened = reopen(&added);
    assert_eq!(prop(&reopened, "/", "sectionCount").as_deref(), Some("2"));
    assert_eq!(
        prop(&reopened, "/section[2]", "name").as_deref(),
        Some("Appendix")
    );
    assert_eq!(text(&reopened, "/section[2]").as_deref(), Some(""));
    assert_eq!(text(&reopened, "/section[1]"), text(&file, "/section[1]"));
    assert_eq!(
        reopened
            .query("section[2] > footer")
            .expect("query")
            .count(),
        file.query("section[1] > footer").expect("query").count()
    );
    // The first section cannot be removed, nor preceded.
    let error = reopened.remove("/section[1]").expect_err("first section");
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add("/", NewNode::new("section").at(Placement::Start))
        .expect_err("before the first section");
    assert_eq!(error.code(), Code::InvalidArgument);
    // Text written into the new section joins the first one on removal.
    let written = reopened
        .set("/section[2]", props! { "text" => "Tail" })
        .expect("set");
    let removed = written.remove("/section[2]").expect("remove section");
    let removed = reopen(&removed);
    assert_eq!(prop(&removed, "/", "sectionCount").as_deref(), Some("1"));
    assert_eq!(
        text(&removed, "/section[1]"),
        text(&file, "/section[1]").map(|text| format!("{text}Tail"))
    );
}

#[test]
fn body_tables_are_added_with_seeded_text() {
    let file = read("pages/body-table-visible.pages");
    let tables: usize = prop(&file, "/", "tableCount")
        .expect("tables")
        .parse()
        .expect("count");
    let paragraphs_before = paragraphs(&file).len();
    let added = file
        .add(
            "/body",
            NewNode::new("tbl")
                .prop("name", "Budget")
                .prop("data", "Item,Cost;Paper,3")
                .at(Placement::Position(litchi_core::Position::new(0))),
        )
        .expect("table");
    let reopened = reopen(&added.file);
    let path = added.path.to_string();
    assert!(path.starts_with("/body/tbl["), "{path}");
    assert_eq!(prop(&reopened, &path, "name").as_deref(), Some("Budget"));
    assert_eq!(prop(&reopened, &path, "rows").as_deref(), Some("2"));
    assert_eq!(prop(&reopened, &path, "columns").as_deref(), Some("2"));
    assert_eq!(
        text(&reopened, &format!("{path}/tr[2]/tc[1]")).as_deref(),
        Some("Paper")
    );
    assert_eq!(
        prop(&reopened, "/", "tableCount").as_deref(),
        Some((tables + 1).to_string().as_str())
    );
    // The table anchors in a paragraph of its own, at the placement.
    assert_eq!(paragraphs(&reopened).len(), paragraphs_before + 1);
    assert_eq!(
        prop(&reopened, "/body/p[1]", "objects").as_deref(),
        Some("1")
    );
    let edited = file
        .edit(|tx| {
            tx.add("/body", NewNode::new("tbl").prop("rows", 2).prop("cols", 2))?;
            Ok(())
        })
        .expect("edit");
    let restored = edited
        .patch()
        .inverse()
        .apply(edited.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    // Removing the new table's paragraph removes the table with it.
    let removed = reopened.remove("/body/p[1]").expect("remove");
    assert_eq!(
        prop(&reopen(&removed), "/", "tableCount").as_deref(),
        Some(tables.to_string().as_str())
    );
    // A document without a body table has no structure to copy.
    let error = read("pages/basic.pages")
        .add("/body", NewNode::new("tbl").prop("rows", 2).prop("cols", 2))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .add(
            "/body/p[1]",
            NewNode::new("tbl").prop("rows", 2).prop("cols", 2),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
}
