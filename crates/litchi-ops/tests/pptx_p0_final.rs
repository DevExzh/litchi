//! The last PPTX P0 rows (record 0346): legacy comment author, initials and
//! date writes and `direction` on add; modern comment add and set; notes
//! and placeholder text on add; slide title, text and layout writes. Every
//! case runs on both owners.
#![cfg(feature = "pptx")]

use std::io::{Cursor, Read as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode};

#[path = "support/pptx_presentation.rs"]
#[allow(
    dead_code,
    reason = "the shared fixture module serves several test targets"
)]
mod presentation;

fn files(bytes: &[u8]) -> [File; 2] {
    let source = litchi_ops::open_bytes(bytes.to_vec()).expect("source-backed PPTX");
    let eager = litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::default().source_mode(SourceMode::Eager),
    )
    .expect("eager PPTX");
    [source, eager]
}

fn value(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

/// One package member's text.
fn member(file: &File, name: &str) -> Option<String> {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name(name).ok()?;
    let mut text = String::new();
    member.read_to_string(&mut text).unwrap();
    Some(text)
}

#[test]
fn legacy_comment_date_and_attribution_are_set() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let dated = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("date", "2026-10-19T08:00:00Z"),
            )
            .unwrap();
        assert_eq!(
            value(&dated, "/slide[1]/comment[1]", "date").as_deref(),
            Some("2026-10-19T08:00:00Z")
        );
        // Ada's only comment: her author record is renamed in place.
        let renamed = dated
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new()
                    .with("author", "Ada King")
                    .with("initials", "AK"),
            )
            .unwrap();
        assert_eq!(
            value(&renamed, "/slide[1]/comment[1]", "author").as_deref(),
            Some("Ada King")
        );
        assert_eq!(
            value(&renamed, "/slide[1]/comment[1]", "initials").as_deref(),
            Some("AK")
        );
        let authors = member(&renamed, "ppt/commentAuthors.xml").unwrap();
        assert!(authors.contains(r#"name="Ada King""#), "{authors}");
        assert!(!authors.contains("Ada Lovelace"), "{authors}");
        // The text and the other comment are untouched.
        assert_eq!(text(&renamed, "/slide[1]/comment[1]"), "Check the numbers");
        assert_eq!(
            value(&renamed, "/slide[1]/comment[2]", "author").as_deref(),
            Some("Alan Turing")
        );
        // An equal assignment is an exact no-op.
        let same = renamed
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("author", "Ada King"),
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), renamed.to_bytes().unwrap());
    }
}

#[test]
fn a_shared_author_is_forked_not_renamed() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        // A second comment by Ada makes her author record shared.
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "Another"),
            )
            .unwrap()
            .file;
        let comments = |file: &File| member(file, "ppt/comments/comment1.xml").unwrap();
        assert!(comments(&added).contains(r#"authorId="0""#));
        let forked = added
            .set(
                "/slide[1]/comment[3]",
                litchi_ops::Props::new().with("initials", "A."),
            )
            .unwrap();
        // The first comment keeps Ada's record; the third moves to a new
        // author of the same name with the new initials, at its index 1.
        assert_eq!(
            value(&forked, "/slide[1]/comment[1]", "initials").as_deref(),
            Some("AL")
        );
        assert_eq!(
            value(&forked, "/slide[1]/comment[3]", "initials").as_deref(),
            Some("A.")
        );
        assert_eq!(
            value(&forked, "/slide[1]/comment[3]", "author").as_deref(),
            Some("Ada Lovelace")
        );
        let xml = comments(&forked);
        assert_eq!(xml.matches(r#"authorId="2""#).count(), 1, "{xml}");
        assert_eq!(text(&forked, "/slide[1]/comment[3]"), "Another");
        let authors = member(&forked, "ppt/commentAuthors.xml").unwrap();
        assert_eq!(authors.matches("<p:cmAuthor ").count(), 3, "{authors}");
        // The forked author is now the comment's own: a later change
        // renames it in place, as OfficeCLI does.
        let back = forked
            .set(
                "/slide[1]/comment[3]",
                litchi_ops::Props::new().with("initials", "AL"),
            )
            .unwrap();
        assert_eq!(
            value(&back, "/slide[1]/comment[3]", "initials").as_deref(),
            Some("AL")
        );
        assert_eq!(comments(&back), comments(&forked));
    }
}

#[test]
fn legacy_comment_writes_refuse_bad_values() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("date", "yesterday"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("author", ""),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("direction", "rtl"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }
}

#[test]
fn comment_direction_leads_right_to_left_text_with_a_mark() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "مرحبا")
                    .prop("dir", "rtl"),
            )
            .unwrap()
            .file;
        assert_eq!(text(&added, "/slide[1]/comment[3]"), "\u{200F}مرحبا");
        let ltr = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "plain")
                    .prop("direction", "ltr"),
            )
            .unwrap()
            .file;
        assert_eq!(text(&ltr, "/slide[1]/comment[3]"), "plain");
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("direction", "up"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }
}

#[test]
fn modern_comment_reads_initials_resolution_and_parent() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let thread = "/slide[1]/moderncomment[1]";
        assert_eq!(value(&file, thread, "initials").as_deref(), Some("AL"));
        assert_eq!(value(&file, thread, "resolved").as_deref(), Some("true"));
        assert_eq!(value(&file, thread, "parent"), None);
        // The native GUID is raw: not in the view, but a raw query selects
        // by it.
        assert_eq!(value(&file, thread, "id"), None);
        let raw = litchi_ops::query::raw::query(
            "moderncomment[@id=\"{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C02}\"]",
        )
        .unwrap();
        let found: Vec<String> = file
            .query(raw)
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect();
        assert_eq!(found, ["/slide[1]/moderncomment[2]"]);
        assert_eq!(
            value(&file, "/slide[1]/moderncomment[2]", "resolved").as_deref(),
            Some("false")
        );
        let reply = "/slide[1]/moderncomment[1]/reply[1]";
        assert_eq!(value(&file, reply, "initials").as_deref(), Some("GH"));
        assert_eq!(
            value(&file, reply, "parent").as_deref(),
            Some("/slide[1]/moderncomment[1]")
        );
    }
}

fn thread(author: &str, text: &str) -> NewNode<'static> {
    NewNode::new("modernComment")
        .prop("author", author.to_owned())
        .prop("text", text.to_owned())
        .prop("created", "2026-10-19T09:30:00Z")
}

#[test]
fn modern_comments_and_replies_are_added() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let added = file
            .add(
                "/slide[1]",
                thread("Ada Lovelace", "New thread").prop("initials", "AL"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/moderncomment[3]");
        let file2 = added.file;
        let path = "/slide[1]/moderncomment[3]";
        assert_eq!(text(&file2, path), "New thread");
        assert_eq!(
            value(&file2, path, "author").as_deref(),
            Some("Ada Lovelace")
        );
        assert_eq!(value(&file2, path, "resolved").as_deref(), Some("false"));
        assert_eq!(
            value(&file2, path, "created").as_deref(),
            Some("2026-10-19T09:30:00Z")
        );
        // Ada's author record is reused, not duplicated.
        let authors = member(&file2, "ppt/modernAuthors.xml").unwrap();
        assert_eq!(authors.matches("<p188:author ").count(), 2, "{authors}");
        // A new author is created, with initials derived from the name.
        let resolved = file2
            .add(
                "/slide[1]",
                thread("Charles Babbage", "Done").prop("resolved", true),
            )
            .unwrap()
            .file;
        let path = "/slide[1]/moderncomment[4]";
        assert_eq!(value(&resolved, path, "initials").as_deref(), Some("CB"));
        assert_eq!(value(&resolved, path, "resolved").as_deref(), Some("true"));
        let authors = member(&resolved, "ppt/modernAuthors.xml").unwrap();
        assert!(authors.contains(r#"name="Charles Babbage""#), "{authors}");
        // A reply through `parent`.
        let reply = resolved
            .add(
                "/slide[1]",
                thread("Grace Hopper", "Agreed").prop("parent", "/slide[1]/moderncomment[3]"),
            )
            .unwrap();
        assert_eq!(
            reply.path.to_string(),
            "/slide[1]/moderncomment[3]/reply[1]"
        );
        let path = "/slide[1]/moderncomment[3]/reply[1]";
        assert_eq!(text(&reply.file, path), "Agreed");
        assert_eq!(
            value(&reply.file, path, "author").as_deref(),
            Some("Grace Hopper")
        );
        assert_eq!(
            value(&reply.file, path, "parent").as_deref(),
            Some("/slide[1]/moderncomment[3]")
        );
        // Equal inputs produce equal bytes: no clock or random source.
        let again = file
            .add(
                "/slide[1]",
                thread("Ada Lovelace", "New thread").prop("initials", "AL"),
            )
            .unwrap()
            .file;
        assert_eq!(again.to_bytes().unwrap(), file2.to_bytes().unwrap());
        // A new deck gains a comment part and an author part.
        let fresh = file
            .add("/slide[2]", thread("Ada Lovelace", "On slide two"))
            .unwrap()
            .file;
        assert_eq!(text(&fresh, "/slide[2]/moderncomment[1]"), "On slide two");
    }
}

#[test]
fn modern_comment_add_refusals() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let missing = file
            .add(
                "/slide[1]",
                NewNode::new("modernComment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "x"),
            )
            .unwrap_err();
        assert_eq!(missing.code(), Code::MissingProperty);
        assert_eq!(missing.property(), Some("created"));
        let missing = file
            .add(
                "/slide[1]",
                NewNode::new("modernComment")
                    .prop("text", "x")
                    .prop("created", "2026-10-19T09:30:00Z"),
            )
            .unwrap_err();
        assert_eq!(missing.code(), Code::MissingProperty);
        let bad = file
            .add(
                "/slide[1]",
                NewNode::new("modernComment")
                    .prop("author", "Ada Lovelace")
                    .prop("created", "soon"),
            )
            .unwrap_err();
        assert_eq!(bad.code(), Code::InvalidValue);
        let resolved_reply = file
            .add(
                "/slide[1]",
                thread("Ada Lovelace", "x")
                    .prop("parent", "/slide[1]/moderncomment[1]")
                    .prop("resolved", true),
            )
            .unwrap_err();
        assert_eq!(resolved_reply.code(), Code::InvalidArgument);
        let wrong_slide = file
            .add(
                "/slide[2]",
                thread("Ada Lovelace", "x").prop("parent", "/slide[1]/moderncomment[1]"),
            )
            .unwrap_err();
        assert_eq!(wrong_slide.code(), Code::InvalidValue);
        let not_thread = file
            .add(
                "/slide[1]",
                thread("Ada Lovelace", "x").prop("parent", "/slide[1]/comment[1]"),
            )
            .unwrap_err();
        assert_eq!(not_thread.code(), Code::InvalidValue);
    }
}

#[test]
fn modern_comments_and_replies_are_set() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let thread = "/slide[1]/moderncomment[1]";
        let set = file
            .set(
                thread,
                litchi_ops::Props::new()
                    .with("text", "Reworded")
                    .with("resolved", false)
                    .with("created", "2026-10-02T08:00:00Z"),
            )
            .unwrap();
        assert_eq!(text(&set, thread), "Reworded");
        assert_eq!(value(&set, thread, "resolved").as_deref(), Some("false"));
        assert_eq!(value(&set, thread, "status").as_deref(), Some("active"));
        assert_eq!(
            value(&set, thread, "created").as_deref(),
            Some("2026-10-02T08:00:00Z")
        );
        // The reply and the other thread are untouched.
        assert_eq!(
            text(&set, "/slide[1]/moderncomment[1]/reply[1]"),
            "Updated the model"
        );
        assert_eq!(text(&set, "/slide[1]/moderncomment[2]"), "Looks ready");
        // Ada writes only this thread: her record is renamed in place.
        let renamed = file
            .set(thread, litchi_ops::Props::new().with("author", "Ada King"))
            .unwrap();
        assert_eq!(
            value(&renamed, thread, "author").as_deref(),
            Some("Ada King")
        );
        let authors = member(&renamed, "ppt/modernAuthors.xml").unwrap();
        assert_eq!(authors.matches("<p188:author ").count(), 2, "{authors}");
        // Grace writes the reply and the second thread: changing the reply's
        // initials forks a new author for it.
        let reply = "/slide[1]/moderncomment[1]/reply[1]";
        let forked = file
            .set(reply, litchi_ops::Props::new().with("initials", "G."))
            .unwrap();
        assert_eq!(value(&forked, reply, "initials").as_deref(), Some("G."));
        assert_eq!(
            value(&forked, reply, "author").as_deref(),
            Some("Grace Hopper")
        );
        assert_eq!(
            value(&forked, "/slide[1]/moderncomment[2]", "initials").as_deref(),
            Some("GH")
        );
        let authors = member(&forked, "ppt/modernAuthors.xml").unwrap();
        assert_eq!(authors.matches("<p188:author ").count(), 3, "{authors}");
        // Resolution is thread-level.
        let error = file
            .set(reply, litchi_ops::Props::new().with("resolved", true))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        // `parent` is add-only.
        let error = file
            .set(
                thread,
                litchi_ops::Props::new().with("parent", "/slide[1]/moderncomment[2]"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        // Equal values are exact no-ops.
        let same = file
            .set(
                thread,
                litchi_ops::Props::new()
                    .with("text", "Please review the forecast")
                    .with("author", "Ada Lovelace")
                    .with("resolved", true),
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    }
}

/// A deck of the builder's default layouts and no slides.
fn layouts_deck() -> Vec<u8> {
    litchi_pptx::builder::Builder::new()
        .unwrap()
        .build()
        .unwrap()
        .to_bytes()
        .unwrap()
}

fn paths(file: &File, query: &str) -> Vec<String> {
    file.query(query)
        .unwrap()
        .map(|node| node.unwrap().path().to_string())
        .collect()
}

#[test]
fn slide_title_and_text_fill_the_layout_placeholders() {
    let bytes = layouts_deck();
    for file in files(&bytes) {
        let plain = file
            .add(
                "/",
                NewNode::new("slide").prop("layout", "Title and Content"),
            )
            .unwrap()
            .file;
        let added = file
            .add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Title and Content")
                    .prop("title", "Quarterly review")
                    .prop("text", "Revenue\nCosts"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]");
        let file2 = added.file;
        // The layout's own placeholders take the text: no shape is added.
        assert_eq!(
            paths(&file2, "/slide[1]/placeholder"),
            paths(&plain, "/slide[1]/placeholder")
        );
        assert_eq!(text(&file2, "/slide[1]/placeholder[1]"), "Quarterly review");
        assert_eq!(text(&file2, "/slide[1]/placeholder[2]"), "Revenue\nCosts");
        assert_eq!(
            value(&file2, "/slide[1]", "layout").as_deref(),
            Some("Title and Content")
        );
        // The title of a title-slide layout is its centred title.
        let title_slide = file
            .add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Title Slide")
                    .prop("title", "Welcome"),
            )
            .unwrap()
            .file;
        assert_eq!(
            value(&title_slide, "/slide[1]/placeholder[1]", "phType").as_deref(),
            Some("ctrTitle")
        );
        assert_eq!(text(&title_slide, "/slide[1]/placeholder[1]"), "Welcome");
    }
}

#[test]
fn slide_title_and_text_add_placeholders_a_blank_layout_lacks() {
    let bytes = layouts_deck();
    for file in files(&bytes) {
        let added = file
            .add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Blank")
                    .prop("title", "Agenda")
                    .prop("text", "One"),
            )
            .unwrap()
            .file;
        assert_eq!(paths(&added, "/slide[1]/placeholder").len(), 2);
        assert_eq!(
            value(&added, "/slide[1]/placeholder[1]", "phType").as_deref(),
            Some("title")
        );
        assert_eq!(text(&added, "/slide[1]/placeholder[1]"), "Agenda");
        assert_eq!(
            value(&added, "/slide[1]/placeholder[2]", "phType").as_deref(),
            Some("body")
        );
        assert_eq!(text(&added, "/slide[1]/placeholder[2]"), "One");
        // Unbound placeholders get a rectangle, so they render.
        assert!(value(&added, "/slide[1]/placeholder[1]", "width").is_some());
        let slide = member(&added, "ppt/slides/slide1.xml").unwrap();
        assert!(slide.contains(r#"<p:ph type="body" idx="1"/>"#), "{slide}");
    }
}

#[test]
fn slide_layout_is_set() {
    let bytes = layouts_deck();
    for file in files(&bytes) {
        let deck = file
            .add(
                "/",
                NewNode::new("slide").prop("layout", "Title and Content"),
            )
            .unwrap()
            .file;
        let changed = deck
            .set(
                "/slide[1]",
                litchi_ops::Props::new().with("layout", "Title Only"),
            )
            .unwrap();
        assert_eq!(
            value(&changed, "/slide[1]", "layout").as_deref(),
            Some("Title Only")
        );
        // The slide's shapes are kept, as OfficeCLI keeps them.
        assert_eq!(
            paths(&changed, "/slide[1]/placeholder"),
            paths(&deck, "/slide[1]/placeholder")
        );
        let same = changed
            .set(
                "/slide[1]",
                litchi_ops::Props::new().with("layout", "Title Only"),
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), changed.to_bytes().unwrap());
        let missing = deck
            .set(
                "/slide[1]",
                litchi_ops::Props::new().with("layout", "No Such Layout"),
            )
            .unwrap_err();
        assert_eq!(missing.code(), Code::NotFound);
    }
}

#[test]
fn placeholders_are_added_with_text() {
    let bytes = layouts_deck();
    for file in files(&bytes) {
        let deck = file
            .add("/", NewNode::new("slide").prop("layout", "Blank"))
            .unwrap()
            .file;
        let added = deck
            .add(
                "/slide[1]",
                NewNode::new("placeholder")
                    .prop("phType", "title")
                    .prop("text", "Heading"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/shape[1]");
        assert_eq!(text(&added.file, "/slide[1]/placeholder[1]"), "Heading");
        let twice = added
            .file
            .add(
                "/slide[1]",
                NewNode::new("placeholder").prop("type", "title"),
            )
            .unwrap_err();
        assert_eq!(twice.code(), Code::InvalidArgument);
        let footer = added
            .file
            .add(
                "/slide[1]",
                NewNode::new("placeholder")
                    .prop("phType", "footer")
                    .prop("text", "Confidential"),
            )
            .unwrap();
        assert_eq!(footer.path.to_string(), "/slide[1]/shape[2]");
        // The Blank layout's footer slot binds the new placeholder.
        let slide = member(&footer.file, "ppt/slides/slide1.xml").unwrap();
        assert!(slide.contains(r#"type="ftr" idx="11""#), "{slide}");
        assert_eq!(
            text(&footer.file, "/slide[1]/placeholder[2]"),
            "Confidential"
        );
        let missing = deck
            .add("/slide[1]", NewNode::new("placeholder").prop("text", "x"))
            .unwrap_err();
        assert_eq!(missing.code(), Code::MissingProperty);
    }
}

#[test]
fn notes_are_added() {
    let bytes = layouts_deck();
    for file in files(&bytes) {
        let deck = file
            .add("/", NewNode::new("slide").prop("layout", "Blank"))
            .unwrap()
            .file;
        assert!(paths(&deck, "/slide[1]/notes").is_empty());
        let added = deck
            .add(
                "/slide[1]",
                NewNode::new("notes").prop("text", "Emphasize slide 3 data"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/notes");
        assert_eq!(
            text(&added.file, "/slide[1]/notes"),
            "Emphasize slide 3 data"
        );
        // Adding again replaces the text, as OfficeCLI does.
        let again = added
            .file
            .add("/slide[1]", NewNode::new("notes").prop("text", "Second"))
            .unwrap()
            .file;
        assert_eq!(text(&again, "/slide[1]/notes"), "Second");
    }
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let replaced = file
            .add("/slide[1]", NewNode::new("notes").prop("text", "New notes"))
            .unwrap()
            .file;
        assert_eq!(text(&replaced, "/slide[1]/notes"), "New notes");
    }
}

// Record 0346 (record 0333 F3): cell, notes and paragraph-text batches
// coalesce into one native commit.

type Plan = Vec<(String, litchi_ops::Props<'static>)>;

/// The batch's result file, per-item changes and commit count.
fn batch(file: &File, items: &Plan) -> (File, Vec<bool>, usize) {
    let commit = file
        .batch(
            items
                .iter()
                .map(|(path, props)| {
                    let mut item = litchi_ops::BatchItem::new("set");
                    item.path = Some(path.clone().into());
                    item.props = Some(props.clone());
                    item
                })
                .collect::<Vec<_>>(),
            litchi_ops::BatchMode::Atomic,
        )
        .unwrap();
    let changes = commit
        .value()
        .iter()
        .map(|result| match result {
            litchi_ops::ItemResult::Applied(applied) => applied.changed,
            other => panic!("unexpected result {other:?}"),
        })
        .collect();
    (
        commit.file().clone(),
        changes,
        commit.diagnostics().commits(),
    )
}

/// The same items applied one `set` at a time.
fn sequential(file: &File, items: &Plan) -> Vec<u8> {
    let mut current = file.clone();
    for (path, props) in items {
        current = current.set(path.as_str(), props.clone()).unwrap();
    }
    current.to_bytes().unwrap()
}

/// A one-slide deck holding a `size`×`size` table.
fn table_deck(size: usize) -> Vec<u8> {
    let data = (0..size)
        .map(|row| {
            (0..size)
                .map(|column| format!("r{row}c{column}"))
                .collect::<Vec<_>>()
                .join(",")
        })
        .collect::<Vec<_>>()
        .join(";");
    let [_, eager] = files(&layouts_deck());
    let deck = eager
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file;
    deck.add(
        "/slide[1]",
        NewNode::new("table")
            .prop("rows", i64::try_from(size).unwrap())
            .prop("cols", i64::try_from(size).unwrap())
            .prop("data", data),
    )
    .unwrap()
    .file
    .to_bytes()
    .unwrap()
}

fn cell(row: usize, column: usize) -> String {
    format!("/slide[1]/table[1]/tr[{row}]/tc[{column}]")
}

#[test]
fn cell_batches_are_one_commit() {
    let bytes = table_deck(10);
    for file in files(&bytes) {
        let mut items: Plan = Vec::new();
        for row in 1..=10 {
            for column in 1..=10 {
                items.push((
                    cell(row, column),
                    litchi_ops::Props::new()
                        .with("fill", "FFEECC")
                        .with("bold", true),
                ));
            }
        }
        let (after, changes, commits) = batch(&file, &items);
        assert_eq!(commits, 1);
        assert!(changes.iter().all(|changed| *changed));
        assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
        assert_eq!(
            value(&after, &cell(10, 10), "fill").as_deref(),
            Some("#FFEECC")
        );
        let (_, changes, _) = batch(&after, &items);
        assert!(changes.iter().all(|changed| !*changed));

        // Cell paragraphs and runs, and cell text.
        let items: Plan = (1..=10)
            .flat_map(|row| {
                [
                    (
                        format!("{}/paragraph[1]", cell(row, 1)),
                        litchi_ops::Props::new().with("align", "center"),
                    ),
                    (
                        format!("{}/paragraph[1]/run[1]", cell(row, 2)),
                        litchi_ops::Props::new().with("italic", true),
                    ),
                    (
                        cell(row, 3),
                        litchi_ops::Props::new().with("text", format!("new {row}")),
                    ),
                ]
            })
            .collect();
        let (after, changes, commits) = batch(&file, &items);
        assert_eq!(commits, 1);
        assert!(changes.iter().all(|changed| *changed));
        assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
        assert_eq!(text(&after, &cell(4, 3)), "new 4");
    }
}

#[test]
fn overlapping_cell_plans_keep_sequential_semantics() {
    let bytes = table_deck(3);
    for file in files(&bytes) {
        let items: Plan = vec![
            (
                format!("{}/paragraph[1]/run[1]", cell(1, 1)),
                litchi_ops::Props::new().with("bold", true),
            ),
            // The same cell: closes the run first.
            (
                cell(1, 1),
                litchi_ops::Props::new()
                    .with("bold", true)
                    .with("valign", "center"),
            ),
            (
                format!("{}/paragraph[1]", cell(2, 2)),
                litchi_ops::Props::new().with("text", "two\nlines"),
            ),
            (
                format!("{}/paragraph[1]/run[1]", cell(2, 2)),
                litchi_ops::Props::new().with("italic", true),
            ),
        ];
        let (after, changes, commits) = batch(&file, &items);
        assert_eq!(changes, [true, true, true, true]);
        assert!(commits < items.len());
        assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
    }
}

#[test]
fn paragraph_text_batches_are_one_commit() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let items: Plan = vec![
            (
                "/slide[1]/shape[1]/paragraph[1]".into(),
                litchi_ops::Props::new().with("text", "Retitled"),
            ),
            (
                "/slide[1]/shape[2]/paragraph[1]".into(),
                litchi_ops::Props::new()
                    .with("text", "Body\nline")
                    .with("bold", true),
            ),
        ];
        let (after, changes, commits) = batch(&file, &items);
        assert_eq!(commits, 1);
        assert_eq!(changes, [true, true]);
        assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
        let (_, changes, _) = batch(&after, &items);
        assert_eq!(changes, [false, false]);
    }
}

#[test]
fn notes_batches_are_one_commit_on_eager_owners() {
    let bytes = presentation::bytes();
    for (eager, file) in files(&bytes).into_iter().enumerate() {
        let mut file = file
            .add("/slide[1]", NewNode::new("notes").prop("text", "one"))
            .unwrap()
            .file;
        for text in ["two", "three"] {
            file = file
                .add(
                    "/slide[1]/notes",
                    NewNode::new("paragraph").prop("text", text),
                )
                .unwrap()
                .file;
        }
        // A paragraph and its own run overlap, so each kind is one run.
        let items: Plan = (1..=3)
            .map(|paragraph| {
                (
                    format!("/slide[1]/notes/paragraph[{paragraph}]"),
                    litchi_ops::Props::new().with("align", "right"),
                )
            })
            .chain((1..=3).map(|paragraph| {
                (
                    format!("/slide[1]/notes/paragraph[{paragraph}]/run[1]"),
                    litchi_ops::Props::new().with("bold", true),
                )
            }))
            .collect();
        let (after, changes, commits) = batch(&file, &items);
        assert!(changes.iter().all(|changed| *changed));
        assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
        // The notes add promoted the source owner to an eager package: both
        // owners coalesce.
        assert_eq!(commits, 2, "{eager}");
        assert_eq!(
            value(&after, "/slide[1]/notes/paragraph[3]/run[1]", "bold").as_deref(),
            Some("true")
        );
    }
}

#[test]
fn source_owned_notes_keep_the_ordinary_writer() {
    let bytes = presentation::bytes();
    let [source, _] = files(&bytes);
    let items: Plan = vec![
        (
            "/slide[1]/notes/paragraph[1]".into(),
            litchi_ops::Props::new().with("align", "center"),
        ),
        (
            "/slide[1]/notes/paragraph[1]/run[1]".into(),
            litchi_ops::Props::new().with("italic", true),
        ),
    ];
    let (after, changes, _) = batch(&source, &items);
    assert_eq!(changes, [true, true]);
    assert_eq!(after.to_bytes().unwrap(), sequential(&source, &items));
}

// Record 0348: a cell plan's change is the leaf's report from its batched
// staging; the table is indexed, not read, before planning.

/// Each item's change and the file after it, applied one `set` at a time.
fn sequential_changes(file: &File, items: &Plan) -> (Vec<bool>, Vec<u8>) {
    let mut current = file.clone();
    let mut changes = Vec::new();
    for (path, props) in items {
        let commit = current
            .edit(|edit| edit.set(path.as_str(), props.clone()))
            .unwrap();
        changes.push(commit.diagnostics().changed());
        current = commit.file().clone();
    }
    (changes, current.to_bytes().unwrap())
}

#[test]
fn cell_batch_changes_are_the_items_own() {
    let bytes = table_deck(4);
    for file in files(&bytes) {
        let first: Plan = vec![
            (cell(1, 1), litchi_ops::Props::new().with("fill", "FFEECC")),
            (
                format!("{}/paragraph[1]", cell(2, 2)),
                litchi_ops::Props::new().with("align", "center"),
            ),
            (
                format!("{}/paragraph[1]/run[1]", cell(3, 3)),
                litchi_ops::Props::new().with("bold", true),
            ),
        ];
        let (after, ..) = batch(&file, &first);
        // Repeats of the first batch (no change) between new changes.
        let items: Plan = first
            .iter()
            .cloned()
            .chain([
                (cell(1, 2), litchi_ops::Props::new().with("fill", "FFEECC")),
                (
                    format!("{}/paragraph[1]", cell(4, 4)),
                    litchi_ops::Props::new().with("align", "center"),
                ),
                (cell(4, 1), litchi_ops::Props::new().with("text", "r3c0")),
                (cell(4, 2), litchi_ops::Props::new().with("text", "new")),
            ])
            .collect();
        let (expected, sequential) = sequential_changes(&after, &items);
        assert_eq!(
            expected,
            [false, false, false, true, true, false, true],
            "the fixture"
        );
        let (batched, changes, commits) = batch(&after, &items);
        assert_eq!(changes, expected);
        assert_eq!(commits, 1);
        assert_eq!(batched.to_bytes().unwrap(), sequential);
    }
}

#[test]
fn cell_batch_misses_fail_as_the_item_alone() {
    let bytes = table_deck(2);
    for file in files(&bytes) {
        for path in [
            "/slide[1]/table[1]/tr[3]/tc[1]".to_owned(),
            "/slide[1]/table[1]/tr[1]/tc[3]".to_owned(),
            "/slide[1]/table[2]/tr[1]/tc[1]".to_owned(),
            format!("{}/paragraph[2]", cell(1, 1)),
            format!("{}/paragraph[1]/run[2]", cell(1, 1)),
        ] {
            let alone = file
                .set(path.as_str(), litchi_ops::Props::new().with("bold", true))
                .expect_err("a miss fails alone");
            let items: Vec<litchi_ops::BatchItem<'static>> = [cell(2, 2), path.clone()]
                .into_iter()
                .map(|path| {
                    let mut item = litchi_ops::BatchItem::new("set");
                    item.path = Some(path.into());
                    item.props = Some(litchi_ops::Props::new().with("bold", true));
                    item
                })
                .collect();
            let error = file
                .batch(items, litchi_ops::BatchMode::Atomic)
                .expect_err("a miss fails the batch");
            assert_eq!(error.code(), alone.code(), "{path}");
            assert_eq!(error.item(), Some(1), "{path}");
        }
    }
}

#[test]
fn notes_batches_past_one_chunk_equal_the_items() {
    let bytes = presentation::bytes();
    let [_, eager] = files(&bytes);
    let mut file = eager
        .add("/slide[1]", NewNode::new("notes").prop("text", "n0"))
        .unwrap()
        .file;
    for index in 1..300 {
        file = file
            .add(
                "/slide[1]/notes",
                NewNode::new("paragraph").prop("text", format!("n{index}")),
            )
            .unwrap()
            .file;
    }
    let items: Plan = (1..=300)
        .map(|paragraph| {
            (
                format!("/slide[1]/notes/paragraph[{paragraph}]"),
                litchi_ops::Props::new().with("align", "right"),
            )
        })
        .collect();
    let (after, changes, commits) = batch(&file, &items);
    assert_eq!(commits, 1);
    assert!(changes.iter().all(|changed| *changed));
    assert_eq!(after.to_bytes().unwrap(), sequential(&file, &items));
}

#[test]
fn notes_batch_changes_and_misses_are_the_items_own() {
    let bytes = presentation::bytes();
    let [_, eager] = files(&bytes);
    let mut file = eager
        .add("/slide[1]", NewNode::new("notes").prop("text", "one"))
        .unwrap()
        .file;
    file = file
        .add(
            "/slide[1]/notes",
            NewNode::new("paragraph").prop("text", "two"),
        )
        .unwrap()
        .file;
    let notes = |path: &str| format!("/slide[1]/notes/{path}");
    let first: Plan = vec![(
        notes("paragraph[1]"),
        litchi_ops::Props::new().with("align", "right"),
    )];
    let (after, ..) = batch(&file, &first);
    let items: Plan = vec![
        // A repeat: no change.
        (
            notes("paragraph[1]"),
            litchi_ops::Props::new().with("align", "right"),
        ),
        (
            notes("paragraph[2]/run[1]"),
            litchi_ops::Props::new().with("bold", true),
        ),
        (
            notes("paragraph[2]"),
            litchi_ops::Props::new().with("text", "two"),
        ),
    ];
    let (expected, sequential) = sequential_changes(&after, &items);
    assert_eq!(expected.first(), Some(&false));
    let (batched, changes, _) = batch(&after, &items);
    assert_eq!(changes, expected);
    assert_eq!(batched.to_bytes().unwrap(), sequential);
    for path in [notes("paragraph[3]"), notes("paragraph[1]/run[2]")] {
        let alone = after
            .set(path.as_str(), litchi_ops::Props::new().with("bold", true))
            .expect_err("a miss fails alone");
        let items: Vec<litchi_ops::BatchItem<'static>> = [notes("paragraph[1]"), path.clone()]
            .into_iter()
            .map(|path| {
                let mut item = litchi_ops::BatchItem::new("set");
                item.path = Some(path.into());
                item.props = Some(litchi_ops::Props::new().with("bold", true));
                item
            })
            .collect();
        let error = after
            .batch(items, litchi_ops::BatchMode::Atomic)
            .expect_err("a miss fails the batch");
        assert_eq!(error.code(), alone.code(), "{path}");
        assert_eq!(error.item(), Some(1), "{path}");
    }
}
