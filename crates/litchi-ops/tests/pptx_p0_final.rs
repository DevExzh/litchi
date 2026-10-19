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
