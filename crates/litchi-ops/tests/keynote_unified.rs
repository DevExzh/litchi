//! Unified Keynote reads and writes (record 0298): slides with their
//! layouts, placeholders, speaker notes, paragraphs, runs, drawables and
//! tables; `set` of texts, paragraphs, frames and visibility; `add` of
//! slides (from a layout or as a copy), paragraphs, text boxes, shapes and
//! pictures; `remove`, `move` and `swap`; creation; coalesced batches and
//! their atomicity.

#![cfg(feature = "keynote")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Placement, create, open_bytes, props};

/// A 2 x 3 RGB PNG as a data URI.
const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAAA2iEnWAAAAEUlEQVR4nGP4z8AARGACTgMAPtYF+1iCz0YAAAAASUVORK5CYII=";

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote")
        .join(name)
}

fn read(name: &str) -> File {
    open_bytes(std::fs::read(fixture(name)).expect("read fixture")).expect("open")
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn count(file: &File, selector: &str) -> usize {
    file.query(selector).expect("query").count()
}

fn texts(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
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

fn titles(file: &File) -> Vec<String> {
    (1..=count(file, "slide"))
        .map(|slide| prop(file, &format!("/slide[{slide}]"), "title").unwrap_or_default())
        .collect()
}

/// How a frame property prints `points`.
fn pt(points: i32) -> String {
    litchi_ops::value::Length::from_points(points).to_string()
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

const BODY: &str = "/slide[2]/placeholder[body]";

#[test]
fn reads_paragraphs_runs_drawables_and_layouts() {
    let file = read("multi-slide.key");
    assert!(matches!(file, File::Keynote(_)));
    assert_eq!(count(&file, "slide"), 5);
    assert!(prop(&file, "/slide[2]", "layout").is_some());
    assert_eq!(
        texts(&file, &format!("{BODY}/paragraph")),
        ["First topic", "Second topic", "Third topic"]
    );
    assert_eq!(
        text(&file, &format!("{BODY}/p[2]")).as_deref(),
        Some("Second topic")
    );
    assert_eq!(
        texts(&file, "/slide[1]/notes/paragraph"),
        ["Welcome everyone."]
    );
    assert_eq!(
        text(&file, "/slide[1]/shape[1]").as_deref(),
        Some("2026-08-07")
    );
    assert_eq!(
        prop(&file, "/slide[1]/shape[1]", "zorder").as_deref(),
        Some("0")
    );
    assert!(prop(&file, "/slide[1]/shape[1]", "x").is_some());
    assert!(prop(&file, "/slide[1]/placeholder[title]", "width").is_some());
    // Character runs join under their paragraphs.
    assert_eq!(
        texts(&file, "/slide[1]/placeholder[title]/paragraph[1]/run"),
        ["Litchi multi-slide fixture"]
    );
    assert_eq!(
        texts(&file, &format!("{BODY}/paragraph[3]/run")),
        ["Third topic"]
    );

    let tables = read("slide-table-read-native.key");
    assert_eq!(
        prop(&tables, "/slide[1]/table[1]", "rows").as_deref(),
        Some("5")
    );
    assert_eq!(
        prop(&tables, "/slide[1]/table[1]", "cols").as_deref(),
        Some("4")
    );
    assert_eq!(
        text(&tables, "/slide[1]/table[1]/tr[2]/tc[1]").as_deref(),
        Some("Café 北京")
    );
}

#[test]
fn sets_texts_paragraphs_frames_and_visibility_and_round_trips() {
    let file = read("multi-slide.key");
    let edited = file
        .set("/slide[1]/notes", props! { "text" => "Hello, audience." })
        .expect("notes")
        .set(
            &format!("{BODY}/p[2]"),
            props! { "text" => "Changed topic" },
        )
        .expect("paragraph")
        .set("/slide[1]/shape[1]", props! { "text" => "Q4\nreview" })
        .expect("shape text")
        .set(
            "/slide[1]/shape[1]",
            props! { "x" => "72pt", "width" => "300pt" },
        )
        .expect("frame")
        .set("/slide[1]/placeholder[title]", props! { "height" => "2in" })
        .expect("placeholder frame")
        .set("/slide[4]", props! { "hidden" => true })
        .expect("hidden");
    let reopened = reopen(&edited);
    assert_eq!(
        text(&reopened, "/slide[1]/notes").as_deref(),
        Some("Hello, audience.")
    );
    assert_eq!(
        texts(&reopened, &format!("{BODY}/paragraph")),
        ["First topic", "Changed topic", "Third topic"]
    );
    assert_eq!(
        texts(&reopened, "/slide[1]/shape[1]/paragraph"),
        ["Q4", "review"]
    );
    assert_eq!(prop(&reopened, "/slide[1]/shape[1]", "x"), Some(pt(72)));
    assert_eq!(
        prop(&reopened, "/slide[1]/shape[1]", "width"),
        Some(pt(300))
    );
    assert_eq!(
        prop(&reopened, "/slide[1]/placeholder[title]", "height"),
        Some(pt(144))
    );
    assert_eq!(
        prop(&reopened, "/slide[4]", "hidden").as_deref(),
        Some("true")
    );
}

#[test]
fn exact_noop_sets_return_the_same_bytes() {
    let file = read("multi-slide.key");
    let bytes = file.to_bytes().unwrap();
    for (path, props) in [
        (
            "/slide[1]/placeholder[title]",
            props! { "text" => "Litchi multi-slide fixture" },
        ),
        (
            "/slide[2]/placeholder[body]/p[1]",
            props! { "text" => "First topic" },
        ),
        ("/slide[1]/notes", props! { "text" => "Welcome everyone." }),
        ("/slide[1]/shape[1]", props! { "text" => "2026-08-07" }),
        ("/slide[3]", props! { "hidden" => true }),
    ] {
        let edited = file.set(path, props).expect(path);
        assert_eq!(edited.to_bytes().unwrap(), bytes, "{path}");
    }
    // A frame set to the frame it has is a no-op too.
    let moved = file
        .set("/slide[1]/shape[1]", props! { "x" => "72pt" })
        .expect("move");
    let again = moved
        .set(
            "/slide[1]/shape[1]",
            props! { "x" => "72pt", "rotation" => 0 },
        )
        .expect("same frame");
    assert_eq!(again.to_bytes().unwrap(), moved.to_bytes().unwrap());
}

#[test]
fn adds_slides_from_layouts_and_copies() {
    let file = read("multi-slide.key");
    let layout = prop(&file, "/slide[2]", "layout").expect("layout");
    let added = file
        .add(
            "/",
            NewNode::new("slide")
                .prop("layout", layout.as_str())
                .at(Placement::Position(litchi_core::Position::new(1))),
        )
        .expect("add from layout");
    assert_eq!(added.path.to_string(), "/slide[2]");
    let reopened = reopen(&added.file);
    assert_eq!(count(&reopened, "slide"), 6);
    assert_eq!(
        prop(&reopened, "/slide[2]", "layout").as_deref(),
        Some(layout.as_str())
    );
    // The theme's default layout, at the end.
    let added = file.add("/", NewNode::new("slide")).expect("add default");
    assert_eq!(added.path.to_string(), "/slide[6]");
    // A copy of the agenda slide, after it.
    let copied = file
        .add(
            "/",
            NewNode::new("slide")
                .prop("from", "/slide[2]")
                .at(Placement::after("/slide[2]").unwrap()),
        )
        .expect("duplicate");
    assert_eq!(copied.path.to_string(), "/slide[3]");
    let reopened = reopen(&copied.file);
    assert_eq!(
        titles(&reopened)[..4],
        [
            "Litchi multi-slide fixture",
            "Agenda",
            "Agenda",
            "Hidden section"
        ]
    );
    assert_eq!(
        texts(&reopened, "/slide[3]/placeholder[body]/paragraph"),
        ["First topic", "Second topic", "Third topic"]
    );
    let error = file
        .add(
            "/",
            NewNode::new("slide")
                .prop("from", "/slide[1]")
                .prop("layout", "x"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add("/", NewNode::new("slide").prop("layout", "No such layout"))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn adds_paragraphs_text_boxes_shapes_and_pictures() {
    let file = read("multi-slide.key");
    let added = file
        .add(
            BODY,
            NewNode::new("paragraph")
                .prop("text", "Inserted topic")
                .at(Placement::Position(litchi_core::Position::new(1))),
        )
        .expect("paragraph");
    assert_eq!(
        added.path.to_string(),
        "/slide[2]/placeholder[2]/paragraph[2]"
    );
    assert_eq!(
        texts(&reopen(&added.file), &format!("{BODY}/paragraph")),
        [
            "First topic",
            "Inserted topic",
            "Second topic",
            "Third topic"
        ]
    );

    let text_box = file
        .add(
            "/slide[1]",
            NewNode::new("textbox")
                .prop("text", "Fresh box")
                .prop("x", "1in")
                .prop("y", "1in")
                .prop("width", "4in")
                .prop("height", "1in"),
        )
        .expect("text box");
    assert_eq!(text_box.path.to_string(), "/slide[1]/shape[2]");
    let reopened = reopen(&text_box.file);
    assert_eq!(
        text(&reopened, "/slide[1]/shape[2]").as_deref(),
        Some("Fresh box")
    );
    assert_eq!(
        prop(&reopened, "/slide[1]/shape[2]", "width"),
        Some(pt(288))
    );

    let shape = file
        .add(
            "/slide[2]",
            NewNode::new("shape")
                .prop("preset", "ellipse")
                .prop("text", "Oval"),
        )
        .expect("shape");
    assert_eq!(
        text(&reopen(&shape.file), &shape.path.to_string()).as_deref(),
        Some("Oval")
    );

    let picture = file
        .add(
            "/slide[1]",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("width", "2in")
                .prop("height", "3in"),
        )
        .expect("picture");
    assert_eq!(picture.path.to_string(), "/slide[1]/picture[1]");
    let reopened = reopen(&picture.file);
    assert_eq!(count(&reopened, "/slide[1]/picture"), 1);
    assert_eq!(
        prop(&reopened, "/slide[1]/picture[1]", "height"),
        Some(pt(216))
    );
    let error = file.add("/slide[1]", NewNode::new("picture")).unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    let error = file
        .add(
            "/slide[1]",
            NewNode::new("picture").prop("src", "data:image/png;base64,AAAA"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn removes_slides_paragraphs_and_drawables() {
    let file = read("multi-slide.key");
    // F8 (record 0296) is closed: a multi-slide deck deletes slides.
    let removed = file.remove("/slide[3]").expect("remove slide");
    assert_eq!(
        titles(&reopen(&removed)),
        [
            "Litchi multi-slide fixture",
            "Agenda",
            "Closing",
            "Agenda (continued)"
        ]
    );
    let removed = file
        .remove(&format!("{BODY}/p[2]"))
        .expect("remove paragraph");
    assert_eq!(
        texts(&reopen(&removed), &format!("{BODY}/paragraph")),
        ["First topic", "Third topic"]
    );
    let removed = file.remove("/slide[1]/shape[1]").expect("remove shape");
    assert_eq!(count(&reopen(&removed), "/slide[1]/shape"), 0);
    let error = file
        .remove("/slide[1]/placeholder[title]/p[1]")
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    // A title placeholder is hidden on its slide (record 0309).
    file.remove("/slide[1]/placeholder[title]")
        .expect("hide the title placeholder");
    let tables = read("slide-table-read-native.key");
    let error = tables.remove("/slide[1]/table[1]").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let single = read("basic.key");
    let error = single.remove("/slide[1]").unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[test]
fn moves_and_swaps_slides_paragraphs_and_drawables() {
    let file = read("multi-slide.key");
    let moved = file
        .move_to("/slide[1]", Placement::End)
        .expect("move slide");
    assert_eq!(titles(&moved)[4], "Litchi multi-slide fixture");
    let swapped = file.swap("/slide[1]", "/slide[3]").expect("swap slides");
    assert_eq!(
        titles(&reopen(&swapped))[..3],
        ["Hidden section", "Agenda", "Litchi multi-slide fixture"]
    );
    let moved = file
        .move_to(&format!("{BODY}/p[3]"), Placement::Start)
        .expect("move paragraph");
    assert_eq!(
        texts(&reopen(&moved), &format!("{BODY}/paragraph")),
        ["Third topic", "First topic", "Second topic"]
    );
    let swapped = file
        .swap(&format!("{BODY}/p[1]"), &format!("{BODY}/p[2]"))
        .expect("swap paragraphs");
    assert_eq!(
        texts(&reopen(&swapped), &format!("{BODY}/paragraph")),
        ["Second topic", "First topic", "Third topic"]
    );
    // The date text box moves from the back to the front of slide 1.
    let moved = file
        .move_to("/slide[1]/shape[1]", Placement::End)
        .expect("move shape");
    let reopened = reopen(&moved);
    assert_eq!(
        prop(&reopened, "/slide[1]/shape[1]", "zorder").as_deref(),
        Some("2")
    );
    let swapped = file
        .swap("/slide[1]/shape[1]", "/slide[1]/placeholder[body]")
        .expect("swap drawables");
    let reopened = reopen(&swapped);
    assert_eq!(
        prop(&reopened, "/slide[1]/shape[1]", "zorder").as_deref(),
        Some("2")
    );
    assert_eq!(
        prop(&reopened, "/slide[1]/placeholder[body]", "zorder").as_deref(),
        Some("0")
    );
    let error = file
        .swap("/slide[1]/shape[1]", "/slide[2]/shape[1]")
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[test]
fn create_builds_a_presentation() {
    let file = create(FileFormat::Keynote).expect("create");
    assert!(matches!(file, File::Keynote(_)));
    assert_eq!(
        create(FileFormat::Keynote).unwrap().to_bytes().unwrap(),
        file.to_bytes().unwrap(),
        "creation is deterministic"
    );
    let slides = count(&file, "slide");
    let added = file.add("/", NewNode::new("slide")).expect("add slide");
    let reopened = reopen(&added.file);
    assert_eq!(count(&reopened, "slide"), slides + 1);
    let capabilities = Capabilities::of(FileFormat::Keynote);
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
fn unsupported_writes_are_typed_refusals() {
    let file = read("multi-slide.key");
    let error = file
        .set(&format!("{BODY}/p[1]"), props! { "text" => "two\nlines" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set("/slide[1]/shape[1]", props! { "width" => "-1pt" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set("/slide[1]/shape[1]", props! { "zorder" => 2 })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // Slide 3's body placeholder has no frame on its slide.
    let error = file
        .set("/slide[3]/placeholder[body]", props! { "x" => "1in" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // A formula cell refuses a value (no recalculation).
    let tables = read("slide-table-read-native.key");
    let error = tables
        .set("/slide[1]/table[1]/tr[2]/tc[4]", props! { "text" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // A deck without a slide table has no table structure to copy.
    let error = file
        .add(
            "/slide[1]",
            NewNode::new("table").prop("rows", 2).prop("cols", 2),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = tables.add("/slide[1]", NewNode::new("tbl")).unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    let error = tables
        .add(
            "/slide[1]",
            NewNode::new("tbl").prop("rows", 0).prop("cols", 2),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::OutOfRange);
    let error = tables
        .add(
            "/slide[1]",
            NewNode::new("tbl")
                .prop("rows", 1)
                .prop("cols", 1)
                .prop("data", "a,b"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::OutOfRange);
}

#[test]
fn slide_tables_are_added_with_seeded_text() {
    let file = read("slide-table-read-native.key");
    let before = count(&file, "/slide[1]/table");
    let added = file
        .add(
            "/slide[1]",
            NewNode::new("tbl")
                .prop("name", "Forecast")
                .prop("data", "Region,Q1;North,12")
                .prop("x", "1in")
                .prop("width", "6in"),
        )
        .expect("table");
    assert_eq!(
        added.path.to_string(),
        format!("/slide[1]/table[{}]", before + 1)
    );
    let reopened = reopen(&added.file);
    let path = added.path.to_string();
    assert_eq!(prop(&reopened, &path, "name").as_deref(), Some("Forecast"));
    assert_eq!(prop(&reopened, &path, "rows").as_deref(), Some("2"));
    assert_eq!(prop(&reopened, &path, "cols").as_deref(), Some("2"));
    assert_eq!(prop(&reopened, &path, "width"), Some(pt(432)));
    assert_eq!(
        text(&reopened, &format!("{path}/tr[2]/tc[1]")).as_deref(),
        Some("North")
    );
    // The new table's cells take values like any other.
    let written = reopened
        .set(&format!("{path}/tr[1]/tc[2]"), props! { "value" => 7 })
        .expect("a cell");
    assert_eq!(
        prop(&reopen(&written), &format!("{path}/tr[1]/tc[2]"), "type").as_deref(),
        Some("number")
    );
    let edited = file
        .edit(|tx| {
            tx.add(
                "/slide[1]",
                NewNode::new("tbl").prop("rows", 2).prop("cols", 2),
            )?;
            Ok(())
        })
        .expect("edit");
    let restored = edited
        .patch()
        .inverse()
        .apply(edited.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    // In a batch, a table add takes its own commit beside other items.
    let batch = file
        .edit(|tx| {
            tx.add(
                "/slide[1]",
                NewNode::new("table").prop("rows", 2).prop("cols", 3),
            )?;
            tx.add("/slide[1]", NewNode::new("textbox").prop("text", "Note"))?;
            Ok(())
        })
        .expect("batch");
    let reopened = reopen(batch.file());
    assert_eq!(count(&reopened, "/slide[1]/table"), before + 1);
    assert_eq!(
        prop(
            &reopened,
            &format!("/slide[1]/table[{}]", before + 1),
            "cols"
        )
        .as_deref(),
        Some("3")
    );
}

#[test]
fn a_batch_of_sets_publishes_one_commit_per_family() {
    let file = read("multi-slide.key");
    let commit = file
        .edit(|tx| {
            for round in 0..100 {
                for line in 1..=3 {
                    tx.set(
                        &format!("{BODY}/p[{line}]"),
                        props! { "text" => format!("round {round} topic {line}") },
                    )?;
                }
                tx.set(
                    "/slide[1]/shape[1]",
                    props! { "x" => format!("{}pt", 10 + round) },
                )?;
            }
            Ok(())
        })
        .expect("edit");
    assert_eq!(commit.diagnostics().operations(), 400);
    // One transaction step for 400 operations, whose leaf work is one
    // slide-drawable commit (the frame) and one slide-text commit (the body
    // text, one span replacement).
    assert_eq!(commit.diagnostics().commits(), 1);
    let reopened = reopen(commit.file());
    assert_eq!(
        texts(&reopened, &format!("{BODY}/paragraph")),
        ["round 99 topic 1", "round 99 topic 2", "round 99 topic 3"]
    );
    assert_eq!(prop(&reopened, "/slide[1]/shape[1]", "x"), Some(pt(109)));
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
}

/// A long run of paragraph sets on one text (record 0340: staged paragraph
/// by paragraph, joined once) publishes one commit and what its items
/// publish one at a time, including sets that bring a paragraph back to its
/// base text, empty paragraphs, repeated targets and the parent's other
/// spellings.
#[test]
fn long_paragraph_runs_equal_their_items_one_at_a_time() {
    let lines = (0..120).map(|i| format!("line {i}")).collect::<Vec<_>>();
    let file = reopen(
        &read("multi-slide.key")
            .set(BODY, props! { "text" => lines.join("\n") })
            .expect("body"),
    );
    let mut sets = Vec::new();
    for i in 0..120_usize {
        let value = match i % 7 {
            0 => String::new(),
            1 => format!("line {i}"),
            _ => format!("edited \u{e9} {i}"),
        };
        sets.push((format!("{BODY}/p[{}]", i + 1), value));
        if i % 11 == 0 {
            // Back to the base text, then again.
            sets.push((format!("{BODY}/p[{}]", i + 1), format!("line {i}")));
        }
        if i % 13 == 0 {
            sets.push((
                format!("{BODY}/paragraph[{}]", i / 2 + 1),
                format!("other {i}"),
            ));
        }
    }
    let mut one_by_one = file.clone();
    for (path, value) in &sets {
        one_by_one = one_by_one
            .set(path.as_str(), props! { "text" => value.clone() })
            .expect("set");
    }
    let items = sets.iter().map(|(path, value)| {
        let mut item = BatchItem::new("set");
        item.path = Some(path.clone().into());
        item.props = Some(litchi_ops::Props::new().with("text", value.clone()));
        item
    });
    let batch = file.batch(items, BatchMode::Atomic).expect("batch");
    assert_eq!(
        batch.file().to_bytes().expect("publish"),
        one_by_one.to_bytes().expect("publish")
    );
    // A run without cancelling sets is one leaf commit.
    let items = (0..120).map(|i| {
        let mut item = BatchItem::new("set");
        item.path = Some(format!("{BODY}/p[{}]", i + 1).into());
        item.props = Some(litchi_ops::Props::new().with("text", format!("batch {i}")));
        item
    });
    let batch = file.batch(items, BatchMode::Atomic).expect("batch");
    assert_eq!(batch.diagnostics().commits(), 1);
    assert_eq!(
        texts(&reopen(batch.file()), &format!("{BODY}/paragraph")),
        (0..120).map(|i| format!("batch {i}")).collect::<Vec<_>>()
    );
}

#[test]
fn consecutive_new_slides_and_drawables_coalesce() {
    let file = read("multi-slide.key");
    let commit = file
        .edit(|tx| {
            let first = tx.add("/", NewNode::new("slide"))?;
            let second = tx.add(
                "/",
                NewNode::new("slide").at(Placement::Position(litchi_core::Position::new(0))),
            )?;
            let third = tx.add("/", NewNode::new("slide").prop("from", "/slide[2]"))?;
            Ok((first, second, third))
        })
        .expect("slides");
    assert_eq!(commit.diagnostics().commits(), 1);
    let (first, second, third) = commit.value();
    assert_eq!(first.to_string(), "/slide[6]");
    assert_eq!(second.to_string(), "/slide[1]");
    assert_eq!(third.to_string(), "/slide[8]");
    let reopened = reopen(commit.file());
    assert_eq!(count(&reopened, "slide"), 8);
    assert_eq!(titles(&reopened)[2], "Agenda");

    let commit = file
        .edit(|tx| {
            let mut paths = Vec::new();
            for index in 0..3 {
                paths.push(tx.add(
                    "/slide[1]",
                    NewNode::new("textbox").prop("text", format!("box {index}")),
                )?);
            }
            paths.push(tx.add("/slide[1]", NewNode::new("picture").prop("src", PNG))?);
            tx.set("/slide[1]/shape[1]", props! { "y" => "10pt" })?;
            Ok(paths)
        })
        .expect("drawables");
    assert_eq!(commit.diagnostics().commits(), 1);
    let reopened = reopen(commit.file());
    for (index, path) in commit.value()[..3].iter().enumerate() {
        assert_eq!(
            text(&reopened, &path.to_string()),
            Some(format!("box {index}"))
        );
    }
    assert_eq!(commit.value()[3].to_string(), "/slide[1]/picture[1]");
    assert_eq!(count(&reopened, "/slide[1]/picture"), 1);

    // A set after a new slide publishes the slides first, so it sees them.
    let commit = file
        .edit(|tx| {
            tx.add(
                "/",
                NewNode::new("slide")
                    .prop("from", "/slide[2]")
                    .at(Placement::Start),
            )?;
            tx.set("/slide[1]/placeholder[title]", props! { "text" => "Copy" })?;
            Ok(())
        })
        .expect("mixed");
    let reopened = reopen(commit.file());
    assert_eq!(
        titles(&reopened)[..3],
        ["Copy", "Litchi multi-slide fixture", "Agenda"]
    );
}

#[test]
fn a_failing_item_leaves_the_file_untouched() {
    let file = read("multi-slide.key");
    let bytes = file.to_bytes().unwrap();
    // Refused while staging: the second item.
    let error = file
        .edit(|tx| {
            tx.set(&format!("{BODY}/p[1]"), props! { "text" => "staged" })?;
            tx.set("/slide[1]/shape[1]", props! { "height" => "0pt" })?;
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(error.item(), Some(1));
    // Refused by the leaf when the run publishes: the third item.
    let error = file
        .edit(|tx| {
            tx.set(&format!("{BODY}/p[1]"), props! { "text" => "staged" })?;
            tx.set("/slide[1]/shape[1]", props! { "x" => "5pt" })?;
            tx.set(
                "/slide[1]/shape[1]",
                props! { "text" => "object \u{fffc} marker" },
            )?;
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.item(), Some(2));
    assert_eq!(file.to_bytes().unwrap(), bytes);

    let mut first = BatchItem::new("set");
    first.path = Some(format!("{BODY}/p[1]").into());
    first.props = Some(litchi_ops::Props::new().with("text", "batch"));
    let mut second = BatchItem::new("set");
    second.path = Some("/slide[1]/shape[1]".into());
    second.props = Some(litchi_ops::Props::new().with("zorder", 2_i64));
    let error = file.batch([first, second], BatchMode::Atomic).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(file.to_bytes().unwrap(), bytes);
}

#[test]
fn runs_of_shapes_take_character_properties() {
    let file = read("multi-slide.key");
    let added = file
        .add(
            "/slide[1]",
            NewNode::new("shape").prop("text", "Styled words here"),
        )
        .expect("add a text box");
    let shape = added.path.to_string();
    let styled = added
        .file
        .set(
            &format!("{shape}/p[1]/run[1]"),
            props! { "bold" => true, "color" => "C80000", "size" => "24pt", "font.latin" => "Helvetica" },
        )
        .expect("style the run");
    let reopened = reopen(&styled);
    let run = format!("{shape}/p[1]/run[1]");
    assert_eq!(prop(&reopened, &run, "bold").as_deref(), Some("true"));
    assert_eq!(
        prop(&reopened, &run, "font.latin").as_deref(),
        Some("Helvetica")
    );
    assert!(prop(&reopened, &run, "color").is_some());
    assert_eq!(
        text(&reopened, &shape).as_deref(),
        Some("Styled words here")
    );
    // A second write builds on the first.
    let italic = reopened
        .set(&run, props! { "italic" => true })
        .expect("italic");
    let reopened = reopen(&italic);
    assert_eq!(prop(&reopened, &run, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&reopened, &run, "italic").as_deref(), Some("true"));
    // Undo restores the bytes.
    let commit = added
        .file
        .edit(|tx| tx.set(&run, props! { "underline" => true }))
        .expect("underline");
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), added.file.to_bytes().unwrap());
    // A text change, then a run style on the new text, in one edit.
    let both = added
        .file
        .edit(|tx| {
            tx.set(&shape, props! { "text" => "changed" })?;
            tx.set(&run, props! { "bold" => true })?;
            Ok(())
        })
        .expect("text then style");
    let reopened = reopen(both.file());
    assert_eq!(text(&reopened, &shape).as_deref(), Some("changed"));
    assert_eq!(prop(&reopened, &run, "bold").as_deref(), Some("true"));
}

#[test]
fn slide_table_cells_take_values() {
    let file = read("table-discovery.key");
    let cell = file
        .query("tc")
        .expect("query")
        .next()
        .expect("a cell")
        .expect("node")
        .path()
        .to_string();
    let written = file
        .edit(|tx| tx.set(&cell, props! { "text" => "Updated" }))
        .expect("write a cell");
    let reopened = reopen(written.file());
    assert_eq!(text(&reopened, &cell).as_deref(), Some("Updated"));
    let restored = written
        .patch()
        .inverse()
        .apply(written.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    let numeric = reopened
        .set(&cell, props! { "value" => 3 })
        .expect("a number");
    assert_eq!(
        prop(&reopen(&numeric), &cell, "type").as_deref(),
        Some("number")
    );
}

#[test]
fn slide_table_cell_comments_round_trip_through_the_unified_route() {
    let file = read("slide-table-read-native.key");
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

    let same = file
        .set(&cell, props! { "comment" => existing.as_str() })
        .expect("same comment");
    assert_eq!(same.to_bytes().expect("same bytes"), original);

    let commit = file
        .edit(|tx| tx.set(&cell, props! { "comment" => "D7 unified Keynote root" }))
        .expect("set comment");
    let edited = reopen(commit.file());
    assert_eq!(
        prop(&edited, &cell, "comment").as_deref(),
        Some("D7 unified Keynote root")
    );
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

    let error = file
        .set(
            "/slide[999]/table[1]/tr[1]/tc[1]",
            props! { "comment" => "must refuse" },
        )
        .expect_err("missing slide");
    assert_eq!(error.code(), Code::PathNotFound);
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
fn keynote_cell_comment_schema_is_clearable() {
    let property = litchi_ops::schema(FileFormat::Keynote)
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
fn movies_and_audio_are_removed() {
    for name in [
        "media-lifecycle-audio-remove-native.key",
        "slide-movie-creation-focused-native.key",
    ] {
        let file = read(name);
        let before = count(&file, "video") + count(&file, "audio");
        assert!(before > 0, "{name} has media");
        let path = file
            .query("video, audio")
            .or_else(|_| file.query("audio"))
            .expect("query")
            .next()
            .expect("media")
            .expect("node")
            .path()
            .to_string();
        let commit = file.edit(|tx| tx.remove(&path)).expect("remove media");
        let reopened = reopen(commit.file());
        assert_eq!(
            count(&reopened, "video") + count(&reopened, "audio"),
            before - 1,
            "{name}"
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse");
        assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    }
}

#[test]
fn placeholders_are_hidden_and_shown() {
    use litchi_keynote::slide::placeholder::{Kind, State};
    let file = read("multi-slide.key");
    let leaf = |file: &File| {
        litchi_keynote::Package::from_bytes(&file.to_bytes().expect("bytes")).expect("leaf")
    };
    let removed = file.remove(BODY).expect("hide the body placeholder");
    assert_eq!(
        leaf(&removed)
            .slide_placeholder_visibility(litchi_keynote::SlideSelector::index(1), Kind::Body)
            .expect("read"),
        Some(State::Hidden)
    );
    let shown = removed
        .set(BODY, props! { "hidden" => false })
        .expect("show it again");
    assert_eq!(
        leaf(&shown)
            .slide_placeholder_visibility(litchi_keynote::SlideSelector::index(1), Kind::Body)
            .expect("read"),
        Some(State::Visible)
    );
    // The unified `hidden` reads the per-slide state (record 0312).
    assert_eq!(prop(&file, BODY, "hidden").as_deref(), Some("false"));
    assert_eq!(
        prop(&reopen(&removed), BODY, "hidden").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&reopen(&shown), BODY, "hidden").as_deref(),
        Some("false")
    );
}

#[test]
fn placeholder_runs_take_character_properties() {
    // Record 0312: title and body placeholders style their slide's own
    // storage, as text boxes do.
    let file = read("multi-slide.key");
    let run = "/slide[1]/placeholder[title]/p[1]/run[1]";
    let before = text(&file, run);
    let styled = file
        .set(run, props! { "bold" => true, "size" => "40pt" })
        .expect("style a placeholder run");
    let reopened = reopen(&styled);
    assert_eq!(prop(&reopened, run, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&reopened, run, "size").as_deref(), Some("40pt"));
    assert_eq!(text(&reopened, run), before);
    assert_eq!(
        text(&reopened, "/slide[1]/placeholder[title]"),
        text(&file, "/slide[1]/placeholder[title]")
    );
    // In a batch, the run addresses the text the earlier item wrote.
    let commit = file
        .edit(|tx| {
            tx.set(
                "/slide[1]/placeholder[title]",
                props! { "text" => "Changed" },
            )?;
            tx.set(run, props! { "italic" => true })?;
            Ok(())
        })
        .expect("batch");
    let reopened = reopen(commit.file());
    assert_eq!(text(&reopened, run).as_deref(), Some("Changed"));
    assert_eq!(prop(&reopened, run, "italic").as_deref(), Some("true"));
}

#[test]
fn notes_runs_take_character_properties() {
    // Record 0329: speaker notes list their runs, which take character
    // properties as placeholder runs do.
    let file = read("multi-slide.key")
        .set(
            "/slide[2]/notes",
            props! { "text" => "Remember\nthe second point" },
        )
        .expect("notes");
    let file = reopen(&file);
    let run = "/slide[2]/notes/p[2]/run[1]";
    assert_eq!(text(&file, run).as_deref(), Some("the second point"));
    let styled = file
        .set(run, props! { "bold" => true, "color" => "C01020" })
        .expect("style a notes run");
    let reopened = reopen(&styled);
    assert_eq!(prop(&reopened, run, "bold").as_deref(), Some("true"));
    assert_eq!(
        text(&reopened, "/slide[2]/notes"),
        text(&file, "/slide[2]/notes")
    );
    let restored = file
        .edit(|tx| tx.set(run, props! { "italic" => true }))
        .expect("edit");
    let undone = restored
        .patch()
        .inverse()
        .apply(restored.file())
        .expect("inverse");
    assert_eq!(undone.to_bytes().unwrap(), file.to_bytes().unwrap());
    // In a batch, the run addresses the notes text the earlier item wrote.
    let commit = file
        .edit(|tx| {
            tx.set(
                "/slide[2]/notes",
                props! { "text" => "Recall\nthe other point" },
            )?;
            tx.set(run, props! { "italic" => true })?;
            Ok(())
        })
        .expect("batch");
    let reopened = reopen(commit.file());
    assert_eq!(text(&reopened, run).as_deref(), Some("the other point"));
    assert_eq!(prop(&reopened, run, "italic").as_deref(), Some("true"));
}
