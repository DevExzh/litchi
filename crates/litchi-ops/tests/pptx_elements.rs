//! PPTX element gaps closed for OfficeCLI tier-1 parity (record 0321): the
//! theme, table columns, hyperlinks, media, equations and slide masters and
//! layouts through the unified `File` API, for source and eager owners,
//! with round trips, exact no-ops and refusals that leave the input intact.
#![cfg(feature = "pptx")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn deck() -> File {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
}

/// Both owners of `file`'s bytes: source-backed and eager.
fn owners(file: &File) -> [File; 2] {
    let bytes = file.to_bytes().unwrap();
    [SourceMode::Auto, SourceMode::Eager].map(|mode| {
        litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().source_mode(mode)).unwrap()
    })
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[test]
fn theme_get_set_round_trip_and_no_op() {
    for file in owners(&deck()) {
        let node = file.get("/theme").unwrap();
        assert_eq!(node.view().kind(), "theme");
        let accent = prop(&file, "/theme", "accent1").expect("accent1");
        assert!(accent.starts_with('#'), "{accent}");
        assert!(prop(&file, "/theme", "bodyFont").is_some());
        assert!(prop(&file, "/theme", "name").is_some());
        // An equal value is an exact no-op.
        let bytes = file.to_bytes().unwrap();
        let same = file
            .set("/theme", props! { "accent1" => accent.as_str() })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        let changed = file
            .set(
                "/theme",
                props! {
                    "accent1" => "FF0000",
                    "dark2" => "#112233",
                    "majorFont" => "Georgia",
                    "bodyFont.ea" => "Yu Gothic",
                    "name" => "Custom",
                },
            )
            .unwrap();
        assert_eq!(file.to_bytes().unwrap(), bytes, "input unchanged");
        let read = reopen(&changed);
        assert_eq!(prop(&read, "/theme", "accent1").as_deref(), Some("#FF0000"));
        assert_eq!(prop(&read, "/theme", "dk2").as_deref(), Some("#112233"));
        assert_eq!(
            prop(&read, "/theme", "headingFont").as_deref(),
            Some("Georgia")
        );
        assert_eq!(
            prop(&read, "/theme", "bodyFont.ea").as_deref(),
            Some("Yu Gothic")
        );
        assert_eq!(prop(&read, "/theme", "name").as_deref(), Some("Custom"));
        assert_eq!(
            prop(&read, "/theme", "accent2"),
            prop(&file, "/theme", "accent2")
        );
        assert_eq!(read.query("theme").unwrap().count(), 1);
    }
}

#[test]
fn theme_refusals_leave_the_input_unchanged() {
    let file = deck();
    let bytes = file.to_bytes().unwrap();
    let error = file
        .set("/theme", props! { "accent9" => "FF0000" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    let error = file
        .set("/theme", props! { "accent1" => "accent2" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(file.to_bytes().unwrap(), bytes);
    assert!(file.add("/", NewNode::new("theme")).is_err());
    assert!(file.remove("/theme").is_err());
}

#[test]
fn master_and_layout_names_and_placeholder_text() {
    for file in owners(&deck()) {
        let layout = prop(&file, "/slidemaster[1]/slidelayout[1]", "name").unwrap();
        let bytes = file.to_bytes().unwrap();
        // An equal name is an exact no-op.
        let same = file
            .set(
                "/slidemaster[1]/slidelayout[1]",
                props! { "name" => layout.as_str() },
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        let renamed = file
            .set("/slidemaster[1]", props! { "name" => "Brand" })
            .unwrap()
            .set("/slidelayout[2]", props! { "name" => "Agenda" })
            .unwrap();
        assert_eq!(file.to_bytes().unwrap(), bytes, "input unchanged");
        let read = reopen(&renamed);
        assert_eq!(
            prop(&read, "/slidemaster[1]", "name").as_deref(),
            Some("Brand")
        );
        assert_eq!(
            prop(&read, "/slidemaster[1]/slidelayout[2]", "name").as_deref(),
            Some("Agenda")
        );
        // The slide keeps its layout binding.
        assert_eq!(
            prop(&read, "/slide[1]", "layout"),
            prop(&file, "/slide[1]", "layout")
        );
        // Placeholder text on a layout.
        let layout_ph = read
            .query("slidelayout placeholder")
            .unwrap()
            .next()
            .expect("a layout placeholder")
            .unwrap()
            .view()
            .path()
            .to_string();
        let edited = read
            .set(layout_ph.as_str(), props! { "text" => "Click to add" })
            .unwrap();
        let edited = reopen(&edited);
        assert_eq!(
            edited
                .get(layout_ph.as_str())
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Click to add")
        );
        let error = read
            .set(layout_ph.as_str(), props! { "bold" => true })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = read
            .set("/slidemaster[1]", props! { "preserved" => true })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }
}

/// A deck whose first slide holds a 2×2 table (1in columns).
fn table_deck() -> File {
    let source = litchi_pptx::Package::new().unwrap();
    let mut edit = source.edit().unwrap();
    let slide = edit
        .insert_slide(litchi_core::Position::new(0), "Blank")
        .unwrap();
    let cells = vec![
        vec!["a".to_owned(), "b".to_owned()],
        vec!["c".to_owned(), "d".to_owned()],
    ];
    edit.add_table(slide.name(), &cells, (0, 0, 1_828_800, 914_400))
        .unwrap();
    let bytes = edit.commit().unwrap().into_package().to_bytes().unwrap();
    litchi_ops::open_bytes(bytes).unwrap()
}

fn cell_text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
        .unwrap_or_default()
}

#[test]
fn table_columns_get_add_set_remove() {
    for file in owners(&table_deck()) {
        let width = prop(&file, "/slide[1]/table[1]/col[2]", "width").unwrap();
        assert_eq!(width, "2.54cm");
        assert_eq!(file.query("col").unwrap().count(), 2);
        assert_eq!(
            file.get("/slide[1]/table[1]/col[2]")
                .unwrap()
                .view()
                .path()
                .to_string(),
            "/slide[1]/table[1]/col[2]"
        );
        let bytes = file.to_bytes().unwrap();
        // An equal width is an exact no-op.
        let same = file
            .set("/slide[1]/table[1]/col[1]", props! { "width" => "1in" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        let resized = file
            .set("/slide[1]/table[1]/col[1]", props! { "width" => "2in" })
            .unwrap();
        let read = reopen(&resized);
        assert_eq!(
            prop(&read, "/slide[1]/table[1]/col[1]", "width").as_deref(),
            Some("5.08cm")
        );
        assert_eq!(
            prop(&read, "/slide[1]/table[1]", "width").as_deref(),
            Some("7.62cm")
        );
        // Insert a seeded column in the middle.
        let added = file
            .add(
                "/slide[1]/table[1]",
                NewNode::new("column")
                    .prop("width", "3cm")
                    .prop("text", "new")
                    .at(litchi_ops::Placement::Position(litchi_core::Position::new(
                        1,
                    ))),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/table[1]/col[2]");
        let read = reopen(&added.file);
        assert_eq!(
            prop(&read, "/slide[1]/table[1]", "cols").as_deref(),
            Some("3")
        );
        assert_eq!(cell_text(&read, "/slide[1]/table[1]/tr[1]/tc[2]"), "new");
        assert_eq!(cell_text(&read, "/slide[1]/table[1]/tr[2]/tc[3]"), "d");
        assert_eq!(
            prop(&read, "/slide[1]/table[1]/col[2]", "width").as_deref(),
            Some("3cm")
        );
        // Remove it again: the original bytes come back.
        let removed = read.remove("/slide[1]/table[1]/col[2]").unwrap();
        let read = reopen(&removed);
        assert_eq!(
            prop(&read, "/slide[1]/table[1]", "cols").as_deref(),
            Some("2")
        );
        assert_eq!(cell_text(&read, "/slide[1]/table[1]/tr[1]/tc[2]"), "b");
        // Refusals leave the input unchanged.
        let one = read.remove("/slide[1]/table[1]/col[1]").unwrap();
        let one = reopen(&one);
        let before = one.to_bytes().unwrap();
        assert!(one.remove("/slide[1]/table[1]/col[1]").is_err());
        assert_eq!(
            one.add(
                "/slide[1]/table[1]",
                NewNode::new("col").at(litchi_ops::Placement::Position(
                    litchi_core::Position::new(5)
                ))
            )
            .unwrap_err()
            .code(),
            Code::OutOfRange
        );
        assert!(
            one.set("/slide[1]/table[1]/col[1]", props! { "width" => "-1cm" })
                .is_err()
        );
        assert_eq!(one.to_bytes().unwrap(), before);
    }
}

/// A deck with two slides, each with one text box.
fn two_slides() -> File {
    let mut file = deck()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file;
    for slide in ["/slide[1]", "/slide[2]"] {
        file = file
            .add(
                slide,
                NewNode::new("shape")
                    .prop("text", "Docs")
                    .prop("width", "2in")
                    .prop("height", "1in"),
            )
            .unwrap()
            .file;
    }
    file
}

#[test]
fn hyperlinks_add_get_set_remove() {
    for file in owners(&two_slides()) {
        let bytes = file.to_bytes().unwrap();
        assert_eq!(
            file.get("/slide[1]/shape[1]/hyperlink").unwrap_err().code(),
            Code::PathNotFound
        );
        let added = file
            .add(
                "/slide[1]/shape[1]",
                NewNode::new("hyperlink")
                    .prop("url", "https://example.com/docs")
                    .prop("tooltip", "Open docs"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/shape[1]/hyperlink");
        assert_eq!(file.to_bytes().unwrap(), bytes, "input unchanged");
        let read = reopen(&added.file);
        let link = "/slide[1]/shape[1]/hyperlink";
        assert_eq!(
            prop(&read, link, "link").as_deref(),
            Some("https://example.com/docs")
        );
        assert_eq!(prop(&read, link, "tooltip").as_deref(), Some("Open docs"));
        // The shape-wide link is on every run too.
        assert_eq!(
            prop(
                &read,
                "/slide[1]/shape[1]/paragraph[1]/run[1]/hyperlink",
                "link"
            )
            .as_deref(),
            Some("https://example.com/docs")
        );
        assert_eq!(read.query("hyperlink").unwrap().count(), 2);
        // The same link again is an exact no-op.
        let again = read
            .set(link, props! { "link" => "https://example.com/docs" })
            .unwrap();
        assert_eq!(again.to_bytes().unwrap(), read.to_bytes().unwrap());
        // A slide jump, then a named action with the tooltip kept.
        let jump = reopen(&read.set(link, props! { "link" => "slide[2]" }).unwrap());
        assert_eq!(prop(&jump, link, "link").as_deref(), Some("slide[2]"));
        assert_eq!(prop(&jump, link, "tooltip").as_deref(), Some("Open docs"));
        let next = reopen(&jump.set(link, props! { "link" => "nextslide" }).unwrap());
        assert_eq!(prop(&next, link, "link").as_deref(), Some("nextslide"));
        // A tooltip-only set keeps the target.
        let tip = reopen(&next.set(link, props! { "tooltip" => "Next" }).unwrap());
        assert_eq!(prop(&tip, link, "link").as_deref(), Some("nextslide"));
        assert_eq!(prop(&tip, link, "tooltip").as_deref(), Some("Next"));
        // A run link on the second slide.
        let run = reopen(
            &tip.add(
                "/slide[2]/shape[1]/paragraph[1]/run[1]",
                NewNode::new("hyperlink").prop("link", "slide[1]"),
            )
            .unwrap()
            .file,
        );
        assert_eq!(
            prop(
                &run,
                "/slide[2]/shape[1]/paragraph[1]/run[1]/hyperlink",
                "link"
            )
            .as_deref(),
            Some("slide[1]")
        );
        // Remove and `link=none`.
        let removed = reopen(&run.remove(link).unwrap());
        assert!(removed.get(link).is_err());
        assert!(
            removed
                .get("/slide[1]/shape[1]/paragraph[1]/run[1]/hyperlink")
                .is_err()
        );
        let none = reopen(
            &run.set(
                "/slide[2]/shape[1]/paragraph[1]/run[1]/hyperlink",
                props! { "link" => "none" },
            )
            .unwrap(),
        );
        assert_eq!(none.query("hyperlink").unwrap().count(), 2);
        // Refusals.
        let before = run.to_bytes().unwrap();
        assert_eq!(
            run.set(link, props! { "link" => "slide[9]" })
                .unwrap_err()
                .code(),
            Code::OutOfRange
        );
        assert_eq!(
            run.set(link, props! { "link" => "javascript:alert(1)" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        assert_eq!(
            run.add("/slide[1]/shape[1]", NewNode::new("hyperlink"))
                .unwrap_err()
                .code(),
            Code::MissingProperty
        );
        assert_eq!(run.to_bytes().unwrap(), before);
    }
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC";
/// Not a playable file: the leaf stores media payloads inert.
const MP4: &str = "data:video/mp4;base64,AAAAGGZ0eXBtcDQyAAAAAG1wNDJpc29t";
const MP3: &str = "data:audio/mpeg;base64,SUQzAwAAAAAAAA==";

#[test]
fn media_add_get_query_remove() {
    for file in owners(&deck()) {
        let bytes = file.to_bytes().unwrap();
        let video = file
            .add(
                "/slide[1]",
                NewNode::new("video")
                    .prop("src", MP4)
                    .prop("poster", PNG)
                    .prop("x", "1in")
                    .prop("width", "4in")
                    .prop("height", "3in"),
            )
            .unwrap();
        assert_eq!(video.path.to_string(), "/slide[1]/video[1]");
        assert_eq!(file.to_bytes().unwrap(), bytes, "input unchanged");
        let audio = video
            .file
            .add(
                "/slide[1]",
                NewNode::new("media").prop("src", MP3).prop("poster", PNG),
            )
            .unwrap();
        assert_eq!(audio.path.to_string(), "/slide[1]/media[2]");
        let read = reopen(&audio.file);
        assert_eq!(
            read.get("/slide[1]/video[1]").unwrap().view().kind(),
            "video"
        );
        assert_eq!(
            prop(&read, "/slide[1]/video[1]", "width").as_deref(),
            Some("10.16cm")
        );
        assert_eq!(
            prop(&read, "/slide[1]/audio[1]", "name").as_deref(),
            Some("Audio 2")
        );
        assert_eq!(read.query("media").unwrap().count(), 2);
        assert_eq!(read.query("video").unwrap().count(), 1);
        // Media pictures stay pictures of the slide.
        assert_eq!(read.query("picture").unwrap().count(), 2);
        let removed = reopen(&read.remove("/slide[1]/video[1]").unwrap());
        assert!(removed.get("/slide[1]/video[1]").is_err());
        assert_eq!(removed.query("audio").unwrap().count(), 1);
        // Refusals leave the input unchanged.
        let before = read.to_bytes().unwrap();
        assert_eq!(
            read.add("/slide[1]", NewNode::new("video").prop("src", MP4))
                .unwrap_err()
                .code(),
            Code::MissingProperty
        );
        assert_eq!(
            read.add(
                "/slide[1]",
                NewNode::new("audio").prop("src", MP4).prop("poster", PNG)
            )
            .unwrap_err()
            .code(),
            Code::InvalidValue
        );
        assert_eq!(read.to_bytes().unwrap(), before);
    }
}

const MCE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx");

#[test]
fn equations_read_from_the_choice_branch() {
    let file = litchi_ops::open_bytes(MCE.to_vec()).unwrap();
    for file in owners(&file) {
        let path = "/slide[1]/shape[1]/equation[1]";
        assert_eq!(file.get(path).unwrap().view().kind(), "equation");
        assert_eq!(
            prop(&file, path, "formula").as_deref(),
            Some("𝑎=\\frac{𝑏}{𝑐}")
        );
        assert_eq!(
            prop(&file, "/slide[1]/equation[1]", "formula").as_deref(),
            Some("𝑎=\\frac{𝑏}{𝑐}")
        );
        let found: Vec<String> = file
            .query("equation")
            .unwrap()
            .map(|node| node.unwrap().view().path().to_string())
            .collect();
        assert_eq!(found, vec![path.to_owned()]);
        assert!(file.get("/slide[1]/shape[1]/equation[2]").is_err());
        assert!(file.set(path, props! { "formula" => "x" }).is_err());
    }
}

#[test]
fn star_queries_and_views_keep_their_shape() {
    // Media pictures, hyperlinks and equations are reached by name only.
    let file = litchi_ops::open_bytes(MCE.to_vec()).unwrap();
    let all: Vec<String> = file
        .query("*")
        .unwrap()
        .map(|node| node.unwrap().view().kind().to_owned())
        .collect();
    assert!(!all.iter().any(|kind| kind == "equation"), "{all:?}");
}

#[path = "support/pptx_presentation.rs"]
#[allow(dead_code, reason = "the shared fixture module serves several tests")]
mod presentation;

#[test]
fn shape_run_picture_and_group_links_as_properties() {
    // Record 0325: `link` and `tooltip` as properties (OfficeCLI's usual
    // write path), on add and set, read back on both owners.
    let file = two_slides()
        .add(
            "/slide[1]",
            NewNode::new("shape")
                .prop("text", "Go")
                .prop("width", "1in")
                .prop("height", "1in")
                .prop("link", "https://example.com/a")
                .prop("tooltip", "Open"),
        )
        .unwrap()
        .file
        .add(
            "/slide[1]",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("link", "slide[2]")
                .prop("tooltip", "Next"),
        )
        .unwrap()
        .file;
    for file in owners(&file) {
        assert_eq!(
            prop(&file, "/slide[1]/shape[2]", "link").as_deref(),
            Some("https://example.com/a")
        );
        assert_eq!(
            prop(&file, "/slide[1]/shape[2]", "tooltip").as_deref(),
            Some("Open")
        );
        // A shape-wide link is on every run too, as OfficeCLI writes it.
        assert_eq!(
            prop(&file, "/slide[1]/shape[2]/paragraph[1]/run[1]", "link").as_deref(),
            Some("https://example.com/a")
        );
        assert_eq!(
            prop(&file, "/slide[1]/picture[1]", "link").as_deref(),
            Some("slide[2]")
        );
        assert_eq!(
            prop(&file, "/slide[1]/picture[1]", "tooltip").as_deref(),
            Some("Next")
        );
        assert_eq!(prop(&file, "/slide[1]/shape[1]", "link"), None);
        // set: a named action with text in one call, then the tooltip alone.
        let set = file
            .set(
                "/slide[1]/shape[1]",
                props! { "text" => "Next", "link" => "nextslide" },
            )
            .unwrap();
        assert_eq!(
            prop(&set, "/slide[1]/shape[1]", "link").as_deref(),
            Some("nextslide")
        );
        let tipped = set
            .set("/slide[1]/shape[1]", props! { "tooltip" => "Advance" })
            .unwrap();
        assert_eq!(
            prop(&tipped, "/slide[1]/shape[1]", "tooltip").as_deref(),
            Some("Advance")
        );
        assert_eq!(
            prop(&tipped, "/slide[1]/shape[1]", "link").as_deref(),
            Some("nextslide")
        );
        // A run's own link; `none` removes a picture's.
        let run = tipped
            .set(
                "/slide[2]/shape[1]/paragraph[1]/run[1]",
                props! { "link" => "slide[1]", "tooltip" => "Back" },
            )
            .unwrap();
        assert_eq!(
            prop(&run, "/slide[2]/shape[1]/paragraph[1]/run[1]", "link").as_deref(),
            Some("slide[1]")
        );
        let removed = run
            .set("/slide[1]/picture[1]", props! { "link" => "none" })
            .unwrap();
        assert_eq!(prop(&removed, "/slide[1]/picture[1]", "link"), None);
        // An equal link is an exact no-op.
        let bytes = removed.to_bytes().unwrap();
        let same = removed
            .set(
                "/slide[1]/shape[1]",
                props! { "link" => "nextslide", "tooltip" => "Advance" },
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        // Refusals leave the input unchanged.
        let error = file
            .set("/slide[2]/shape[1]", props! { "tooltip" => "x" })
            .unwrap_err();
        assert_eq!(error.code(), Code::MissingProperty);
        let error = file
            .set(
                "/slide[1]/shape[1]",
                props! { "link" => "javascript:alert(1)" },
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(
            file.to_bytes().unwrap(),
            owners(&file)[0].to_bytes().unwrap()
        );
    }
    // Groups take a link on their non-visual properties.
    let deck = litchi_ops::open_bytes(presentation::bytes()).unwrap();
    let linked = deck
        .set(
            "/slide[1]/group[1]",
            props! { "link" => "https://example.com/g", "tooltip" => "Group" },
        )
        .unwrap();
    for file in owners(&linked) {
        assert_eq!(
            prop(&file, "/slide[1]/group[1]", "link").as_deref(),
            Some("https://example.com/g")
        );
        assert_eq!(
            prop(&file, "/slide[1]/group[1]", "tooltip").as_deref(),
            Some("Group")
        );
        // The group's members keep theirs.
        assert_eq!(prop(&file, "/slide[1]/group[1]/shape[1]", "link"), None);
    }
}
