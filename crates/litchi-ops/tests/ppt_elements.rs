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
    assert!(noop.patch().is_noop());
    assert_eq!(noop.file().to_bytes().unwrap(), before);

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
        assert!(noop.patch().is_noop());

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

const SAMPLE: &str = "ole/ppt/SampleShow.ppt";

#[test]
fn shape_hyperlinks_add_set_and_remove_like_pptx() {
    each_owner(&fixture(SAMPLE), |source| {
        let before = source.to_bytes().unwrap();
        assert_eq!(
            source
                .get("/slide[1]/shape[1]/hyperlink")
                .unwrap_err()
                .code(),
            Code::PathNotFound
        );
        let added = source
            .add(
                "/slide[1]/shape[1]",
                NewNode::new("hyperlink")
                    .prop("url", "https://example.com/")
                    .prop("tooltip", "Example"),
            )
            .expect("add hyperlink");
        assert_eq!(added.path.to_string(), "/slide[1]/shape[1]/hyperlink");
        let read = reopen(&added.file);
        let link = "/slide[1]/shape[1]/hyperlink";
        assert_eq!(
            prop(&read, link, "link").as_deref(),
            Some("https://example.com/")
        );
        assert_eq!(prop(&read, link, "tooltip").as_deref(), Some("Example"));
        assert_eq!(paths(&read, "hyperlink"), [link]);
        assert_eq!(source.to_bytes().unwrap(), before);

        // A tooltip-only set keeps the target; a slide jump; a named action.
        let tip = reopen(&read.set(link, props! { "tooltip" => "Other" }).unwrap());
        assert_eq!(
            prop(&tip, link, "link").as_deref(),
            Some("https://example.com/")
        );
        assert_eq!(prop(&tip, link, "tooltip").as_deref(), Some("Other"));
        let jump = reopen(&tip.set(link, props! { "link" => "slide[2]" }).unwrap());
        assert_eq!(prop(&jump, link, "link").as_deref(), Some("slide[2]"));
        let named = reopen(
            &jump
                .add(
                    "/slide[2]/shape[1]",
                    NewNode::new("hyperlink").prop("link", "NextSlide"),
                )
                .unwrap()
                .file,
        );
        assert_eq!(
            prop(&named, "/slide[2]/shape[1]/hyperlink", "link").as_deref(),
            Some("nextslide")
        );
        let cleared = reopen(&named.set(link, props! { "link" => "none" }).unwrap());
        assert_eq!(cleared.get(link).unwrap_err().code(), Code::PathNotFound);
        let removed = reopen(&named.remove("/slide[2]/shape[1]/hyperlink").unwrap());
        assert_eq!(
            removed
                .get("/slide[2]/shape[1]/hyperlink")
                .unwrap_err()
                .code(),
            Code::PathNotFound
        );
        assert_eq!(prop(&removed, link, "link").as_deref(), Some("slide[2]"));
    });
}

#[test]
fn hyperlink_refusals_are_typed() {
    let source = eager(&fixture(SAMPLE));
    let refused =
        |node: NewNode<'static>| source.add("/slide[1]/shape[1]", node).unwrap_err().code();
    assert_eq!(
        refused(NewNode::new("hyperlink").prop("link", "javascript:alert(1)")),
        Code::InvalidValue
    );
    assert_eq!(refused(NewNode::new("hyperlink")), Code::MissingProperty);
    assert_eq!(
        refused(NewNode::new("hyperlink").prop("link", "slide[9]")),
        Code::OutOfRange
    );
    assert_eq!(
        source
            .add(
                "/slide[1]",
                NewNode::new("hyperlink").prop("link", "https://a/")
            )
            .unwrap_err()
            .code(),
        Code::UnsupportedType
    );
}

/// A deck with one text box of two paragraphs and a line break.
fn text_deck() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(
            slide,
            10,
            10,
            400,
            60,
            "Hello brave\u{b}new world\rSecond line",
        )
        .unwrap();
    writer.add_slide().unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn runs_and_line_breaks_read_like_pptx() {
    let deck = eager(&text_deck());
    let paragraph = "/slide[1]/shape[1]/paragraph[1]";
    assert_eq!(text(&deck, &format!("{paragraph}/run[1]")), "Hello brave");
    assert_eq!(text(&deck, &format!("{paragraph}/run[2]")), "new world");
    assert!(deck.get(&format!("{paragraph}/linebreak[1]")).is_ok());
    assert_eq!(
        deck.get(&format!("{paragraph}/linebreak[2]"))
            .unwrap_err()
            .code(),
        Code::PathNotFound
    );
    assert_eq!(
        text(&deck, "/slide[1]/shape[1]/paragraph[2]/run[1]"),
        "Second line"
    );
    assert_eq!(
        paths(&deck, &format!("{paragraph} > run")),
        [format!("{paragraph}/run[1]"), format!("{paragraph}/run[2]")]
    );

    let sample = eager(&fixture(SAMPLE));
    let italic = "/slide[1]/shape[@name=Another Subtitle]/paragraph[3]/run[1]";
    assert_eq!(text(&sample, italic), "This bit is in italic green");
    assert_eq!(prop(&sample, italic, "italic").as_deref(), Some("true"));
}

#[test]
fn run_formatting_and_text_set_round_trip() {
    each_owner(&text_deck(), |source| {
        let before = source.to_bytes().unwrap();
        let run = "/slide[1]/shape[1]/paragraph[1]/run[2]";
        let changed = source
            .set(
                run,
                props! { "bold" => true, "size" => "28pt", "color" => "1F4E79", "underline" => "single" },
            )
            .expect("run format");
        let read = reopen(&changed);
        assert_eq!(text(&read, run), "new world");
        assert_eq!(prop(&read, run, "bold").as_deref(), Some("true"));
        assert_eq!(prop(&read, run, "size").as_deref(), Some("28pt"));
        assert_eq!(prop(&read, run, "color").as_deref(), Some("#1F4E79"));
        assert_eq!(prop(&read, run, "underline").as_deref(), Some("single"));
        assert_eq!(
            prop(&read, "/slide[1]/shape[1]/paragraph[1]/run[1]", "bold"),
            None
        );
        assert_eq!(source.to_bytes().unwrap(), before);

        // The same text length keeps every run; a changed length over
        // several runs refuses with a typed error.
        let same = reopen(&read.set(run, props! { "text" => "NEW WORLD" }).unwrap());
        assert_eq!(text(&same, run), "NEW WORLD");
        assert_eq!(prop(&same, run, "bold").as_deref(), Some("true"));
        assert!(
            read.set(run, props! { "text" => "much longer text" })
                .is_err()
        );

        // A run hyperlink.
        let linked = source
            .add(
                "/slide[1]/shape[1]/paragraph[2]/run[1]",
                NewNode::new("hyperlink").prop("link", "https://example.com/run"),
            )
            .expect("run hyperlink");
        assert_eq!(
            linked.path.to_string(),
            "/slide[1]/shape[1]/paragraph[2]/run[1]/hyperlink"
        );
        let linked = reopen(&linked.file);
        assert_eq!(
            prop(
                &linked,
                "/slide[1]/shape[1]/paragraph[2]/run[1]/hyperlink",
                "link"
            )
            .as_deref(),
            Some("https://example.com/run")
        );
        let removed = reopen(
            &linked
                .remove("/slide[1]/shape[1]/paragraph[2]/run[1]/hyperlink")
                .unwrap(),
        );
        assert_eq!(
            removed
                .get("/slide[1]/shape[1]/paragraph[2]/run[1]/hyperlink")
                .unwrap_err()
                .code(),
            Code::PathNotFound
        );
    });
}

#[test]
fn theme_reads_the_round_trip_theme_like_pptx() {
    let deck = eager(&fixture("poi/test-data/slideshow/headers_footers_2007.ppt"));
    assert_eq!(prop(&deck, "/theme", "name").as_deref(), Some("Office"));
    assert_eq!(prop(&deck, "/theme", "dk2").as_deref(), Some("#1F497D"));
    assert_eq!(
        prop(&deck, "/theme", "hyperlink").as_deref(),
        Some("#0000FF")
    );
    assert_eq!(
        prop(&deck, "/theme", "headingFont").as_deref(),
        Some("Calibri")
    );
    assert_eq!(paths(&deck, "theme"), ["/theme"]);
    assert_eq!(
        eager(&fixture(NO_COMMENTS))
            .get("/theme")
            .unwrap_err()
            .code(),
        Code::PathNotFound
    );
    assert!(deck.set("/theme", props! { "dk1" => "FF0000" }).is_err());
}

/// A deck with a bar and a pie Microsoft Graph chart.
fn chart_deck() -> Vec<u8> {
    use litchi_ppt::builder::{Chart, ChartKind};
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    for kind in [ChartKind::Bar, ChartKind::Pie] {
        let mut chart = Chart::new(kind);
        chart.set_title("Quarterly sales");
        chart.set_categories(["Q1", "Q2", "Q3", "Q4"]);
        chart
            .add_series(Some("2024"), vec![1.5, 2.5, 3.5, 4.5])
            .unwrap();
        writer.add_chart(slide, 50, 50, 400, 240, chart).unwrap();
    }
    let mut output = std::io::Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn embedded_graph_charts_read_like_pptx_charts() {
    let deck = eager(&chart_deck());
    assert_eq!(
        paths(&deck, "chart"),
        ["/slide[1]/chart[1]", "/slide[1]/chart[2]"]
    );
    assert_eq!(
        prop(&deck, "/slide[1]/chart[1]", "chartType").as_deref(),
        Some("bar")
    );
    assert_eq!(
        prop(&deck, "/slide[1]/chart[1]", "title").as_deref(),
        Some("Quarterly sales")
    );
    assert_eq!(
        prop(&deck, "/slide[1]/chart[2]", "chartType").as_deref(),
        Some("pie")
    );
    assert!(prop(&deck, "/slide[1]/chart[1]", "width").is_some());
    assert_eq!(
        deck.get("/slide[1]/chart[3]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        deck.get("/slide[1]/chart[1]/series[1]").unwrap_err().code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn ole_and_media_frames_read_their_object_kinds() {
    let ole = eager(&fixture("poi/test-data/slideshow/testPPT_oleWorkbook.ppt"));
    let frame = "/slide[2]/ole[1]";
    assert!(
        prop(&ole, frame, "progId").is_some_and(|program| program.starts_with("Excel.Sheet")),
        "{:?}",
        prop(&ole, frame, "progId")
    );
    assert_eq!(prop(&ole, frame, "linked").as_deref(), Some("false"));
    assert!(prop(&ole, frame, "width").is_some());
    assert_eq!(paths(&ole, "ole"), [frame]);

    let sound = eager(&fixture("poi/test-data/slideshow/sound.ppt"));
    assert_eq!(paths(&sound, "audio"), ["/slide[1]/audio[1]"]);
    assert_eq!(paths(&sound, "media"), ["/slide[1]/media[1]"]);
    assert!(paths(&sound, "video").is_empty());
    assert!(
        sound
            .add("/slide[1]", NewNode::new("audio").prop("src", "a.wav"))
            .is_err()
    );
}
