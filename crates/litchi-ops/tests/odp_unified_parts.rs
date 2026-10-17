//! Unified ODP elements (record 0328): shape and text hyperlinks, line
//! breaks, table columns, video and audio media, formula objects and chart
//! series and axes, read and written through the PPTX paths and property
//! names of record 0321.

#![cfg(feature = "odp")]

use litchi_ops::{Code, File, NewNode, open_bytes, props};

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .expect("read fixture")
}

fn parts() -> File {
    open_bytes(fixture("unified-ops/0294/parts.odp")).expect("open ODP")
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
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

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .expect("text")
        .unwrap_or_default()
        .into_owned()
}

fn count(file: &File, query: &str) -> usize {
    file.query(query)
        .unwrap_or_else(|error| panic!("{query}: {error}"))
        .count()
}

#[test]
fn hyperlinks_read_shape_clicks_and_link_runs() {
    let file = parts();
    let run = "/slide[1]/shape[1]/paragraph[1]/run[2]/hyperlink";
    assert_eq!(
        prop(&file, run, "link").as_deref(),
        Some("https://example.com/")
    );
    assert_eq!(prop(&file, run, "tooltip").as_deref(), Some("Example"));
    assert_eq!(text(&file, run), "example");
    assert_eq!(
        prop(&file, "/slide[1]/shape[2]/hyperlink", "link").as_deref(),
        Some("https://example.org/")
    );
    assert_eq!(
        prop(&file, "/slide[1]/shape[3]/hyperlink", "link").as_deref(),
        Some("nextslide")
    );
    assert_eq!(
        prop(&file, "/slide[1]/shape[4]/hyperlink", "link").as_deref(),
        Some("https://example.net/")
    );
    // The text box's hyperlink is its first link run's.
    assert_eq!(
        prop(&file, "/slide[1]/shape[1]/hyperlink", "link").as_deref(),
        Some("https://example.com/")
    );
    let error = file.get("/slide[1]/shape[5]/hyperlink").unwrap_err();
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(count(&file, "hyperlink"), 5);
}

#[test]
fn line_breaks_columns_media_equations_and_charts_read() {
    let file = parts();
    let paragraph = "/slide[1]/shape[1]/paragraph[2]";
    assert_eq!(count(&file, &format!("{paragraph}/linebreak")), 2);
    file.get(format!("{paragraph}/br[2]").as_str()).unwrap();
    assert_eq!(count(&file, "/slide[1]/table[1]/col"), 3);
    assert_eq!(
        prop(&file, "/slide[1]/table[1]/col[2]", "width").as_deref(),
        Some("4cm")
    );
    assert_eq!(
        prop(&file, "/slide[1]/table[1]/col[3]", "width").as_deref(),
        Some("2.5cm")
    );
    assert_eq!(
        prop(&file, "/slide[1]/video[1]", "name").as_deref(),
        Some("Movie")
    );
    assert_eq!(
        prop(&file, "/slide[1]/video[1]", "mediaType").as_deref(),
        Some("video")
    );
    assert_eq!(
        prop(&file, "/slide[1]/video[1]", "loop").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&file, "/slide[1]/audio[1]", "name").as_deref(),
        Some("Sound")
    );
    assert_eq!(
        prop(&file, "/slide[1]/audio[1]", "mimeType").as_deref(),
        Some("audio/x-wav")
    );
    assert_eq!(count(&file, "/slide[1]/media"), 2);
    assert_eq!(count(&file, "video"), 1);
    assert_eq!(count(&file, "equation"), 2);
    assert_eq!(
        prop(&file, "/slide[1]/equation[1]", "formula").as_deref(),
        Some("E = m c^2")
    );
    assert_eq!(text(&file, "/slide[1]/equation[2]"), "y = 2");
    assert_eq!(
        prop(&file, "/slide[1]/equation[2]", "name").as_deref(),
        Some("InlineFormula")
    );
    let chart = "/slide[1]/chart[1]";
    assert_eq!(count(&file, &format!("{chart}/series")), 2);
    assert_eq!(
        prop(&file, &format!("{chart}/series[1]"), "values").as_deref(),
        Some("local-table.$B$2:.$B$3")
    );
    assert_eq!(
        prop(&file, &format!("{chart}/series[1]"), "pointCount").as_deref(),
        Some("2")
    );
    assert_eq!(
        prop(&file, &format!("{chart}/axis[1]"), "role").as_deref(),
        Some("category")
    );
    assert_eq!(
        prop(&file, &format!("{chart}/axis[1]"), "title").as_deref(),
        Some("Quarter")
    );
    assert_eq!(text(&file, &format!("{chart}/axis[@role=value]")), "");
}

#[test]
fn views_and_wildcards_list_what_they_listed() {
    let file = parts();
    let slide: Vec<String> = file
        .get("/slide[1]")
        .unwrap()
        .view()
        .children()
        .map(|child| child.unwrap().kind().to_owned())
        .collect();
    assert!(
        !slide
            .iter()
            .any(|kind| matches!(kind.as_str(), "video" | "audio" | "equation"))
    );
    assert_eq!(slide.iter().filter(|kind| *kind == "media").count(), 2);
    let paragraph = file.get("/slide[1]/shape[1]/paragraph[2]").unwrap();
    assert!(
        paragraph
            .view()
            .children()
            .all(|child| child.unwrap().kind() == "run")
    );
    // `*` lists the slides' direct children, as the views do.
    let second = file.get("/slide[2]").unwrap().view().children().count();
    assert_eq!(count(&file, "*"), slide.len() + second);
}

#[test]
fn text_hyperlinks_set_add_and_remove() {
    let file = parts();
    let run = "/slide[1]/shape[1]/paragraph[1]/run[2]/hyperlink";
    let edited = reopen(
        &file
            .set(
                run,
                props! { "link" => "https://litchi.example/", "tooltip" => "Litchi" },
            )
            .unwrap(),
    );
    assert_eq!(
        prop(&edited, run, "link").as_deref(),
        Some("https://litchi.example/")
    );
    assert_eq!(prop(&edited, run, "tooltip").as_deref(), Some("Litchi"));
    // Setting what the link holds is an exact no-op.
    let same = file
        .set(
            run,
            props! { "link" => "https://example.com/", "tooltip" => "Example" },
        )
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    // A script link is refused.
    let error = file
        .set(run, props! { "link" => "javascript:alert(1)" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    // Removing the link keeps its text.
    let removed = reopen(&file.remove(run).unwrap());
    assert_eq!(
        text(&removed, "/slide[1]/shape[1]/paragraph[1]"),
        "Visit example now"
    );
    assert_eq!(
        count(&removed, "/slide[1]/shape[1]/paragraph[1]/run/hyperlink"),
        0
    );
    // A plain run becomes a link.
    let added = file
        .add(
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            NewNode::new("hyperlink").prop("link", "https://first.example/"),
        )
        .unwrap();
    assert_eq!(
        added.path.to_string(),
        "/slide[1]/shape[1]/paragraph[1]/run[1]/hyperlink"
    );
    let added = reopen(&added.file);
    assert_eq!(
        prop(
            &added,
            "/slide[1]/shape[1]/paragraph[1]/run[1]/hyperlink",
            "link"
        )
        .as_deref(),
        Some("https://first.example/")
    );
    assert_eq!(
        text(&added, "/slide[1]/shape[1]/paragraph[1]"),
        "Visit example now"
    );
}

#[test]
fn shape_hyperlinks_set_add_and_remove() {
    let file = parts();
    // A listener's URL changes in place.
    let edited = reopen(
        &file
            .set(
                "/slide[1]/shape[2]/hyperlink",
                props! { "link" => "slide[2]" },
            )
            .unwrap(),
    );
    assert_eq!(
        prop(&edited, "/slide[1]/shape[2]/hyperlink", "link").as_deref(),
        Some("slide[2]")
    );
    // A shape without one gains a click action.
    let added = file
        .add(
            "/slide[1]/shape[5]",
            NewNode::new("hyperlink").prop("link", "lastslide"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/slide[1]/shape[5]/hyperlink");
    let added = reopen(&added.file);
    assert_eq!(
        prop(&added, "/slide[1]/shape[5]/hyperlink", "link").as_deref(),
        Some("lastslide")
    );
    assert_eq!(text(&added, "/slide[1]/shape[5]"), "Plain");
    // A second add is refused; `set` changes it.
    let error = added
        .add(
            "/slide[1]/shape[5]",
            NewNode::new("hyperlink").prop("link", "nextslide"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    // Removing a click action, and a draw:a link.
    let removed = reopen(&added.remove("/slide[1]/shape[5]/hyperlink").unwrap());
    assert!(removed.get("/slide[1]/shape[5]/hyperlink").is_err());
    let unwrapped = reopen(&file.remove("/slide[1]/shape[4]/hyperlink").unwrap());
    assert!(unwrapped.get("/slide[1]/shape[4]/hyperlink").is_err());
    assert_eq!(text(&unwrapped, "/slide[1]/shape[4]"), "Linked box");
    let relinked = reopen(
        &file
            .set(
                "/slide[1]/shape[4]/hyperlink",
                props! { "link" => "https://other.example/" },
            )
            .unwrap(),
    );
    assert_eq!(
        prop(&relinked, "/slide[1]/shape[4]/hyperlink", "link").as_deref(),
        Some("https://other.example/")
    );
    // A tooltip on a click action is refused, typed.
    let error = file
        .set("/slide[1]/shape[2]/hyperlink", props! { "tooltip" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn line_breaks_add_and_remove() {
    let file = parts();
    let paragraph = "/slide[1]/shape[1]/paragraph[2]";
    let added = file.add(paragraph, NewNode::new("linebreak")).unwrap();
    assert_eq!(added.path.to_string(), format!("{paragraph}/linebreak[3]"));
    let added = reopen(&added.file);
    assert_eq!(count(&added, &format!("{paragraph}/linebreak")), 3);
    let removed = reopen(
        &file
            .remove(format!("{paragraph}/linebreak[1]").as_str())
            .unwrap(),
    );
    assert_eq!(count(&removed, &format!("{paragraph}/linebreak")), 1);
    assert_eq!(text(&removed, paragraph), "Line oneLine two\nLine three");
}

#[test]
fn read_only_elements_refuse_writes() {
    let file = parts();
    for path in [
        "/slide[1]/table[1]/col[1]",
        "/slide[1]/chart[1]/series[1]",
        "/slide[1]/chart[1]/axis[1]",
        "/slide[1]/equation[1]",
        "/slide[1]/video[1]",
    ] {
        let error = file.set(path, props! { "name" => "x" }).unwrap_err();
        assert!(
            matches!(
                error.code(),
                Code::UnsupportedCapability | Code::UnsupportedProperty
            ),
            "{path}: {error}"
        );
    }
}
