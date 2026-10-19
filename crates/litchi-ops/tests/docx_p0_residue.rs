//! The DOCX P0 residue of record 0335 through the unified API: picture
//! `set`, `remove` and `query`; comment `set` (replies refused); SmartArt
//! `set`, `remove` and its nodes; and the series and axis nodes of DOCX
//! charts. Every write is read back on both owners after a byte round
//! trip, refusals leave the input unchanged, and an identical set is an
//! exact no-op.
#![cfg(feature = "docx")]

use base64::Engine as _;
use litchi_core::FileFormat;
use litchi_ops::capabilities::Capabilities;
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode, props};

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

fn ok<T>(result: litchi_ops::Result<T>) -> T {
    result.unwrap_or_else(|error| panic!("{error} ({:?})", error.property()))
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

fn document() -> File {
    litchi_ops::create(FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Logo: "))
        .unwrap()
        .file
        .add("/body", NewNode::new("paragraph").prop("text", "Second"))
        .unwrap()
        .file
}

/// A PNG header of `width` x `height` pixels, as a data URI.
fn png(width: u32, height: u32) -> String {
    let mut png = vec![0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13];
    png.extend_from_slice(b"IHDR");
    png.extend_from_slice(&width.to_be_bytes());
    png.extend_from_slice(&height.to_be_bytes());
    png.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    format!(
        "data:image/png;base64,{}",
        base64::engine::general_purpose::STANDARD.encode(png)
    )
}

fn with_pictures() -> File {
    let file = document();
    let file = ok(file.add(
        "/body/p[1]",
        NewNode::new("picture")
            .prop("src", png(2, 1))
            .prop("alt", "first"),
    ))
    .file;
    ok(file.add(
        "/body/p[1]",
        NewNode::new("picture")
            .prop("src", png(4, 4))
            .prop("name", "Second"),
    ))
    .file
}

#[test]
fn pictures_are_set_queried_and_removed() {
    for eager in [false, true] {
        let file = reopen(&with_pictures(), eager);
        assert_eq!(file.query("picture").unwrap().count(), 2);
        let edited = ok(file.set(
            "/body/p[1]/picture[2]",
            props! {
                "width" => "2in",
                "alt" => "The logo",
                "name" => "Logo",
            },
        ));
        for eager in [false, true] {
            let edited = reopen(&edited, eager);
            let path = "/body/p[1]/picture[2]";
            assert_eq!(prop(&edited, path, "width").as_deref(), Some("5.08cm"));
            // One side keeps the other.
            assert_eq!(prop(&edited, path, "height").as_deref(), Some("3pt"));
            assert_eq!(prop(&edited, path, "alt").as_deref(), Some("The logo"));
            assert_eq!(prop(&edited, path, "name").as_deref(), Some("Logo"));
            assert_eq!(
                prop(&edited, "/body/p[1]/picture[1]", "alt").as_deref(),
                Some("first")
            );
            assert_eq!(text(&edited, "/body/p[1]").as_deref(), Some("Logo: "));
        }
        // A new image: a new relationship, the extent kept.
        let before = prop(&edited, "/body/p[1]/picture[1]", "relId");
        let replaced = ok(edited.set(
            "/body/p[1]/picture[1]",
            props! { "src" => png(8, 8).as_str() },
        ));
        let replaced = reopen(&replaced, eager);
        assert_ne!(prop(&replaced, "/body/p[1]/picture[1]", "relId"), before);
        assert_eq!(
            prop(&replaced, "/body/p[1]/picture[1]", "width"),
            prop(&edited, "/body/p[1]/picture[1]", "width")
        );
        // Remove: the first picture goes, the second becomes the first.
        let removed = ok(replaced.remove("/body/p[1]/picture[1]"));
        let removed = reopen(&removed, eager);
        assert_eq!(removed.query("picture").unwrap().count(), 1);
        assert_eq!(
            prop(&removed, "/body/p[1]/picture[1]", "name").as_deref(),
            Some("Logo")
        );
        assert_eq!(text(&removed, "/body/p[1]").as_deref(), Some("Logo: "));
    }
}

#[test]
fn picture_refusals_and_no_ops_leave_bytes_unchanged() {
    let file = with_pictures();
    let before = file.to_bytes().unwrap();
    for (props, code) in [
        (props! { "width" => "0in" }, Code::InvalidValue),
        (
            props! { "src" => "data:image/png;base64,bm90IGFuIGltYWdl" },
            Code::InvalidValue,
        ),
        (props! { "relId" => "rId9" }, Code::UnsupportedProperty),
        (props! { "rotation" => "90" }, Code::UnsupportedProperty),
    ] {
        let error = file.set("/body/p[1]/picture[1]", props).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
    }
    assert!(file.remove("/body/p[1]/picture[3]").is_err());
    assert_eq!(file.to_bytes().unwrap(), before);
    // The same values again publish the same bytes.
    let same = ok(file.set("/body/p[1]/picture[1]", props! { "alt" => "first" }));
    assert_eq!(same.to_bytes().unwrap(), before);
}

fn with_comment() -> File {
    ok(document().add(
        "/body/p[1]",
        NewNode::new("comment")
            .prop("author", "Ada")
            .prop("initials", "AL")
            .prop("text", "Draft"),
    ))
    .file
}

#[test]
fn comments_are_set() {
    for eager in [false, true] {
        let file = reopen(&with_comment(), eager);
        let edited = ok(file.set(
            "/comments/comment[1]",
            props! {
                "author" => "Grace",
                "initials" => "GH",
                "date" => "2026-10-17T00:00:00Z",
                "text" => "Final",
            },
        ));
        for eager in [false, true] {
            let edited = reopen(&edited, eager);
            let path = "/comments/comment[1]";
            assert_eq!(prop(&edited, path, "author").as_deref(), Some("Grace"));
            assert_eq!(prop(&edited, path, "initials").as_deref(), Some("GH"));
            assert_eq!(
                prop(&edited, path, "date").as_deref(),
                Some("2026-10-17T00:00:00Z")
            );
            assert_eq!(prop(&edited, path, "text").as_deref(), Some("Final"));
            assert_eq!(text(&edited, "/body/p[1]").as_deref(), Some("Logo: "));
        }
        let before = edited.to_bytes().unwrap();
        let same = ok(edited.set("/comments/comment[1]", props! { "author" => "Grace" }));
        assert_eq!(same.to_bytes().unwrap(), before);
    }
}

/// Record 0345 writes replies and resolution (`docx_p0_final.rs`); what
/// stays refused: a reply parent on `set`, a value of the wrong type, an
/// empty author and the identity.
#[test]
fn comment_replies_and_bad_values_are_refused() {
    let file = with_comment();
    let before = file.to_bytes().unwrap();
    let error = file
        .set("/comments/comment[1]", props! { "parentId" => "1" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    for key in ["done", "resolved"] {
        let error = file
            .set("/comments/comment[1]", props! { key => "maybe" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue, "{error}");
    }
    let error = file
        .set("/comments/comment[1]", props! { "author" => "" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    let error = file
        .set("/comments/comment[1]", props! { "id" => "4" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    assert_eq!(file.to_bytes().unwrap(), before);
}

fn with_smartart() -> File {
    ok(document().add(
        "/body/p[2]",
        NewNode::new("smartart")
            .prop("layout", "process")
            .prop("text", "Plan\nBuild\n\tTest\nShip"),
    ))
    .file
}

#[test]
fn smartart_graphics_and_nodes_are_edited() {
    for eager in [false, true] {
        let file = reopen(&with_smartart(), eager);
        // Nodes read in outline order, nested under their parents.
        assert_eq!(
            text(&file, "/smartart[1]/node[2]").as_deref(),
            Some("Build")
        );
        assert_eq!(
            text(&file, "/smartart[1]/node[2]/node[1]").as_deref(),
            Some("Test")
        );
        assert_eq!(
            prop(&file, "/smartart[1]/node[2]/node[1]", "level").as_deref(),
            Some("2")
        );
        assert_eq!(file.query("node").unwrap().count(), 4);
        // Node text, node add and node remove.
        let edited = ok(file.set("/smartart[1]/node[1]", props! { "text" => "Design" }));
        let added = ok(edited.add("/smartart[1]", NewNode::new("node").prop("text", "Review")));
        assert_eq!(added.path.to_string(), "/smartart[1]/node[4]");
        let child = ok(added.file.add(
            "/smartart[1]/node[2]",
            NewNode::new("node").prop("text", "Lint"),
        ));
        assert_eq!(child.path.to_string(), "/smartart[1]/node[2]/node[2]");
        let removed = ok(child.file.remove("/smartart[1]/node[3]"));
        for eager in [false, true] {
            let removed = reopen(&removed, eager);
            assert_eq!(
                prop(&removed, "/smartart[1]", "text").as_deref(),
                Some("Design\nBuild\n\tTest\n\tLint\nReview")
            );
            assert_eq!(
                prop(&removed, "/smartart[1]", "nodeCount").as_deref(),
                Some("5")
            );
        }
        // The whole outline and the layout.
        let rebuilt = ok(removed.set(
            "/smartart[1]",
            props! { "layout" => "cycle", "text" => "A\nB\nC" },
        ));
        let rebuilt = reopen(&rebuilt, eager);
        assert_eq!(
            prop(&rebuilt, "/smartart[1]", "layout").as_deref(),
            Some("cycle")
        );
        assert_eq!(
            prop(&rebuilt, "/smartart[1]", "text").as_deref(),
            Some("A\nB\nC")
        );
        // Removing the graphic removes its drawing and parts.
        let gone = ok(rebuilt.remove("/smartart[1]"));
        let gone = reopen(&gone, eager);
        assert_eq!(gone.query("smartart").unwrap().count(), 0);
        assert_eq!(text(&gone, "/body/p[2]").as_deref(), Some("Second"));
        assert!(gone.get("/smartart[1]").is_err());
    }
}

#[test]
fn smartart_refusals_and_no_ops_leave_bytes_unchanged() {
    let file = with_smartart();
    let before = file.to_bytes().unwrap();
    for (path, props, code) in [
        (
            "/smartart[1]",
            props! { "layout" => "spiral" },
            Code::InvalidValue,
        ),
        (
            "/smartart[1]",
            props! { "name" => "X" },
            Code::UnsupportedCapability,
        ),
        (
            "/smartart[1]/node[1]",
            props! { "level" => "2" },
            Code::UnsupportedCapability,
        ),
    ] {
        let error = file.set(path, props).unwrap_err();
        assert!(
            error.code() == code || error.code() == Code::UnsupportedProperty,
            "{path}: {error}"
        );
    }
    let error = file
        .add(
            "/smartart[1]",
            NewNode::new("node")
                .prop("text", "Late")
                .at(Placement::Position(litchi_core::Position::new(9))),
        )
        .unwrap_err();
    assert!(error.code() == Code::OutOfRange, "{error}");
    assert!(file.get("/smartart[1]/node[9]").is_err());
    assert_eq!(file.to_bytes().unwrap(), before);
    let same = ok(file.set("/smartart[1]/node[1]", props! { "text" => "Plan" }));
    assert_eq!(same.to_bytes().unwrap(), before);
}

fn with_chart() -> File {
    ok(document().add(
        "/body/p[2]",
        NewNode::new("chart")
            .prop("chartType", "column")
            .prop("title", "Sales")
            .prop("categories", "Q1,Q2,Q3")
            .prop("data", "East:1,2,3;West:4,5,6")
            .prop("catTitle", "Quarter"),
    ))
    .file
}

#[test]
fn chart_series_and_axes_are_nodes() {
    for eager in [false, true] {
        let file = reopen(&with_chart(), eager);
        for chart in ["/body/p[2]/chart[1]", "/chart[1]"] {
            let series = format!("{chart}/series[2]");
            assert_eq!(prop(&file, &series, "name").as_deref(), Some("West"));
            assert_eq!(prop(&file, &series, "values").as_deref(), Some("4,5,6"));
            assert_eq!(
                prop(&file, &series, "categories").as_deref(),
                Some("Q1,Q2,Q3")
            );
            let axis = format!("{chart}/axis[@role=category]");
            assert_eq!(prop(&file, &axis, "title").as_deref(), Some("Quarter"));
            assert_eq!(prop(&file, &axis, "role").as_deref(), Some("category"));
        }
        assert_eq!(file.query("series").unwrap().count(), 2);
        assert_eq!(file.query("axis").unwrap().count(), 2);
        let children: Vec<String> = file
            .get("/body/p[2]/chart[1]")
            .unwrap()
            .children()
            .map(|child| child.unwrap().path().to_string())
            .collect();
        assert_eq!(
            children,
            [
                "/body/p[2]/chart[1]/series[1]",
                "/body/p[2]/chart[1]/series[2]",
                "/body/p[2]/chart[1]/axis[1]",
                "/body/p[2]/chart[1]/axis[2]",
            ]
        );
        let edited = ok(file.set(
            "/chart[1]/series[1]",
            props! { "name" => "North", "values" => "7,8,9", "color" => "70AD47" },
        ));
        let edited = ok(edited.set(
            "/body/p[2]/chart[1]/axis[@role=value]",
            props! { "title" => "Units", "min" => "0", "max" => "20" },
        ));
        for eager in [false, true] {
            let edited = reopen(&edited, eager);
            let series = "/chart[1]/series[1]";
            assert_eq!(prop(&edited, series, "name").as_deref(), Some("North"));
            assert_eq!(prop(&edited, series, "values").as_deref(), Some("7,8,9"));
            assert_eq!(prop(&edited, series, "color").as_deref(), Some("70AD47"));
            let axis = "/chart[1]/axis[@role=value]";
            assert_eq!(prop(&edited, axis, "title").as_deref(), Some("Units"));
            assert_eq!(prop(&edited, axis, "min").as_deref(), Some("0"));
            assert_eq!(prop(&edited, axis, "max").as_deref(), Some("20"));
            assert_eq!(
                prop(&edited, "/chart[1]", "title").as_deref(),
                Some("Sales")
            );
        }
        let before = edited.to_bytes().unwrap();
        for (path, props, code) in [
            (
                "/chart[1]/series[1]",
                props! { "categories" => "A" },
                Code::UnsupportedCapability,
            ),
            (
                "/chart[1]/series[1]",
                props! { "color" => "notacolor" },
                Code::InvalidValue,
            ),
            (
                "/chart[1]/axis[@role=category]",
                props! { "min" => "1" },
                Code::UnsupportedCapability,
            ),
            (
                "/chart[1]/series[3]",
                props! { "name" => "x" },
                Code::PathNotFound,
            ),
        ] {
            let error = edited.set(path, props).unwrap_err();
            assert!(
                error.code() == code || error.code() == Code::UnsupportedProperty,
                "{path}: {error}"
            );
        }
        assert_eq!(edited.to_bytes().unwrap(), before);
    }
}

#[test]
fn capabilities_claim_the_new_cells() {
    let capabilities = Capabilities::of(FileFormat::Docx);
    let operations = |name: &str| {
        capabilities
            .elements()
            .find(|element| element.name() == name)
            .unwrap_or_else(|| panic!("{name}"))
            .operations()
    };
    let picture = operations("picture");
    assert!(picture.add && picture.set && picture.get && picture.query && picture.remove);
    let comment = operations("comment");
    assert!(comment.add && comment.set && comment.remove);
    let smartart = operations("smartart");
    assert!(smartart.add && smartart.set && smartart.remove);
    let node = operations("node");
    assert!(node.add && node.set && node.get && node.query && node.remove);
    for name in ["series", "axis"] {
        let element = operations(name);
        assert!(element.set && element.get && element.query, "{name}");
        assert!(!element.add && !element.remove, "{name}");
    }
}

/// A Strict Word document with a SmartArt graphic Word made (LibreOffice
/// `sw/qa/extras/ooxmlexport/data/strict.docx`, MPL-2.0, already tracked).
fn foreign() -> File {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx");
    litchi_ops::open_bytes(std::fs::read(path).unwrap()).unwrap()
}

#[test]
fn a_foreign_smartart_takes_text_in_place_and_refuses_structure() {
    let file = foreign();
    let before = file.to_bytes().unwrap();
    let first = text(&file, "/smartart[1]/node[1]").unwrap();
    let edited = ok(file.set("/smartart[1]/node[1]", props! { "text" => "Edited" }));
    let edited = reopen(&edited, false);
    assert_eq!(
        text(&edited, "/smartart[1]/node[1]").as_deref(),
        Some("Edited")
    );
    assert_ne!(first, "Edited");
    // Nodes of a graphic Litchi did not author are not added or removed.
    for error in [
        file.add("/smartart[1]", NewNode::new("node").prop("text", "New"))
            .unwrap_err(),
        file.remove("/smartart[1]/node[1]").unwrap_err(),
    ] {
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
    }
    assert_eq!(file.to_bytes().unwrap(), before);
    // A layout set rebuilds it as a Litchi graphic from its outline, after
    // which nodes can be added.
    let rebuilt = ok(file.set("/smartart[1]", props! { "layout" => "list" }));
    let rebuilt = ok(rebuilt.add("/smartart[1]", NewNode::new("node").prop("text", "New")));
    let rebuilt = reopen(&rebuilt.file, true);
    assert_eq!(
        prop(&rebuilt, "/smartart[1]", "layout").as_deref(),
        Some("list")
    );
    assert!(
        prop(&rebuilt, "/smartart[1]", "text")
            .unwrap()
            .ends_with("\nNew")
    );
    // The whole graphic goes with its parts.
    let gone = ok(file.remove("/smartart[1]"));
    let gone = reopen(&gone, false);
    assert_eq!(
        gone.query("smartart").unwrap().count(),
        file.query("smartart").unwrap().count() - 1
    );
}
