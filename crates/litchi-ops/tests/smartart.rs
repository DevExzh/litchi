//! Typed SmartArt CRUD through the unified API (record 0284): add, get, set
//! and remove of PPTX SmartArt graphics and their nodes, Litchi-authored and
//! foreign, with round trips, no-ops, refusals and the capability report.
#![cfg(feature = "pptx")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode, props};

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
}

fn deck() -> File {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
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

fn smartart() -> NewNode<'static> {
    NewNode::new("smartart")
        .prop("layout", "process")
        .prop("text", "Plan\n\tScope\n\tBudget\nBuild\nShip")
}

const FOREIGN: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/smartart-accent-process.pptx");

#[test]
fn smartart_add_get_set_remove_round_trip() {
    let source = deck();
    let before = source.to_bytes().unwrap();
    let added = source
        .add(
            "/slide[1]",
            smartart()
                .prop("x", "1in")
                .prop("width", "8in")
                .prop("height", "4in"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/slide[1]/smartart[1]");
    assert_eq!(source.to_bytes().unwrap(), before);
    // Deterministic.
    let again = deck()
        .add(
            "/slide[1]",
            smartart()
                .prop("x", "1in")
                .prop("width", "8in")
                .prop("height", "4in"),
        )
        .unwrap();
    assert_eq!(
        again.file.to_bytes().unwrap(),
        added.file.to_bytes().unwrap()
    );
    for eager in [false, true] {
        let file = reopen(&added.file, eager);
        let art = "/slide[1]/smartart[1]";
        assert_eq!(prop(&file, art, "layout").as_deref(), Some("process"));
        assert_eq!(
            prop(&file, art, "layoutId").as_deref(),
            Some("urn:microsoft.com/office/officeart/2005/8/layout/process1")
        );
        assert_eq!(prop(&file, art, "nodeCount").as_deref(), Some("5"));
        assert_eq!(prop(&file, art, "name").as_deref(), Some("Diagram 1"));
        assert_eq!(prop(&file, art, "width").as_deref(), Some("20.32cm"));
        assert_eq!(
            text(&file, art).as_deref(),
            Some("Plan\n\tScope\n\tBudget\nBuild\nShip")
        );
        assert_eq!(file.get(art).unwrap().children().count(), 3);
        assert_eq!(
            text(&file, "/slide[1]/smartart[1]/node[1]/node[2]").as_deref(),
            Some("Budget")
        );
        assert_eq!(
            prop(&file, "/slide[1]/smartart[1]/node[1]/node[2]", "level").as_deref(),
            Some("2")
        );
        assert_eq!(
            file.get("/slide[1]/smartart[1]/node[3]")
                .unwrap()
                .path()
                .to_string(),
            "/slide[1]/smartart[1]/node[3]"
        );
        assert_eq!(file.query("node").unwrap().count(), 5);

        // Node text.
        let edited = file
            .set("/slide[1]/smartart[1]/node[2]", props! { "text" => "Make" })
            .unwrap();
        assert_eq!(
            text(&edited, "/slide[1]/smartart[1]/node[2]").as_deref(),
            Some("Make")
        );
        // No-op.
        let same = edited
            .set("/slide[1]/smartart[1]/node[2]", props! { "text" => "Make" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), edited.to_bytes().unwrap());

        // Add nodes: at the end, at a position, as a child.
        let grown = edited
            .add(
                "/slide[1]/smartart[1]",
                NewNode::new("node").prop("text", "Learn"),
            )
            .unwrap();
        assert_eq!(grown.path.to_string(), "/slide[1]/smartart[1]/node[4]");
        let grown = grown
            .file
            .add(
                "/slide[1]/smartart[1]",
                NewNode::new("node")
                    .prop("text", "Idea")
                    .at(Placement::Start),
            )
            .unwrap();
        assert_eq!(grown.path.to_string(), "/slide[1]/smartart[1]/node[1]");
        let grown = grown
            .file
            .add(
                "/slide[1]/smartart[1]/node[3]",
                NewNode::new("node").prop("text", "Prototype"),
            )
            .unwrap();
        assert_eq!(
            grown.path.to_string(),
            "/slide[1]/smartart[1]/node[3]/node[1]"
        );
        let grown = reopen(&grown.file, eager);
        assert_eq!(
            text(&grown, "/slide[1]/smartart[1]").as_deref(),
            Some("Idea\nPlan\n\tScope\n\tBudget\nMake\n\tPrototype\nShip\nLearn")
        );

        // Remove a node and its children.
        let shrunk = grown.remove("/slide[1]/smartart[1]/node[2]").unwrap();
        assert_eq!(
            text(&shrunk, "/slide[1]/smartart[1]").as_deref(),
            Some("Idea\nMake\n\tPrototype\nShip\nLearn")
        );

        // Layout and whole-outline changes.
        let cycled = shrunk
            .set("/slide[1]/smartart[1]", props! { "layout" => "cycle" })
            .unwrap();
        assert_eq!(prop(&cycled, art, "layout").as_deref(), Some("cycle"));
        assert_eq!(
            text(&cycled, art).as_deref(),
            Some("Idea\nMake\n\tPrototype\nShip\nLearn")
        );
        let replaced = cycled
            .set(art, props! { "text" => "A\nB", "layout" => "hierarchy" })
            .unwrap();
        assert_eq!(text(&replaced, art).as_deref(), Some("A\nB"));
        assert_eq!(prop(&replaced, art, "layout").as_deref(), Some("hierarchy"));

        // Removing the frame removes every diagram part.
        let removed = replaced.remove(art).unwrap();
        assert!(removed.get(art).is_err());
    }
}

#[test]
fn every_layout_is_authored() {
    for layout in [
        "list",
        "process",
        "cycle",
        "hierarchy",
        "pyramid",
        "matrix",
        "venn",
    ] {
        let file = deck()
            .add(
                "/slide[1]",
                NewNode::new("smartart")
                    .prop("layout", layout)
                    .prop("text", "A\n\tA1\nB\nC\nD"),
            )
            .unwrap()
            .file;
        let file = reopen(&file, true);
        assert_eq!(
            prop(&file, "/slide[1]/smartart[1]", "layout").as_deref(),
            Some(layout)
        );
        assert_eq!(
            prop(&file, "/slide[1]/smartart[1]", "nodeCount").as_deref(),
            Some("5")
        );
    }
}

#[test]
fn foreign_smartart_reads_and_edits_text_in_place() {
    let file = litchi_ops::open_bytes(FOREIGN.to_vec()).unwrap();
    let art = "/slide[1]/smartart[1]";
    assert_eq!(prop(&file, art, "layout").as_deref(), Some("process3"));
    assert_eq!(text(&file, art).as_deref(), Some("a\n\tb\ncc\n\td"));
    // A text edit keeps everything else.
    let edited = file
        .set(
            "/slide[1]/smartart[1]/node[2]/node[1]",
            props! { "text" => "dd" },
        )
        .unwrap();
    let edited = reopen(&edited, false);
    assert_eq!(text(&edited, art).as_deref(), Some("a\n\tb\ncc\n\tdd"));
    assert_eq!(prop(&edited, art, "layout").as_deref(), Some("process3"));

    // Structural edits of a foreign graphic are refused; the input stays.
    let before = file.to_bytes().unwrap();
    let refused = file
        .add(art, NewNode::new("node").prop("text", "x"))
        .unwrap_err();
    assert_eq!(refused.code(), Code::UnsupportedCapability);
    let refused = file.remove("/slide[1]/smartart[1]/node[1]").unwrap_err();
    assert_eq!(refused.code(), Code::UnsupportedCapability);
    assert_eq!(file.to_bytes().unwrap(), before);

    // A layout change rebuilds it as a Litchi graphic; then nodes can be added.
    let rebuilt = file.set(art, props! { "layout" => "process" }).unwrap();
    assert_eq!(text(&rebuilt, art).as_deref(), Some("a\n\tb\ncc\n\td"));
    let grown = rebuilt
        .add(art, NewNode::new("node").prop("text", "e"))
        .unwrap()
        .file;
    assert_eq!(text(&grown, art).as_deref(), Some("a\n\tb\ncc\n\td\ne"));
}

#[test]
fn refusals_leave_the_input_unchanged() {
    let file = deck();
    let before = file.to_bytes().unwrap();
    for (node, code) in [
        (
            NewNode::new("smartart").prop("layout", "list"),
            Code::MissingProperty,
        ),
        (
            NewNode::new("smartart")
                .prop("layout", "spiral")
                .prop("text", "A"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("smartart").prop("text", "A\n\t\tB"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("smartart").prop("text", "  \n"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("smartart")
                .prop("text", "A")
                .prop("colour", "red"),
            Code::UnsupportedProperty,
        ),
    ] {
        let error = file.add("/slide[1]", node).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
    }
    assert_eq!(file.to_bytes().unwrap(), before);
    let file = file.add("/slide[1]", smartart()).unwrap().file;
    let error = file
        .set("/slide[1]/smartart[1]/node[1]", props! { "level" => 3 })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    let error = file
        .add(
            "/slide[1]/smartart[1]",
            NewNode::new("node")
                .prop("text", "x")
                .at(Placement::Position(litchi_core::Position::new(9))),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::OutOfRange, "{error}");
    assert!(file.get("/slide[1]/smartart[1]/node[9]").is_err());
}

#[test]
fn smartart_capabilities_report_the_registry() {
    use litchi_ops::capabilities::Capabilities;
    let caps = Capabilities::of(FileFormat::Pptx);
    let art = caps.element("smartart").expect("smartart element");
    assert!(art.operations().add && art.operations().set && art.operations().remove);
    let layout = art.property("layout").expect("layout");
    assert!(layout.add() && layout.set() && layout.get());
    let node = caps.element("node").expect("node element");
    assert!(node.operations().add && node.operations().set && node.operations().remove);
    assert!(node.property("text").expect("text").set());
}
