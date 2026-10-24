//! Public PPTX unified-read coverage (M1 R4c/R5).
//!
//! The fixture deliberately contains mixed shape-tree kinds, nested groups,
//! repeated native identities, text bodies, notes and legacy comments. Every
//! behavioral assertion goes through `File`, so the source-backed and eager
//! owners exercise the same dispatch boundary.

#![cfg(feature = "pptx")]

use std::fmt;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_ops::node::{Node, Value, ViewOptions};
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, OpenOptions, SourceMode};

#[path = "support/pptx_presentation.rs"]
mod presentation;

fn files(bytes: Vec<u8>) -> [File; 2] {
    let source = litchi_ops::open_bytes(bytes.clone()).expect("source-backed PPTX");
    let eager = litchi_ops::open_bytes_with(
        bytes,
        &OpenOptions::default().source_mode(SourceMode::Eager),
    )
    .expect("eager PPTX");
    assert!(matches!(&source, File::Pptx(pptx) if pptx.is_source_backed()));
    assert!(matches!(&eager, File::Pptx(pptx) if !pptx.is_source_backed()));
    [source, eager]
}

fn text(node: &Node<'_>) -> Option<String> {
    node.view()
        .text()
        .expect("node text")
        .map(|text| text.into_owned())
}

fn property(node: &Node<'_>, key: &str) -> Option<String> {
    node.view()
        .format()
        .expect("node format")
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

fn property_value(node: &Node<'_>, key: &str) -> Option<Value<'static>> {
    node.view()
        .format()
        .expect("node format")
        .find(|property| property.key() == key)
        .map(|property| property.into_value().into_owned())
}

fn query_paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| {
            node.unwrap_or_else(|error| panic!("{selector}: {error}"))
                .positional_path()
                .to_string()
        })
        .collect()
}

fn view_json(node: &Node<'_>, depth: u32) -> String {
    serde_json::to_string(&node.view_with(ViewOptions::default().depth(depth)))
        .expect("serialize node view")
}

struct MutatingSink {
    path: PathBuf,
    output: String,
    mutated: bool,
}

impl fmt::Write for MutatingSink {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        if !self.mutated {
            fs::write(&self.path, b"changed source").expect("mutate source from sink");
            self.mutated = true;
        }
        self.output.push_str(text);
        Ok(())
    }
}

#[test]
fn public_file_source_and_eager_owners_have_the_same_projection() {
    let bytes = presentation::bytes();
    let mut snapshots = Vec::new();

    for file in files(bytes) {
        assert_eq!(file.format(), Some(litchi_core::FileFormat::Pptx));
        assert_eq!(file.root().expect("root").kind(), "presentation");
        assert!(
            file.text()
                .expect("presentation text")
                .contains("Quarterly Review")
        );

        let selected = [
            "/slide[1]",
            "/slide[1]/shape[3]",
            "/slide[1]/shape[3]/paragraph[1]/run[1]",
            "/slide[1]/notes",
            "/slide[1]/comment[2]",
            "/slide[2]/smartart[1]",
        ];
        let rows: Vec<_> = selected
            .into_iter()
            .map(|path| {
                let node = file
                    .get(path)
                    .unwrap_or_else(|error| panic!("{path}: {error}"));
                (
                    node.kind(),
                    node.positional_path().to_string(),
                    text(&node),
                    view_json(&node, 1),
                )
            })
            .collect();
        snapshots.push((
            file.text().expect("text"),
            file.view(Mode::Text).expect("text view"),
            rows,
        ));
    }

    assert_eq!(snapshots[0], snapshots[1]);
}

#[test]
fn shape_kinds_have_independent_ordinals_and_keep_z_order() {
    for file in files(presentation::bytes()) {
        let slide = file.get("/slide[1]").expect("first slide");
        let direct: Vec<_> = slide
            .children()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("direct shape-tree children");
        assert_eq!(
            direct.iter().map(Node::kind).collect::<Vec<_>>(),
            [
                "title",
                "placeholder",
                "shape",
                "picture",
                "table",
                "group",
                "connector",
            ]
        );

        // `zorder` is the position in the containing shape tree, while the
        // path ordinal is per kind. The nested group has no root z-order.
        for (position, node) in direct.iter().enumerate() {
            assert_eq!(
                property(node, "zorder"),
                Some((position + 1).to_string()),
                "{}",
                node.positional_path()
            );
        }
        let nested = file
            .get("/slide[1]/group[1]/group[1]/shape[1]")
            .expect("nested shape");
        assert_eq!(property(&nested, "zorder"), None);

        assert_eq!(
            text(&file.get("/slide[1]/shape[1]").expect("title")),
            Some("Quarterly Review".to_owned())
        );
        assert_eq!(
            text(&file.get("/slide[1]/picture[1]").expect("picture")),
            None
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/table[1]/tr[2]/tc[2]")
                    .expect("table cell"),
            ),
            Some("1200".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/group[1]/group[1]/shape[1]")
                    .expect("deep group shape"),
            ),
            Some("Deep text".to_owned())
        );
        assert_eq!(
            text(&file.get("/slide[2]/smartart[1]").expect("SmartArt")),
            None
        );
        assert_eq!(text(&file.get("/slide[2]/ole[1]").expect("OLE")), None);
        assert_eq!(
            query_paths(&file, "group shape"),
            [
                "/slide[1]/group[1]/shape[1]",
                "/slide[1]/group[1]/group[1]/shape[1]",
            ]
        );
        assert_eq!(
            query_paths(&file, "slide > shape"),
            [
                "/slide[1]/shape[1]",
                "/slide[1]/shape[2]",
                "/slide[1]/shape[3]",
                "/slide[2]/shape[1]",
                "/slide[2]/shape[2]",
                "/slide[2]/shape[3]",
                "/slide[2]/shape[4]",
                "/slide[2]/shape[5]",
            ]
        );
    }
}

#[test]
fn generic_frames_fail_with_an_explicit_unsupported_capability() {
    for file in files(presentation::bytes_with_generic_frame()) {
        let error = file
            .get("/slide[1]/shape[1]")
            .expect_err("generic frames have no unified projection");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Pptx));
        assert!(error.to_string().contains("unified projection"));
    }
}

#[test]
fn names_are_semantic_but_native_ids_require_the_raw_escape() {
    for file in files(presentation::bytes()) {
        let morph = file
            .get("/slide[2]/shape[@name='morph box']")
            .expect("case-insensitive morph name");
        assert_eq!(text(&morph), Some("Gamma".to_owned()));

        let ambiguous = file
            .get("/slide[2]/shape[@name='Repeated A']")
            .expect_err("duplicate semantic names are ambiguous");
        assert_eq!(ambiguous.code(), Code::InvalidPath);
        assert!(
            !ambiguous.to_string().contains("Repeated A"),
            "authored names must not leak from ambiguity diagnostics: {ambiguous}"
        );

        let ordinary_id = file
            .get("/slide[2]/shape[@id=5]")
            .expect_err("native IDs are not ordinary path selectors");
        assert_eq!(ordinary_id.code(), Code::InvalidPath);

        let raw_id =
            litchi_ops::path::raw::path("/slide[2]/shape[@id=6]").expect("raw native-id path");
        let first = file.get(raw_id).expect("raw native-id lookup");
        assert_eq!(text(&first), Some("Gamma".to_owned()));
        assert_eq!(first.positional_path().to_string(), "/slide[2]/shape[3]");

        let keys: Vec<_> = first
            .view()
            .format()
            .expect("ordinary shape properties")
            .map(|property| property.key())
            .collect();
        assert!(!keys.contains(&"id"));
        assert!(!keys.contains(&"native_id"));
        let json = view_json(&first, 0);
        assert!(!json.contains("\"id\""), "native identity leaked: {json}");

        // Query predicates keep semantic names ordinary and can intentionally
        // return every matching node; ambiguity is a path-resolution rule.
        assert_eq!(
            query_paths(&file, "shape[name='Repeated A']"),
            ["/slide[2]/shape[1]", "/slide[2]/shape[2]"]
        );
    }
}

#[test]
fn paragraphs_runs_notes_and_comments_are_reachable() {
    for file in files(presentation::bytes()) {
        let paragraph = file
            .get("/slide[1]/shape[3]/paragraph[1]")
            .expect("text paragraph");
        assert!(
            text(&paragraph)
                .expect("paragraph text")
                .contains("Revenue grew")
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[1]/run[1]")
                    .expect("first run"),
            ),
            Some("Revenue grew".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[1]/run[2]")
                    .expect("second run"),
            ),
            Some(" 25%".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[1]/linebreak[1]")
                    .expect("line break"),
            ),
            Some(String::new())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[2]")
                    .expect("second text paragraph"),
            ),
            Some("Follow-up detail".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[2]/run[1]")
                    .expect("first run of the second paragraph"),
            ),
            Some("Follow-up".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/paragraph[2]/run[2]")
                    .expect("second run of the second paragraph"),
            ),
            Some(" detail".to_owned())
        );
        assert_eq!(
            text(
                &file
                    .get("/slide[1]/shape[3]/run[3]")
                    .expect("shape-global run in the second paragraph"),
            ),
            Some("Follow-up".to_owned())
        );

        let notes = file.get("/slide[1]/notes").expect("speaker notes");
        assert_eq!(text(&notes), Some("Remember the forecast".to_owned()));

        let first = file.get("/slide[1]/comment[1]").expect("first comment");
        let second = file.get("/slide[1]/comment[2]").expect("second comment");
        assert_eq!(text(&first), Some("Check the numbers".to_owned()));
        assert_eq!(text(&second), Some("Looks good".to_owned()));
        assert_eq!(property(&first, "author"), Some("Ada Lovelace".to_owned()));
        assert_eq!(property(&second, "author"), Some("Alan Turing".to_owned()));
        // The producer comment index is native metadata and stays out of the
        // ordinary projection; raw access is the explicit escape hatch.
        assert_eq!(property_value(&first, "index"), None);
        assert_eq!(query_paths(&file, "notes"), ["/slide[1]/notes"]);
        assert_eq!(
            query_paths(&file, "comment"),
            ["/slide[1]/comment[1]", "/slide[1]/comment[2]"]
        );
    }
}

#[test]
fn modern_comments_replies_and_master_layout_paths_match_source_and_eager() {
    let mut snapshots = Vec::new();

    for file in files(presentation::bytes()) {
        let master = file
            .get("/slideMaster[1]")
            .expect("canonical slide-master path");
        assert_eq!(master.kind(), "slideMaster");
        assert_eq!(
            property(&master, "name"),
            Some("Corporate Master".to_owned())
        );
        assert_eq!(
            property_value(&master, "preserved"),
            Some(Value::Bool(true))
        );
        let root_children = file
            .root()
            .expect("root")
            .children()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("presentation children");
        assert_eq!(
            root_children
                .iter()
                .filter(|child| child.kind() == "slideMaster")
                .count(),
            1
        );

        let layout = file
            .get("/slideMaster[1]/slideLayout[1]")
            .expect("nested slide-layout path");
        assert_eq!(layout.kind(), "slideLayout");
        assert_eq!(property(&layout, "name"), Some("Title Slide".to_owned()));
        assert_eq!(property(&layout, "type"), Some("title".to_owned()));
        let master_children = master
            .children()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("master children");
        assert_eq!(
            master_children.iter().map(Node::kind).collect::<Vec<_>>(),
            ["slideLayout"]
        );

        // The root layout alias is a presentation-order view of the same
        // master-owned layout. Canonical camel-case and lowercase aliases
        // must resolve identically.
        let root_layout = file.get("/slideLayout[1]").expect("root slide-layout path");
        let alias_layout = file
            .get("/slidelayout[1]")
            .expect("lowercase slide-layout alias");
        assert_eq!(root_layout.positional_path(), layout.positional_path());
        assert_eq!(alias_layout.positional_path(), layout.positional_path());

        let modern = file
            .get("/slide[1]/modernComment[1]")
            .expect("canonical modern-comment path");
        let modern_alias = file
            .get("/slide[1]/moderncomment[1]")
            .expect("lowercase modern-comment alias");
        assert_eq!(modern.kind(), "modernComment");
        assert_eq!(modern_alias.positional_path(), modern.positional_path());
        assert_eq!(text(&modern), Some("Please review the forecast".to_owned()));
        assert_eq!(property(&modern, "author"), Some("Ada Lovelace".to_owned()));
        assert_eq!(
            property(&modern, "created"),
            Some("2026-10-01T09:30:00.000".to_owned())
        );
        assert_eq!(property(&modern, "status"), Some("resolved".to_owned()));
        assert_eq!(property(&modern, "title"), Some("Forecast".to_owned()));
        let modern_json = view_json(&modern, 0);
        assert!(
            !modern_json.contains("6C7C3B3A"),
            "modern native IDs must stay behind raw access: {modern_json}"
        );

        let reply = file
            .get("/slide[1]/modernComment[1]/reply[1]")
            .expect("modern reply path");
        assert_eq!(reply.kind(), "reply");
        assert_eq!(text(&reply), Some("Updated the model".to_owned()));
        assert_eq!(property(&reply, "author"), Some("Grace Hopper".to_owned()));
        assert_eq!(
            property(&reply, "created"),
            Some("2026-10-01T10:00:00.000".to_owned())
        );
        assert_eq!(property(&reply, "status"), Some("active".to_owned()));

        let reply_children = modern
            .children()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("modern comment replies");
        assert_eq!(reply_children.len(), 1);
        assert_eq!(reply_children[0].positional_path(), reply.positional_path());
        assert_eq!(
            query_paths(&file, "modernComment"),
            ["/slide[1]/moderncomment[1]", "/slide[1]/moderncomment[2]",]
        );
        assert_eq!(
            query_paths(&file, "reply"),
            ["/slide[1]/moderncomment[1]/reply[1]"]
        );

        snapshots.push(vec![
            (
                master.kind(),
                master.positional_path().to_string(),
                view_json(&master, 1),
            ),
            (
                layout.kind(),
                layout.positional_path().to_string(),
                view_json(&layout, 1),
            ),
            (
                modern.kind(),
                modern.positional_path().to_string(),
                view_json(&modern, 1),
            ),
            (
                reply.kind(),
                reply.positional_path().to_string(),
                view_json(&reply, 1),
            ),
        ]);
    }

    assert_eq!(snapshots[0], snapshots[1]);
}

#[test]
fn eager_set_still_admits_the_native_graph_for_empty_and_equal_properties() {
    let bytes = presentation::bytes_with_malformed_tail();
    let native = litchi_pptx::Package::from_bytes(&bytes).unwrap();
    let mut edit = native
        .edit()
        .expect("deferred native edit construction must remain infallible");
    assert!(
        edit.slides().is_err(),
        "the first native semantic use must reject the malformed tail"
    );
    let eager = files(bytes)[1].clone();
    let before = eager.to_bytes().unwrap();
    let path = "/slide[1]/shape[1]";
    let text = eager
        .get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned();
    assert!(eager.set(path, litchi_ops::Props::new()).is_err());
    assert!(
        eager
            .set(path, litchi_ops::props! { "text" => text.as_str() })
            .is_err()
    );
    assert_eq!(eager.to_bytes().unwrap(), before);
}

#[test]
fn query_limit_does_not_decode_a_malformed_tail() {
    let file = files(presentation::bytes_with_malformed_tail())[0].clone();
    let first = file
        .query("shape")
        .expect("compile shape query")
        .first()
        .expect("first query result")
        .expect("first shape");
    assert_eq!(first.positional_path().to_string(), "/slide[1]/shape[1]");

    let mut matches = file.query("shape").expect("compile full query");
    let mut successful = 0;
    let mut failed = false;
    for item in matches.by_ref() {
        match item {
            Ok(_) => successful += 1,
            Err(_) => {
                failed = true;
                break;
            },
        }
    }
    assert!(successful > 0);
    assert!(
        failed,
        "the malformed tail must be observed without a limit"
    );
    assert!(matches.next().is_none(), "a failed query is fused");
    assert!(matches.next().is_none(), "a fused query stays exhausted");
}

#[cfg(any(unix, windows))]
#[test]
fn source_file_reads_refuse_after_the_source_revision_changes() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("revision.pptx");
    fs::write(&path, presentation::bytes()).expect("write PPTX");
    let file = litchi_ops::open(Path::new(&path)).expect("source-backed path open");
    assert!(file.get("/slide[1]").is_ok());

    fs::write(&path, b"changed source").expect("change source");
    let error = file
        .get("/slide[1]")
        .expect_err("changed source must refuse reads");
    assert_eq!(error.code(), Code::IoError);
}

#[cfg(any(unix, windows))]
#[test]
fn view_to_refuses_a_source_revision_changed_by_its_sink() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("view-to-revision.pptx");
    fs::write(&path, presentation::bytes()).expect("write PPTX");
    let file = litchi_ops::open(Path::new(&path)).expect("source-backed path open");
    let mut sink = MutatingSink {
        path,
        output: String::new(),
        mutated: false,
    };

    let error = file
        .view_to(Mode::Text, &mut sink)
        .expect_err("a sink-side source mutation must fail the final fence");
    assert_eq!(error.code(), Code::IoError);
    assert!(sink.mutated);
}

#[cfg(any(unix, windows))]
#[test]
fn cached_nodes_and_children_iterators_fence_source_revision_changes() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("cached-node-revision.pptx");
    fs::write(&path, presentation::bytes()).expect("write PPTX");
    let file = litchi_ops::open(Path::new(&path)).expect("source-backed path open");
    let root = file.root().expect("root");
    let node = file.get("/slide[1]/shape[3]").expect("cached text shape");

    // Populate both node-side caches before changing the underlying path.
    let cached_view = node.view();
    assert!(cached_view.text().expect("initial text").is_some());
    assert!(
        cached_view
            .format()
            .expect("initial format")
            .next()
            .is_some()
    );
    let mut children = root.children();

    fs::write(&path, b"changed source").expect("change source");

    let text_error = cached_view
        .text()
        .expect_err("cached node text must recheck source freshness");
    assert_eq!(text_error.code(), Code::IoError);
    let format_error = cached_view
        .format()
        .expect_err("cached node format must recheck source freshness");
    assert_eq!(format_error.code(), Code::IoError);
    let child_error = children
        .next()
        .expect("stale children iterator must report an error")
        .expect_err("stale children iterator must refuse the cached source");
    assert_eq!(child_error.code(), Code::IoError);
    assert!(
        children.next().is_none(),
        "failed children iterator is fused"
    );
}
