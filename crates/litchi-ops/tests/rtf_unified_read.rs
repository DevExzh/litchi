#![cfg(feature = "rtf")]
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure"
)]

use litchi_ops::capabilities::Verb;
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, open_bytes};
use std::borrow::Cow;

fn fixture() -> File {
    open_bytes(br"{\rtf1\ansi{\header Not body}First\line second\par\par A{\b B}C\par last}")
        .unwrap()
}
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}
#[test]
fn paths_use_native_paragraph_boundaries_and_preserve_empty_paragraphs() {
    let file = fixture();
    assert_eq!(file.root().unwrap().kind(), "document");
    if let File::Rtf(handle) = &file {
        assert_eq!(handle.paragraph_count().unwrap(), 4);
    }
    assert_eq!(text(&file, "/body/p[1]"), "First\nsecond");
    assert_eq!(text(&file, "/body/p[2]"), "");
    assert_eq!(text(&file, "/body/paragraph[3]"), "ABC");
    assert_eq!(text(&file, "/body/p[last()]"), "last");
    assert!(file.find("/body/p[5]").unwrap().is_none());
    assert_eq!(
        file.get("/body/p[0]").unwrap_err().code(),
        Code::InvalidPath
    );
    let node = file.get("/body/p[3]").unwrap();
    assert_eq!(node.path().to_string(), "/body/p[3]");
    assert_eq!(
        node.as_rtf().unwrap().paragraph_position().unwrap().get(),
        2
    );
}
#[test]
fn paragraph_nodes_borrow_contiguous_native_text_and_materialize_fragments() {
    let file = fixture();
    let simple = file.get("/body/p[4]").unwrap();
    assert!(matches!(
        simple.view().text().unwrap(),
        Some(Cow::Borrowed("last"))
    ));
    let fragmented = file.get("/body/p[3]").unwrap();
    assert_eq!(fragmented.view().text().unwrap().unwrap(), "ABC");
    assert!(!format!("{fragmented:?}").contains("ABC"));
}
#[test]
fn queries_and_children_follow_body_order_and_stop_at_limit() {
    let file = fixture();
    let paths = file
        .query("p")
        .unwrap()
        .map(|node| node.unwrap().path().to_string())
        .collect::<Vec<_>>();
    assert_eq!(
        paths,
        ["/body/p[1]", "/body/p[2]", "/body/p[3]", "/body/p[4]"]
    );
    assert_eq!(file.query("body > p").unwrap().limit(2).count(), 2);
    let body = file.get("/body").unwrap();
    assert_eq!(body.children().count(), 4);
    assert_eq!(file.query("p[text~=ABC]").unwrap().count(), 1);
    let first = file.query("p").unwrap().first().unwrap().unwrap();
    assert_eq!(first.view().text().unwrap().unwrap(), "First\nsecond");
}
#[test]
fn views_and_capabilities_agree_on_supported_read_surface() {
    let file = fixture();
    for verb in [Verb::Get, Verb::Find, Verb::Query] {
        assert!(file.capabilities().verb(verb).is_supported());
    }
    assert!(file.capabilities().verb(Verb::Set).is_supported());
    assert_eq!(
        file.view(Mode::Text).unwrap(),
        "First\nsecond\n\nABC\nlast\n"
    );
    let json = serde_json::to_value(file.view_json(Mode::Text).unwrap()).unwrap();
    assert_eq!(json["totalElements"], 4);
    assert_eq!(json["elements"][2]["text"], "ABC");
    let stats = serde_json::to_value(file.view_json(Mode::Stats).unwrap()).unwrap();
    assert_eq!(stats["paragraphs"], 4);
    assert_eq!(stats["emptyParagraphs"], 1);
    assert!(
        file.view(Mode::Annotated)
            .unwrap()
            .contains("[/body/p[3]] ABC")
    );
    assert_eq!(text(&file, "/body/p[1]/r[1]"), "First");
}

#[test]
fn native_substories_are_read_and_unprojected_concepts_stay_pending() {
    let file = fixture();
    // Record 0275 projects header/footer stories; the main-story `p` axis
    // still excludes their paragraphs.
    assert_eq!(text(&file, "/header[1]"), "Not body");
    assert_eq!(text(&file, "/header[1]/p[1]"), "Not body");
    assert_eq!(file.query("header").unwrap().count(), 1);
    assert_eq!(file.query("p").unwrap().count(), 4);
    // Record 0328 reads shapes as body children, not paragraph children;
    // equations stay a capability gap.
    assert_eq!(
        file.get("/body/p[1]/shape[1]").unwrap_err().code(),
        Code::UnsupportedPath
    );
    assert!(
        file.capabilities()
            .gaps()
            .any(|gap| gap.subject() == litchi_ops::capabilities::Subject::Element("equation"))
    );
}

#[test]
fn retained_corpus_paragraphs_match_native_views_exactly() {
    for name in [
        "typed_hyperlinks_and_references.rtf",
        "testNegativeUnicode.rtf",
        "test950Encoding.rtf",
        "chtoutline.rtf",
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/rtf")
            .join(name);
        let bytes = std::fs::read(path).unwrap();
        let native = litchi_rtf::Package::from_bytes(&bytes).unwrap();
        let file = open_bytes(bytes.as_slice()).unwrap();
        let actual = file
            .query("p")
            .unwrap()
            .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
            .collect::<Vec<_>>();
        let expected = native
            .body()
            .paragraphs()
            .map(|paragraph| paragraph.to_text())
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "{name}");
    }
}
