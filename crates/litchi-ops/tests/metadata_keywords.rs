//! The structured keyword list (record 0354): `Metadata::keywords` is a
//! `Vec<String>` in every format, the node view writes it as a JSON array,
//! a set takes a list or text split at `;` and `,`, and an unchanged list
//! keeps the source's own keyword text.
#![cfg(all(
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "doc",
    feature = "xls",
    feature = "ppt",
    feature = "odt",
    feature = "ods",
    feature = "odp",
))]

use litchi_core::FileFormat;
use litchi_ops::{BatchMode, Code, File, PropValue, create, open_bytes, props};

const FORMATS: [FileFormat; 10] = [
    FileFormat::Docx,
    FileFormat::Xlsx,
    FileFormat::Pptx,
    FileFormat::Xlsb,
    FileFormat::Doc,
    FileFormat::Xls,
    FileFormat::Ppt,
    FileFormat::Odt,
    FileFormat::Ods,
    FileFormat::Odp,
];

fn keywords(file: &File) -> Vec<String> {
    file.metadata()
        .expect("read metadata")
        .unwrap_or_default()
        .keywords
}

fn reopened(file: &File) -> File {
    open_bytes(file.to_bytes().expect("bytes")).expect("reopen")
}

fn list(items: &[&str]) -> PropValue<'static> {
    PropValue::from(items.to_vec())
}

#[test]
fn lists_and_separated_text_round_trip_in_every_format() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let listed = file
            .set(
                "/metadata",
                props! { "keywords" => list(&["q4", "review", "two words"]) },
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(
            keywords(&reopened(&listed)),
            ["q4", "review", "two words"],
            "{format:?}"
        );

        let texted = listed
            .set("/metadata", props! { "keywords" => "x; y,z ," })
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(keywords(&reopened(&texted)), ["x", "y", "z"], "{format:?}");

        let cleared = texted
            .set("/metadata", props! { "keywords" => PropValue::Null })
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert!(keywords(&reopened(&cleared)).is_empty(), "{format:?}");
    }
}

#[test]
fn the_root_node_view_writes_a_json_array() {
    for format in [FileFormat::Docx, FileFormat::Xlsx, FileFormat::Pptx] {
        let file = create(format)
            .expect("create")
            .set("/metadata", props! { "keywords" => "plan, review" })
            .expect("set");
        let json = serde_json::to_value(file.get("/").expect("root").view()).expect("view");
        assert_eq!(
            json["format"]["keywords"],
            serde_json::json!(["plan", "review"]),
            "{format:?}"
        );
    }
}

#[test]
fn query_comparison_tests_each_keyword() {
    use litchi_core::selector::CompareOp;
    use litchi_ops::value::{Scalar, Test, ValueType};

    let items = ["plan".to_owned(), "Review".to_owned()];
    let list = Some(Scalar::List(&items));
    let ty = Some(ValueType::List);
    assert!(Test::with_text(CompareOp::Eq, "review", ty).eval(list));
    assert!(!Test::with_text(CompareOp::Eq, "plan, review", ty).eval(list));
    assert!(Test::with_text(CompareOp::Contains, "evi", ty).eval(list));
    assert!(Test::with_text(CompareOp::Ne, "other", ty).eval(list));
    assert!(!Test::with_text(CompareOp::Ne, "plan", ty).eval(list));
    assert!(Test::exists().eval(list));
    assert!(!Test::exists().eval(Some(Scalar::List(&[]))));
}

#[test]
fn metadata_serializes_keywords_as_an_array() {
    let file = create(FileFormat::Docx)
        .expect("create")
        .set("/metadata", props! { "keywords" => list(&["a", "b"]) })
        .expect("set");
    let metadata = file.metadata().expect("read").expect("present");
    let json = serde_json::to_value(&metadata).expect("json");
    assert_eq!(json["keywords"], serde_json::json!(["a", "b"]));
    let none = serde_json::to_value(litchi_core::Metadata::default()).expect("json");
    assert!(none.get("keywords").is_none(), "an empty list is skipped");
    let read: litchi_core::Metadata =
        serde_json::from_value(serde_json::json!({ "title": "t" })).expect("read");
    assert!(read.keywords.is_empty(), "a missing list reads as empty");
}

#[test]
fn a_batch_takes_a_json_array() {
    let file = create(FileFormat::Xlsx).expect("create");
    let items = litchi_ops::batch::from_bytes(
        br#"[{"command":"set","path":"/metadata","props":{"keywords":["finance","2026"]}}]"#,
        &litchi_ops::batch::SourceLimits::default(),
    )
    .expect("decode");
    let commit = file.batch(items, BatchMode::Atomic).expect("batch");
    assert_eq!(keywords(&reopened(commit.file())), ["finance", "2026"]);
    assert!(
        litchi_ops::batch::from_bytes(
            br#"[{"command":"set","path":"/metadata","props":{"keywords":["a",1]}}]"#,
            &litchi_ops::batch::SourceLimits::default()
        )
        .is_err(),
        "a list holds only strings"
    );
}

#[test]
fn an_unchanged_list_keeps_the_source_keyword_text() {
    for format in [FileFormat::Docx, FileFormat::Xlsx, FileFormat::Doc] {
        // The list the source reads as, set again (as a list or as text with
        // other separators), writes nothing. The codec unit tests cover
        // sources spelled with other separators byte for byte.
        let source = reopened(
            &create(format)
                .expect("create")
                .set("/metadata", props! { "keywords" => "a, b" })
                .expect("seed"),
        );
        let source_bytes = source.to_bytes().expect("bytes");
        let same = source
            .set("/metadata", props! { "keywords" => list(&["a", "b"]) })
            .expect("set");
        assert_eq!(same.to_bytes().expect("bytes"), source_bytes, "{format:?}");
        let same = source
            .set("/metadata", props! { "keywords" => "a;b" })
            .expect("set");
        assert_eq!(same.to_bytes().expect("bytes"), source_bytes, "{format:?}");
    }
}

#[test]
fn string_formats_refuse_a_keyword_their_string_cannot_keep() {
    for format in [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt] {
        let error = create(format)
            .expect("create")
            .set("/metadata", props! { "keywords" => list(&["a, b"]) })
            .expect_err("refused");
        assert_eq!(error.code(), Code::InvalidValue, "{format:?}");
    }
    // OOXML keeps such a keyword whole as a `cp:value` child.
    let file = create(FileFormat::Docx)
        .expect("create")
        .set("/metadata", props! { "keywords" => list(&["a, b", "c"]) })
        .expect("set");
    assert_eq!(keywords(&reopened(&file)), ["a, b", "c"]);
    // ODF stores one element per keyword.
    let file = create(FileFormat::Odt)
        .expect("create")
        .set("/metadata", props! { "keywords" => list(&["a; b", "c"]) })
        .expect("set");
    assert_eq!(keywords(&reopened(&file)), ["a; b", "c"]);
}
