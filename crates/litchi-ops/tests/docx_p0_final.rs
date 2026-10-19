//! The last DOCX P0 writes of record 0345 through the unified API, after
//! OfficeCLI's word schemas: the hyperlink screen tip, cell insertion and
//! cell text on `add`, the table `col` element, and comment replies,
//! resolution and anchors (`commentsExtended`). Each write is read back
//! after a byte round trip on both owners.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _};

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

fn ok<T>(result: litchi_ops::Result<T>) -> T {
    result.unwrap_or_else(|error| panic!("{error} ({:?})", error.property()))
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn part(file: &File, name: &str) -> Option<String> {
    let bytes = file.to_bytes().expect("publish");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name(name).ok()?;
    let mut contents = String::new();
    member.read_to_string(&mut contents).expect("read part");
    Some(contents)
}

fn document_xml(file: &File) -> String {
    part(file, "word/document.xml").expect("main document")
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

fn with_paragraph(text: &str) -> File {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    ok(file.add("/body", NewNode::new("paragraph").prop("text", text))).file
}

#[allow(dead_code)]
fn with_table(rows: &str, cols: &str) -> File {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    ok(file.add(
        "/body",
        NewNode::new("table").prop("rows", rows).prop("cols", cols),
    ))
    .file
}

// ---------------------------------------------------------------------------
// Hyperlink tooltip
// ---------------------------------------------------------------------------

/// `hyperlink.json`: `tooltip` seeds the new link's `w:tooltip`; Litchi
/// also sets and clears it.
#[test]
fn hyperlink_tooltip_is_added_set_and_cleared() {
    for eager in [true, false] {
        let file = reopen(&with_paragraph("See "), eager);
        let added = ok(file.add(
            "/body/p[1]",
            NewNode::new("hyperlink")
                .prop("url", "https://example.test/")
                .prop("text", "the report")
                .prop("tooltip", "Opens the report"),
        ));
        assert_eq!(added.path.to_string(), "/body/p[1]/hyperlink[1]");
        let file = reopen(&added.file, eager);
        assert_eq!(
            prop(&file, "/body/p[1]/hyperlink[1]", "tooltip").as_deref(),
            Some("Opens the report")
        );
        assert!(document_xml(&file).contains(r#"w:tooltip="Opens the report""#));

        let file = reopen(
            &ok(file.set(
                "/body/p[1]/hyperlink[1]",
                props! { "tooltip" => "Q3 & Q4 <draft>" },
            )),
            eager,
        );
        assert_eq!(
            prop(&file, "/body/p[1]/hyperlink[1]", "tooltip").as_deref(),
            Some("Q3 & Q4 <draft>")
        );
        assert!(document_xml(&file).contains(r#"w:tooltip="Q3 &amp; Q4 &lt;draft&gt;""#));

        // An identical value publishes the input bytes.
        let before = file.to_bytes().unwrap();
        let same = ok(file.set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => "Q3 & Q4 <draft>" },
        ));
        assert_eq!(same.to_bytes().unwrap(), before);

        let cleared = ok(file.set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => litchi_ops::PropValue::Null },
        ));
        let cleared = reopen(&cleared, eager);
        assert_eq!(prop(&cleared, "/body/p[1]/hyperlink[1]", "tooltip"), None);
        assert!(!document_xml(&cleared).contains("tooltip"));
    }
}

/// A hyperlink without a tooltip is written exactly as before.
#[test]
fn hyperlink_without_tooltip_has_no_tooltip_attribute() {
    let file = with_paragraph("See ");
    let added = ok(file.add(
        "/body/p[1]",
        NewNode::new("hyperlink")
            .prop("anchor", "Intro")
            .prop("text", "intro"),
    ));
    assert!(!document_xml(&added.file).contains("tooltip"));
    let error = added
        .file
        .set(
            "/body/p[1]/hyperlink[1]",
            props! { "tooltip" => "bad\u{1}tip" },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
}
