//! DOCX find/replace through `set` with `find=` (record 0273): cross-run
//! matches collapse into the run where they start, format properties apply
//! to exactly the replaced (or matched) text, the scope is the target's
//! subtree, and zero matches are an exact no-op. Every case runs on both
//! owners and survives a save and reopen.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Props, SourceMode, props};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

fn package(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/" TargetMode="External"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", document_relationships),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(contents.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Three paragraphs (the first split over formatted runs), a table and a
/// content control.
fn fixture() -> Vec<u8> {
    package(concat!(
        r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hel</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">lo wor</w:t></w:r><w:r><w:t>ld</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t xml:space="preserve">draft </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>draft link</w:t></w:r></w:hyperlink><w:r><w:tab/><w:t>Draft</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell draft</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        r#"<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>control draft</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    ))
}

fn open(eager: bool, bytes: Vec<u8>) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .unwrap()
    } else {
        litchi_ops::open_bytes(bytes).unwrap()
    }
}

fn document_xml(file: &File) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name("word/document.xml").unwrap();
    let mut contents = String::new();
    member.read_to_string(&mut contents).unwrap();
    contents
}

fn find(file: &File, path: &str, props: Props<'_>) -> (File, Option<usize>) {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    let commit = file
        .apply(&item)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    (commit.file().clone(), commit.value().matched())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

#[test]
fn a_cross_run_match_takes_the_formatting_of_its_first_run() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let (after, matched) = find(
            &file,
            "/body/p[1]",
            Props::new().with("find", "llo w").with("replace", "LLO-W"),
        );
        assert_eq!(matched, Some(1));
        let xml = document_xml(&after);
        assert!(
            xml.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>HeLLO-W</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">or</w:t></w:r><w:r><w:t>ld</w:t></w:r>"#),
            "{xml}"
        );
        assert_eq!(text(&after, "/body/p[1]"), "HeLLO-World");
        // The rest of the document is untouched, byte for byte.
        let before = document_xml(&file);
        let tail = before
            .split("</w:p>")
            .skip(1)
            .collect::<Vec<_>>()
            .join("</w:p>");
        assert!(xml.ends_with(&tail));
        // Save and reopen.
        let reopened = open(eager, after.to_bytes().unwrap());
        assert_eq!(text(&reopened, "/body/p[1]"), "HeLLO-World");
    }
}

#[test]
fn the_root_searches_every_paragraph_cell_control_and_hyperlink() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let (after, matched) = find(
            &file,
            "/",
            Props::new().with("find", "draft").with("replace", "final"),
        );
        assert_eq!(matched, Some(4));
        let xml = document_xml(&after);
        for expected in [
            r#"<w:t xml:space="preserve">final </w:t>"#,
            "<w:t>final link</w:t>",
            "<w:t>Draft</w:t>",
            "<w:t>cell final</w:t>",
            "<w:t>control final</w:t>",
        ] {
            assert!(xml.contains(expected), "{expected} in {xml}");
        }
        // Case-insensitive matching reaches the capitalized word too.
        let (_, matched) = find(
            &file,
            "/body",
            Props::new()
                .with("find", "DRAFT")
                .with("replace", "x")
                .with("matchCase", false),
        );
        assert_eq!(matched, Some(5));
        // Whole words skip the hyperlink's "draft link"? No: it is a word.
        let (_, matched) = find(
            &file,
            "/",
            Props::new()
                .with("find", "raft")
                .with("replace", "x")
                .with("wholeWord", true),
        );
        assert_eq!(matched, Some(0));
    }
}

#[test]
fn zero_matches_are_an_exact_no_op() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let before = file.to_bytes().unwrap();
        let (after, matched) = find(
            &file,
            "/",
            Props::new().with("find", "absent").with("replace", "x"),
        );
        assert_eq!(matched, Some(0));
        assert_eq!(after.to_bytes().unwrap(), before);
        let commit = file
            .edit(|tx| tx.set("/", props! { "find" => "absent", "replace" => "x" }))
            .unwrap();
        assert!(!commit.diagnostics().changed());
        // `File::set` takes the same path and drops the count.
        let set = file
            .set("/", props! { "find" => "absent", "replace" => "x" })
            .unwrap();
        assert_eq!(set.to_bytes().unwrap(), before);
    }
}

#[test]
fn format_properties_apply_to_the_replacement_only() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let (after, matched) = find(
            &file,
            "/body/p[1]",
            Props::new()
                .with("find", "lo wor")
                .with("replace", "LO WOR")
                .with("underline", "single"),
        );
        assert_eq!(matched, Some(1));
        let xml = document_xml(&after);
        assert!(
            xml.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Hel</w:t></w:r><w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">LO WOR</w:t></w:r><w:r><w:t>ld</w:t></w:r>"#),
            "{xml}"
        );
        // Without `replace`, the matched text is formatted in place, split
        // across the runs it covers.
        let (after, matched) = find(
            &file,
            "/body/p[1]",
            Props::new().with("find", "llo w").with("color", "FF0000"),
        );
        assert_eq!(matched, Some(1));
        let xml = document_xml(&after);
        assert!(
            xml.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>He</w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>l</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">lo w</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">or</w:t></w:r>"#),
            "{xml}"
        );
        assert_eq!(text(&after, "/body/p[1]"), "Hello world");
    }
}

#[test]
fn a_run_or_hyperlink_scope_matches_only_its_own_text() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        // OfficeCLI's run scope: a match must lie inside the run.
        let (after, matched) = find(
            &file,
            "/body/p[2]/r[1]",
            Props::new().with("find", "draft").with("replace", "X"),
        );
        assert_eq!(matched, Some(1));
        let xml = document_xml(&after);
        assert!(xml.contains(r#"<w:t xml:space="preserve">X </w:t>"#));
        assert!(xml.contains("<w:t>draft link</w:t>"));
        let (after, matched) = find(
            &file,
            "/body/p[2]/hyperlink[1]",
            Props::new().with("find", "draft").with("replace", "X"),
        );
        assert_eq!(matched, Some(1));
        let xml = document_xml(&after);
        assert!(
            xml.contains("<w:t>X link</w:t>")
                && xml.contains(r#"<w:t xml:space="preserve">draft </w:t>"#)
        );
    }
}

#[test]
fn tabs_and_hyperlink_edges_separate_text() {
    let file = open(true, fixture());
    let (_, matched) = find(
        &file,
        "/",
        Props::new().with("find", "linkDraft").with("replace", "x"),
    );
    assert_eq!(matched, Some(0));
    let (_, matched) = find(
        &file,
        "/",
        Props::new()
            .with("find", "draft draft")
            .with("replace", "x"),
    );
    assert_eq!(matched, Some(0));
}

#[test]
fn finds_join_atomic_batches_and_invert() {
    let file = open(false, fixture());
    let items = litchi_ops::batch::from_bytes(
        br#"[
          {"command":"set","path":"/","props":{"find":"draft","replace":"final"}},
          {"command":"set","path":"/body/p[1]/r[1]","props":{"italic":true}},
          {"command":"set","path":"/body/p[2]","props":{"find":"absent","replace":"x"}}
        ]"#,
        &litchi_ops::batch::SourceLimits::default(),
    )
    .unwrap();
    let commit = file.batch(items, BatchMode::Atomic).unwrap();
    let counts: Vec<Option<usize>> = commit
        .value()
        .iter()
        .map(|result| match result {
            litchi_ops::ItemResult::Applied(applied) => applied.matched(),
            _ => panic!("failed item"),
        })
        .collect();
    assert_eq!(counts, [Some(4), None, Some(0)]);
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn refusals_name_the_property() {
    let file = open(true, fixture());
    let cases = [
        (
            Props::new().with("find", "x"),
            Code::MissingProperty,
            "replace",
        ),
        (
            Props::new()
                .with("find", "x")
                .with("replace", "y")
                .with("regex", true),
            Code::UnsupportedCapability,
            "regex",
        ),
        (
            Props::new().with("find", "").with("replace", "y"),
            Code::InvalidValue,
            "find",
        ),
        (
            Props::new().with("find", "x").with("sparkle", true),
            Code::UnsupportedProperty,
            "sparkle",
        ),
        (
            Props::new().with("find", "x").with("text", "y"),
            Code::InvalidArgument,
            "text",
        ),
    ];
    for (props, code, property) in cases {
        let error = file.set("/", props).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
        assert_eq!(error.property(), Some(property), "{error}");
    }
    let error = file
        .set("/body/p[1]", props! { "find" => "x", "replace" => "y\n" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}
