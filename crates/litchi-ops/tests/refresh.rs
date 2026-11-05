//! Field and table-of-contents refresh (record 0279): DOCX tables of
//! contents and computable fields, PPTX slide fields, and the exact no-op
//! on documents a word processor left current. Every DOCX case runs on both
//! owners.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::refresh::{Clock, Options, PageNumbers, Status};
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};

const DATA: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/"
);
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn open(eager: bool, bytes: Vec<u8>) -> File {
    let mode = if eager {
        SourceMode::Eager
    } else {
        SourceMode::Auto
    };
    litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(mode)).unwrap()
}

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(format!("{DATA}{path}")).unwrap()
}

fn member(file: &File, name: &str) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name(name).unwrap();
    let mut contents = String::new();
    member.read_to_string(&mut contents).unwrap();
    contents
}

/// A package with heading styles, a title and `body`.
fn package(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1000" w:right="1000" w:top="1000" w:bottom="1000"/></w:sectPr></w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Appendix"><w:name w:val="Appendix"/><w:basedOn w:val="Heading1"/></w:style></w:styles>"#
    );
    let core = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly report</dc:title><dc:creator>Ada</dc:creator></cp:coreProperties>"#;
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#;
    let document_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", document_relationships),
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", styles.as_str()),
        ("docProps/core.xml", core),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(contents.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn heading(style: &str, text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn complex(instruction: &str, result: &str) -> String {
    format!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
    )
}

/// OfficeCLI's `add --type toc` placeholder, then headings and fields.
fn report() -> Vec<u8> {
    let mut body = String::new();
    body.push_str(&complex(
        r#" TOC \o "1-2" \h "#,
        "Update field to see table of contents",
    ));
    body.push_str(&heading("Heading1", "Introduction"));
    body.push_str(r#"<w:p><w:bookmarkStart w:id="7" w:name="Scope"/><w:r><w:t>The scope</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>"#);
    body.push_str(&heading("Heading2", "Background"));
    body.push_str(&complex(r" SEQ Figure \* ARABIC ", "9"));
    body.push_str(&complex(r" SEQ Figure \* ROMAN ", "9"));
    body.push_str(&complex(r" REF Scope \h ", "stale"));
    body.push_str(&complex(r#" DATE \@ "yyyy-MM-dd" "#, "2000-01-01"));
    body.push_str(&complex(" TITLE ", "old"));
    body.push_str(&complex(" PAGE ", "4"));
    body.push_str(&complex(" REF Missing ", "x"));
    body.push_str(&heading("Appendix", "Appendix A"));
    body.push_str(&heading("Heading3", "Not a level"));
    package(&body)
}

#[test]
fn word_and_libreoffice_tables_of_contents_are_exact_no_ops() {
    for name in [
        "sw/qa/extras/ooxmlexport/data/multi-page-toc.docx",
        "sw/qa/extras/ooxmlexport/data/testTOCFlag_u.docx",
        "sw/qa/extras/ooxmlexport/data/custom-styles-TOC-comma.docx",
    ] {
        for eager in [true, false] {
            let file = open(eager, fixture(name));
            let commit = file.refresh(&Options::new()).unwrap();
            assert!(commit.patch().is_noop(), "{name}");
            assert!(!commit.diagnostics().changed(), "{name}");
            assert_eq!(commit.value().count(Status::Unchanged), 1, "{name}");
        }
    }
}

#[test]
fn a_changed_heading_rebuilds_the_table_and_keeps_page_numbers() {
    for eager in [true, false] {
        let file = open(
            eager,
            fixture("sw/qa/extras/ooxmlexport/data/multi-page-toc.docx"),
        );
        // Change the heading alone: a find over the body would also change
        // the table's entry, a field result it searches (record 0285).
        let file = file
            .set(
                "/body/p[2]",
                props! { "find" => "Chapter 1", "replace" => "Chapter One" },
            )
            .unwrap();
        let commit = file.refresh(&Options::new()).unwrap_or_else(|error| {
            let mut chain = String::new();
            let mut source = std::error::Error::source(&error);
            while let Some(inner) = source {
                chain.push_str(&format!(" / {inner}"));
                source = inner.source();
            }
            panic!("{error}{chain}")
        });
        assert_eq!(commit.value().refreshed(), 1);
        assert_eq!(commit.value().missing_page_numbers(), 0);
        assert_eq!(commit.value().bookmarks(), 0);
        let xml = member(commit.file(), "word/document.xml");
        assert!(xml.contains("Chapter One"), "{xml}");
        assert!(xml.contains(r#"PAGEREF _Toc400458847 \h"#), "{xml}");
        assert!(!xml.contains("dirty"), "{xml}");
        // The refreshed table is current; the saved file reopens current.
        let again = commit.file().refresh(&Options::new()).unwrap();
        assert!(again.patch().is_noop());
        let reopened = open(eager, commit.file().to_bytes().unwrap());
        assert!(reopened.refresh(&Options::new()).unwrap().patch().is_noop());
    }
}

#[test]
fn a_placeholder_table_and_fields_are_computed() {
    for eager in [true, false] {
        let file = open(eager, report());
        let commit = file.refresh(&Options::new()).unwrap();
        let summary = commit.value();
        let statuses: Vec<(&str, Status)> = summary
            .fields()
            .iter()
            .map(|field| (field.kind(), field.status()))
            .collect();
        assert_eq!(
            statuses,
            [
                ("TOC", Status::Refreshed),
                ("SEQ", Status::Refreshed),
                ("SEQ", Status::Refreshed),
                ("REF", Status::Refreshed),
                ("DATE", Status::NeedsClock),
                ("TITLE", Status::Refreshed),
                ("PAGE", Status::NeedsLayout),
                ("REF", Status::Unresolved),
            ]
        );
        assert_eq!(summary.fields()[0].location(), Some("/body/p[1]"));
        assert_eq!(summary.bookmarks(), 3);
        assert_eq!(summary.missing_page_numbers(), 3);
        let xml = member(commit.file(), "word/document.xml");
        for expected in [
            r#"<w:fldChar w:fldCharType="begin" w:dirty="true"/>"#,
            r#"<w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9906"/></w:tabs>"#,
            r#"<w:hyperlink w:anchor="_Toc100000001" w:history="1">"#,
            r#"<w:bookmarkStart w:id="8" w:name="_Toc100000001"/>"#,
            ">Introduction<",
            ">Background<",
            ">Appendix A<",
            ">II<",
            ">The scope<",
            ">Quarterly report<",
        ] {
            assert!(xml.contains(expected), "{expected}: {xml}");
        }
        assert!(!xml.contains("Not a level") || xml.matches("Not a level").count() == 1);
        assert!(!xml.contains("Update field"), "{xml}");
        // Every computed result is current now.
        let again = commit.file().refresh(&Options::new()).unwrap();
        assert!(again.patch().is_noop());
        assert_eq!(again.value().count(Status::Unchanged), 5);
        // The source handle is untouched.
        assert!(member(&file, "word/document.xml").contains("Update field"));
    }
}

#[test]
fn an_explicit_clock_fills_date_fields() {
    let file = open(true, report());
    let clock = Clock::parse("2026-10-14T09:30:00").unwrap();
    let commit = file.refresh(&Options::new().clock(clock)).unwrap();
    assert_eq!(commit.value().count(Status::NeedsClock), 0);
    assert!(member(commit.file(), "word/document.xml").contains(">2026-10-14<"));
}

#[test]
fn unknown_page_numbers_can_be_refused_or_left_clean() {
    let file = open(true, report());
    let error = file
        .refresh(&Options::new().page_numbers(PageNumbers::Require))
        .unwrap_err();
    assert_eq!(error.code(), Code::RefreshFailed);
    let commit = file.refresh(&Options::new().mark_dirty(false)).unwrap();
    assert!(!member(commit.file(), "word/document.xml").contains("dirty"));
}

#[test]
fn a_refresh_composes_in_a_transaction() {
    let file = open(true, report());
    let commit = file
        .edit(|edit| {
            let summary = edit.refresh(&Options::new())?;
            assert_eq!(summary.refreshed(), 5);
            // The table now holds three entries and its end paragraph.
            edit.set(
                "/body/p[6]",
                props! { "find" => "Introduction", "replace" => "Overview" },
            )?;
            let summary = edit.refresh(&Options::new())?;
            assert_eq!(summary.refreshed(), 1);
            Ok(())
        })
        .unwrap();
    let xml = member(commit.file(), "word/document.xml");
    assert!(xml.contains(">Overview<") && !xml.contains(">Introduction<"));
    assert_eq!(xml.matches(">Overview<").count(), 2);
}

#[cfg(feature = "pptx")]
#[test]
fn slide_fields_take_their_number_and_the_clock() {
    let bytes = fixture("oox/qa/unit/data/testTdf132557_footerCustomShapes.pptx");
    for eager in [true, false] {
        let file = open(eager, bytes.clone());
        let commit = file.refresh(&Options::new()).unwrap();
        assert!(commit.patch().is_noop());
        let statuses: Vec<(&str, Status)> = commit
            .value()
            .fields()
            .iter()
            .map(|field| (field.kind(), field.status()))
            .collect();
        assert!(
            statuses.contains(&("SLIDENUM", Status::Unchanged)),
            "{statuses:?}"
        );
        assert!(
            statuses.contains(&("DATETIME1", Status::NeedsClock)),
            "{statuses:?}"
        );
        let clock = Clock::parse("2026-10-14").unwrap();
        let commit = file.refresh(&Options::new().clock(clock)).unwrap();
        assert_eq!(commit.value().refreshed(), 1);
        let slide = member(commit.file(), "ppt/slides/slide1.xml");
        assert!(slide.contains("<a:t>10/14/2026</a:t>"), "{slide}");
        assert!(
            commit
                .file()
                .refresh(&Options::new().clock(clock))
                .unwrap()
                .patch()
                .is_noop()
        );
    }
}

#[test]
fn other_formats_refuse() {
    #[cfg(feature = "rtf")]
    {
        let file = litchi_ops::open_bytes(b"{\\rtf1 hello}".to_vec()).unwrap();
        assert_eq!(
            file.refresh(&Options::new()).unwrap_err().code(),
            Code::UnsupportedCapability
        );
    }
    let file = open(true, report());
    assert!(
        file.capabilities()
            .verb(litchi_ops::capabilities::Verb::Refresh)
            .is_supported()
    );
}
