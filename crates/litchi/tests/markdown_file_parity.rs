//! `File` Markdown is byte-identical to the facade types' Markdown (M0a step
//! S7).
//!
//! The Markdown adapter moved into `litchi-ops`, where `ToMarkdown` is
//! implemented for `File`. The first tests replay the fixtures and goldens of
//! `markdown_docx_lists.rs`, `markdown_docx_semantics.rs` and
//! `markdown_real_fixtures.rs` through `litchi::open` and
//! `litchi::open_bytes`. The corpus test converts every DOCX, DOC, RTF, ODT,
//! PPTX, PPT and ODP fixture under `test-data` with both the facade types and
//! `File`, under several option sets and on the path and byte routes, and
//! requires the same output, or a refusal from both.

#![cfg(all(
    feature = "markdown",
    feature = "docx",
    feature = "doc",
    feature = "rtf",
    feature = "odt",
    feature = "pptx",
    feature = "ppt",
    feature = "odp"
))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::collections::BTreeMap;
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};

use litchi::markdown::{MarkdownOptions, TableStyle, ToMarkdown};
use litchi::{Code, Document};
use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// A DOCX package with the given main document, numbering part and extra
/// related parts, built the way the facade's Markdown tests build theirs.
struct DocxParts<'a> {
    content_types: &'a str,
    document_rels: &'a str,
    document: &'a str,
    parts: &'a [(&'a str, &'a [u8])],
}

impl DocxParts<'_> {
    fn build(&self) -> Vec<u8> {
        const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        for (name, contents) in [
            ("[Content_Types].xml", self.content_types.as_bytes()),
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("word/document.xml", self.document.as_bytes()),
            (
                "word/_rels/document.xml.rels",
                self.document_rels.as_bytes(),
            ),
        ]
        .into_iter()
        .chain(self.parts.iter().copied())
        {
            archive.start_file(name, options).unwrap();
            archive.write_all(contents).unwrap();
        }
        archive.finish().unwrap().into_inner()
    }
}

/// Both routes of the facade type and of `File` for one in-memory DOCX:
/// `(route, facade Markdown, File Markdown)`.
fn both_routes(
    bytes: &[u8],
    options: &MarkdownOptions,
) -> Vec<(
    &'static str,
    litchi::common::Result<String>,
    litchi::Result<String>,
)> {
    let temporary = tempfile::Builder::new().suffix(".docx").tempfile().unwrap();
    std::fs::write(temporary.path(), bytes).unwrap();
    vec![
        (
            "bytes",
            Document::from_bytes(bytes.to_vec())
                .and_then(|document| document.to_markdown_with_options(options)),
            litchi::open_bytes(bytes.to_vec())
                .and_then(|file| file.to_markdown_with_options(options)),
        ),
        (
            "path",
            Document::open(temporary.path())
                .and_then(|document| document.to_markdown_with_options(options)),
            litchi::open(temporary.path()).and_then(|file| file.to_markdown_with_options(options)),
        ),
    ]
}

#[track_caller]
fn assert_golden(bytes: &[u8], expected: &str) {
    for (route, old, new) in both_routes(bytes, &MarkdownOptions::new()) {
        assert_eq!(old.unwrap(), expected, "facade, {route}");
        assert_eq!(new.unwrap(), expected, "File, {route}");
    }
}

// ------------------------------------------- markdown_docx_lists.rs fixtures --

const LIST_CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#;
const LIST_DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

fn list_docx(document: &str, numbering: &str) -> Vec<u8> {
    DocxParts {
        content_types: LIST_CONTENT_TYPES,
        document_rels: LIST_DOCUMENT_RELS,
        document,
        parts: &[
            ("word/numbering.xml", numbering.as_bytes()),
            ("word/styles.xml", STYLES.as_bytes()),
        ],
    }
    .build()
}

#[test]
fn resolved_numbering_semantics_match() {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>bullet &amp; *</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="20"/></w:numPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="20"/></w:numPr></w:pPr><w:r><w:t>nested</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="20"/></w:numPr></w:pPr><w:r><w:t>nested two</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="21"/></w:numPr></w:pPr><w:r><w:t>restart</w:t></w:r></w:p></w:body></w:document>"#;
    let numbering = r#"<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum><w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="20"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="21"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num></w:numbering>"#;
    assert_golden(
        &list_docx(document, numbering),
        "- bullet \\& \\*\n\n1. one\n\n  1. nested\n\n  2. nested two\n\n3. restart\n\n",
    );
}

#[test]
fn a_missing_numbering_definition_is_refused_by_both() {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="999"/></w:numPr></w:pPr><w:r><w:t>missing</w:t></w:r></w:p></w:body></w:document>"#;
    let numbering = r#"<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    for (route, old, new) in both_routes(&list_docx(document, numbering), &MarkdownOptions::new()) {
        assert!(old.is_err(), "facade, {route}");
        assert!(new.is_err(), "File, {route}");
    }
}

// --------------------------------------- markdown_docx_semantics.rs fixtures --

const SEMANTIC_CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>"#;
const SEMANTIC_DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/a%20b" TargetMode="External"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/><Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>"#;
const FOOTNOTES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note &amp; *body*</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
const ENDNOTES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>end note</w:t></w:r></w:p></w:endnote></w:endnotes>"#;
const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";

fn semantic_docx(document: &str) -> Vec<u8> {
    DocxParts {
        content_types: SEMANTIC_CONTENT_TYPES,
        document_rels: SEMANTIC_DOCUMENT_RELS,
        document,
        parts: &[
            ("word/styles.xml", STYLES.as_bytes()),
            ("word/footnotes.xml", FOOTNOTES.as_bytes()),
            ("word/endnotes.xml", ENDNOTES.as_bytes()),
            ("word/media/p.png", PNG),
            ("word/embeddings/oleObject1.bin", b"ole payload"),
        ],
    }
    .build()
}

#[test]
fn package_semantics_match_and_a_changed_source_is_refused() {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rIdLink" w:tooltip="tip &quot;q&quot;"><w:r><w:rPr><w:b/></w:rPr><w:t>site [x]</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="Picture &quot;one&quot;" descr="alt [image]"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:p><w:r><w:t>With note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> and end</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p></w:body></w:document>"#;
    let bytes = semantic_docx(document);
    assert_golden(
        &bytes,
        "Before [**site \\[x\\]**](<https://example.test/a%20b> \"tip \\\"q\\\"\") after\n\n![alt \\[image\\]](<data:image/png;base64,iVBORw0KGgo=> \"Picture &quot;one&quot;\")\n\nWith note[^fn-1] and end[^en-2]\n\n[^fn-1]: note \\& \\*body\\*\n[^en-2]: end note\n",
    );

    let temporary = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(temporary.path(), &bytes).unwrap();
    let file = litchi::open(temporary.path()).unwrap();
    file.to_markdown().unwrap();
    let mut source = std::fs::OpenOptions::new()
        .append(true)
        .open(temporary.path())
        .unwrap();
    source.write_all(b"source mutation").unwrap();
    let error = file.to_markdown().expect_err("a changed source is refused");
    assert_eq!(error.code(), Code::IoError, "{error:?}");
}

#[test]
fn the_parallel_hint_keeps_the_serial_output() {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>"#;
    let file = litchi::open_bytes(semantic_docx(document)).unwrap();
    let serial = file
        .to_markdown_with_options(&MarkdownOptions::new().with_parallel(false))
        .unwrap();
    let parallel = file
        .to_markdown_with_options(&MarkdownOptions::new().with_parallel(true))
        .unwrap();
    assert_eq!(parallel, serial);
    assert_eq!(serial, "First\n\nSecond\n\n");
}

#[test]
fn tables_and_markup_compatibility_match_on_both_routes() {
    for document in [
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>"#,
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#,
    ] {
        let bytes = semantic_docx(document);
        let routes = both_routes(&bytes, &MarkdownOptions::new());
        let expected = routes[0].1.as_ref().unwrap().clone();
        for (route, old, new) in routes {
            assert_eq!(old.unwrap(), expected, "facade, {route}");
            assert_eq!(new.unwrap(), expected, "File, {route}");
        }
    }
}

#[test]
fn unplaceable_semantics_are_refused_by_both() {
    for (document, needle) in [
        (
            r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rIdLink" w:tgtFrame="_blank"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#,
            "w:tgtFrame",
        ),
        (
            r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:t>before</w:t><w:drawing><wp:inline><wp:docPr id="1" name="p" descr="a"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#,
            "interleaved DOCX text and embedded semantics",
        ),
    ] {
        for (route, old, new) in both_routes(&semantic_docx(document), &MarkdownOptions::new()) {
            assert!(
                matches!(&old, Err(litchi::common::Error::Unsupported(message)) if message.contains(needle)),
                "facade, {route}: {old:?}"
            );
            let error = new.expect_err("File refuses it too");
            assert_eq!(error.code(), Code::UnsupportedElement, "File, {route}");
            assert!(error.message().contains(needle), "File, {route}: {error:?}");
            assert_eq!(error.format(), Some(litchi::FileFormat::Docx));
        }
    }
}

#[test]
fn embedded_objects_floating_drawings_and_bookmarks_match() {
    for (document, expected) in [
        (
            r##"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Solve </w:t></w:r><w:r><w:object><v:shape id="_x0000_i1025" type="#_x0000_t75" alt="equation" style="width:16.7pt;height:14pt" o:ole=""><v:imagedata r:id="rIdImage" o:title="equation"/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1468075725" r:id="rIdOle"/></w:object></w:r><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:p></w:body></w:document>"##,
            "Solve ![equation](<data:image/png;base64,iVBORw0KGgo=> \"equation\") now\n\n",
        ),
        (
            r##"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Value: </w:t></w:r><w:r><w:object><v:shape id="_x0000_i1025" type="#_x0000_t75" alt="equation preview" style="width:16.7pt;height:14pt" o:ole=""><v:imagedata r:id="rIdMissing" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1468075725" r:id="rIdOle"/></w:object></w:r></w:p></w:body></w:document>"##,
            "Value: ![equation preview]\n\n",
        ),
        (
            r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:positionH relativeFrom="page"><wp:posOffset>11430000</wp:posOffset></wp:positionH><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="floater"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:body></w:document>"#,
            "![](<data:image/png;base64,iVBORw0KGgo=> \"floater\")\n\n",
        ),
        (
            r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>kept</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p></w:body></w:document>"#,
            "kept\n\n",
        ),
    ] {
        assert_golden(&semantic_docx(document), expected);
    }
}

// ----------------------------------------- markdown_real_fixtures.rs fixtures --

#[test]
fn real_docx_plain_paragraphs_match() {
    for (name, expected) in [
        ("documentProperties.docx", "Hello World\\!\n\n"),
        ("documentProtection_no_protection.docx", "Non protetto.\n\n"),
        (
            "documentProtection_comments_no_password.docx",
            "Comments senza password.\n\n",
        ),
    ] {
        let path = fixture("ooxml/docx").join(name);
        assert_eq!(
            Document::open(&path).unwrap().to_markdown().unwrap(),
            expected,
            "facade {name}"
        );
        assert_eq!(
            litchi::open(&path).unwrap().to_markdown().unwrap(),
            expected,
            "File {name}"
        );
    }
}

// ------------------------------------------------------------------ corpus --

/// The family of the facade type that opens a fixture, by extension.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Family {
    Document,
    Presentation,
}

fn family(path: &Path) -> Option<Family> {
    let extension = path.extension()?.to_str()?.to_ascii_lowercase();
    match extension.as_str() {
        "docx" | "dotx" | "docm" | "dotm" | "doc" | "rtf" | "odt" => Some(Family::Document),
        "pptx" | "pptm" | "potx" | "ppsx" | "ppt" | "odp" => Some(Family::Presentation),
        _ => None,
    }
}

fn corpus(root: &Path, files: &mut Vec<(PathBuf, Family)>) {
    let mut entries: Vec<_> = std::fs::read_dir(root)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .collect();
    entries.sort();
    for path in entries {
        if path.is_dir() {
            corpus(&path, files);
        } else if let Some(family) = family(&path) {
            files.push((path, family));
        }
    }
}

/// The facade type's Markdown, or the debug form of its refusal.
fn facade_markdown(
    path: &Path,
    family: Family,
    bytes: bool,
    options: &MarkdownOptions,
) -> Result<String, String> {
    let result = match (family, bytes) {
        (Family::Document, false) => {
            Document::open(path).and_then(|document| document.to_markdown_with_options(options))
        },
        (Family::Document, true) => Document::from_bytes(std::fs::read(path).unwrap())
            .and_then(|document| document.to_markdown_with_options(options)),
        (Family::Presentation, false) => litchi::Presentation::open(path)
            .and_then(|presentation| presentation.to_markdown_with_options(options)),
        (Family::Presentation, true) => {
            litchi::Presentation::from_bytes(std::fs::read(path).unwrap())
                .and_then(|presentation| presentation.to_markdown_with_options(options))
        },
    };
    result.map_err(|error| format!("{error:?}"))
}

/// `File`'s Markdown, or the code of its refusal.
fn file_markdown(path: &Path, bytes: bool, options: &MarkdownOptions) -> Result<String, String> {
    let file = if bytes {
        litchi::open_bytes(std::fs::read(path).unwrap())
    } else {
        litchi::open(path)
    };
    file.and_then(|file| file.to_markdown_with_options(options))
        .map_err(|error| error.code().as_str().to_owned())
}

/// The facade error's variant name, for the summary.
fn variant(debug: &str) -> &str {
    debug
        .split(|character: char| !character.is_alphanumeric())
        .next()
        .unwrap_or(debug)
}

#[test]
fn file_markdown_matches_the_facade_types_over_the_corpus() {
    let mut files = Vec::new();
    corpus(&fixture(""), &mut files);
    assert!(
        files.len() > 300,
        "the corpus shrank: {} files",
        files.len()
    );

    let option_sets = [
        ("default", MarkdownOptions::new()),
        ("plain", MarkdownOptions::new().with_styles(false)),
        (
            "metadata+html",
            MarkdownOptions::new()
                .with_metadata(true)
                .with_table_style(TableStyle::StyledHtml),
        ),
    ];

    let mut identical = 0usize;
    let mut refused = BTreeMap::<(String, String), usize>::new();
    let mut differences = Vec::new();
    for (path, family) in &files {
        let relative = path
            .strip_prefix(fixture(""))
            .unwrap_or(path)
            .display()
            .to_string();
        for (name, options) in &option_sets {
            for bytes in [false, true] {
                let route = if bytes { "bytes" } else { "path" };
                let old = facade_markdown(path, *family, bytes, options);
                let new = file_markdown(path, bytes, options);
                match (&old, &new) {
                    (Ok(old), Ok(new)) if old == new => identical += 1,
                    (Err(old), Err(new)) => {
                        *refused
                            .entry((variant(old).to_owned(), new.clone()))
                            .or_default() += 1;
                    },
                    _ => differences.push(format!(
                        "{relative} [{name}, {route}]: facade {:?} / File {:?}",
                        old.as_ref().map(String::len),
                        new.as_ref().map(String::len)
                    )),
                }
            }
        }
    }
    eprintln!(
        "Markdown parity over {} fixtures x {} option sets x 2 routes: {identical} identical, {} refused by both, {} differences",
        files.len(),
        option_sets.len(),
        refused.values().sum::<usize>(),
        differences.len()
    );
    for ((old, new), count) in &refused {
        eprintln!("  refused by both: facade {old} / File {new}: {count}");
    }
    assert!(
        differences.is_empty(),
        "File Markdown differs from the facade types:\n{}",
        differences.join("\n")
    );
}

// ---------------------------------------------- goldens for M0a step S9 --
//
// S9 deletes the facade types this file compares `File` against, and this
// file with them. Before that, these tests pin the golden files under
// `tests/golden/` to the facade types' output, so the `File`-only tests that
// read the goldens keep this file's semantics, and those of the parity tests
// in `unified_open_facade.rs`, once the facade types are gone:
//
// - `markdown_corpus.tsv` <- the corpus test above; read by
//   `markdown_corpus_golden.rs`.
// - `tolerance_reports.tsv` <- `unified_open_facade.rs`'s S7 parity of
//   tolerance reports; read there.
// - `pptx_slide_catalog.tsv` <- `unified_open_facade.rs`'s S7 parity of the
//   PPTX slide catalog and slide size; read there.
//
// `LITCHI_BLESS_GOLDEN=1` rewrites the goldens from the facade types instead
// of comparing (record `docs/unified-ops/s8c-facade-tests.md`).

/// The path of a golden file.
fn golden(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden")
        .join(name)
}

/// Compare `actual` with the golden file `name`, or rewrite the golden when
/// `LITCHI_BLESS_GOLDEN=1`.
#[track_caller]
fn assert_golden_file(name: &str, actual: &str) {
    let path = golden(name);
    if std::env::var_os("LITCHI_BLESS_GOLDEN").is_some_and(|value| value == "1") {
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, actual).unwrap();
        return;
    }
    let expected = std::fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
    if expected != actual {
        let first = expected
            .lines()
            .zip(actual.lines())
            .position(|(expected, actual)| expected != actual)
            .map_or_else(
                || "a missing or extra line".to_owned(),
                |line| format!("line {}", line + 1),
            );
        panic!(
            "{} differs from the facade types' output at {first}",
            path.display()
        );
    }
}

/// The fixture's path below `test-data`, with `/` separators.
fn relative(path: &Path) -> String {
    path.strip_prefix(fixture(""))
        .unwrap_or(path)
        .components()
        .map(|component| component.as_os_str().to_string_lossy().into_owned())
        .collect::<Vec<_>>()
        .join("/")
}

/// FNV-1a over 64 bits: a stable digest of a Markdown output, so the
/// corpus golden stays small. No hashing crate is a facade dependency.
fn fnv1a64(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// One corpus outcome as the golden records it: `<length>:<digest>` for
/// Markdown, `!<variant>` for a refusal.
fn golden_outcome(result: &Result<String, String>) -> String {
    match result {
        Ok(markdown) => format!("{}:{:016x}", markdown.len(), fnv1a64(markdown.as_bytes())),
        Err(debug) => format!("!{}", variant(debug)),
    }
}

#[test]
fn the_markdown_corpus_golden_is_the_facade_types_output() {
    let mut files = Vec::new();
    corpus(&fixture(""), &mut files);
    let option_sets = [
        MarkdownOptions::new(),
        MarkdownOptions::new().with_styles(false),
        MarkdownOptions::new()
            .with_metadata(true)
            .with_table_style(TableStyle::StyledHtml),
    ];
    let mut golden = String::from(
        "# Markdown of every DOCX, DOC, RTF, ODT, PPTX, PPT and ODP fixture under test-data, as\n\
         # the facade's Document and Presentation types produced it (M0a step S8c). Columns:\n\
         # fixture, then default/path, default/bytes, plain/path, plain/bytes,\n\
         # metadata+html/path, metadata+html/bytes, where `plain` is `with_styles(false)` and\n\
         # `metadata+html` is `with_metadata(true)` plus `TableStyle::StyledHtml`.\n\
         # An outcome is `<bytes>:<FNV-1a 64 hex>` for Markdown or `!<facade error variant>`\n\
         # for a refusal; `=` repeats the outcome to its left.\n",
    );
    for (path, family) in &files {
        golden.push_str(&relative(path));
        let mut previous = String::new();
        for options in &option_sets {
            for bytes in [false, true] {
                let outcome = golden_outcome(&facade_markdown(path, *family, bytes, options));
                golden.push('\t');
                golden.push_str(if outcome == previous { "=" } else { &outcome });
                previous = outcome;
            }
        }
        golden.push('\n');
    }
    assert_golden_file("markdown_corpus.tsv", &golden);
}

/// One golden row per tolerance report: `family`, `route`, the fixture and
/// the report's debug form.
#[cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
fn tolerance_row(
    family: &str,
    route: &str,
    fixture: &str,
    report: &litchi::ToleranceReport,
) -> String {
    format!("{family}\t{route}\t{fixture}\t{report:?}\n")
}

/// The tolerance reports of the facade types over the fixtures of
/// `unified_open_facade.rs`'s S7 parity tests, read as those tests read
/// them: documents by path and by bytes after their metadata, presentations
/// by path after their metadata, workbooks by path after their text and
/// metadata.
#[cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
#[test]
fn the_tolerance_report_golden_is_the_facade_types_output() {
    let mut golden = String::from(
        "# Tolerance reports of the facade's Document, Presentation and Workbook types (M0a\n\
         # step S8c), as unified_open_facade.rs's S7 parity tests read them. Columns: family,\n\
         # route, fixture, the report's debug form.\n",
    );
    for relative in [
        "ole/doc/watermark.doc",
        "libreoffice-core/embeddedobj/qa/cppunit/data/insert-file-config.doc",
        "poi/test-data/openxml4j/OPCCompliance_CoreProperties_LimitedXSITypeAttribute_NotPresentFAIL.docx",
        "libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx",
        "libreoffice-core/sw/qa/extras/rtfexport/data/tdf108955.rtf",
        "libreoffice-core/sw/qa/extras/layout/data/A020-min.rtf",
        "odfpy/tests/examples/simpletable.odt",
    ] {
        let path = fixture(relative);
        let document = Document::open(&path).unwrap();
        let _ = document.metadata().unwrap();
        golden.push_str(&tolerance_row(
            "document",
            "path",
            relative,
            &document.tolerance_report().unwrap(),
        ));
        let document = Document::from_bytes(std::fs::read(&path).unwrap()).unwrap();
        let _ = document.metadata().unwrap();
        golden.push_str(&tolerance_row(
            "document",
            "bytes",
            relative,
            &document.tolerance_report().unwrap(),
        ));
    }
    for relative in [
        "poi/test-data/slideshow/45543.ppt",
        "ole/ppt/empty.ppt",
        "libreoffice-core/oox/qa/unit/data/shape-text-alignment.pptx",
        "libreoffice-core/sd/qa/unit/data/odp/ole_icon.odp",
    ] {
        let presentation = litchi::Presentation::open(fixture(relative)).unwrap();
        let _ = presentation.metadata().unwrap();
        golden.push_str(&tolerance_row(
            "presentation",
            "path",
            relative,
            &presentation.tolerance_report().unwrap(),
        ));
    }
    for relative in [
        "ole/xls/WithCustomViews.xls",
        "libreoffice-core/chart2/qa/extras/data/xls/chart.xls",
        "ooxml/xlsx/universal-content.xlsx",
        "libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx",
        "ooxml/xlsb/universal-content.xlsb",
        "libreoffice-core/sc/qa/unit/data/xlsb/tdf108017_calcProtection.xlsb",
        "libreoffice-core/sc/qa/extras/testdocuments/ScDatabaseRangeObj.ods",
    ] {
        let workbook = litchi::Workbook::open(fixture(relative)).unwrap();
        let _ = workbook.text().unwrap();
        let _ = workbook.metadata().unwrap();
        golden.push_str(&tolerance_row(
            "workbook",
            "path",
            relative,
            &workbook.tolerance_report().unwrap(),
        ));
    }
    assert_golden_file("tolerance_reports.tsv", &golden);
}

/// The slide catalog and slide size the facade's `Presentation` reports for
/// every `test-data/ooxml/pptx` fixture it opens.
#[test]
fn the_pptx_slide_catalog_golden_is_the_facade_types_output() {
    let mut golden = String::from(
        "# Slide catalog and slide size of every test-data/ooxml/pptx fixture the facade's\n\
         # Presentation opens (M0a step S8c). Columns: file name, slide width and height in\n\
         # EMU (`-` when absent), then `position:slide id` per slide.\n",
    );
    let mut paths: Vec<_> = std::fs::read_dir(fixture("ooxml/pptx"))
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().and_then(|extension| extension.to_str()) == Some("pptx"))
        .collect();
    paths.sort();
    for path in paths {
        let Ok(presentation) = litchi::Presentation::open(&path) else {
            continue;
        };
        let size =
            |value: Option<i64>| value.map_or_else(|| "-".to_owned(), |value| value.to_string());
        let catalog = presentation
            .slide_catalog()
            .unwrap()
            .iter()
            .map(|slide| format!("{}:{}", slide.position(), slide.slide_id()))
            .collect::<Vec<_>>()
            .join(",");
        golden.push_str(&format!(
            "{}\t{}\t{}\t{catalog}\n",
            path.file_name().unwrap().to_string_lossy(),
            size(presentation.slide_width().unwrap()),
            size(presentation.slide_height().unwrap()),
        ));
    }
    assert_golden_file("pptx_slide_catalog.tsv", &golden);
}
