//! DOCX equations through the format-neutral `File` facade: display math
//! (`oMathPara[N]`) and inline math (`p[N]/oMath[M]`) are read, added, set
//! and removed in the body, table cells, headers and notes, on both owners,
//! and every byte outside the edited equation is kept.
#![cfg(feature = "docx")]

use std::collections::BTreeMap;
use std::io::{Cursor, Read as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// A Word document with one inline equation (beside a run) and one display
/// equation.
const MATH: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/microsoft-office/word/word-math-navigation.docx"
);
const INLINE: &str = "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>";

fn open(eager: bool, bytes: Vec<u8>) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager fixture")
    } else {
        litchi_ops::open_bytes(bytes).expect("source fixture")
    }
}

fn fixture(eager: bool) -> File {
    open(eager, std::fs::read(MATH).expect("fixture"))
}

fn reopen(file: &File, eager: bool) -> File {
    open(eager, file.to_bytes().expect("publish"))
}

fn paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| {
            node.unwrap_or_else(|error| panic!("{selector}: {error}"))
                .view()
                .path()
                .to_string()
        })
        .collect()
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

/// Every member of a saved package.
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).expect("member");
            let mut contents = Vec::new();
            member.read_to_end(&mut contents).expect("read member");
            (member.name().expect("member name").into_owned(), contents)
        })
        .collect()
}

fn main_xml(bytes: &[u8]) -> String {
    String::from_utf8(members(bytes).remove("word/document.xml").expect("main")).expect("UTF-8")
}

/// The fixture's inline equation, exactly as its part holds it.
fn fixture_inline(xml: &str) -> &str {
    let start = xml
        .find("OMML_INLINE_005: </w:t></w:r>")
        .expect("inline paragraph")
        + "OMML_INLINE_005: </w:t></w:r>".len();
    let end = xml[start..].find("</m:oMath>").expect("inline end") + start + "</m:oMath>".len();
    &xml[start..end]
}

/// Assert that `after` differs from `before` only in the main part, by
/// replacing `old` with `new` once.
fn assert_only_change(before: &[u8], after: &[u8], old: &str, new: &str) {
    let before = members(before);
    let after = members(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, contents) in &before {
        if name != "word/document.xml" {
            assert_eq!(contents, &after[name], "{name} changed");
        }
    }
    let before_main = String::from_utf8(before["word/document.xml"].clone()).unwrap();
    let after_main = String::from_utf8(after["word/document.xml"].clone()).unwrap();
    assert_eq!(before_main.matches(old).count(), 1);
    assert_eq!(after_main, before_main.replacen(old, new, 1));
}

#[test]
fn equations_are_listed_with_their_mode_and_omml() {
    for eager in [false, true] {
        let file = fixture(eager);
        let found = paths(&file, "equation");
        assert_eq!(found.len(), 2, "{found:?}");
        let inline = found
            .iter()
            .find(|path| path.ends_with("/oMath[1]"))
            .expect("inline equation");
        assert!(inline.starts_with("/body/p["), "{inline}");
        assert!(found.contains(&"/body/oMathPara[1]".to_owned()));
        assert_eq!(property(&file, inline, "mode").as_deref(), Some("inline"));
        assert_eq!(
            property(&file, "/body/oMathPara[1]", "mode").as_deref(),
            Some("display")
        );
        let omml = property(&file, inline, "omml").expect("inline omml");
        assert!(omml.starts_with("<m:oMath><m:sSup>"), "{omml}");
        let omml = property(&file, "/body/oMathPara[1]", "omml").expect("display omml");
        assert!(omml.starts_with("<m:oMathPara><m:oMath><m:nary>"), "{omml}");
        // The paragraph lists its inline equation.
        let paragraph = inline.rsplit_once("/oMath").expect("parent").0;
        let children = file.get(paragraph).expect("paragraph");
        // A paragraph's view lists its runs only, as OfficeCLI's does; its
        // equations are reached by path and query.
        assert!(
            !children.view().children().any(|child| child
                .expect("child")
                .view()
                .path()
                .to_string()
                == *inline)
        );
        #[cfg(feature = "equation")]
        {
            let latex = property(&file, inline, "latex").expect("inline latex");
            assert!(latex.contains(r"\frac{1}{2}"), "{latex}");
            let latex = property(&file, "/body/oMathPara[1]", "latex").expect("display latex");
            assert!(latex.contains(r"\sum"), "{latex}");
        }
    }
}

#[test]
fn setting_an_equation_rewrites_only_its_markup() {
    for eager in [false, true] {
        let file = fixture(eager);
        let before = file.to_bytes().expect("source bytes");
        let old = fixture_inline(&main_xml(&before)).to_owned();
        let inline = paths(&file, "equation")
            .into_iter()
            .find(|path| path.ends_with("/oMath[1]"))
            .expect("inline equation");
        let set = file
            .set(inline.as_str(), props! { "omml" => INLINE })
            .unwrap_or_else(|error| panic!("set omml: {error}"));
        assert_eq!(property(&set, &inline, "omml").as_deref(), Some(INLINE));
        let after = set.to_bytes().expect("edited bytes");
        assert_only_change(&before, &after, &old, INLINE);
        // The source handle is an unchanged snapshot.
        assert_eq!(file.to_bytes().expect("republish"), before);

        // A display equation keeps its math paragraph.
        let set = file
            .set("/body/oMathPara[1]", props! { "omml" => INLINE })
            .expect("set display omml");
        assert_eq!(
            property(&set, "/body/oMathPara[1]", "omml").as_deref(),
            Some(format!("<m:oMathPara>{INLINE}</m:oMathPara>").as_str())
        );
        let reopened = reopen(&set, eager);
        assert_eq!(paths(&reopened, "equation").len(), 2);
    }
}

#[cfg(feature = "equation")]
#[test]
fn latex_writes_word_omml() {
    for eager in [false, true] {
        let file = fixture(eager)
            .set("/body/oMathPara[1]", props! { "latex" => r"\frac{a}{b}" })
            .expect("set latex");
        let omml = property(&file, "/body/oMathPara[1]", "omml").expect("omml");
        assert!(omml.starts_with("<m:oMathPara><m:oMath><m:f>"), "{omml}");
        assert!(!omml.contains("xmlns"), "{omml}");
        assert_eq!(
            property(&file, "/body/oMathPara[1]", "latex").as_deref(),
            Some(r"\frac{a}{b}")
        );
        let added = file
            .add("/body", NewNode::new("equation").prop("latex", "E=mc^2"))
            .expect("add display latex");
        assert_eq!(added.path.to_string(), "/body/oMathPara[2]");
        let file = reopen(&added.into_file(), eager);
        let latex = property(&file, "/body/oMathPara[2]", "latex").expect("latex");
        assert!(latex.contains("^{2}"), "{latex}");
    }
}

#[cfg(not(feature = "equation"))]
#[test]
fn latex_needs_the_equation_feature() {
    let error = fixture(true)
        .set("/body/oMathPara[1]", props! { "latex" => "x" })
        .expect_err("no converter");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn equations_are_added_in_the_body_cells_headers_and_notes() {
    let base = litchi_ops::create(litchi_core::FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Area: "))
        .unwrap()
        .file
        .add(
            "/body",
            NewNode::new("table").prop("rows", 1).prop("cols", 1),
        )
        .unwrap()
        .file
        .add("/", NewNode::new("header").prop("text", "Top"))
        .unwrap()
        .file;
    let added = base
        .add(
            "/body/p[1]",
            NewNode::new("footnote").prop("text", "Note text"),
        )
        .unwrap();
    let note = added.path.to_string();
    let base = added.file;
    for eager in [false, true] {
        let mut file = reopen(&base, eager);
        let note_paragraph = format!("{note}/p[1]");
        let note_equation = format!("{note_paragraph}/oMath[1]");
        // Stories first: a source-backed handle's story edit materializes
        // the package from the handle's source.
        for (parent, mode, wanted) in [
            ("/header[1]/p[1]", "inline", "/header[1]/p[1]/oMath[1]"),
            ("/header[1]", "display", "/header[1]/oMathPara[1]"),
            (note_paragraph.as_str(), "inline", note_equation.as_str()),
            ("/body/p[1]", "inline", "/body/p[1]/oMath[1]"),
            ("/body", "display", "/body/oMathPara[1]"),
            (
                "/body/tbl[1]/tr[1]/tc[1]/p[1]",
                "inline",
                "/body/tbl[1]/tr[1]/tc[1]/p[1]/oMath[1]",
            ),
            (
                "/body/tbl[1]/tr[1]/tc[1]",
                "display",
                "/body/tbl[1]/tr[1]/tc[1]/oMathPara[1]",
            ),
        ] {
            let added = file
                .add(
                    parent,
                    NewNode::new("equation")
                        .prop("mode", mode)
                        .prop("omml", INLINE),
                )
                .unwrap_or_else(|error| panic!("add under {parent}: {error}"));
            assert_eq!(added.path.to_string(), wanted);
            file = added.into_file();
            assert_eq!(property(&file, wanted, "mode").as_deref(), Some(mode));
        }
        let file = reopen(&file, eager);
        for path in [
            "/body/p[1]/oMath[1]",
            "/body/oMathPara[1]",
            "/body/tbl[1]/tr[1]/tc[1]/p[1]/oMath[1]",
            "/body/tbl[1]/tr[1]/tc[1]/oMathPara[1]",
            "/header[1]/p[1]/oMath[1]",
            "/header[1]/oMathPara[1]",
            note_equation.as_str(),
        ] {
            let omml = property(&file, path, "omml").unwrap_or_else(|| panic!("{path}"));
            assert!(omml.contains("<m:t>y</m:t>"), "{path}: {omml}");
        }
        // Set and remove in the stories.
        let replaced = "<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>";
        let file = file
            .set("/header[1]/p[1]/oMath[1]", props! { "omml" => replaced })
            .expect("set header equation")
            .set("/header[1]/oMathPara[1]", props! { "omml" => replaced })
            .expect("set header display")
            .set(note_equation.as_str(), props! { "omml" => replaced })
            .expect("set note equation");
        // These story parts declare no math namespace on their roots, so
        // each equation keeps its own declaration.
        let declared = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>z</m:t></m:r></m:oMath>"#;
        for path in ["/header[1]/p[1]/oMath[1]", note_equation.as_str()] {
            assert_eq!(property(&file, path, "omml").as_deref(), Some(declared));
        }
        let file = file
            .remove("/header[1]/p[1]/oMath[1]")
            .expect("remove header inline")
            .remove("/header[1]/oMathPara[1]")
            .expect("remove header display")
            .remove(note_equation.as_str())
            .expect("remove note equation");
        assert!(file.get("/header[1]/oMathPara[1]").is_err());
        assert!(file.get("/header[1]/p[1]/oMath[1]").is_err());
        assert!(file.get(note_equation.as_str()).is_err());
        assert_eq!(paths(&file, "equation").len(), 4);
        // The cell's display equation is removed with its paragraph.
        let file = file
            .remove("/body/tbl[1]/tr[1]/tc[1]/oMathPara[1]")
            .expect("remove cell display");
        assert_eq!(paths(&file, "equation").len(), 3);
    }
}

#[test]
fn removing_equations_takes_only_what_belongs_to_them() {
    for eager in [false, true] {
        let file = fixture(eager);
        let before = file.to_bytes().expect("source bytes");
        let old = fixture_inline(&main_xml(&before)).to_owned();
        let inline = paths(&file, "equation")
            .into_iter()
            .find(|path| path.ends_with("/oMath[1]"))
            .expect("inline equation");
        let removed = file.remove(inline.as_str()).expect("remove inline");
        assert_only_change(&before, &removed.to_bytes().unwrap(), &old, "");
        // A display equation's paragraph goes with it.
        let paragraphs = paths(&file, "paragraph").len();
        let removed = file.remove("/body/oMathPara[1]").expect("remove display");
        assert!(
            paths(&removed, "equation")
                .iter()
                .all(|path| path.contains("/oMath["))
        );
        assert_eq!(paths(&removed, "paragraph").len(), paragraphs);
        let xml = main_xml(&removed.to_bytes().unwrap());
        assert!(!xml.contains("<m:oMathPara>"));
        // Removing every equation by query.
        let query = litchi_core::selector::Query::parse("equation")
            .unwrap()
            .with_scope(litchi_core::selector::Scope::Path(Box::new(
                litchi_core::selector::Path::parse("/").unwrap(),
            )));
        let removed = file.remove(query).expect("remove all");
        assert!(paths(&removed, "equation").is_empty());
    }
}

#[test]
fn invalid_equation_writes_are_refused_with_typed_errors() {
    let file = fixture(true);
    let inline = paths(&file, "equation")
        .into_iter()
        .find(|path| path.ends_with("/oMath[1]"))
        .expect("inline equation");
    let error = file
        .add("/body/p[1]", NewNode::new("equation"))
        .expect_err("no content");
    assert_eq!(error.code(), Code::MissingProperty);
    let error = file
        .add(
            "/body/p[1]",
            NewNode::new("equation")
                .prop("omml", INLINE)
                .prop("latex", "x"),
        )
        .expect_err("two contents");
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add(
            "/body/p[1]",
            NewNode::new("equation")
                .prop("mode", "display")
                .prop("omml", INLINE),
        )
        .expect_err("display under a paragraph");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .set(inline.as_str(), props! { "omml" => "<m:oMath>" })
        .expect_err("malformed OMML");
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set(
            inline.as_str(),
            props! { "omml" => "<m:oMathPara><m:oMath/></m:oMathPara>" },
        )
        .expect_err("a math paragraph is not inline math");
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set(inline.as_str(), props! { "mode" => "display" })
        .expect_err("mode is chosen when adding");
    assert_eq!(error.code(), Code::UnsupportedProperty);
    // OMML holding what an equation does not own is an invalid value.
    for omml in [
        r#"<m:oMath xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><m:r><v:shape><v:imagedata o:relid="rId3"/></v:shape></m:r></m:oMath>"#,
        r#"<m:oMath><m:r><x:junk xmlns:x="urn:x"/></m:r></m:oMath>"#,
        r#"<m:oMath><m:r><w:footnoteReference xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="9"/></m:r></m:oMath>"#,
    ] {
        let error = file
            .set(inline.as_str(), props! { "omml" => omml })
            .expect_err("foreign content");
        assert_eq!(error.code(), Code::InvalidValue, "{omml}: {error}");
        // The reason is in the message the caller sees.
        assert!(error.to_string().contains("Office Math"), "{error}");
        assert_eq!(error.property(), Some("omml"));
    }
    // The element name chooses the kind of math.
    let error = file
        .add("/body/p[1]", NewNode::new("oMathPara").prop("omml", INLINE))
        .expect_err("display math is not added inside a paragraph");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .add("/body", NewNode::new("oMath").prop("omml", INLINE))
        .expect_err("inline math is added under a paragraph");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .add(
            "/body",
            NewNode::new("oMathPara")
                .prop("mode", "inline")
                .prop("omml", INLINE),
        )
        .expect_err("the name and the mode disagree");
    assert_eq!(error.code(), Code::InvalidArgument);
    let added = file
        .add("/body", NewNode::new("oMathPara").prop("omml", INLINE))
        .expect("display math by name");
    assert!(added.path.to_string().starts_with("/body/oMathPara["));
}

/// A DOCX package holding `body` in a Word main part that declares `w` and
/// `m` on its root.
fn package(body: &str) -> Vec<u8> {
    use std::io::Write as _;
    const MAIN: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{MAIN}"/></Types>"#
    );
    let relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", relationships),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options).expect("member");
        zip.write_all(contents.as_bytes()).expect("write member");
    }
    zip.finish().expect("finish").into_inner()
}

#[test]
fn display_math_beside_other_paragraph_content_is_an_equation() {
    // A bookmark around display math, as Word writes a cross-referenced
    // equation.
    let bytes = package(
        r#"<w:p><w:bookmarkStart w:id="0" w:name="eq1"/><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath></m:oMathPara><w:bookmarkEnd w:id="0"/></w:p>"#,
    );
    for eager in [false, true] {
        let file = open(eager, bytes.clone());
        assert_eq!(paths(&file, "equation"), ["/body/p[1]/oMathPara[1]"]);
        let path = "/body/p[1]/oMathPara[1]";
        assert_eq!(property(&file, path, "mode").as_deref(), Some("display"));
        let set = file
            .set(path, props! { "omml" => INLINE })
            .expect("set display math of a paragraph");
        assert_eq!(
            property(&set, path, "omml").as_deref(),
            Some(
                r#"<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>"#
            )
        );
        let removed = set
            .remove(path)
            .expect("remove display math of a paragraph");
        assert!(paths(&removed, "equation").is_empty());
        let xml = main_xml(&removed.to_bytes().unwrap());
        assert!(
            xml.contains(
                r#"<w:p><w:bookmarkStart w:id="0" w:name="eq1"/><w:bookmarkEnd w:id="0"/></w:p>"#
            ),
            "{xml}"
        );
    }
}

#[test]
fn a_cell_keeps_a_paragraph_when_its_display_equation_goes() {
    let bytes = package(
        "<w:tbl><w:tr><w:tc><w:tcPr/><w:sdt><w:sdtPr/><w:sdtContent><w:p><m:oMathPara><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath></m:oMathPara></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>",
    );
    for eager in [false, true] {
        let file = open(eager, bytes.clone());
        let path = "/body/tbl[1]/tr[1]/tc[1]/sdt[1]/oMathPara[1]";
        let removed = file.remove(path).expect("remove the display equation");
        let xml = main_xml(&removed.to_bytes().unwrap());
        assert!(
            xml.contains("<w:sdtContent><w:p></w:p></w:sdtContent>"),
            "{xml}"
        );
    }
}

#[test]
fn strict_omml_is_written_in_a_transitional_document_s_namespaces() {
    let bytes = package("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    let strict = r#"<m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"><m:r><m:t>q</m:t></m:r></m:oMath>"#;
    let file = open(true, bytes)
        .add("/body/p[1]", NewNode::new("equation").prop("omml", strict))
        .expect("add Strict OMML")
        .into_file();
    let xml = main_xml(&file.to_bytes().unwrap());
    assert!(
        xml.contains("<m:oMath><m:r><m:t>q</m:t></m:r></m:oMath></w:p>"),
        "{xml}"
    );
}

#[cfg(feature = "equation")]
#[test]
fn latex_that_does_not_convert_is_an_invalid_value() {
    let file = fixture(true);
    for latex in [
        format!("{}x{}", "{".repeat(10_000), "}".repeat(10_000)),
        "x\u{1}".to_owned(),
        r"\frac{a".to_owned(),
    ] {
        let error = file
            .set("/body/oMathPara[1]", props! { "latex" => latex.as_str() })
            .expect_err("unconvertible LaTeX");
        assert_eq!(error.code(), Code::InvalidValue, "{latex:.20}: {error}");
    }
}
