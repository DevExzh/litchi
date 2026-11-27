//! DOCX OLE equation objects (MathType, Equation Editor 3.0) through the
//! format-neutral `File` facade: `…/p[N]/equation[M]` nodes are read
//! beside the Office Math ones, added, set and removed on both owners, with
//! read-back through the leaf reader and the MTEF reader, every untouched
//! part kept, and typed refusals.
//!
//! With `LITCHI_REPORT_DIR` set, the files the tests write are saved there
//! for inspection in an office application.
#![cfg(feature = "docx")]

#[cfg(feature = "equation")]
use std::collections::BTreeMap;
use std::io::{Cursor, Read as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// Word 2010: one `Equation.DSMT4` object with a WMF preview, its run
/// lowered by `w:position`.
const MATHTYPE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sw/qa/core/data/ooxml/pass/fdo78906.docx");
/// Word 2007: one `Equation.3` object with an EMF preview.
const EQUATION_EDITOR: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/mathtype.docx"
);
/// A Word document with one inline and one display Office Math equation.
const OFFICE_MATH: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/word/word-math-navigation.docx");
const OBJECT: &str = "/body/p[1]/equation[1]";

fn open(bytes: &[u8], eager: bool) -> File {
    let mode = if eager {
        SourceMode::Eager
    } else {
        SourceMode::Auto
    };
    litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode)).unwrap()
}

#[cfg(feature = "equation")]
fn reopen(file: &File, eager: bool) -> File {
    open(&file.to_bytes().expect("publish"), eager)
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| node.expect("node").view().path().to_string())
        .collect()
}

#[cfg(feature = "equation")]
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

#[cfg(feature = "equation")]
/// Save `bytes` as `name` under `LITCHI_REPORT_DIR`, when it is set.
fn keep(name: &str, bytes: &[u8]) {
    if let Some(directory) = std::env::var_os("LITCHI_REPORT_DIR") {
        let directory = std::path::PathBuf::from(directory);
        std::fs::create_dir_all(&directory).expect("report directory");
        std::fs::write(directory.join(name), bytes).expect("write report file");
    }
}

#[cfg(feature = "equation")]
/// Whether a member is one an equation-object edit may change.
fn owned_by_objects(name: &str) -> bool {
    name == "word/document.xml"
        || name == "word/_rels/document.xml.rels"
        || name == "[Content_Types].xml"
        || name.starts_with("word/embeddings/")
        || name.starts_with("word/media/")
}

#[cfg(feature = "equation")]
fn decoded_mtef(file: &File, path: &str) -> Vec<u8> {
    use base64::Engine as _;
    let text = property(file, path, "mtef").unwrap_or_else(|| panic!("{path} mtef"));
    base64::engine::general_purpose::STANDARD
        .decode(text)
        .expect("base64 mtef")
}

#[test]
fn word_equation_objects_are_equation_nodes() {
    for eager in [false, true] {
        for (bytes, prog_id) in [
            (MATHTYPE, "Equation.DSMT4"),
            (EQUATION_EDITOR, "Equation.3"),
        ] {
            let file = open(bytes, eager);
            assert_eq!(paths(&file, "equation"), [OBJECT]);
            assert_eq!(file.get(OBJECT).unwrap().kind(), "equation");
            assert_eq!(property(&file, OBJECT, "progId").as_deref(), Some(prog_id));
            assert_eq!(property(&file, OBJECT, "mode").as_deref(), Some("inline"));
            assert_eq!(property(&file, OBJECT, "omml"), None);
            assert_eq!(property(&file, OBJECT, "error"), None);
            #[cfg(feature = "equation")]
            {
                let mtef = decoded_mtef(&file, OBJECT);
                let read =
                    litchi_equation::read_mtef(&mtef, &litchi_equation::Limits::DEFAULT).unwrap();
                let version = if prog_id == "Equation.3" { 3 } else { 5 };
                assert_eq!(read.info().unwrap().version, version);
                let latex = property(&file, OBJECT, "latex").expect("latex");
                assert_eq!(latex, read.latex());
            }
            #[cfg(not(feature = "equation"))]
            assert_eq!(property(&file, OBJECT, "mtef"), None);
        }
    }
}

#[test]
fn office_math_paths_are_unchanged_beside_objects() {
    for eager in [false, true] {
        let file = open(OFFICE_MATH, eager);
        let found = paths(&file, "equation");
        assert_eq!(found.len(), 2, "{found:?}");
        assert!(found.iter().any(|path| path.ends_with("/oMath[1]")));
        assert!(found.contains(&"/body/oMathPara[1]".to_owned()));
        // `equation[M]` of a paragraph names its objects, of which this one
        // has none.
        let inline = found
            .iter()
            .find(|path| path.ends_with("/oMath[1]"))
            .unwrap();
        let paragraph = inline.rsplit_once("/oMath").unwrap().0;
        assert!(
            file.get(format!("{paragraph}/equation[1]").as_str())
                .is_err()
        );
        // Under the body, `equation[N]` stays an alias of display math, which
        // reads back by its canonical `oMathPara` path.
        let display = file.get("/body/equation[1]").expect("body equation alias");
        assert_eq!(display.view().path().to_string(), "/body/oMathPara[1]");
        assert_eq!(
            property(&file, "/body/equation[1]", "mode").as_deref(),
            Some("display")
        );
        assert_eq!(property(&file, "/body/equation[1]", "progId"), None);
        // Under a paragraph it names an OLE object, by its own canonical path.
        let objects = open(MATHTYPE, eager);
        let object = objects.get(OBJECT).expect("object");
        assert_eq!(object.view().path().to_string(), OBJECT);
        assert_eq!(
            property(&objects, OBJECT, "mode").as_deref(),
            Some("inline")
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn objects_are_added_from_latex_omml_and_mtef() {
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
        .file;
    let mtef3 = property(&open(EQUATION_EDITOR, true), OBJECT, "mtef").unwrap();
    for eager in [false, true] {
        let file = reopen(&base, eager);
        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("latex", r"\frac{a}{b}"),
            )
            .expect("add from latex");
        assert_eq!(added.path.to_string(), OBJECT);
        let added = added
            .into_file()
            .add(
                "/body/p[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("omml", "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"),
            )
            .expect("add from omml");
        assert_eq!(added.path.to_string(), "/body/p[1]/equation[2]");
        let added = added
            .into_file()
            .add(
                "/body/tbl[1]/tr[1]/tc[1]/p[1]",
                NewNode::new("equation").prop("mtef", mtef3.as_str()),
            )
            .expect("add from mtef");
        assert_eq!(
            added.path.to_string(),
            "/body/tbl[1]/tr[1]/tc[1]/p[1]/equation[1]"
        );
        let bytes = added.into_file().to_bytes().expect("publish");
        if eager {
            keep("docx-equation-objects-added.docx", &bytes);
        }
        let file = open(&bytes, eager);
        assert_eq!(
            property(&file, OBJECT, "latex").as_deref(),
            Some(r"\frac{a}{b}")
        );
        assert_eq!(
            property(&file, OBJECT, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        assert_eq!(
            property(&file, "/body/p[1]/equation[2]", "latex").as_deref(),
            Some("x")
        );
        let cell = "/body/tbl[1]/tr[1]/tc[1]/p[1]/equation[1]";
        assert_eq!(
            property(&file, cell, "progId").as_deref(),
            Some("Equation.3")
        );
        assert_eq!(
            property(&file, cell, "mtef").as_deref(),
            Some(mtef3.as_str())
        );
        // The leaf reader sees the same objects and parts.
        let package = litchi_docx::Package::from_bytes(&bytes).unwrap();
        let objects = package.document().unwrap().paragraphs().unwrap()[0]
            .equation_objects()
            .unwrap();
        assert_eq!(objects.len(), 2);
        let members = members(&bytes);
        for name in [
            "word/embeddings/oleObject1.bin",
            "word/embeddings/oleObject2.bin",
            "word/embeddings/oleObject3.bin",
            "word/media/image3.png",
        ] {
            assert!(members.contains_key(name), "{name}");
        }
        let document = String::from_utf8(members["word/document.xml"].clone()).unwrap();
        assert_eq!(document.matches("<v:shapetype ").count(), 1);
        for shape in ["_x0000_i1025", "_x0000_i1026", "_x0000_i1027"] {
            assert_eq!(
                document.matches(format!("id=\"{shape}\"").as_str()).count(),
                1,
                "{shape}"
            );
        }
        assert_eq!(paths(&file, "equation").len(), 3);
    }
}

#[cfg(feature = "equation")]
#[test]
fn setting_an_object_keeps_its_kind_and_every_other_part() {
    for eager in [false, true] {
        for (bytes, version) in [(MATHTYPE, 5u8), (EQUATION_EDITOR, 3)] {
            let file = open(bytes, eager);
            let before = members(&file.to_bytes().unwrap());
            let set = file
                .set(OBJECT, props! { "latex" => r"x^{2}+1" })
                .expect("set latex");
            let after_bytes = set.to_bytes().unwrap();
            if eager {
                keep(
                    &format!("docx-equation-object-set-mtef{version}.docx"),
                    &after_bytes,
                );
            }
            let set = open(&after_bytes, eager);
            let mtef = decoded_mtef(&set, OBJECT);
            let read =
                litchi_equation::read_mtef(&mtef, &litchi_equation::Limits::DEFAULT).unwrap();
            assert_eq!(read.info().unwrap().version, version);
            assert_eq!(
                property(&set, OBJECT, "latex").map(|latex| latex.replace(' ', "")),
                Some("x^{2}+1".to_owned())
            );
            let after = members(&after_bytes);
            for (name, contents) in &before {
                if !owned_by_objects(name) {
                    assert_eq!(Some(contents), after.get(name), "{name} changed");
                }
            }
            // The Office preview went with its relationship; a PNG took its
            // place.
            assert!(
                !after
                    .keys()
                    .any(|name| name.ends_with(".wmf") || name.ends_with(".emf"))
            );
            assert!(after.keys().any(|name| name.ends_with(".png")));
            // Office Math and MTEF content both convert; `omml` writes too.
            let set = set
                .set(
                    OBJECT,
                    props! { "omml" => "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>" },
                )
                .expect("set omml");
            assert_eq!(property(&set, OBJECT, "latex").as_deref(), Some("y"));
            let mtef = property(&open(bytes, true), OBJECT, "mtef").unwrap();
            let set = set
                .set(OBJECT, props! { "mtef" => mtef.as_str() })
                .expect("set mtef");
            assert_eq!(property(&set, OBJECT, "mtef"), Some(mtef));
        }
    }
}

#[cfg(feature = "equation")]
#[test]
fn removing_an_object_takes_its_parts() {
    for eager in [false, true] {
        let file = open(MATHTYPE, eager);
        let before = members(&file.to_bytes().unwrap());
        let removed = file.remove(OBJECT).expect("remove");
        let bytes = removed.to_bytes().unwrap();
        if eager {
            keep("docx-equation-object-removed.docx", &bytes);
        }
        let after = members(&bytes);
        assert!(!after.contains_key("word/embeddings/oleObject1.bin"));
        assert!(!after.contains_key("word/media/image1.wmf"));
        let relationships =
            String::from_utf8(after["word/_rels/document.xml.rels"].clone()).unwrap();
        assert!(!relationships.contains("oleObject"), "{relationships}");
        assert!(!relationships.contains("image1.wmf"), "{relationships}");
        let document = String::from_utf8(after["word/document.xml"].clone()).unwrap();
        assert!(!document.contains("w:object"));
        for (name, contents) in &before {
            if !owned_by_objects(name) {
                assert_eq!(Some(contents), after.get(name), "{name} changed");
            }
        }
        assert!(paths(&open(&bytes, eager), "equation").is_empty());
    }
}

#[test]
fn invalid_object_writes_are_refused_with_typed_errors() {
    for eager in [false, true] {
        let file = open(MATHTYPE, eager);
        let error = file
            .add(
                "/body/p[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.3")
                    .prop("latex", "x"),
            )
            .expect_err("only MathType objects are made");
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("progId"));
        let error = file
            .add("/body", NewNode::new("equation").prop("mtef", "AAAA"))
            .expect_err("an object goes under a paragraph");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add(
                "/body/p[1]",
                NewNode::new("equation").prop("mtef", "not base64!"),
            )
            .expect_err("mtef is base64");
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .set(OBJECT, props! { "progId" => "Equation.DSMT4" })
            .expect_err("the ProgID is chosen when adding");
        assert_eq!(error.code(), Code::UnsupportedProperty);
        #[cfg(feature = "equation")]
        {
            // Raw MTEF of another version than the object's.
            let mtef3 = property(&open(EQUATION_EDITOR, true), OBJECT, "mtef").unwrap();
            let error = file
                .set(OBJECT, props! { "mtef" => mtef3.as_str() })
                .expect_err("MTEF 3 into a MathType object");
            assert_eq!(error.code(), Code::InvalidValue);
            let error = file
                .set(OBJECT, props! { "mtef" => "AAAA" })
                .expect_err("not MTEF");
            assert_eq!(error.code(), Code::InvalidValue);
        }
        // The input stays intact.
        assert_eq!(
            file.to_bytes().unwrap(),
            open(MATHTYPE, eager).to_bytes().unwrap()
        );
        // MTEF is not Office Math.
        let math = open(OFFICE_MATH, eager);
        let inline = paths(&math, "equation")
            .into_iter()
            .find(|path| path.ends_with("/oMath[1]"))
            .unwrap();
        let error = math
            .set(inline.as_str(), props! { "mtef" => "AAAA" })
            .expect_err("mtef on Office Math");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.property(), Some("mtef"));
    }
}

/// `MATHTYPE` with its embedded compound file replaced by `payload`.
fn with_embedding(payload: &[u8]) -> Vec<u8> {
    use std::io::Write as _;
    let mut source = zip::ZipArchive::new(Cursor::new(MATHTYPE)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut member = source.by_index(index).unwrap();
        let name = member.name().unwrap().into_owned();
        let mut contents = Vec::new();
        member.read_to_end(&mut contents).unwrap();
        if name == "word/embeddings/oleObject1.bin" {
            contents = payload.to_vec();
        }
        out.start_file(name.as_str(), zip::write::SimpleFileOptions::default())
            .unwrap();
        out.write_all(&contents).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn an_unreadable_embedding_is_a_per_object_error() {
    let garbled = with_embedding(b"not a compound file");
    for eager in [false, true] {
        let file = open(&garbled, eager);
        assert_eq!(paths(&file, "equation"), [OBJECT]);
        assert_eq!(
            property(&file, OBJECT, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        assert_eq!(property(&file, OBJECT, "mtef"), None);
        assert_eq!(property(&file, OBJECT, "latex"), None);
        #[cfg(feature = "equation")]
        {
            let error = property(&file, OBJECT, "error").expect("error");
            assert!(error.contains("OLE equation object"), "{error}");
            // A rewrite writes a new object of the ProgID's kind.
            let set = file
                .set(OBJECT, props! { "latex" => "x" })
                .expect("rewrite an unreadable embedding");
            assert_eq!(property(&set, OBJECT, "error"), None);
            assert_eq!(property(&set, OBJECT, "latex").as_deref(), Some("x"));
            assert_eq!(
                property(&set, OBJECT, "progId").as_deref(),
                Some("Equation.DSMT4")
            );
            let read = litchi_equation::read_mtef(
                &decoded_mtef(&set, OBJECT),
                &litchi_equation::Limits::DEFAULT,
            )
            .unwrap();
            assert_eq!(read.info().unwrap().version, 5);
            // It can also be removed.
            let removed = file.remove(OBJECT).expect("remove");
            assert!(paths(&removed, "equation").is_empty());
        }
    }
}

/// `MATHTYPE` with a default header holding a copy of the body's object,
/// relating the same embedding and preview from the header part.
fn with_header_object() -> Vec<u8> {
    use std::io::Write as _;
    const OBJECT_MARKUP: &str = concat!(
        r#"<w:r><w:object w:dxaOrig="7839" w:dyaOrig="999"><v:shape id="_x0000_i1030" style="width:43pt;height:5pt" o:ole="">"#,
        r#"<v:imagedata r:id="rId2" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" "#,
        r#"ShapeID="_x0000_i1030" DrawAspect="Content" ObjectID="_1461848873" r:id="rId1"/></w:object></w:r>"#
    );
    let header = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:p>{OBJECT_MARKUP}</w:p></w:hdr>"#
    );
    let header_relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.wmf"/></Relationships>"#;
    let mut source = zip::ZipArchive::new(Cursor::new(MATHTYPE)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for index in 0..source.len() {
        let mut member = source.by_index(index).unwrap();
        let name = member.name().unwrap().into_owned();
        let mut contents = Vec::new();
        member.read_to_end(&mut contents).unwrap();
        let text = String::from_utf8_lossy(&contents).into_owned();
        let contents = match name.as_str() {
            "word/document.xml" => text
                .replace(
                    r#"<w:sectPr w:rsidR="008B24F1">"#,
                    r#"<w:sectPr w:rsidR="008B24F1"><w:headerReference w:type="default" r:id="rId99"/>"#,
                )
                .into_bytes(),
            "word/_rels/document.xml.rels" => text
                .replace(
                    "</Relationships>",
                    r#"<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#,
                )
                .into_bytes(),
            "[Content_Types].xml" => text
                .replace(
                    "</Types>",
                    r#"<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#,
                )
                .into_bytes(),
            _ => contents,
        };
        out.start_file(name.as_str(), options).unwrap();
        out.write_all(&contents).unwrap();
    }
    out.start_file("word/header1.xml", options).unwrap();
    out.write_all(header.as_bytes()).unwrap();
    out.start_file("word/_rels/header1.xml.rels", options)
        .unwrap();
    out.write_all(header_relationships.as_bytes()).unwrap();
    out.finish().unwrap().into_inner()
}

#[test]
fn objects_of_headers_share_parts_with_the_body_safely() {
    let bytes = with_header_object();
    let path = "/header[1]/p[1]/equation[1]";
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert_eq!(
            property(&file, path, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        assert_eq!(property(&file, path, "error"), None);
        #[cfg(feature = "equation")]
        {
            assert_eq!(
                property(&file, path, "mtef"),
                property(&file, OBJECT, "mtef")
            );
            // The header's object shares its embedding with the body's: a
            // replacement writes a new embedding and keeps the body's.
            let set = file
                .set(path, props! { "latex" => "z" })
                .expect("set header object");
            assert_eq!(property(&set, path, "latex").as_deref(), Some("z"));
            assert_eq!(
                property(&set, OBJECT, "mtef"),
                property(&file, OBJECT, "mtef")
            );
            let after = members(&set.to_bytes().unwrap());
            assert!(after.contains_key("word/embeddings/oleObject1.bin"));
            assert!(after.contains_key("word/embeddings/oleObject2.bin"));
            // Removing the header's object keeps the parts the body uses.
            let removed = file.remove(path).expect("remove header object");
            let after = members(&removed.to_bytes().unwrap());
            assert!(after.contains_key("word/embeddings/oleObject1.bin"));
            assert!(after.contains_key("word/media/image1.wmf"));
            let header_relationships =
                String::from_utf8(after["word/_rels/header1.xml.rels"].clone()).unwrap();
            assert!(!header_relationships.contains("oleObject"));
            assert!(property(&removed, OBJECT, "error").is_none());
            assert!(removed.get(path).is_err());
        }
    }
}

#[cfg(feature = "equation")]
#[test]
fn objects_are_written_in_headers_footers_and_notes() {
    let base = litchi_ops::create(litchi_core::FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Claim"))
        .unwrap()
        .file
        .add("/", NewNode::new("header").prop("text", "Top"))
        .unwrap()
        .file
        .add("/", NewNode::new("footer").prop("text", "Bottom"))
        .unwrap()
        .file;
    let added = base
        .add("/body/p[1]", NewNode::new("footnote").prop("text", "Foot"))
        .unwrap();
    let footnote = added.path.to_string();
    let added = added
        .into_file()
        .add("/body/p[1]", NewNode::new("endnote").prop("text", "End"))
        .unwrap();
    let endnote = added.path.to_string();
    let base = added.into_file();
    for eager in [false, true] {
        let mut file = reopen(&base, eager);
        let parents = [
            "/header[1]/p[1]".to_owned(),
            "/footer[1]/p[1]".to_owned(),
            format!("{footnote}/p[1]"),
            format!("{endnote}/p[1]"),
        ];
        for parent in &parents {
            let added = file
                .add(
                    parent.as_str(),
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", "a^{2}"),
                )
                .unwrap_or_else(|error| panic!("add under {parent}: {error}"));
            assert_eq!(added.path.to_string(), format!("{parent}/equation[1]"));
            file = added.into_file();
        }
        let bytes = file.to_bytes().unwrap();
        if eager {
            keep("docx-equation-objects-stories.docx", &bytes);
        }
        let mut file = open(&bytes, eager);
        for parent in &parents {
            let path = format!("{parent}/equation[1]");
            assert_eq!(
                property(&file, &path, "latex").as_deref(),
                Some("a^{2}"),
                "{path}"
            );
            file = file
                .set(path.as_str(), props! { "latex" => "b" })
                .unwrap_or_else(|error| panic!("set {path}: {error}"));
            assert_eq!(property(&file, &path, "latex").as_deref(), Some("b"));
        }
        // Shape numbers stay unique across the stories.
        let saved = members(&file.to_bytes().unwrap());
        let mut shapes: Vec<String> = saved
            .iter()
            .filter(|(name, _)| name.ends_with(".xml"))
            .flat_map(|(_, contents)| {
                String::from_utf8_lossy(contents)
                    .split("<v:shape id=\"")
                    .skip(1)
                    .filter_map(|tail| tail.split('"').next().map(str::to_owned))
                    .collect::<Vec<_>>()
            })
            .collect();
        shapes.sort();
        shapes.dedup();
        assert_eq!(shapes.len(), 4, "{shapes:?}");
        for parent in &parents {
            let path = format!("{parent}/equation[1]");
            file = file
                .remove(path.as_str())
                .unwrap_or_else(|error| panic!("remove {path}: {error}"));
        }
        let saved = members(&file.to_bytes().unwrap());
        assert!(
            !saved
                .keys()
                .any(|name| name.starts_with("word/embeddings/") || name.ends_with(".png")),
            "{:?}",
            saved.keys().collect::<Vec<_>>()
        );
        assert!(paths(&file, "equation").is_empty());
    }
}
