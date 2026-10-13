//! DOCX writes beyond the direct body through the format-neutral `File`
//! facade (record 0251a): paragraphs and runs inside table cells, content
//! controls and hyperlinks; hyperlink and cell text; and table, row and cell
//! structure. Every case runs on both owners, reads back through `get`, and
//! survives a save and reopen.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_core::Position;
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode, props};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

fn package(document: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/" TargetMode="External"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root_relationships),
        ("word/_rels/document.xml.rels", document_relationships),
        ("word/document.xml", document),
    ] {
        zip.start_file(name, options).expect("start fixture member");
        zip.write_all(contents.as_bytes())
            .expect("write fixture member");
    }
    zip.finish().expect("finish fixture package").into_inner()
}

/// A paragraph with a hyperlink, a table with two rows, a block content
/// control and a final section.
fn fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Lead </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#,
    )
}

fn open(eager: bool, bytes: Vec<u8>) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager fixture")
    } else {
        litchi_ops::open_bytes(bytes).expect("source fixture")
    }
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn has_format(file: &File, path: &str, key: &str, wanted: &str) -> bool {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .any(|property| property.key() == key && property.to_string() == wanted)
}

fn document_xml(file: &File) -> String {
    let bytes = file.to_bytes().expect("publish");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name("word/document.xml").expect("main document");
    let mut contents = String::new();
    member.read_to_string(&mut contents).expect("read main");
    contents
}

fn reopen(file: &File, eager: bool) -> File {
    open(eager, file.to_bytes().expect("publish"))
}

#[test]
fn paragraphs_and_runs_inside_cells_and_controls_accept_formatting_and_text() {
    for eager in [false, true] {
        let original = open(eager, fixture());
        let before = original.to_bytes().expect("original bytes");
        let file = original
            .set(
                "/body/tbl[1]/tr[2]/tc[1]/p[1]",
                props! { "align" => "center", "text" => "A2 revised" },
            )
            .unwrap_or_else(|e| panic!("cell paragraph set: {e}"))
            .set(
                "/body/tbl[1]/tr[1]/tc[2]/p[1]/r[1]",
                props! { "bold" => true },
            )
            .expect("cell run set")
            .set("/body/sdt[1]/p[1]", props! { "text" => "Control revised" })
            .expect("control paragraph set");
        assert_eq!(text(&file, "/body/tbl[1]/tr[2]/tc[1]/p[1]"), "A2 revised");
        assert!(has_format(
            &file,
            "/body/tbl[1]/tr[2]/tc[1]/p[1]",
            "align",
            "center"
        ));
        assert!(has_format(
            &file,
            "/body/tbl[1]/tr[1]/tc[2]/p[1]/r[1]",
            "bold",
            "true"
        ));
        assert_eq!(text(&file, "/body/sdt[1]/p[1]"), "Control revised");
        // The source handle is an unchanged snapshot.
        assert_eq!(original.to_bytes().expect("republish"), before);
        let reopened = reopen(&file, eager);
        assert_eq!(
            text(&reopened, "/body/tbl[1]/tr[2]/tc[1]/p[1]"),
            "A2 revised"
        );
        assert!(has_format(
            &reopened,
            "/body/tbl[1]/tr[1]/tc[2]/p[1]/r[1]",
            "bold",
            "true"
        ));
    }
}

#[test]
fn hyperlink_runs_and_hyperlink_text_are_writable() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        // `r[2]` is the run inside the hyperlink.
        let file = file
            .set("/body/p[1]/r[2]", props! { "italic" => true })
            .expect("hyperlink run format")
            .set("/body/p[1]/hyperlink[1]", props! { "text" => "new link" })
            .expect("hyperlink text");
        assert!(has_format(&file, "/body/p[1]/r[2]", "italic", "true"));
        assert_eq!(text(&file, "/body/p[1]/r[2]"), "new link");
        let xml = document_xml(&file);
        assert!(xml.contains(r#"<w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:i/></w:rPr><w:t>new link</w:t></w:r></w:hyperlink>"#), "{xml}");
        // Hyperlinks accept only their text.
        let error = file
            .set(
                "/body/p[1]/hyperlink[1]",
                props! { "url" => "https://other.test/" },
            )
            .expect_err("url is not writable here");
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }
}

#[test]
fn cell_text_replaces_the_cell_content() {
    for eager in [false, true] {
        let file = open(eager, fixture())
            .set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "Header" })
            .expect("cell text");
        assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]/p[1]"), "Header");
        assert_eq!(
            text(&reopen(&file, eager), "/body/tbl[1]/tr[1]/tc[1]"),
            "Header"
        );
    }
}

#[test]
fn paragraphs_and_runs_are_added_inside_cells_and_controls() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let added = file
            .add(
                "/body/tbl[1]/tr[1]/tc[1]",
                NewNode::new("paragraph").prop("text", "second line"),
            )
            .unwrap_or_else(|e| panic!("cell paragraph add: {e}"));
        assert_eq!(added.path.to_string(), "/body/tbl[1]/tr[1]/tc[1]/p[2]");
        let file = added.into_file();
        assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]/p[2]"), "second line");
        let added = file
            .add(
                "/body/sdt[1]",
                NewNode::new("paragraph")
                    .prop("text", "first")
                    .at(Placement::Start),
            )
            .expect("control paragraph add");
        assert_eq!(added.path.to_string(), "/body/sdt[1]/p[1]");
        let file = added.into_file();
        let added = file
            .add(
                "/body/tbl[1]/tr[2]/tc[2]/p[1]",
                NewNode::new("run")
                    .prop("text", "pre")
                    .prop("bold", true)
                    .at(Placement::Start),
            )
            .expect("cell run add");
        assert_eq!(added.path.to_string(), "/body/tbl[1]/tr[2]/tc[2]/p[1]/r[1]");
        let file = added.into_file();
        assert_eq!(text(&file, "/body/tbl[1]/tr[2]/tc[2]/p[1]"), "preB2");
        // A run after the hyperlink in a body paragraph.
        let added = file
            .add("/body/p[1]", NewNode::new("run").prop("text", " tail"))
            .expect("run after hyperlink");
        assert_eq!(added.path.to_string(), "/body/p[1]/r[3]");
        let reopened = reopen(&added.into_file(), eager);
        assert_eq!(text(&reopened, "/body/p[1]/r[3]"), " tail");
        assert_eq!(text(&reopened, "/body/sdt[1]/p[1]"), "first");
    }
}

#[test]
fn tables_and_rows_are_added_removed_and_moved() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let added = file
            .add(
                "/body",
                NewNode::new("table")
                    .prop("rows", 2)
                    .prop("cols", 3)
                    .prop("data", "a,b,c;d,e,f")
                    .at(Placement::Start),
            )
            .expect("table add");
        assert_eq!(added.path.to_string(), "/body/tbl[1]");
        let file = added.into_file();
        assert_eq!(text(&file, "/body/tbl[1]/tr[2]/tc[3]"), "f");
        assert_eq!(text(&file, "/body/tbl[2]/tr[1]/tc[1]"), "A1");

        let added = file
            .add("/body/tbl[2]", NewNode::new("row").at(Placement::Start))
            .expect("row add");
        assert_eq!(added.path.to_string(), "/body/tbl[2]/tr[1]");
        let file = added.into_file();
        assert_eq!(text(&file, "/body/tbl[2]/tr[2]/tc[1]"), "A1");
        let file = file
            .set("/body/tbl[2]/tr[1]/tc[2]", props! { "text" => "new" })
            .expect("new row cell text");

        // Move the original first table row to the end and remove a cell.
        let file = file
            .move_to("/body/tbl[2]/tr[2]", Placement::End)
            .expect("row move");
        assert_eq!(text(&file, "/body/tbl[2]/tr[3]/tc[1]"), "A1");
        let file = file
            .remove("/body/tbl[2]/tr[3]/tc[2]")
            .expect("cell remove");
        assert!(
            file.find("/body/tbl[2]/tr[3]/tc[2]")
                .expect("find")
                .is_none()
        );

        // Move the added table after the content control, then remove it.
        let file = file
            .move_to(
                "/body/tbl[1]",
                Placement::after("/body/sdt[1]").expect("anchor"),
            )
            .expect("table move");
        assert_eq!(text(&file, "/body/tbl[2]/tr[1]/tc[1]"), "a");
        let file = file.remove("/body/tbl[2]").expect("table remove");
        assert!(file.find("/body/tbl[2]").expect("find").is_none());
        let reopened = reopen(&file, eager);
        assert_eq!(text(&reopened, "/body/tbl[1]/tr[1]/tc[2]"), "new");
        assert_eq!(text(&reopened, "/body/tbl[1]/tr[3]/tc[1]"), "A1");
    }
}

#[test]
fn nested_removals_refuse_the_last_paragraph_of_a_cell() {
    let file = open(true, fixture());
    let error = file
        .remove("/body/tbl[1]/tr[1]/tc[1]/p[1]")
        .expect_err("a cell keeps one paragraph");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let file = file
        .add(
            "/body/tbl[1]/tr[1]/tc[1]",
            NewNode::new("paragraph").prop("text", "x"),
        )
        .expect("second paragraph")
        .into_file()
        .remove("/body/tbl[1]/tr[1]/tc[1]/p[1]")
        .expect("now removable");
    assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]"), "x");
}

#[test]
fn nested_positions_honor_explicit_placement() {
    let file = open(false, fixture());
    let added = file
        .add(
            "/body/tbl[1]/tr[1]/tc[1]",
            NewNode::new("paragraph")
                .prop("text", "zero")
                .at(Placement::Position(Position::new(0))),
        )
        .unwrap_or_else(|e| panic!("positioned add: {e}"));
    assert_eq!(added.path.to_string(), "/body/tbl[1]/tr[1]/tc[1]/p[1]");
    let file = added.into_file();
    assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]/p[2]"), "A1");
}
