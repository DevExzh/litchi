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
        // Run formatting of an existing hyperlink goes through its runs.
        let error = file
            .set("/body/p[1]/hyperlink[1]", props! { "bold" => true })
            .expect_err("bold seeds a new hyperlink only");
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

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

#[test]
fn hyperlinks_are_added_and_retargeted_on_both_owners() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        let added = file
            .add(
                "/body/tbl[1]/tr[1]/tc[1]/p[1]",
                NewNode::new("hyperlink")
                    .prop("url", "https://example.test/new?x=1&y=2")
                    .prop("text", " more"),
            )
            .unwrap_or_else(|error| panic!("hyperlink add: {error}"));
        assert_eq!(
            added.path.to_string(),
            "/body/tbl[1]/tr[1]/tc[1]/p[1]/hyperlink[1]"
        );
        let file = added.into_file();
        let link = "/body/tbl[1]/tr[1]/tc[1]/p[1]/hyperlink[1]";
        assert_eq!(
            property(&file, link, "url").as_deref(),
            Some("https://example.test/new?x=1&y=2")
        );
        assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]/p[1]"), "A1 more");
        assert_eq!(property(&file, link, "color").as_deref(), Some("#0563C1"));

        // A bookmark hyperlink in the body, then a retarget to a URL.
        let file = file
            .add(
                "/body/p[1]",
                NewNode::new("link")
                    .prop("anchor", "Top")
                    .prop("text", "top")
                    .prop("bold", true)
                    .at(Placement::Start),
            )
            .unwrap_or_else(|error| panic!("anchor hyperlink add: {error}"))
            .into_file();
        assert_eq!(
            property(&file, "/body/p[1]/hyperlink[1]", "anchor").as_deref(),
            Some("Top")
        );
        assert_eq!(text(&file, "/body/p[1]"), "topLead link");
        let file = file
            .set(
                "/body/p[1]/hyperlink[1]",
                props! { "url" => "https://example.test/top" },
            )
            .unwrap_or_else(|error| panic!("hyperlink retarget: {error}"));
        assert_eq!(
            property(&file, "/body/p[1]/hyperlink[1]", "url").as_deref(),
            Some("https://example.test/top")
        );
        assert_eq!(property(&file, "/body/p[1]/hyperlink[1]", "anchor"), None);
        let reopened = reopen(&file, eager);
        assert_eq!(
            property(&reopened, link, "url").as_deref(),
            Some("https://example.test/new?x=1&y=2")
        );
        assert_eq!(
            property(&reopened, "/body/p[1]/hyperlink[1]", "url").as_deref(),
            Some("https://example.test/top")
        );
        // Both targets are required to be exclusive.
        let error = reopened
            .add(
                "/body/p[1]",
                NewNode::new("hyperlink").prop("text", "nowhere"),
            )
            .expect_err("a hyperlink needs a target");
        assert_eq!(error.code(), Code::MissingProperty);
    }
}

#[test]
fn comments_are_added_read_and_removed_on_both_owners() {
    for eager in [false, true] {
        let file = open(eager, fixture());
        assert!(file.get("/comments").is_err(), "no comments part yet");
        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("comment")
                    .prop("author", "Ada")
                    .prop("initials", "AL")
                    .prop("date", "2026-10-13T00:00:00Z")
                    .prop("text", "Check the lead"),
            )
            .unwrap_or_else(|error| panic!("comment add: {error}"));
        assert_eq!(added.path.to_string(), "/comments/comment[1]");
        let file = added.into_file();
        let first = "/comments/comment[1]";
        assert_eq!(text(&file, first), "Check the lead");
        assert_eq!(property(&file, first, "author").as_deref(), Some("Ada"));
        assert_eq!(property(&file, first, "initials").as_deref(), Some("AL"));
        assert_eq!(
            property(&file, first, "date").as_deref(),
            Some("2026-10-13T00:00:00Z")
        );
        let id = property(&file, first, "id").expect("comment id");
        let raw = format!("/comments/comment[@commentId={id}]");
        let by_id = file
            .get(litchi_ops::path::raw::path(raw.as_str()).expect("raw native-id path"))
            .unwrap_or_else(|error| panic!("{raw}: {error}"));
        assert_eq!(
            by_id.view().text().expect("comment text").as_deref(),
            Some("Check the lead")
        );
        // The commented paragraph keeps its text.
        assert_eq!(text(&file, "/body/p[1]"), "Lead link");

        // A second comment, anchored on a cell run, joins the part.
        let added = file
            .add(
                "/body/tbl[1]/tr[2]/tc[1]/p[1]/r[1]",
                NewNode::new("comment")
                    .prop("author", "Bo")
                    .prop("text", "Cell note"),
            )
            .unwrap_or_else(|error| panic!("run comment add: {error}"));
        assert_eq!(added.path.to_string(), "/comments/comment[2]");
        let file = added.into_file();
        let xml = document_xml(&file);
        assert_eq!(xml.matches("<w:commentReference").count(), 2, "{xml}");
        let reopened = reopen(&file, eager);
        assert_eq!(text(&reopened, "/comments/comment[2]"), "Cell note");
        assert_eq!(
            property(&reopened, "/comments/comment[2]", "author").as_deref(),
            Some("Bo")
        );

        // Removal strips the entry and its anchors.
        let removed = reopened
            .remove("/comments/comment[1]")
            .unwrap_or_else(|error| panic!("comment remove: {error}"));
        assert_eq!(text(&removed, "/comments/comment[1]"), "Cell note");
        assert!(removed.get("/comments/comment[2]").is_err());
        let xml = document_xml(&removed);
        assert_eq!(xml.matches("<w:commentReference").count(), 1, "{xml}");
        assert_eq!(text(&removed, "/body/p[1]"), "Lead link");

        // An author is required.
        let error = removed
            .add("/body/p[1]", NewNode::new("comment").prop("text", "anon"))
            .expect_err("a comment needs an author");
        assert_eq!(error.code(), Code::MissingProperty);
    }
}

/// A 4x2 PNG.
const PNG_URI: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC";

#[test]
fn pictures_are_added_and_read_on_both_owners() {
    use base64::Engine as _;
    let png = base64::engine::general_purpose::STANDARD
        .decode(PNG_URI.split_once(',').unwrap().1)
        .unwrap();
    let directory = tempfile::tempdir().expect("temporary directory");
    let file_source = directory.path().join("logo.png");
    std::fs::write(&file_source, &png).expect("write the picture source");
    for eager in [false, true] {
        let file = open(eager, fixture());
        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("picture")
                    .prop("src", PNG_URI)
                    .prop("alt", "Logo")
                    .at(Placement::Start),
            )
            .unwrap_or_else(|error| panic!("picture add: {error}"));
        assert_eq!(added.path.to_string(), "/body/p[1]/picture[1]");
        let file = added.into_file();
        let picture = "/body/p[1]/picture[1]";
        // 4x2 pixels at 96 DPI.
        assert_eq!(property(&file, picture, "width").as_deref(), Some("3pt"));
        assert_eq!(property(&file, picture, "alt").as_deref(), Some("Logo"));
        assert_eq!(text(&file, "/body/p[1]"), "Lead link");
        // A second picture at the end of the paragraph follows the first.
        let added = file
            .add("/body/p[1]", NewNode::new("img").prop("src", PNG_URI))
            .unwrap_or_else(|error| panic!("second picture add: {error}"));
        assert_eq!(added.path.to_string(), "/body/p[1]/picture[2]");
        let file = added.into_file();
        assert_eq!(
            property(&file, "/body/p[1]/picture[1]", "alt").as_deref(),
            Some("Logo")
        );

        // A file source with an explicit width keeps the aspect ratio.
        let added = file
            .add(
                "/body/tbl[1]/tr[1]/tc[1]/p[1]",
                NewNode::new("image")
                    .prop("src", file_source.to_str().unwrap())
                    .prop("width", "2in")
                    .prop("name", "Cell logo"),
            )
            .unwrap_or_else(|error| panic!("cell picture add: {error}"));
        assert_eq!(
            added.path.to_string(),
            "/body/tbl[1]/tr[1]/tc[1]/p[1]/picture[1]"
        );
        let reopened = reopen(&added.into_file(), eager);
        let cell = "/body/tbl[1]/tr[1]/tc[1]/p[1]/picture[1]";
        assert_eq!(
            property(&reopened, cell, "width").as_deref(),
            Some("5.08cm")
        );
        assert_eq!(
            property(&reopened, cell, "height").as_deref(),
            Some("2.54cm")
        );
        assert_eq!(
            property(&reopened, cell, "name").as_deref(),
            Some("Cell logo")
        );
        assert_eq!(property(&reopened, picture, "alt").as_deref(), Some("Logo"));
        assert_ne!(
            property(&reopened, picture, "relId"),
            property(&reopened, cell, "relId")
        );

        let error = reopened
            .add("/body/p[1]", NewNode::new("picture").prop("alt", "nothing"))
            .expect_err("a picture needs a source");
        assert_eq!(error.code(), Code::MissingProperty);
        let error = reopened
            .add(
                "/body/p[1]",
                NewNode::new("picture").prop("src", "data:text/plain;base64,aGVsbG8gd29ybGQ="),
            )
            .expect_err("a picture needs an image");
        assert_eq!(error.code(), Code::InvalidValue);
    }
}

fn numbering_xml(file: &File) -> Option<String> {
    let bytes = file.to_bytes().expect("publish");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name("word/numbering.xml").ok()?;
    let mut contents = String::new();
    member
        .read_to_string(&mut contents)
        .expect("read numbering");
    Some(contents)
}

#[test]
fn list_styles_create_continue_and_remove_lists_on_both_owners() {
    let paragraphs = |count: usize| {
        let body: String = (1..=count)
            .map(|index| format!("<w:p><w:r><w:t>Item {index}</w:t></w:r></w:p>"))
            .collect();
        package(&format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
        ))
    };
    for eager in [false, true] {
        let file = open(eager, paragraphs(5));
        assert!(numbering_xml(&file).is_none());
        // One write over two adjacent paragraphs: one new list, continued.
        let file = file
            .set("/body/p[1]", props! { "listStyle" => "bullet" })
            .unwrap_or_else(|error| panic!("bullet p1: {error}"));
        let file = file
            .set("/body/p[2]", props! { "listStyle" => "ul" })
            .unwrap_or_else(|error| panic!("bullet p2: {error}"));
        let first = property(&file, "/body/p[1]", "numId").expect("p1 numId");
        assert_eq!(property(&file, "/body/p[2]", "numId"), Some(first.clone()));
        assert_eq!(
            property(&file, "/body/p[2]", "numLevel").as_deref(),
            Some("0")
        );
        // A numbered list after the bullets starts a new list.
        let file = file
            .set("/body/p[3]", props! { "listStyle" => "ordered" })
            .unwrap_or_else(|error| panic!("ordered p3: {error}"));
        let ordered = property(&file, "/body/p[3]", "numId").expect("p3 numId");
        assert_ne!(ordered, first);
        // `start` forces a fresh list even after a same-kind paragraph.
        let file = file
            .set(
                "/body/p[4]",
                props! { "listStyle" => "decimal", "start" => 3, "numLevel" => "1" },
            )
            .unwrap_or_else(|error| panic!("restart p4: {error}"));
        let restarted = property(&file, "/body/p[4]", "numId").expect("p4 numId");
        assert_ne!(restarted, ordered);
        assert_eq!(
            property(&file, "/body/p[4]", "numLevel").as_deref(),
            Some("1")
        );
        let numbering = numbering_xml(&file).expect("a numbering part");
        assert_eq!(
            numbering.matches("<w:abstractNum ").count(),
            3,
            "{numbering}"
        );
        assert!(numbering.contains(r#"<w:start w:val="3"/>"#));
        // `none` removes the numbering.
        let file = file
            .set("/body/p[2]", props! { "listStyle" => "none" })
            .unwrap_or_else(|error| panic!("none p2: {error}"));
        assert_eq!(property(&file, "/body/p[2]", "numId"), None);
        let reopened = reopen(&file, eager);
        assert_eq!(property(&reopened, "/body/p[1]", "numId"), Some(first));
        assert_eq!(property(&reopened, "/body/p[3]", "numId"), Some(ordered));
        // A query over adjacent paragraphs joins them into one new list.
        let joined = open(eager, paragraphs(3))
            .set("/body > p", props! { "listStyle" => "bullet" })
            .unwrap_or_else(|error| panic!("query bullets: {error}"));
        let id = property(&joined, "/body/p[1]", "numId");
        assert!(id.is_some());
        assert_eq!(property(&joined, "/body/p[3]", "numId"), id);
    }
}
