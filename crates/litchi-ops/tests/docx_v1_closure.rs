//! The DOCX v1-closure writes of record 0342 through the unified API:
//! table borders (`set` and `add` of `border.*`), row height and flags, and
//! `listStyle`/`start` on a new paragraph, following OfficeCLI's
//! `word/tables.sh` and `word/paragraph-formatting.sh`. Each write is read
//! back after a byte round trip on both owners; an identical set is an
//! exact no-op, and a refused value leaves the input unchanged.
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

fn document_xml(file: &File) -> String {
    let bytes = file.to_bytes().expect("publish");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name("word/document.xml").expect("main document");
    let mut contents = String::new();
    member.read_to_string(&mut contents).expect("read main");
    contents
}

/// The attributes of each `w:tblBorders` side of the first table, by side,
/// in document order; attribute order is not compared.
fn borders(xml: &str, table: usize) -> Vec<(String, Vec<String>)> {
    let Some(start) = xml
        .match_indices("<w:tblBorders>")
        .nth(table)
        .map(|(at, _)| at)
    else {
        return Vec::new();
    };
    let end = start
        + xml[start..]
            .find("</w:tblBorders>")
            .expect("closed borders");
    xml[start + "<w:tblBorders>".len()..end]
        .split("<w:")
        .filter(|side| !side.is_empty())
        .map(|side| {
            let side = side.trim_end_matches("/>");
            let mut parts = side.split(' ');
            let name = parts.next().unwrap_or_default().to_owned();
            let mut attributes: Vec<String> = parts.map(str::to_owned).collect();
            attributes.sort();
            (name, attributes)
        })
        .collect()
}

fn side(name: &str, style: &str, size: &str, color: &str) -> (String, Vec<String>) {
    let mut attributes = vec![
        format!("w:color=\"{color}\""),
        format!("w:sz=\"{size}\""),
        format!("w:val=\"{style}\""),
    ];
    attributes.sort();
    (name.to_owned(), attributes)
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

fn with_table() -> File {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    ok(file.add(
        "/body",
        NewNode::new("table").prop("rows", "3").prop("cols", "4"),
    ))
    .file
}

/// `word/tables.sh:189-197`: outer and inside borders after creation, then
/// a repeated header row of exact height.
#[test]
fn table_borders_and_row_properties_are_set() {
    for eager in [true, false] {
        let mut file = reopen(&with_table(), eager);
        for (key, value) in [
            ("border.top", "double;8;1F3864"),
            ("border.bottom", "double;8;1F3864"),
            ("border.left", "double;8;1F3864"),
            ("border.right", "double;8;1F3864"),
            ("border.horizontal", "single;4;9DC3E6"),
            ("border.vertical", "single;4;9DC3E6"),
        ] {
            file = ok(file.set("/body/tbl[1]", props! { key => value }));
        }
        file = ok(file.set(
            "/body/tbl[1]/tr[1]",
            props! { "header" => true, "height.exact" => "400" },
        ));
        let xml = document_xml(&file);
        assert_eq!(
            borders(&xml, 0),
            [
                side("top", "double", "8", "1F3864"),
                side("left", "double", "8", "1F3864"),
                side("bottom", "double", "8", "1F3864"),
                side("right", "double", "8", "1F3864"),
                side("insideH", "single", "4", "9DC3E6"),
                side("insideV", "single", "4", "9DC3E6"),
            ],
            "{xml}"
        );
        for read in [reopen(&file, true), reopen(&file, false)] {
            assert_eq!(
                prop(&read, "/body/tbl[1]/tr[1]", "header").as_deref(),
                Some("true")
            );
            assert_eq!(
                prop(&read, "/body/tbl[1]/tr[1]", "height.rule").as_deref(),
                Some("exact")
            );
            assert_eq!(prop(&read, "/body/tbl[1]/tr[2]", "header"), None);
        }
        // An identical set is an exact no-op.
        let same = ok(file.set("/body/tbl[1]", props! { "border.top" => "double;8;1F3864" }));
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        // `header=false` removes the flag; `null` removes a border.
        let cleared = ok(file.set("/body/tbl[1]/tr[1]", props! { "header" => false }));
        assert_eq!(prop(&cleared, "/body/tbl[1]/tr[1]", "header"), None);
        let cleared = ok(cleared.set(
            "/body/tbl[1]",
            props! { "border.all" => litchi_ops::PropValue::Null },
        ));
        assert!(!document_xml(&cleared).contains("<w:tblBorders>"));
        // A bad value is refused and changes nothing.
        let before = file.to_bytes().unwrap();
        let error = file
            .set("/body/tbl[1]", props! { "border.top" => "zigzag;8;000000" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("border.top"));
        let error = file
            .set("/body/tbl[1]/tr[1]", props! { "height.exact" => "tall" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(file.to_bytes().unwrap(), before);
    }
}

/// `word/tables.sh:179`: `border.all` on `add`, in OfficeCLI's two value
/// forms; a bare row `height` has no rule.
#[test]
fn borders_seed_a_new_table_and_heights_take_their_rule() {
    let file = litchi_ops::create(FileFormat::Docx).unwrap();
    let file = ok(file.add(
        "/body",
        NewNode::new("table")
            .prop("rows", "2")
            .prop("cols", "2")
            .prop("border.all", "single;8;2E74B5"),
    ))
    .file;
    let file = ok(file.add(
        "/body",
        NewNode::new("table")
            .prop("rows", "1")
            .prop("cols", "1")
            .prop("border.all", "1pt dashed FF0000"),
    ))
    .file;
    let xml = document_xml(&file);
    let all = |style: &str, color: &str| {
        ["top", "left", "bottom", "right", "insideH", "insideV"]
            .map(|name| side(name, style, "8", color))
            .to_vec()
    };
    assert_eq!(borders(&xml, 0), all("single", "2E74B5"), "{xml}");
    assert_eq!(borders(&xml, 1), all("dashed", "FF0000"), "{xml}");
    let file = ok(file.set("/body/tbl[1]/tr[2]", props! { "height" => "1cm" }));
    let file = ok(file.set("/body/tbl[1]/tr[1]", props! { "height.atLeast" => "300" }));
    let xml = document_xml(&file);
    assert!(xml.contains(r#"<w:trHeight w:val="567"/>"#), "{xml}");
    assert!(
        xml.contains(r#"<w:trHeight w:val="300" w:hRule="atLeast"/>"#),
        "{xml}"
    );
}

/// `word/paragraph-formatting.sh:153-155`: `listStyle` and `start` on a new
/// paragraph create a list, a following item continues it, and `start`
/// restarts.
#[test]
fn list_style_and_start_number_a_new_paragraph() {
    for eager in [true, false] {
        let file = reopen(&litchi_ops::create(FileFormat::Docx).unwrap(), eager);
        let file = ok(file.add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "Bulleted item")
                .prop("listStyle", "bullet"),
        ))
        .file;
        let file = ok(file.add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "Second bullet")
                .prop("listStyle", "bullet"),
        ))
        .file;
        let file = ok(file.add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "Ordered item starting at 5")
                .prop("listStyle", "ordered")
                .prop("start", 5),
        ))
        .file;
        for read in [reopen(&file, true), reopen(&file, false)] {
            let first = prop(&read, "/body/p[1]", "numId");
            assert!(first.is_some());
            assert_eq!(
                prop(&read, "/body/p[2]", "numId"),
                first,
                "the list continues"
            );
            assert_eq!(
                prop(&read, "/body/p[1]", "listStyle").as_deref(),
                Some("bullet")
            );
            let ordered = prop(&read, "/body/p[3]", "numId");
            assert!(ordered.is_some() && ordered != first);
            assert_eq!(prop(&read, "/body/p[3]", "start").as_deref(), Some("5"));
            assert_eq!(prop(&read, "/body/p[3]", "numLevel").as_deref(), Some("0"));
        }
        // `none` adds a plain paragraph.
        let plain = ok(file.add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "plain")
                .prop("listStyle", "none"),
        ))
        .file;
        assert_eq!(prop(&plain, "/body/p[4]", "numId"), None);
    }
}

/// A paragraph added into a table cell takes its list the same way.
#[test]
fn list_style_numbers_a_paragraph_added_to_a_cell() {
    let file = ok(with_table().add(
        "/body/tbl[1]/tr[1]/tc[1]",
        NewNode::new("paragraph")
            .prop("text", "item")
            .prop("listStyle", "ordered"),
    ))
    .file;
    assert!(prop(&file, "/body/tbl[1]/tr[1]/tc[1]/p[2]", "numId").is_some());
}
