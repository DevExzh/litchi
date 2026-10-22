//! XLSX find/replace through `set` with `find=` and `replace=` (record
//! 0273): the workbook searches every worksheet, a sheet, cell, range, row
//! or column its own cells; formulas and non-text values are not searched;
//! zero matches are an exact no-op; rich-text cells and format properties
//! are refused.
#![cfg(feature = "xlsx")]

use std::io::{Cursor, Write as _};

use litchi_core::FileFormat;
use litchi_ops::{BatchItem, Code, File, OpenOptions, Props, SourceMode, props};

fn find(file: &File, path: &str, props: Props<'_>) -> litchi_ops::Result<(File, Option<usize>)> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    let commit = file.apply(&item)?;
    Ok((commit.file().clone(), commit.value().matched()))
}

fn value(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn workbook() -> File {
    let file = litchi_ops::create(FileFormat::Xlsx).unwrap();
    let file = file
        .add("/", litchi_ops::NewNode::new("sheet").prop("name", "Two"))
        .unwrap()
        .file;
    file.edit(|tx| {
        tx.set("/Sheet1/A1", props! { "value" => "draft one, draft two" })?;
        tx.set("/Sheet1/B2", props! { "value" => "Draft" })?;
        tx.set("/Sheet1/C3", props! { "value" => 42 })?;
        tx.set("/Sheet1/D4", props! { "formula" => "\"draft\"&\"x\"" })?;
        tx.set("/Two/A1", props! { "value" => "draft" })?;
        Ok(())
    })
    .unwrap()
    .file()
    .clone()
}

#[test]
fn the_workbook_and_sheet_scopes_follow_officecli() {
    let file = workbook();
    let (after, matched) = find(
        &file,
        "/",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    assert_eq!(matched, Some(3));
    assert_eq!(value(&after, "/Sheet1/A1"), "final one, final two");
    assert_eq!(value(&after, "/Sheet1/B2"), "Draft");
    assert_eq!(value(&after, "/Two/A1"), "final");
    // The formula keeps its text.
    let formula = after
        .get("/Sheet1/D4")
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == "formula")
        .map(|property| property.value().to_string());
    assert_eq!(formula.as_deref(), Some("\"draft\"&\"x\""));
    let (after, matched) = find(
        &file,
        "/Sheet1",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    assert_eq!(matched, Some(2));
    assert_eq!(value(&after, "/Two/A1"), "draft");
    // Litchi also scopes to cells, ranges and rows.
    let (_, matched) = find(
        &file,
        "/Sheet1/A1:B2",
        Props::new()
            .with("find", "raft")
            .with("replace", "x")
            .with("matchCase", false),
    )
    .unwrap();
    assert_eq!(matched, Some(3));
    let (_, matched) = find(
        &file,
        "/Sheet1/row[2]",
        Props::new().with("find", "Draft").with("replace", "x"),
    )
    .unwrap();
    assert_eq!(matched, Some(1));
    let (_, matched) = find(
        &file,
        "/Sheet1/A1",
        Props::new()
            .with("find", "draft")
            .with("replace", "x")
            .with("wholeWord", true),
    )
    .unwrap();
    assert_eq!(matched, Some(2));
    // Saving and reopening keeps the replacement.
    let (after, _) = find(
        &file,
        "/",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    let reopened = litchi_ops::open_bytes(after.to_bytes().unwrap()).unwrap();
    assert_eq!(value(&reopened, "/Sheet1/A1"), "final one, final two");
}

#[test]
fn zero_matches_are_an_exact_no_op() {
    let file = workbook();
    let bytes = file.to_bytes().unwrap();
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let (after, matched) = find(
            &file,
            "/",
            Props::new().with("find", "absent").with("replace", "x"),
        )
        .unwrap();
        assert_eq!(matched, Some(0));
        assert_eq!(after.to_bytes().unwrap(), bytes);
        // A replacement equal to the match counts and changes nothing.
        let (after, matched) = find(
            &file,
            "/",
            Props::new().with("find", "draft").with("replace", "draft"),
        )
        .unwrap();
        assert_eq!(matched, Some(3));
        assert_eq!(after.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn formatting_and_bare_finds_are_refused() {
    let file = workbook();
    let error = find(
        &file,
        "/",
        Props::new()
            .with("find", "draft")
            .with("replace", "x")
            .with("bold", true),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(error.property(), Some("bold"));
    let error = find(
        &file,
        "/",
        Props::new().with("find", "draft").with("bold", true),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    let error = find(&file, "/", Props::new().with("find", "draft")).unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
}

fn rich_package(inline: bool) -> Vec<u8> {
    let inline = if inline {
        r#"<c r="C1" t="inlineStr"><is><r><rPr><i/></rPr><t>draft</t></r></is></c>"#
    } else {
        ""
    };
    let sheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>{inline}</row></sheetData></worksheet>"#
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#,
        ),
        (
            "xl/sharedStrings.xml",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>plain draft</t></si><si><r><rPr><b/></rPr><t>rich </t></r><r><t>draft</t></r></si></sst>"#,
        ),
        ("xl/worksheets/sheet1.xml", sheet.as_str()),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(contents.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// The worksheet part of a published workbook.
fn sheet_xml(file: &File) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(file.to_bytes().unwrap())).unwrap();
    let mut part = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut part, &mut xml).unwrap();
    xml
}

/// Record 0295: a formatted shared string or an inline rich string is
/// rewritten run by run, each run keeping its properties.
#[test]
fn a_rich_text_cell_keeps_its_runs() {
    let file = litchi_ops::open_bytes(rich_package(false)).unwrap();
    let (after, matched) = find(
        &file,
        "/Sheet1",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    assert_eq!(matched, Some(2));
    assert_eq!(value(&after, "/Sheet1/A1"), "plain final");
    assert_eq!(value(&after, "/Sheet1/B1"), "rich final");
    let xml = sheet_xml(&after);
    assert!(
        xml.contains(r#"<r><rPr><b/></rPr><t xml:space="preserve">rich </t></r><r><t xml:space="preserve">final</t></r>"#),
        "{xml}"
    );
    let reopened = litchi_ops::open_bytes(after.to_bytes().unwrap()).unwrap();
    assert_eq!(value(&reopened, "/Sheet1/B1"), "rich final");
    // A match across runs collapses into the run where it starts.
    let (after, matched) = find(
        &file,
        "/Sheet1/B1",
        Props::new()
            .with("find", "rich draft")
            .with("replace", "one"),
    )
    .unwrap();
    assert_eq!(matched, Some(1));
    assert_eq!(value(&after, "/Sheet1/B1"), "one");
    assert!(sheet_xml(&after).contains(r#"<r><rPr><b/></rPr><t xml:space="preserve">one</t></r>"#));

    // An inline rich string keeps its runs too.
    let inline = litchi_ops::open_bytes(rich_package(true)).unwrap();
    let (after, matched) = find(
        &inline,
        "/Sheet1/C1",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    assert_eq!(matched, Some(1));
    assert_eq!(value(&after, "/Sheet1/C1"), "final");
    assert!(
        sheet_xml(&after).contains(r#"<r><rPr><i/></rPr><t xml:space="preserve">final</t></r>"#)
    );
}

/// Record 0295: merge fills a formatted shared string run by run.
#[test]
fn merge_keeps_rich_runs() {
    let file = litchi_ops::open_bytes(rich_package(false))
        .unwrap()
        .set("/Sheet1/A1", litchi_ops::props! { "value" => "{{a}}" })
        .unwrap();
    let file = find(
        &file,
        "/Sheet1/B1",
        Props::new()
            .with("find", "draft")
            .with("replace", "{{name}}"),
    )
    .unwrap()
    .0;
    let mut record = litchi_ops::merge::Record::new();
    record.insert("name", "Ada");
    let commit = file
        .merge(&record, &litchi_ops::merge::Options::new())
        .unwrap();
    assert_eq!(value(commit.file(), "/Sheet1/B1"), "rich Ada");
    assert!(sheet_xml(commit.file()).contains(r#"<r><rPr><b/></rPr>"#));
}
