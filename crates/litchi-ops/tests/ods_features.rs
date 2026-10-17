//! ODS worksheet features through the unified API (record 0327):
//! hyperlinks, pictures, shapes, conditional formats, content validations,
//! the auto-filter and page breaks, each added, read back on both owners,
//! set, queried and removed, with the XLSX vocabulary of record 0320.
#![cfg(feature = "ods")]
#![allow(clippy::expect_used, clippy::unwrap_used, reason = "test assertions")]

use std::io::Read as _;

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// Unwrap with the whole error chain in the panic message.
trait Must<T> {
    fn must(self) -> T;
}

impl<T> Must<T> for litchi_ops::Result<T> {
    fn must(self) -> T {
        match self {
            Ok(value) => value,
            Err(error) => {
                let mut message = format!("{:?}: {error}", error.code());
                let mut source = std::error::Error::source(&error);
                while let Some(cause) = source {
                    message.push_str(" <- ");
                    message.push_str(&cause.to_string());
                    source = cause.source();
                }
                panic!("{message}")
            },
        }
    }
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .must()
        .view()
        .format()
        .must()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn reopen(file: &File, eager: bool) -> File {
    litchi_ops::open_bytes_with(
        file.to_bytes().must(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .must()
}

/// One part of the package.
fn part(file: &File, name: &str) -> String {
    let bytes = file.to_bytes().must();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name(name).unwrap();
    let mut text = String::new();
    entry.read_to_string(&mut text).unwrap();
    text
}

/// A LibreOffice spreadsheet (with `styles.xml`).
fn calc() -> File {
    litchi_ops::open(format!(
        "{}/../../test-data/odf/corpus/calc-two-sheets.ods",
        env!("CARGO_MANIFEST_DIR")
    ))
    .must()
}

fn workbook() -> File {
    let file = litchi_ops::create(FileFormat::Ods).must();
    file.set("/Sheet1/A1", props! { "value" => 1 })
        .must()
        .set("/Sheet1/A2", props! { "value" => 2 })
        .must()
        .set("/Sheet1/A3", props! { "value" => "x" })
        .must()
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==";

#[test]
fn comments_report_their_anchor() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("comment")
                .prop("ref", "A2")
                .prop("text", "Check"),
        )
        .must()
        .file;
    assert_eq!(
        prop(&file, "/Sheet1/comment[1]", "anchoredTo").as_deref(),
        Some("/Sheet1/A2")
    );
}

#[test]
fn hyperlinks_as_elements() {
    let added = workbook()
        .add(
            "/Sheet1/B4",
            NewNode::new("hyperlink")
                .prop("url", "https://example.com/")
                .prop("tooltip", "Open"),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/hyperlink[1]");
    for eager in [false, true] {
        let file = reopen(&added.file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "ref").as_deref(),
            Some("B4")
        );
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "url").as_deref(),
            Some("https://example.com/")
        );
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "tooltip").as_deref(),
            Some("Open")
        );
        assert_eq!(
            prop(&file, "/Sheet1/B4", "link").as_deref(),
            Some("https://example.com/")
        );
        let internal = file
            .add(
                "/Sheet1",
                NewNode::new("hyperlink")
                    .prop("ref", "A3")
                    .prop("location", "Sheet1!C9"),
            )
            .must();
        assert_eq!(internal.path.to_string(), "/sheet[1]/hyperlink[1]");
        assert_eq!(
            prop(&internal.file, "/Sheet1/hyperlink[1]", "location").as_deref(),
            Some("Sheet1!C9")
        );
        assert_eq!(
            prop(&internal.file, "/Sheet1/hyperlink[1]", "display").as_deref(),
            Some("x")
        );
        let edited = internal
            .file
            .set(
                "/Sheet1/hyperlink[2]",
                props! { "display" => "Docs", "url" => "https://example.org/" },
            )
            .must();
        assert_eq!(
            prop(&edited, "/Sheet1/hyperlink[2]", "display").as_deref(),
            Some("Docs")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/B4", "link").as_deref(),
            Some("https://example.org/")
        );
        let moved = edited
            .set("/Sheet1/hyperlink[2]", props! { "ref" => "D1" })
            .must();
        assert_eq!(
            prop(&moved, "/Sheet1/hyperlink[1]", "ref").as_deref(),
            Some("D1")
        );
        let removed = moved
            .remove("/Sheet1/hyperlink[1]")
            .must()
            .remove("/Sheet1/hyperlink[1]")
            .must();
        assert_eq!(removed.query("hyperlink").must().count(), 0);
        let error = removed
            .add(
                "/Sheet1",
                NewNode::new("hyperlink")
                    .prop("ref", "A1:B2")
                    .prop("url", "https://example.com/"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[test]
fn pictures_and_shapes() {
    let before = part(&workbook(), "content.xml");
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("alt", "Logo")
                .prop("title", "Brand")
                .prop("x", "1cm")
                .prop("y", "2cm")
                .prop("width", "3cm")
                .prop("height", "1.5cm"),
        )
        .must();
    assert_eq!(file.path.to_string(), "/sheet[1]/picture[1]");
    let file = file
        .file
        .add(
            "/Sheet1",
            NewNode::new("shape")
                .prop("geometry", "ellipse")
                .prop("text", "Hello")
                .prop("fill", "#FF0000")
                .prop("line", "#0000FF:2")
                .prop("width", "4cm")
                .prop("height", "2cm"),
        )
        .must()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "alt").as_deref(),
            Some("Logo")
        );
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "title").as_deref(),
            Some("Brand")
        );
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "x").as_deref(),
            Some("360000emu")
        );
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "contentType").as_deref(),
            Some("image/png")
        );
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "anchorMode").as_deref(),
            Some("absolute")
        );
        assert_eq!(
            prop(&file, "/Sheet1/shape[1]", "geometry").as_deref(),
            Some("ellipse")
        );
        assert_eq!(
            prop(&file, "/Sheet1/shape[1]", "text").as_deref(),
            Some("Hello")
        );
        assert_eq!(
            prop(&file, "/Sheet1/shape[1]", "fill").as_deref(),
            Some("#FF0000")
        );
        assert_eq!(
            prop(&file, "/Sheet1/shape[1]", "line").as_deref(),
            Some("#0000FF:2")
        );
        assert_eq!(
            prop(&file, "/Sheet1/shape[1]", "name").as_deref(),
            Some("Shape 1")
        );
        let edited = file
            .set(
                "/Sheet1/picture[1]",
                props! { "width" => "6cm", "name" => "Logo", "alt" => "New" },
            )
            .must();
        assert_eq!(
            prop(&edited, "/Sheet1/picture[1]", "width").as_deref(),
            Some("2160000emu")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/picture[1]", "name").as_deref(),
            Some("Logo")
        );
        let error = edited
            .set("/Sheet1/shape[1]", props! { "rotation" => 45 })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let error = edited
            .set("/Sheet1/shape[1]", props! { "x" => 2 })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let removed = edited
            .remove("/Sheet1/picture[1]")
            .must()
            .remove("/Sheet1/shape[1]")
            .must();
        assert_eq!(part(&removed, "content.xml"), before);
        assert_eq!(removed.query("picture").must().count(), 0);
    }
}

#[test]
fn conditional_formats_of_every_kind() {
    let file = calc();
    let before = part(&file, "content.xml");
    let styles = part(&file, "styles.xml");
    let added = file
        .add(
            "/sheet1/A1:A10",
            NewNode::new("cf")
                .prop("type", "cellIs")
                .prop("operator", "greaterThan")
                .prop("value", "B1")
                .prop("fill", "#FFC7CE")
                .prop("font.bold", true),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/cf[1]");
    let content = part(&added.file, "content.xml");
    assert!(content.contains("calcext:value=\"&gt;[.B1]\""), "{content}");
    assert!(part(&added.file, "styles.xml").contains("Litchi_CF_1"));
    let mut file = added.file;
    for (kind, extra) in [
        ("formula", vec![("formula", "MOD(A1,2)=1")]),
        ("containsText", vec![("text", "x")]),
        ("top10", vec![("rank", "3")]),
        ("belowAverage", vec![]),
        ("duplicateValues", vec![]),
        ("containsBlanks", vec![]),
        ("containsErrors", vec![]),
        ("timePeriod", vec![("period", "last7Days")]),
        ("colorScale", vec![("midColor", "#FFFF00")]),
        ("dataBar", vec![("color", "#00FF00"), ("minLength", "10")]),
        ("iconSet", vec![("iconset", "3Arrows")]),
    ] {
        let mut node = NewNode::new("cf").prop("type", kind).prop("ref", "C1:C5");
        for (key, value) in extra {
            node = node.prop(key, value);
        }
        file = file.add("/sheet1", node).must().file;
    }
    for eager in [false, true] {
        let file = reopen(&file, eager);
        let types: Vec<String> = (1..=12)
            .map(|index| prop(&file, &format!("/sheet1/cf[{index}]"), "type").unwrap())
            .collect();
        assert_eq!(
            types,
            [
                "cellIs",
                "formula",
                "containsText",
                "topN",
                "aboveAverage",
                "duplicateValues",
                "containsBlanks",
                "containsErrors",
                "timePeriod",
                "colorScale",
                "dataBar",
                "iconSet"
            ]
        );
        assert_eq!(prop(&file, "/sheet1/cf[1]", "value").as_deref(), Some("B1"));
        assert_eq!(
            prop(&file, "/sheet1/cf[1]", "operator").as_deref(),
            Some("greaterThan")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[1]", "ref").as_deref(),
            Some("A1:A10")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[1]", "fill").as_deref(),
            Some("#FFC7CE")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[1]", "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[2]", "formula").as_deref(),
            Some("MOD(A1,2)=1")
        );
        assert_eq!(prop(&file, "/sheet1/cf[3]", "text").as_deref(), Some("x"));
        assert_eq!(prop(&file, "/sheet1/cf[4]", "rank").as_deref(), Some("3"));
        assert_eq!(
            prop(&file, "/sheet1/cf[5]", "aboveAverage").as_deref(),
            Some("false")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[9]", "period").as_deref(),
            Some("last7Days")
        );
        assert!(prop(&file, "/sheet1/cf[10]", "midColor").is_some());
        assert_eq!(
            prop(&file, "/sheet1/cf[11]", "minLength").as_deref(),
            Some("10")
        );
        assert_eq!(
            prop(&file, "/sheet1/cf[12]", "iconset").as_deref(),
            Some("3Arrows")
        );
        let edited = file
            .set(
                "/sheet1/cf[1]",
                props! { "value" => "5", "ref" => "A1:A4", "font.italic" => true },
            )
            .must();
        assert_eq!(
            prop(&edited, "/sheet1/cf[1]", "value").as_deref(),
            Some("5")
        );
        assert_eq!(
            prop(&edited, "/sheet1/cf[1]", "ref").as_deref(),
            Some("A1:A4")
        );
        assert_eq!(
            prop(&edited, "/sheet1/cf[1]", "font.italic").as_deref(),
            Some("true")
        );
        let error = edited
            .set("/sheet1/cf[1]", props! { "stopIfTrue" => true })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let mut removed = edited;
        for _ in 0..12 {
            removed = removed.remove("/sheet1/cf[1]").must();
        }
        assert_eq!(part(&removed, "content.xml"), before);
        assert_eq!(part(&removed, "styles.xml"), styles);
    }
}

#[test]
fn validations_bind_cells() {
    let added = workbook()
        .add(
            "/Sheet1/B2:C4",
            NewNode::new("dataValidation")
                .prop("type", "whole")
                .prop("operator", "between")
                .prop("formula1", "1")
                .prop("formula2", "10")
                .prop("prompt", "Pick")
                .prop("errorStyle", "warning"),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/dataValidation[1]");
    let file = added
        .file
        .add(
            "/Sheet1",
            NewNode::new("dataValidation")
                .prop("ref", "A1:A3")
                .prop("type", "list")
                .prop("formula1", "\"a,b,c\""),
        )
        .must()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "ref").as_deref(),
            Some("B2:C4")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "type").as_deref(),
            Some("whole")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "formula2").as_deref(),
            Some("10")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "prompt").as_deref(),
            Some("Pick")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "errorStyle").as_deref(),
            Some("warning")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[2]", "formula1").as_deref(),
            Some("\"a,b,c\"")
        );
        // Cell values survive the binding.
        assert_eq!(prop(&file, "/Sheet1/A3", "type").as_deref(), Some("String"));
        // The worksheet family refuses a cell write that would drop the
        // binding (record 0327 follow-up), so the validation stays whole.
        let error = file.set("/Sheet1/A2", props! { "value" => 7 }).unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let edited = file
            .set(
                "/Sheet1/dataValidation[1]",
                props! { "ref" => "D1:D2", "operator" => "greaterThan", "formula1" => "B1" },
            )
            .must();
        assert_eq!(
            prop(&edited, "/Sheet1/dataValidation[1]", "ref").as_deref(),
            Some("D1:D2")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/dataValidation[1]", "formula1").as_deref(),
            Some("B1")
        );
        let error = edited
            .add(
                "/Sheet1/A2",
                NewNode::new("dataValidation")
                    .prop("type", "decimal")
                    .prop("operator", "lessThan")
                    .prop("formula1", "3"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        let removed = edited
            .remove("/Sheet1/dataValidation[1]")
            .must()
            .remove("/Sheet1/dataValidation[1]")
            .must();
        assert_eq!(removed.query("dataValidation").must().count(), 0);
        assert!(!part(&removed, "content.xml").contains("content-validation"));
    }
}

#[test]
fn auto_filter_and_page_breaks() {
    let file = workbook();
    let before = part(&file, "content.xml");
    let added = file
        .add(
            "/Sheet1",
            NewNode::new("autofilter").prop("range", "A1:C20"),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/autofilter");
    let file = added
        .file
        .add("/Sheet1", NewNode::new("rowbreak").prop("row", 2))
        .must()
        .file
        .add("/Sheet1", NewNode::new("colbreak").prop("col", "B"))
        .must()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/autofilter", "range").as_deref(),
            Some("A1:C20")
        );
        assert_eq!(
            prop(&file, "/Sheet1/rowbreak[1]", "row").as_deref(),
            Some("2")
        );
        assert_eq!(
            prop(&file, "/Sheet1/colbreak[1]", "col").as_deref(),
            Some("2")
        );
        let edited = file
            .set("/Sheet1/autofilter", props! { "range" => "A1:B5" })
            .must()
            .set("/Sheet1/rowbreak[1]", props! { "row" => 3 })
            .must();
        assert_eq!(
            prop(&edited, "/Sheet1/autofilter", "range").as_deref(),
            Some("A1:B5")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/rowbreak[1]", "row").as_deref(),
            Some("3")
        );
        let error = edited
            .set("/Sheet1/rowbreak[1]", props! { "min" => 1 })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let removed = edited
            .remove("/Sheet1/autofilter")
            .must()
            .remove("/Sheet1/rowbreak[1]")
            .must()
            .remove("/Sheet1/colbreak[1]")
            .must();
        assert_eq!(removed.query("rowbreak").must().count(), 0);
        assert_eq!(removed.query("autofilter").must().count(), 0);
        let content = part(&removed, "content.xml");
        assert!(!content.contains("database-range"));
        assert!(!content.contains("fo:break-before=\"page\""));
        if before.len() == content.len() {
            assert_eq!(content, before);
        }
    }
}

#[test]
fn queries_list_features() {
    let file = workbook()
        .add("/Sheet1", NewNode::new("autofilter").prop("range", "A1:A3"))
        .must()
        .file
        .add(
            "/Sheet1/A1:A3",
            NewNode::new("databar").prop("color", "#FF0000"),
        )
        .must()
        .file;
    assert_eq!(file.query("cf").must().count(), 1);
    assert_eq!(file.query("autofilter").must().count(), 1);
    assert_eq!(
        prop(&file, "/Sheet1/cf[1]", "type").as_deref(),
        Some("dataBar")
    );
}
