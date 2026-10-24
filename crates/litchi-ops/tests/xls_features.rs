//! XLS worksheet features through the unified API (record 0327): notes,
//! hyperlinks, pictures, shapes, conditional formats, data validations, the
//! auto-filter and page breaks, each added, read back on both owners, set,
//! queried and removed, with the XLSX vocabulary of record 0320.
#![cfg(feature = "xls")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
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

/// The Workbook stream: what the editor keeps exact (the compound file
/// around it is re-rendered by every XLS write).
fn stream(file: &File) -> Vec<u8> {
    let bytes = file.to_bytes().unwrap();
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}

fn workbook() -> File {
    litchi_ops::create(FileFormat::Xls).unwrap()
}

fn fixture(name: &str) -> File {
    litchi_ops::open(format!(
        "{}/../../test-data/ole/xls/{name}",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap()
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==";

#[test]
fn notes_add_get_set_remove() {
    let file = workbook();
    let before = stream(&file);
    let added = file
        .add(
            "/Sheet1",
            NewNode::new("comment")
                .prop("ref", "B2")
                .prop("text", "Check")
                .prop("author", "Ann"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/comment[1]");
    assert_eq!(stream(&file), before);
    let added = added
        .file
        .add("/Sheet1/A1", NewNode::new("note").prop("text", "First"))
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/comment[1]");
    for eager in [false, true] {
        let file = reopen(&added.file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/comment[1]", "ref").as_deref(),
            Some("A1")
        );
        assert_eq!(
            prop(&file, "/Sheet1/comment[1]", "author").as_deref(),
            Some("Author")
        );
        assert_eq!(
            prop(&file, "/Sheet1/comment[2]", "text").as_deref(),
            Some("Check")
        );
        assert_eq!(
            prop(&file, "/Sheet1/comment[2]", "anchoredTo").as_deref(),
            Some("/Sheet1/B2")
        );
        assert_eq!(file.query("comment").unwrap().count(), 2);
        let edited = file
            .set(
                "/Sheet1/comment[2]",
                props! { "text" => "Done", "ref" => "C3" },
            )
            .unwrap();
        assert_eq!(
            prop(&edited, "/Sheet1/comment[2]", "ref").as_deref(),
            Some("C3")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/comment[2]", "text").as_deref(),
            Some("Done")
        );
        let removed = edited
            .remove("/Sheet1/comment[1]")
            .unwrap()
            .remove("/Sheet1/comment[1]")
            .unwrap();
        assert_eq!(removed.query("comment").unwrap().count(), 0);
        assert_eq!(stream(&removed), before);
    }
    let duplicate_target = added
        .file
        .set("/Sheet1/comment[2]", props! { "ref" => "A1" })
        .unwrap_err();
    assert_eq!(duplicate_target.code(), Code::InvalidValue);
    let error = added
        .file
        .add("/Sheet1", NewNode::new("comment").prop("ref", "B2"))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[test]
fn hyperlinks_as_elements() {
    let file = workbook()
        .add(
            "/Sheet1/A9",
            NewNode::new("hyperlink")
                .prop("url", "https://example.com/")
                .prop("tooltip", "Open"),
        )
        .unwrap();
    assert_eq!(file.path.to_string(), "/sheet[1]/hyperlink[1]");
    let file = file.file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "ref").as_deref(),
            Some("A9")
        );
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "url").as_deref(),
            Some("https://example.com/")
        );
        assert_eq!(
            prop(&file, "/Sheet1/hyperlink[1]", "tooltip").as_deref(),
            Some("Open")
        );
    }
    let file = file
        .add(
            "/Sheet1",
            NewNode::new("hyperlink")
                .prop("ref", "B2")
                .prop("location", "Sheet1!A1")
                .prop("display", "Top"),
        )
        .unwrap()
        .file;
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[2]", "location").as_deref(),
        Some("Sheet1!A1")
    );
    let file = file
        .set("/Sheet1/hyperlink[2]", props! { "display" => "Home" })
        .unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[2]", "display").as_deref(),
        Some("Home")
    );
    // A link on the same cells replaces the existing one.
    let file = file
        .add(
            "/Sheet1",
            NewNode::new("hyperlink")
                .prop("ref", "A9")
                .prop("url", "https://other.test/"),
        )
        .unwrap()
        .file;
    assert_eq!(file.query("hyperlink").unwrap().count(), 2);
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[1]", "url").as_deref(),
        Some("https://other.test/")
    );
    let file = file.remove("/Sheet1/hyperlink[1]").unwrap();
    assert_eq!(file.query("hyperlink").unwrap().count(), 1);
    let file = file.remove("/Sheet1/hyperlink[1]").unwrap();
    assert_eq!(file.query("hyperlink").unwrap().count(), 0);
    let error = workbook()
        .add(
            "/Sheet1",
            NewNode::new("hyperlink")
                .prop("ref", "A1")
                .prop("url", "javascript:alert(1)"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn pictures_and_shapes() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("x", 1)
                .prop("y", 2)
                .prop("width", 3)
                .prop("height", 4)
                .prop("alt", "Logo"),
        )
        .unwrap();
    assert_eq!(file.path.to_string(), "/sheet[1]/picture[1]");
    let file = file
        .file
        .add(
            "/Sheet1",
            NewNode::new("shape")
                .prop("geometry", "roundRect")
                .prop("x", 0)
                .prop("y", 4)
                .prop("width", 2)
                .prop("height", 3)
                .prop("fill", "4472C4")
                .prop("line", "C00000:2")
                .prop("text", "rect")
                .prop("rotation", 30),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        let picture = "/Sheet1/picture[1]";
        assert_eq!(prop(&file, picture, "alt").as_deref(), Some("Logo"));
        assert_eq!(prop(&file, picture, "name").as_deref(), Some("Picture 1"));
        assert_eq!(prop(&file, picture, "x").as_deref(), Some("1"));
        assert_eq!(prop(&file, picture, "height").as_deref(), Some("4"));
        assert_eq!(
            prop(&file, picture, "contentType").as_deref(),
            Some("image/png")
        );
        let shape = "/Sheet1/shape[1]";
        assert_eq!(prop(&file, shape, "geometry").as_deref(), Some("roundRect"));
        assert_eq!(prop(&file, shape, "text").as_deref(), Some("rect"));
        assert_eq!(prop(&file, shape, "fill").as_deref(), Some("#4472C4"));
        assert_eq!(prop(&file, shape, "line").as_deref(), Some("#C00000:2"));
        assert_eq!(prop(&file, shape, "rotation").as_deref(), Some("30"));
        assert_eq!(file.query("shape").unwrap().count(), 1);
        assert_eq!(file.query("picture").unwrap().count(), 1);
    }
    let file = file
        .set(
            "/Sheet1/picture[1]",
            props! { "x" => 5, "alt" => "Mark", "flip" => "h" },
        )
        .unwrap();
    assert_eq!(prop(&file, "/Sheet1/picture[1]", "x").as_deref(), Some("5"));
    assert_eq!(
        prop(&file, "/Sheet1/picture[1]", "alt").as_deref(),
        Some("Mark")
    );
    assert_eq!(
        prop(&file, "/Sheet1/picture[1]", "flip").as_deref(),
        Some("h")
    );
    let file = file
        .remove("/Sheet1/picture[1]")
        .unwrap()
        .remove("/Sheet1/shape[1]")
        .unwrap();
    assert_eq!(file.query("picture").unwrap().count(), 0);
    assert_eq!(file.query("shape").unwrap().count(), 0);
    assert_eq!(stream(&file), stream(&workbook()));
    // Shape text formatting and drawing effects are typed gaps.
    for (key, value) in [("bold", "true"), ("font", "Georgia"), ("shadow", "000000")] {
        let error = workbook()
            .add(
                "/Sheet1",
                NewNode::new("shape").prop("text", "x").prop(key, value),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty, "{key}");
    }
    // Drawings are anchored to cells.
    let error = workbook()
        .add("/Sheet1", NewNode::new("shape").prop("width", "2in"))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn conditional_formats_of_every_kind() {
    let rules: &[&[(&str, &str)]] = &[
        &[
            ("type", "cellIs"),
            ("operator", "between"),
            ("value", "50"),
            ("value2", "70"),
            ("fill", "FFFF00"),
        ],
        &[
            ("type", "containsText"),
            ("text", "error"),
            ("fill", "FF0000"),
        ],
        &[("type", "notContains"), ("text", "error")],
        &[("type", "beginsWith"), ("text", "B")],
        &[("type", "endsWith"), ("text", "E")],
        &[("type", "top10"), ("rank", "3"), ("fill", "00FF00")],
        &[("type", "topPercent"), ("rank", "25"), ("percent", "true")],
        &[("type", "bottom"), ("rank", "3")],
        &[
            ("type", "aboveAverage"),
            ("aboveAverage", "true"),
            ("stdDev", "1"),
            ("equalAverage", "true"),
        ],
        &[("type", "belowAverage")],
        &[
            ("type", "dataBar"),
            ("color", "638EC6"),
            ("min", "auto"),
            ("max", "auto"),
            ("showValue", "true"),
        ],
        &[
            ("type", "colorScale"),
            ("minColor", "F8696B"),
            ("midColor", "FFEB84"),
            ("maxColor", "63BE7B"),
            ("midpoint", "50"),
        ],
        &[
            ("type", "iconSet"),
            ("iconset", "3Arrows"),
            ("reverse", "true"),
        ],
        &[
            ("type", "formula"),
            ("formula", "MOD(A2,2)=1"),
            ("fill", "CCFFFF"),
        ],
        &[("type", "duplicateValues"), ("fill", "FF0000")],
        &[("type", "uniqueValues")],
        &[
            ("type", "dateOccurring"),
            ("period", "thisMonth"),
            ("font.bold", "true"),
        ],
        &[("type", "containsBlanks")],
    ];
    let mut file = workbook();
    for (index, rule) in rules.iter().enumerate() {
        let mut node = NewNode::new("conditionalformatting").prop("ref", "A2:A11");
        for (key, value) in *rule {
            node = node.prop(*key, *value);
        }
        let added = file
            .add("/Sheet1", node)
            .unwrap_or_else(|error| panic!("{rule:?}: {error}"));
        assert_eq!(
            added.path.to_string(),
            format!("/sheet[1]/cf[{}]", index + 1)
        );
        file = added.file;
    }
    let file = file
        .add("/Sheet1", NewNode::new("databar").prop("ref", "B2:B9"))
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(file.query("cf").unwrap().count(), rules.len() + 1);
        assert_eq!(
            prop(&file, "/Sheet1/cf[1]", "type").as_deref(),
            Some("cellIs")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[1]", "value2").as_deref(),
            Some("70")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[1]", "fill").as_deref(),
            Some("#FFFF00")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[2]", "text").as_deref(),
            Some("error")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[2]", "type").as_deref(),
            Some("containsText")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[3]", "type").as_deref(),
            Some("notContainsText")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[6]", "type").as_deref(),
            Some("topN")
        );
        assert_eq!(prop(&file, "/Sheet1/cf[6]", "rank").as_deref(), Some("3"));
        assert_eq!(
            prop(&file, "/Sheet1/cf[7]", "percent").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[8]", "bottom").as_deref(),
            Some("true")
        );
        assert_eq!(prop(&file, "/Sheet1/cf[9]", "stdDev").as_deref(), Some("1"));
        assert_eq!(
            prop(&file, "/Sheet1/cf[10]", "aboveAverage").as_deref(),
            Some("false")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[11]", "color").as_deref(),
            Some("#638EC6")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[12]", "midColor").as_deref(),
            Some("#FFEB84")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[12]", "midpoint").as_deref(),
            Some("50")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[13]", "iconset").as_deref(),
            Some("3Arrows")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[13]", "reverse").as_deref(),
            Some("true")
        );
        // The renderer parenthesizes every binary operation, as XLSB reads.
        assert_eq!(
            prop(&file, "/Sheet1/cf[14]", "formula").as_deref(),
            Some("(MOD(A2,2)=1)")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[15]", "type").as_deref(),
            Some("duplicateValues")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[16]", "type").as_deref(),
            Some("uniqueValues")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[17]", "period").as_deref(),
            Some("thisMonth")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[17]", "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[18]", "type").as_deref(),
            Some("containsBlanks")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[19]", "type").as_deref(),
            Some("dataBar")
        );
    }
    let file = file
        .set(
            "/Sheet1/cf[1]",
            props! { "value" => "40", "fill" => "00FF00", "ref" => "C1:C5" },
        )
        .unwrap();
    assert_eq!(prop(&file, "/Sheet1/cf[1]", "value").as_deref(), Some("40"));
    assert_eq!(
        prop(&file, "/Sheet1/cf[1]", "fill").as_deref(),
        Some("#00FF00")
    );
    assert_eq!(
        prop(&file, "/Sheet1/cf[1]", "ref").as_deref(),
        Some("C1:C5")
    );
    let file = file
        .set("/Sheet1/cf[2]", props! { "ref" => "D1:D3" })
        .unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/cf[2]", "ref").as_deref(),
        Some("D1:D3")
    );
    let file = file.remove("/Sheet1/cf[1]").unwrap();
    assert_eq!(file.query("cf").unwrap().count(), rules.len());
    let error = workbook()
        .add(
            "/Sheet1",
            NewNode::new("cf")
                .prop("type", "dataBar")
                .prop("negativeColor", "FF0000"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    // A formatting color outside the workbook palette is refused.
    let error = workbook()
        .add(
            "/Sheet1",
            NewNode::new("cf")
                .prop("type", "cellIs")
                .prop("value", "1")
                .prop("fill", "FFEB9C"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn validations_autofilter_and_breaks() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("validation")
                .prop("type", "list")
                .prop("ref", "A2:A20")
                .prop("formula1", "Draft,Review"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("validation")
                .prop("type", "whole")
                .prop("ref", "B2:B50")
                .prop("operator", "between")
                .prop("formula1", "1")
                .prop("formula2", "100")
                .prop("errorStyle", "warn")
                .prop("prompt", "1-100"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("autofilter").prop("range", "A1:C10"),
        )
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("rowbreak").prop("row", 20))
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("pagebreak").prop("col", "C"))
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[1]", "formula1").as_deref(),
            Some("\"Draft,Review\"")
        );
        assert_eq!(
            prop(&file, "/Sheet1/validation[2]", "errorStyle").as_deref(),
            Some("warning")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[2]", "prompt").as_deref(),
            Some("1-100")
        );
        assert_eq!(
            prop(&file, "/Sheet1/dataValidation[2]", "formula2").as_deref(),
            Some("100")
        );
        assert_eq!(
            prop(&file, "/Sheet1/autofilter", "range").as_deref(),
            Some("A1:C10")
        );
        assert_eq!(
            prop(&file, "/Sheet1/rowbreak[1]", "row").as_deref(),
            Some("20")
        );
        assert_eq!(
            prop(&file, "/Sheet1/colbreak[1]", "col").as_deref(),
            Some("3")
        );
    }
    let overlap = file
        .add(
            "/Sheet1",
            NewNode::new("validation")
                .prop("type", "custom")
                .prop("ref", "A5")
                .prop("formula1", "ISNUMBER(A5)"),
        )
        .unwrap_err();
    assert_eq!(overlap.code(), Code::InvalidArgument);
    let file = file
        .set("/Sheet1/dataValidation[2]", props! { "formula2" => "50" })
        .unwrap()
        .set("/Sheet1/autofilter", props! { "range" => "A1:D20" })
        .unwrap()
        .set("/Sheet1/rowbreak[1]", props! { "row" => 30 })
        .unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/dataValidation[2]", "formula2").as_deref(),
        Some("50")
    );
    assert_eq!(
        prop(&file, "/Sheet1/autofilter", "range").as_deref(),
        Some("A1:D20")
    );
    assert_eq!(
        prop(&file, "/Sheet1/rowbreak[1]", "row").as_deref(),
        Some("30")
    );
    let file = file
        .remove("/Sheet1/dataValidation[1]")
        .unwrap()
        .remove("/Sheet1/autofilter")
        .unwrap()
        .remove("/Sheet1/rowbreak[1]")
        .unwrap()
        .remove("/Sheet1/colbreak[1]")
        .unwrap();
    assert_eq!(file.query("dataValidation").unwrap().count(), 1);
    assert_eq!(file.query("autofilter").unwrap().count(), 0);
    assert_eq!(file.query("rowbreak").unwrap().count(), 0);
    let error = workbook()
        .add(
            "/Sheet1",
            NewNode::new("autofilter")
                .prop("range", "A1:B5")
                .prop("criteria0.equals", "x"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn features_compose_with_cell_writes() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("cf")
                .prop("type", "cellIs")
                .prop("ref", "A1:A5")
                .prop("value", "1")
                .prop("fill", "FF0000"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("validation")
                .prop("type", "list")
                .prop("ref", "B1:B5")
                .prop("formula1", "a,b"),
        )
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("autofilter").prop("range", "A1:B5"))
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("comment").prop("ref", "C3").prop("text", "x"),
        )
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("shape").prop("text", "box"))
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("rowbreak").prop("row", 5))
        .unwrap()
        .file
        .add(
            "/Sheet1/D4",
            NewNode::new("hyperlink").prop("url", "https://example.test/"),
        )
        .unwrap()
        .file
        .set("/Sheet1/D4", props! { "value" => "Repo" })
        .unwrap()
        .set("/Sheet1/A1", props! { "value" => 3 })
        .unwrap()
        .add(
            "/Sheet1",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("anchor", "F2:H6"),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        for kind in [
            "cf",
            "dataValidation",
            "autofilter",
            "comment",
            "shape",
            "rowbreak",
            "hyperlink",
            "picture",
        ] {
            assert_eq!(file.query(kind).unwrap().count(), 1, "{kind}");
        }
        assert_eq!(
            file.get("/Sheet1/A1")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("3")
        );
        assert_eq!(prop(&file, "/Sheet1/picture[1]", "x").as_deref(), Some("5"));
        assert_eq!(
            prop(&file, "/Sheet1/picture[1]", "width").as_deref(),
            Some("3")
        );
    }
}

#[test]
fn fixture_features_read_and_keep_their_records() {
    let file = fixture("SimpleWithComments.xls");
    assert_eq!(file.query("comment").unwrap().count(), 3);
    let file = fixture("SimpleWithImages.xls");
    assert_eq!(file.query("picture").unwrap().count(), 4);
    assert_eq!(
        prop(&file, "/sheet[1]/picture[1]", "contentType").as_deref(),
        Some("image/jpeg")
    );
    let file = fixture("SimpleWithAutofilter.xls");
    assert_eq!(
        prop(&file, "/sheet[1]/autofilter", "range").as_deref(),
        Some("A1:B4")
    );
    let file = fixture("WithConditionalFormatting.xls");
    assert!(file.query("cf").unwrap().count() > 0);
    let file = fixture("WithHyperlink.xls");
    assert!(file.query("hyperlink").unwrap().count() > 0);
    // An add and a remove keep the Workbook stream byte for byte.
    let before = stream(&file);
    let changed = file
        .add("/sheet[1]", NewNode::new("rowbreak").prop("row", 4))
        .unwrap()
        .file
        .remove("/sheet[1]/rowbreak[1]")
        .unwrap();
    assert_eq!(stream(&changed), before);
}
