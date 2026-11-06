//! XLSX worksheet features through the unified API (record 0320): notes,
//! hyperlinks, pictures, shapes, conditional formats, data validations, the
//! auto-filter and page breaks, each added, read back on both owners, set,
//! queried and removed.
#![cfg(feature = "xlsx")]

use litchi_core::FileFormat;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, EditOptions, File, NewNode, OpenOptions, SourceMode, props};

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

fn workbook() -> File {
    litchi_ops::create(FileFormat::Xlsx).unwrap()
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

#[test]
fn notes_add_get_set_remove() {
    let file = workbook();
    let before = file.to_bytes().unwrap();
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
    assert_eq!(file.to_bytes().unwrap(), before);
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
    }
    let error = added
        .file
        .add("/Sheet1", NewNode::new("comment").prop("ref", "B2"))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[test]
fn note_fonts_and_rich_runs() {
    // Record 0325: `font.*` writes one formatted run, `runs` a rich note.
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("comment")
                .prop("ref", "B2")
                .prop("text", "Bold red")
                .prop("font.bold", true)
                .prop("font.color", "FF0000")
                .prop("font.size", "12pt"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("comment").prop("ref", "C3").prop(
                "runs",
                r#"[{"text":"Alice:","bold":true,"color":"0000FF"},{"text":" check","italic":true,"underline":"double","superscript":true}]"#,
            ),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        let get = |path: &str, key: &str| prop(&file, path, key);
        assert_eq!(
            get("/Sheet1/comment[1]", "text").as_deref(),
            Some("Bold red")
        );
        assert_eq!(
            get("/Sheet1/comment[1]", "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            get("/Sheet1/comment[1]", "font.color").as_deref(),
            Some("#FF0000")
        );
        assert!(get("/Sheet1/comment[1]", "font.size").is_some());
        // Excel's note typeface is not reported.
        assert_eq!(get("/Sheet1/comment[1]", "font.name"), None);
        assert_eq!(get("/Sheet1/comment[1]", "runs"), None);
        assert_eq!(
            get("/Sheet1/comment[2]", "text").as_deref(),
            Some("Alice: check")
        );
        let runs = get("/Sheet1/comment[2]", "runs").unwrap();
        // JSON object ordering depends on serde_json's feature set; the
        // stored text and formatting values are the public contract.
        let runs: serde_json::Value = serde_json::from_str(&runs).unwrap();
        assert_eq!(runs.as_array().unwrap().len(), 2);
        assert_eq!(runs[0]["text"], "Alice:");
        assert_eq!(runs[0]["bold"], true);
        assert_eq!(runs[0]["color"], "#0000FF");
        assert_eq!(runs[1]["text"], " check");
        assert_eq!(runs[1]["italic"], true);
        assert_eq!(runs[1]["underline"], "double");
        assert_eq!(runs[1]["superscript"], true);
        assert_eq!(get("/Sheet1/comment[2]", "font.bold"), None);
        // `font.*` alone keeps the text and rebuilds one run.
        let edited = file
            .set("/Sheet1/comment[2]", props! { "font.italic" => true })
            .unwrap();
        assert_eq!(
            prop(&edited, "/Sheet1/comment[2]", "text").as_deref(),
            Some("Alice: check")
        );
        assert_eq!(
            prop(&edited, "/Sheet1/comment[2]", "font.italic").as_deref(),
            Some("true")
        );
        assert_eq!(prop(&edited, "/Sheet1/comment[2]", "runs"), None);
        // Moving a rich note keeps its runs; an equal set is an exact no-op.
        let moved = file
            .set("/Sheet1/comment[2]", props! { "ref" => "D9" })
            .unwrap();
        assert!(prop(&moved, "/Sheet1/comment[2]", "runs").is_some());
        let same = file
            .set("/Sheet1/comment[2]", props! { "author" => "Author" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        // A new text alone writes plain text.
        let plain = file
            .set("/Sheet1/comment[1]", props! { "text" => "plain" })
            .unwrap();
        assert_eq!(prop(&plain, "/Sheet1/comment[1]", "font.bold"), None);
    }
    for bad in ["not json", "{}", "[1, 2]"] {
        let error = workbook()
            .add(
                "/Sheet1",
                NewNode::new("comment").prop("ref", "A1").prop("runs", bad),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue, "{bad}");
    }
}

#[test]
fn hyperlinks_as_elements_and_cell_properties() {
    let file = workbook()
        .set(
            "/Sheet1/A9",
            props! { "value" => "Repo", "link" => "https://example.com/", "tooltip" => "Open" },
        )
        .unwrap();
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
        assert_eq!(
            file.get("/Sheet1/A9")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Repo")
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
    let file = file.set("/Sheet1/A9", props! { "link" => "none" }).unwrap();
    assert_eq!(file.query("hyperlink").unwrap().count(), 1);
    let file = file.remove("/Sheet1/hyperlink[1]").unwrap();
    assert_eq!(file.query("hyperlink").unwrap().count(), 0);
    let error = workbook()
        .set("/Sheet1/A1", props! { "link" => "javascript:alert(1)" })
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
                .prop("text", "rect")
                .prop("color", "FFFFFF")
                .prop("bold", true)
                .prop("rotation", 30),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        let picture = "/Sheet1/picture[1]";
        assert_eq!(prop(&file, picture, "alt").as_deref(), Some("Logo"));
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
        assert_eq!(prop(&file, shape, "bold").as_deref(), Some("true"));
        assert_eq!(prop(&file, shape, "rotation").as_deref(), Some("30"));
        assert_eq!(file.query("shape").unwrap().count(), 1);
        assert_eq!(file.query("picture").unwrap().count(), 1);
    }
    let file = file
        .set("/Sheet1/picture[1]", props! { "x" => 5, "alt" => "Mark" })
        .unwrap();
    assert_eq!(prop(&file, "/Sheet1/picture[1]", "x").as_deref(), Some("5"));
    assert_eq!(
        prop(&file, "/Sheet1/picture[1]", "alt").as_deref(),
        Some("Mark")
    );
    let file = file
        .remove("/Sheet1/picture[1]")
        .unwrap()
        .remove("/Sheet1/shape[1]")
        .unwrap();
    assert_eq!(file.query("picture").unwrap().count(), 0);
    assert_eq!(file.query("shape").unwrap().count(), 0);
}

#[test]
fn conditional_formats_of_every_kind() {
    let rules: &[&[(&str, &str)]] = &[
        &[
            ("type", "cellIs"),
            ("operator", "between"),
            ("value", "50"),
            ("value2", "70"),
            ("fill", "FFEB9C"),
        ],
        &[
            ("type", "containsText"),
            ("text", "error"),
            ("fill", "FFC7CE"),
        ],
        &[("type", "notContains"), ("text", "error")],
        &[("type", "beginsWith"), ("text", "B")],
        &[("type", "endsWith"), ("text", "E")],
        &[("type", "top10"), ("rank", "3"), ("fill", "C6EFCE")],
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
            ("formula", "ISODD(A2)"),
            ("fill", "BDD7EE"),
        ],
        &[("type", "duplicateValues"), ("fill", "FFC7CE")],
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
            Some("#FFEB9C")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[2]", "text").as_deref(),
            Some("error")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[6]", "type").as_deref(),
            Some("topN")
        );
        assert_eq!(prop(&file, "/Sheet1/cf[6]", "rank").as_deref(), Some("3"));
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
            prop(&file, "/Sheet1/cf[13]", "iconset").as_deref(),
            Some("3Arrows")
        );
        assert_eq!(
            prop(&file, "/Sheet1/cf[14]", "formula").as_deref(),
            Some("ISODD(A2)")
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
}

#[test]
fn removing_everything_restores_the_package() {
    let file = workbook();
    let changed = file
        .add(
            "/Sheet1",
            NewNode::new("comment").prop("ref", "A1").prop("text", "x"),
        )
        .unwrap()
        .file
        .add("/Sheet1", NewNode::new("autofilter").prop("range", "A1:B2"))
        .unwrap()
        .file
        .remove("/Sheet1/comment[1]")
        .unwrap()
        .remove("/Sheet1/autofilter")
        .unwrap();
    assert_eq!(changed.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn sheet_auto_filter_and_shape_styling() {
    let file = workbook()
        .set(
            "/Sheet1",
            props! { "autoFilter" => "A1:B5", "name" => "Data" },
        )
        .unwrap();
    assert_eq!(
        prop(&file, "/Data/autofilter", "range").as_deref(),
        Some("A1:B5")
    );
    let file = file
        .set("/Data", props! { "autoFilter" => "none" })
        .unwrap();
    assert_eq!(file.query("autofilter").unwrap().count(), 0);
    let file = file
        .add(
            "/Data",
            NewNode::new("textbox")
                .prop("anchor", "K25:L28")
                .prop("fill", "accent1")
                .prop("line", "C00000:2:dash")
                .prop("margin", "6")
                .prop("flipBoth", true)
                .prop("align", "center")
                .prop("valign", "top")
                .prop("font", "Georgia")
                .prop("size", "12")
                .prop("text", "inset"),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        let shape = "/Data/shape[1]";
        assert_eq!(prop(&file, shape, "fill").as_deref(), Some("accent1"));
        assert_eq!(prop(&file, shape, "line").as_deref(), Some("#C00000:2"));
        assert_eq!(prop(&file, shape, "margin").as_deref(), Some("6pt"));
        assert_eq!(prop(&file, shape, "flip").as_deref(), Some("both"));
        assert_eq!(prop(&file, shape, "align").as_deref(), Some("center"));
        assert_eq!(prop(&file, shape, "valign").as_deref(), Some("top"));
        assert_eq!(prop(&file, shape, "font").as_deref(), Some("Georgia"));
        assert_eq!(prop(&file, shape, "size").as_deref(), Some("12pt"));
        assert_eq!(prop(&file, shape, "x").as_deref(), Some("10"));
        assert_eq!(prop(&file, shape, "width").as_deref(), Some("2"));
    }
}

#[test]
fn sheet_auto_filter_snapshot_limit_covers_native_patch() {
    let original = workbook();
    let probe = original
        .edit(|transaction| transaction.set("/Sheet1", props! { "autoFilter" => "A1:B5" }))
        .unwrap();
    let charge = probe.diagnostics().snapshot_bytes();
    assert!(charge > 0);

    let strict = SemanticLimits::builder()
        .max_snapshot_bytes(charge - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = original
        .edit_with_options(EditOptions::new().semantic_limits(strict), |transaction| {
            transaction.set("/Sheet1", props! { "autoFilter" => "A1:B5" })
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);

    let exact = SemanticLimits::builder()
        .max_snapshot_bytes(charge)
        .unwrap()
        .build()
        .unwrap();
    let committed = original
        .edit_with_options(EditOptions::new().semantic_limits(exact), |transaction| {
            transaction.set("/Sheet1", props! { "autoFilter" => "A1:B5" })
        })
        .unwrap();
    assert_eq!(
        prop(committed.file(), "/Sheet1/autofilter", "range").as_deref(),
        Some("A1:B5")
    );
}

#[test]
fn combined_sheet_rename_shrinking_and_auto_filter_snapshot_limit_is_exact() {
    let original = workbook();
    let probe = original
        .edit(|transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "Data", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap();
    let charge = probe.diagnostics().snapshot_bytes();
    assert!(charge > 0);

    let strict = SemanticLimits::builder()
        .max_snapshot_bytes(charge - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = original
        .edit_with_options(EditOptions::new().semantic_limits(strict), |transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "Data", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);

    let exact = SemanticLimits::builder()
        .max_snapshot_bytes(charge)
        .unwrap()
        .build()
        .unwrap();
    let committed = original
        .edit_with_options(EditOptions::new().semantic_limits(exact), |transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "Data", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap();
    assert_eq!(
        prop(committed.file(), "/Data/autofilter", "range").as_deref(),
        Some("A1:B5")
    );
}

#[test]
fn combined_sheet_rename_expanding_and_auto_filter_snapshot_limit_is_exact() {
    let original = workbook();
    let probe = original
        .edit(|transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "FilteredDataSheet", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap();
    let charge = probe.diagnostics().snapshot_bytes();
    assert!(charge > 0);

    let strict = SemanticLimits::builder()
        .max_snapshot_bytes(charge - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = original
        .edit_with_options(EditOptions::new().semantic_limits(strict), |transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "FilteredDataSheet", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);

    let exact = SemanticLimits::builder()
        .max_snapshot_bytes(charge)
        .unwrap()
        .build()
        .unwrap();
    let committed = original
        .edit_with_options(EditOptions::new().semantic_limits(exact), |transaction| {
            transaction.set(
                "/Sheet1",
                props! { "name" => "FilteredDataSheet", "autoFilter" => "A1:B5" },
            )
        })
        .unwrap();
    assert_eq!(
        prop(committed.file(), "/FilteredDataSheet/autofilter", "range").as_deref(),
        Some("A1:B5")
    );
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
                .prop("fill", "FFC7CE"),
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
        .set(
            "/Sheet1/D4",
            props! { "value" => "Repo", "link" => "https://example.test/", "font.bold" => true },
        )
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
fn a_new_cell_link_replaces_the_hyperlink_and_none_is_a_no_op_without_one() {
    let file = workbook()
        .set(
            "/Sheet1/A1",
            props! { "link" => "https://a.test/", "tooltip" => "A" },
        )
        .unwrap()
        .set("/Sheet1/A1", props! { "display" => "Shown" })
        .unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[1]", "tooltip").as_deref(),
        Some("A")
    );
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[1]", "display").as_deref(),
        Some("Shown")
    );
    let file = file
        .set("/Sheet1/A1", props! { "link" => "#Sheet1!B2" })
        .unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/hyperlink[1]", "location").as_deref(),
        Some("Sheet1!B2")
    );
    assert_eq!(prop(&file, "/Sheet1/hyperlink[1]", "tooltip"), None);
    let empty = workbook();
    let unchanged = empty
        .set("/Sheet1/B5", props! { "link" => "none" })
        .unwrap();
    assert_eq!(unchanged.to_bytes().unwrap(), empty.to_bytes().unwrap());
    let error = empty
        .set("/Sheet1/B5", props! { "tooltip" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn auto_filter_column_criteria() {
    use litchi_xlsx::auto_filter::{Operator, Payload};
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("autofilter")
                .prop("range", "A1:D20")
                .prop("criteria0.equals", "Red")
                .prop("criteria1.between", "10,20")
                .prop("criteria2.top", "5")
                .prop("criteria3.values", "a, b"),
        )
        .unwrap()
        .file;
    let package = litchi_xlsx::Package::from_bytes(&file.to_bytes().unwrap()).unwrap();
    let filter = package
        .workbook()
        .sheet(0)
        .unwrap()
        .unwrap()
        .auto_filter()
        .unwrap()
        .unwrap();
    let columns = filter.columns();
    assert_eq!(columns.len(), 4);
    match columns[0].payload() {
        Some(Payload::Custom(customs)) => {
            assert_eq!(customs.filters()[0].operator(), Operator::Equal);
            assert_eq!(customs.filters()[0].value(), "Red");
        },
        other => panic!("{other:?}"),
    }
    match columns[1].payload() {
        Some(Payload::Custom(customs)) => {
            assert!(customs.and());
            assert_eq!(customs.filters().len(), 2);
        },
        other => panic!("{other:?}"),
    }
    assert!(matches!(columns[2].payload(), Some(Payload::Top10(_))));
    match columns[3].payload() {
        Some(Payload::Values(values)) => assert_eq!(values.items().len(), 2),
        other => panic!("{other:?}"),
    }
    assert_eq!(
        prop(&file, "/Sheet1/autofilter", "range").as_deref(),
        Some("A1:D20")
    );
    let error = workbook()
        .add(
            "/Sheet1",
            NewNode::new("autofilter")
                .prop("range", "A1:B2")
                .prop("criteria0.near", "x"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}
