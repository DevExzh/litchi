//! XLSX P0 residue (record 0338): the cell `type`, `add` of a cell, and
//! rich-text runs (`/<sheet>/<A1>/run[N]`), each read back on both owners.
#![cfg(feature = "xlsx")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, PropValue, SourceMode, props};

fn workbook() -> File {
    litchi_ops::create(FileFormat::Xlsx).unwrap()
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

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
}

#[test]
fn the_cell_type_converts_the_value() {
    let file = workbook()
        .set(
            "/Sheet1/A1",
            props! { "value" => "0.00E+00", "type" => "string" },
        )
        .unwrap()
        .set(
            "/Sheet1/A2",
            props! { "value" => "1234.5", "type" => "number" },
        )
        .unwrap()
        .set("/Sheet1/A3", props! { "value" => "yes", "type" => "bool" })
        .unwrap()
        .set(
            "/Sheet1/A4",
            props! { "value" => "2024-05-01", "type" => "date" },
        )
        .unwrap()
        .set(
            "/Sheet1/A5",
            props! { "value" => "#N/A", "type" => "error" },
        )
        .unwrap();
    for eager in [true, false] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/Sheet1/A1", "type").as_deref(),
            Some("SharedString")
        );
        assert_eq!(text(&file, "/Sheet1/A1").as_deref(), Some("0.00E+00"));
        assert_eq!(prop(&file, "/Sheet1/A2", "type").as_deref(), Some("Number"));
        assert_eq!(text(&file, "/Sheet1/A2").as_deref(), Some("1234.5"));
        assert_eq!(
            prop(&file, "/Sheet1/A3", "type").as_deref(),
            Some("Boolean")
        );
        assert_eq!(prop(&file, "/Sheet1/A4", "type").as_deref(), Some("Date"));
        assert_eq!(prop(&file, "/Sheet1/A5", "type").as_deref(), Some("Error"));
    }
    // `type` alone converts the stored value.
    let retyped = file
        .set("/Sheet1/A2", props! { "type" => "string" })
        .unwrap();
    assert_eq!(
        prop(&retyped, "/Sheet1/A2", "type").as_deref(),
        Some("SharedString")
    );
    let back = retyped
        .set("/Sheet1/A2", props! { "type" => "number" })
        .unwrap();
    assert_eq!(prop(&back, "/Sheet1/A2", "type").as_deref(), Some("Number"));
    // Refusals: a non-number, a formula, a value that is no boolean.
    for (key, value, ty, code) in [
        ("value", "abc", "number", Code::InvalidValue),
        ("value", "maybe", "boolean", Code::InvalidValue),
        ("formula", "1+1", "number", Code::InvalidArgument),
        ("value", "x", "currency", Code::InvalidValue),
    ] {
        let props = litchi_ops::Props::new().with(key, value).with("type", ty);
        let error = file.set("/Sheet1/B1", props).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
    }
    // An exact no-op stays one.
    let same = file
        .set(
            "/Sheet1/A2",
            props! { "value" => "1234.5", "type" => "number" },
        )
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn a_cell_is_added_under_a_sheet_a_row_or_a_cell() {
    let file = workbook();
    let added = file
        .add(
            "/Sheet1",
            NewNode::new("cell").prop("ref", "B2").prop("value", "x"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/B2");
    let file = added.file;
    assert_eq!(text(&file, "/Sheet1/B2").as_deref(), Some("x"));
    // No `ref`: the first empty cell of row 1, then of the row.
    let first = file
        .add("/Sheet1", NewNode::new("cell").prop("value", "a"))
        .unwrap();
    assert_eq!(first.path.to_string(), "/sheet[1]/A1");
    let next = first
        .file
        .add("/Sheet1", NewNode::new("cell").prop("value", "b"))
        .unwrap();
    assert_eq!(next.path.to_string(), "/sheet[1]/B1");
    let in_row = next
        .file
        .add("/Sheet1/row[2]", NewNode::new("cell").prop("value", "c"))
        .unwrap();
    assert_eq!(in_row.path.to_string(), "/sheet[1]/A2");
    let at_cell = in_row
        .file
        .add("/Sheet1/D4", NewNode::new("cell").prop("formula", "1+1"))
        .unwrap();
    assert_eq!(at_cell.path.to_string(), "/sheet[1]/D4");
    let file = at_cell.file;
    assert_eq!(prop(&file, "/Sheet1/D4", "formula").as_deref(), Some("1+1"));
    // Refusals: an occupied cell, a disagreeing ref, no property, a bad ref.
    let error = file
        .add(
            "/Sheet1",
            NewNode::new("cell").prop("ref", "B2").prop("value", "y"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add(
            "/Sheet1/D4",
            NewNode::new("cell").prop("ref", "E4").prop("value", "y"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add(
            "/Sheet1/row[3]",
            NewNode::new("cell").prop("ref", "E4").prop("value", "y"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add("/Sheet1", NewNode::new("cell").prop("ref", "Z9"))
        .unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    let error = file
        .add(
            "/Sheet1",
            NewNode::new("cell").prop("ref", "nowhere").prop("value", 1),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    // The added cell takes the formatting keys too.
    let styled = file
        .add(
            "/Sheet1",
            NewNode::new("cell")
                .prop("address", "F6")
                .prop("value", "007")
                .prop("type", "string")
                .prop("font.bold", true),
        )
        .unwrap()
        .file;
    assert_eq!(
        prop(&styled, "/Sheet1/F6", "font.bold").as_deref(),
        Some("true")
    );
    assert_eq!(text(&styled, "/Sheet1/F6").as_deref(), Some("007"));
}

const RUNS: &str = r#"[{"text":"Bold + Red  ","bold":true,"color":"C00000"},{"text":"Italic + Blue","italic":true,"color":"2E75B6"},{"text":"  Normal"}]"#;

#[test]
fn rich_text_runs_are_added_read_set_and_removed() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("cell")
                .prop("ref", "A3")
                .prop("type", "richtext")
                .prop("runs", RUNS),
        )
        .unwrap()
        .file;
    for eager in [true, false] {
        let file = reopen(&file, eager);
        assert_eq!(
            text(&file, "/Sheet1/A3").as_deref(),
            Some("Bold + Red  Italic + Blue  Normal")
        );
        let run = file.get("/Sheet1/A3/run[1]").unwrap();
        assert_eq!(run.view().path().to_string(), "/sheet[1]/A3/run[1]");
        assert_eq!(
            text(&file, "/Sheet1/A3/run[1]").as_deref(),
            Some("Bold + Red  ")
        );
        assert_eq!(
            prop(&file, "/Sheet1/A3/run[1]", "bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(&file, "/Sheet1/A3/r[2]", "color").as_deref(),
            Some("#2E75B6")
        );
        assert_eq!(prop(&file, "/Sheet1/A3/run[3]", "bold"), None);
        assert_eq!(
            file.get("/Sheet1/A3/run[4]").unwrap_err().code(),
            Code::PathNotFound
        );
        let runs: Vec<_> = file
            .query("run")
            .unwrap()
            .map(|node| node.unwrap().view().path().to_string())
            .collect();
        assert_eq!(
            runs,
            [
                "/sheet[1]/A3/run[1]",
                "/sheet[1]/A3/run[2]",
                "/sheet[1]/A3/run[3]"
            ]
        );
    }
    // Set a run's text and font; the other runs keep theirs.
    let set = file
        .set(
            "/Sheet1/A3/run[2]",
            props! { "text" => "Under", "underline" => "double", "size" => "14pt", "italic" => PropValue::Null },
        )
        .unwrap();
    assert_eq!(
        text(&set, "/Sheet1/A3").as_deref(),
        Some("Bold + Red  Under  Normal")
    );
    assert_eq!(
        prop(&set, "/Sheet1/A3/run[2]", "underline").as_deref(),
        Some("double")
    );
    assert_eq!(prop(&set, "/Sheet1/A3/run[2]", "italic"), None);
    assert_eq!(
        prop(&set, "/Sheet1/A3/run[2]", "color").as_deref(),
        Some("#2E75B6")
    );
    assert_eq!(
        prop(&set, "/Sheet1/A3/run[1]", "bold").as_deref(),
        Some("true")
    );
    // Add a run: appended, or at a position.
    let added = set
        .add(
            "/Sheet1/A3",
            NewNode::new("run")
                .prop("text", "!")
                .prop("superscript", true),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/A3/run[4]");
    assert_eq!(
        prop(&added.file, "/Sheet1/A3/run[4]", "superscript").as_deref(),
        Some("true")
    );
    // A plain text cell becomes rich: its text is the first run.
    let plain = workbook()
        .set("/Sheet1/B1", props! { "value" => "Hello" })
        .unwrap()
        .add(
            "/Sheet1/B1",
            NewNode::new("run")
                .prop("text", " world")
                .prop("bold", true),
        )
        .unwrap()
        .file;
    assert_eq!(text(&plain, "/Sheet1/B1").as_deref(), Some("Hello world"));
    assert_eq!(text(&plain, "/Sheet1/B1/run[1]").as_deref(), Some("Hello"));
    assert_eq!(
        prop(&plain, "/Sheet1/B1/run[2]", "bold").as_deref(),
        Some("true")
    );
    // Remove a run; as OfficeCLI, one unformatted run left is plain text,
    // and no run left is empty text.
    let removed = added.file.remove("/Sheet1/A3/run[1]").unwrap();
    assert_eq!(
        text(&removed, "/Sheet1/A3").as_deref(),
        Some("Under  Normal!")
    );
    let single = workbook()
        .set(
            "/Sheet1/C1",
            props! { "type" => "richtext", "runs" => r#"[{"text":"only","bold":true},{"text":"x"}]"# },
        )
        .unwrap();
    let one = single.remove("/Sheet1/C1/run[1]").unwrap();
    assert_eq!(text(&one, "/Sheet1/C1").as_deref(), Some("x"));
    assert_eq!(
        one.get("/Sheet1/C1/run[1]").unwrap_err().code(),
        Code::PathNotFound
    );
    let bold = single.remove("/Sheet1/C1/run[2]").unwrap();
    assert_eq!(
        prop(&bold, "/Sheet1/C1/run[1]", "bold").as_deref(),
        Some("true")
    );
    let none = bold.remove("/Sheet1/C1/run[1]").unwrap();
    assert_eq!(text(&none, "/Sheet1/C1").as_deref(), Some(""));
    // Refusals: an unknown run key, runs that are no JSON array, runs with
    // a value, a non-richtext type with runs.
    for props in [
        litchi_ops::Props::new().with("runs", r#"[{"text":"x","glow":true}]"#),
        litchi_ops::Props::new().with("runs", "{}"),
        litchi_ops::Props::new().with("runs", "[]"),
        litchi_ops::Props::new()
            .with("runs", RUNS)
            .with("value", "x"),
        litchi_ops::Props::new()
            .with("runs", RUNS)
            .with("type", "number"),
    ] {
        assert!(workbook().set("/Sheet1/D1", props).is_err());
    }
    let error = file
        .set("/Sheet1/A3/run[1]", props! { "glow" => true })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn capabilities_claim_the_run_element_and_cell_add() {
    let report = litchi_ops::capabilities::Capabilities::of(FileFormat::Xlsx);
    let run = report
        .elements()
        .find(|element| element.name() == "run")
        .unwrap();
    let operations = run.operations();
    assert!(operations.get && operations.set && operations.add && operations.remove);
    assert!(run.gap().is_none());
    for key in ["bold", "italic", "color", "size", "font", "underline"] {
        let property = run.property(key).unwrap();
        assert!(property.get() && property.set() && property.add(), "{key}");
    }
    let cell = report
        .elements()
        .find(|element| element.name() == "cell")
        .unwrap();
    assert!(cell.operations().add && !cell.operations().remove);
    assert!(cell.property("type").unwrap().set());
    assert!(cell.property("ref").unwrap().add());
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";
const JPEG: &str = "data:image/jpeg;base64,/9j/4AAQSkZJRgA=";

#[test]
fn a_pictures_source_is_replaced() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("picture")
                .prop("src", PNG)
                .prop("x", 1)
                .prop("y", 1),
        )
        .unwrap()
        .file;
    let before = file.to_bytes().unwrap();
    assert_eq!(
        prop(&file, "/Sheet1/picture[1]", "contentType").as_deref(),
        Some("image/png")
    );
    let replaced = file
        .set("/Sheet1/picture[1]", props! { "src" => JPEG })
        .unwrap();
    assert_eq!(file.to_bytes().unwrap(), before, "the input is unchanged");
    for eager in [true, false] {
        let replaced = reopen(&replaced, eager);
        assert_eq!(
            prop(&replaced, "/Sheet1/picture[1]", "contentType").as_deref(),
            Some("image/jpeg")
        );
        assert_eq!(
            prop(&replaced, "/Sheet1/picture[1]", "fileSize").as_deref(),
            Some("11")
        );
        // The geometry is kept.
        assert_eq!(
            prop(&replaced, "/Sheet1/picture[1]", "x"),
            prop(&file, "/Sheet1/picture[1]", "x")
        );
    }
    // The old media part is gone from the package.
    let bytes = replaced.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let media: Vec<String> = (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .unwrap()
                .name()
                .unwrap()
                .into_owned()
        })
        .filter(|name| name.starts_with("xl/media/"))
        .collect();
    assert_eq!(media.len(), 1, "{media:?}");
    assert!(media[0].ends_with(".jpeg"), "{media:?}");
    // `path` is the alias; the same image is an exact no-op.
    let again = replaced
        .set("/Sheet1/picture[1]", props! { "path" => JPEG })
        .unwrap();
    assert_eq!(again.to_bytes().unwrap(), replaced.to_bytes().unwrap());
    // A non-image is refused.
    let error = replaced
        .set(
            "/Sheet1/picture[1]",
            props! { "src" => "data:image/png;base64,AAAA" },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // Together with geometry, in one edit.
    let moved = file
        .set(
            "/Sheet1/picture[1]",
            props! { "src" => JPEG, "width" => "3cm" },
        )
        .unwrap();
    assert_eq!(
        prop(&moved, "/Sheet1/picture[1]", "contentType").as_deref(),
        Some("image/jpeg")
    );
}

fn part(file: &File, name: &str) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name(name).unwrap();
    let mut text = String::new();
    std::io::Read::read_to_string(&mut entry, &mut text).unwrap();
    text
}

#[test]
fn chart_series_values_come_from_another_sheet() {
    let mut file = workbook();
    for (cell, value) in [("B2", "120"), ("B3", "135"), ("B4", "148")] {
        file = file
            .set(&format!("/Sheet1/{cell}"), props! { "value" => value })
            .unwrap();
    }
    let file = file
        .add("/", NewNode::new("sheet").prop("name", "1-Line Charts"))
        .unwrap()
        .file;
    let added = file
        .add(
            "/1-Line Charts",
            NewNode::new("chart")
                .prop("chartType", "line")
                .prop("series1.name", "East")
                .prop("series1.values", "Sheet1!B2:B4")
                .prop("series2", "West:1,2,3"),
        )
        .unwrap();
    let file = added.file;
    let xml = part(&file, "xl/charts/chart1.xml");
    assert!(xml.contains("<c:f>Sheet1!$B$2:$B$4</c:f>"), "{xml}");
    for eager in [true, false] {
        let file = reopen(&file, eager);
        assert_eq!(
            prop(&file, "/1-Line Charts/chart[1]/series[1]", "values").as_deref(),
            Some("120,135,148")
        );
        assert_eq!(
            prop(&file, "/1-Line Charts/chart[1]/series[2]", "values").as_deref(),
            Some("1,2,3")
        );
    }
    // A quoted sheet name, on the chart's own sheet, and refusals.
    let local = file
        .add(
            "/Sheet1",
            NewNode::new("chart").prop("series1.values", "'Sheet1'!$B$2:$B$3"),
        )
        .unwrap()
        .file;
    assert!(part(&local, "xl/charts/chart2.xml").contains("<c:f>Sheet1!$B$2:$B$3</c:f>"));
    let error = file
        .add(
            "/Sheet1",
            NewNode::new("chart").prop("series1.values", "Nowhere!A1:A3"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .add(
            "/Sheet1",
            NewNode::new("chart").prop("series1.values", "Sheet1!A1:B3"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .add(
            "/Sheet1",
            NewNode::new("chart")
                .prop("chartType", "scatter")
                .prop("series1.values", "Sheet1!B2:B3"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}
