//! Unified reads of OpenDocument spreadsheets (record 0277): path
//! resolution, queries, node views and the view modes over packaged and
//! flat ODS, on both open routes.

#![cfg(feature = "ods")]
#![allow(clippy::expect_used, clippy::unwrap_used, reason = "test assertions")]

use std::path::PathBuf;

use litchi_core::FileFormat;
use litchi_ops::capabilities::Verb;
use litchi_ops::view::Mode;
use litchi_ops::{Code, File};

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn open(relative: &str) -> File {
    litchi_ops::open(fixture(relative)).expect("open ODS fixture")
}

fn open_bytes(relative: &str) -> File {
    litchi_ops::open_bytes(std::fs::read(fixture(relative)).expect("read")).expect("open bytes")
}

/// Every match of `selector` as `path=text`, one string.
fn listing(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .expect("query compiles")
        .map(|node| {
            let node = node.expect("node");
            let text = node.view().text().unwrap().map(|text| text.into_owned());
            format!("{}={}", node.path(), text.unwrap_or_default())
        })
        .collect()
}

#[test]
fn cells_resolve_by_name_position_and_coordinate() {
    let file = open("odf/corpus/calc-two-sheets.ods");
    assert_eq!(file.format(), Some(FileFormat::Ods));
    let cell = file.get("/sheet1/A1").expect("by name");
    assert_eq!(cell.kind(), "cell");
    assert_eq!(cell.path().to_string(), "/sheet[1]/A1");
    assert_eq!(cell.view().text().unwrap().as_deref(), Some("col A"));

    let by_position = file.get("/sheet[2]/D3").expect("by position");
    assert_eq!(by_position.view().text().unwrap().as_deref(), Some("3"));
    let json = serde_json::to_value(by_position.view()).unwrap();
    assert_eq!(json["format"]["type"], "Number");

    // A coordinate without a stored run resolves as an empty cell.
    let empty = file.get("/sheet1/Z99").expect("virtual cell");
    let json = serde_json::to_value(empty.view()).unwrap();
    assert_eq!(json["format"]["empty"], true);

    // Excel notation reaches the same node.
    assert_eq!(
        file.get("sheet2!B2").unwrap().path().to_string(),
        "/sheet[2]/B2"
    );
    assert_eq!(
        file.find("/missing/A1").unwrap().map(|node| node.kind()),
        None
    );
    assert_eq!(
        file.get("/missing/A1").unwrap_err().code(),
        Code::PathNotFound
    );
}

#[test]
fn rows_columns_and_ranges_list_their_content_cells() {
    let file = open("odf/corpus/calc-two-sheets.ods");
    let row = file.get("/sheet1/row[3]").expect("row");
    assert_eq!(
        row.view().text().unwrap().as_deref(),
        Some("a text another")
    );
    let children: Vec<String> = row
        .children()
        .map(|child| child.unwrap().path().to_string())
        .collect();
    assert_eq!(children, ["/sheet[1]/A3", "/sheet[1]/C3"]);

    let column = file.get("/sheet2/col[D]").expect("column");
    assert_eq!(
        column.view().text().unwrap().as_deref(),
        Some("col D 3 another")
    );

    let range = file.get("/sheet1/A1:B2").expect("range");
    let cells: Vec<String> = range
        .children()
        .map(|child| child.unwrap().path().to_string())
        .collect();
    assert_eq!(
        cells,
        [
            "/sheet[1]/A1",
            "/sheet[1]/B1",
            "/sheet[1]/A2",
            "/sheet[1]/B2"
        ]
    );
    assert_eq!(
        listing(&file, "row")[..3],
        [
            "/sheet[1]/row[1]=col A col B col C",
            "/sheet[1]/row[2]=1 2 3",
            "/sheet[1]/row[3]=a text another"
        ]
    );
}

#[test]
fn queries_filter_by_type_value_and_formula() {
    let file = open("odf/corpus/calc-formulas.ods");
    assert_eq!(
        listing(&file, "cell[formula]"),
        ["/sheet[1]/A3=5", "/sheet[1]/A4=7"]
    );
    assert_eq!(
        listing(&file, "cell[value>4]"),
        ["/sheet[1]/A3=5", "/sheet[1]/A4=7"]
    );
    let formula = file.get("/Sheet1/A4").unwrap();
    let json = serde_json::to_value(formula.view()).unwrap();
    assert_eq!(json["format"]["formula"], "of:=SUM([.A1:.A2])");
    assert_eq!(json["format"]["cachedValue"], "7");

    let two = open("odf/corpus/calc-two-sheets.ods");
    assert_eq!(listing(&two, "cell[type=Number]").len(), 6);
    assert_eq!(
        listing(&two, "cell:contains(another)"),
        ["/sheet[1]/C3=another", "/sheet[2]/D4=another"]
    );
    let first = two.query("cell").unwrap().first().unwrap().unwrap();
    assert_eq!(first.path().to_string(), "/sheet[1]/A1");
    assert_eq!(two.query("cell").unwrap().limit(2).count(), 2);
}

#[test]
fn styles_merges_and_named_ranges_are_projected() {
    let styled = open("odf/corpus/calc-cell-styles.ods");
    let cell = styled.get("/Sheet1/B2").unwrap();
    assert_eq!(cell.view().style().unwrap().as_deref(), Some("ce1"));

    let merged = open("odf/ods/grid-padding-package.ods");
    assert_eq!(
        listing(&merged, "mergeCell"),
        ["/sheet[1]/mergeCell[1]=Beachte die Verwendung als Matrixformel!"]
    );
    let anchor = merged.get("/Tabelle1/E4").unwrap();
    let json = serde_json::to_value(anchor.view()).unwrap();
    assert_eq!(json["format"]["merge"], "E4:E6");
    assert_eq!(json["format"]["mergeAnchor"], true);
    let covered = serde_json::to_value(merged.get("/Tabelle1/E5").unwrap().view()).unwrap();
    assert_eq!(covered["format"]["merge"], "E4:E6");
    assert!(covered["format"].get("mergeAnchor").is_none());
    let range = serde_json::to_value(merged.get("/Tabelle1/E4:E6").unwrap().view()).unwrap();
    assert_eq!(range["format"]["merge"], true);

    let named = open("odfdo/tests/samples/simple_table_named_range.ods");
    let name = named.get("/namedrange[nr_6]").expect("by name");
    let json = serde_json::to_value(name.view()).unwrap();
    assert_eq!(json["format"]["ref"], "$Example1.$D$3:.$F$4");
    assert_eq!(json["format"]["scope"], "workbook");
    assert_eq!(listing(&named, "namedrange").len(), 2);
}

#[test]
fn views_render_each_mode() {
    let file = open("odf/corpus/calc-two-sheets.ods");
    let text = file.view(Mode::Text).unwrap();
    assert!(text.starts_with("=== Sheet: sheet1 ===\n[/sheet[1]/row[1]] A1=col A\tB1=col B"));
    let outline = file.view(Mode::Outline).unwrap();
    assert!(outline.contains("\"sheet2\" (3 rows × 4 cols, 0 formula(s)"));
    let stats = serde_json::to_value(file.view_json(Mode::Stats).unwrap()).unwrap();
    assert_eq!(stats["totalCells"], 16);
    assert_eq!(stats["dataTypeDistribution"]["String"], 10);
    let annotated = file.view(Mode::Annotated).unwrap();
    assert!(annotated.contains("A2: [1] ← Number"));
    assert_eq!(
        file.view(Mode::Html).unwrap_err().code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn gaps_are_typed_per_element() {
    let file = open("odf/corpus/calc-two-sheets.ods");
    assert_eq!(
        file.get("/sheet1/table[1]").unwrap_err().code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        file.get("/sheet1/slicer[1]").unwrap_err().code(),
        Code::UnsupportedPath
    );
    let capabilities = file.capabilities();
    for verb in [Verb::Get, Verb::Find, Verb::Query, Verb::View, Verb::Save] {
        assert!(capabilities.verb(verb).is_supported(), "{verb:?}");
    }
    let cell = capabilities.element("cell").expect("cell element");
    assert!(cell.operations().get && cell.operations().query);
    assert!(capabilities.element("chart").unwrap().gap().is_some());
}

#[test]
fn byte_and_path_routes_answer_alike() {
    for name in [
        "odf/corpus/calc-two-sheets.ods",
        "odf/corpus/calc-formulas.ods",
        "odf/corpus/calc-unicode-chinese.ods",
        "odf/ods/grid-padding-package.ods",
    ] {
        let path = open(name);
        let bytes = open_bytes(name);
        for selector in ["cell", "row", "col", "mergeCell", "namedrange", "sheet"] {
            assert_eq!(
                listing(&path, selector),
                listing(&bytes, selector),
                "{name} {selector}"
            );
        }
        assert_eq!(
            path.view(Mode::Text).unwrap(),
            bytes.view(Mode::Text).unwrap(),
            "{name}"
        );
    }
}

#[test]
fn flat_spreadsheets_open_read_only() {
    let file = open("odf/ods/rich-text-hyperlinks.fods");
    assert_eq!(file.format(), Some(FileFormat::Ods));
    let cells = listing(&file, "cell");
    assert_eq!(cells.len(), 2);
    let comments = file.query("comment").unwrap().next().unwrap();
    assert_eq!(comments.unwrap_err().code(), Code::UnsupportedCapability);
    // Publication of an unchanged flat file is the source, byte for byte.
    let bytes = std::fs::read(fixture("odf/ods/rich-text-hyperlinks.fods")).unwrap();
    assert_eq!(file.to_bytes().unwrap(), bytes);
}

#[test]
fn repeated_queries_are_deterministic() {
    let file = open("odf/ods/grid-padding-package.ods");
    let first = listing(&file, "cell");
    for _ in 0..3 {
        assert_eq!(
            listing(&open("odf/ods/grid-padding-package.ods"), "cell"),
            first
        );
        assert_eq!(listing(&file, "cell"), first);
    }
}
