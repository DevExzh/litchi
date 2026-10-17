//! Path resolution and OfficeCLI-shaped node views (record 0010 §2, §4.7).

use super::{both, check_golden, find, get, json, kitchen, open_path, path_of};
use crate::Code;
use crate::node::{Node, ViewOptions};

/// The canonical path and type every input path resolves to, on both
/// owners.
#[test]
fn workbook_paths_resolve_to_canonical_nodes() {
    let cases = [
        ("/", "/", "workbook"),
        ("/Data", "/sheet[1]", "sheet"),
        ("/data", "/sheet[1]", "sheet"),
        ("/DATA/a1", "/sheet[1]/A1", "cell"),
        ("/sheet[2]", "/sheet[2]", "sheet"),
        ("/sheet[last()]", "/sheet[4]", "sheet"),
        ("/sheet", "/sheet[1]", "sheet"),
        ("/ärger/A1", "/sheet[4]/A1", "cell"),
        ("/ÄRGER", "/sheet[4]", "sheet"),
        ("/My Data/B2", "/sheet[3]/B2", "cell"),
        ("/\"My Data\"/B2", "/sheet[3]/B2", "cell"),
        ("Data!B2", "/sheet[1]/B2", "cell"),
        ("'My Data'!A1", "/sheet[3]/A1", "cell"),
        ("Data!3:3", "/sheet[1]/row[3]", "row"),
        ("Data!b:b", "/sheet[1]/col[B]", "column"),
        ("Data!A1:C5", "/sheet[1]/A1:C5", "range"),
        ("/Data/C5:A1", "/sheet[1]/A1:C5", "range"),
        ("/Data/B2:B2", "/sheet[1]/B2", "range"),
        ("/Data/row[7]", "/sheet[1]/row[7]", "row"),
        ("/Data/row[100]", "/sheet[1]/row[100]", "row"),
        ("/Data/row[last()]", "/sheet[1]/row[7]", "row"),
        ("/Data/col[C]", "/sheet[1]/col[C]", "column"),
        ("/Data/col[3]", "/sheet[1]/col[C]", "column"),
        ("/Data/column[last()]", "/sheet[1]/col[C]", "column"),
        ("/Data/table[1]", "/sheet[1]/table[1]", "table"),
        ("/Data/table[@name=sales]", "/sheet[1]/table[1]", "table"),
        ("/Data/listobject[last()]", "/sheet[1]/table[1]", "table"),
        ("/table[@name=SALES]", "/sheet[1]/table[1]", "table"),
        ("/table[1]", "/sheet[1]/table[1]", "table"),
        ("/namedrange[2]", "/namedrange[2]", "namedrange"),
        ("/namedrange[total]", "/namedrange[2]", "namedrange"),
        (
            "/definedname[@name=Volatile]",
            "/namedrange[3]",
            "namedrange",
        ),
        ("/Data/Z99", "/sheet[1]/Z99", "cell"),
        ("/Data/XFD1048576", "/sheet[1]/XFD1048576", "cell"),
    ];
    for file in both(&kitchen()) {
        for (input, canonical, kind) in cases {
            let node = get(&file, input).unwrap_or_else(|error| panic!("{input}: {error}"));
            assert_eq!(path_of(&node), canonical, "{input}");
            assert_eq!(node.kind(), kind, "{input}");
            assert_eq!(node.positional_path().to_string(), canonical, "{input}");
        }
    }
}

#[test]
fn misses_and_refusals_carry_one_code_each() {
    let cases = [
        ("/Nope", Code::PathNotFound),
        ("/Nope/A1", Code::PathNotFound),
        ("/sheet[9]", Code::PathNotFound),
        ("/Data/table[2]", Code::PathNotFound),
        ("/Data/table[@name=Missing]", Code::PathNotFound),
        ("/namedrange[Missing]", Code::PathNotFound),
        ("/namedrange[9]", Code::PathNotFound),
        ("/Notes/table[1]", Code::PathNotFound),
        ("/Data/comment[1]", Code::PathNotFound),
        ("/Data/chart[1]", Code::PathNotFound),
        ("/Data/pivottable[1]", Code::UnsupportedCapability),
        ("/Data/validation[1]", Code::PathNotFound),
        ("/Data/sparkline[1]", Code::UnsupportedCapability),
        // Record 0338: runs are read; this cell has none.
        ("/Data/A1/run[1]", Code::PathNotFound),
        ("/slide[1]", Code::UnsupportedPath),
        ("/body/p[1]", Code::UnsupportedPath),
        ("/Data/namedrange[1]", Code::UnsupportedPath),
        ("/Data/XFE1", Code::InvalidPath),
        ("/Data/A1048577", Code::InvalidPath),
        ("/Data/row[0]", Code::InvalidPath),
        ("/Data/col[XFE]", Code::InvalidPath),
        ("/Data/row[Sales>5]", Code::InvalidPath),
        ("/Data/table[@ref=A1]", Code::InvalidPath),
        ("Data", Code::InvalidPath),
    ];
    for file in both(&kitchen()) {
        for (input, code) in cases {
            let error = get(&file, input).expect_err(input);
            assert_eq!(error.code(), code, "{input}: {error}");
        }
    }
}

#[test]
fn misses_suggest_what_exists_and_find_returns_none() {
    let file = &both(&kitchen())[0];
    let error = get(file, "/Nope/A1").expect_err("no sheet");
    let suggestion = error.suggestion().expect("sheet position suggestion");
    assert!(suggestion.contains("sheet[1]") && suggestion.contains("sheet[4]"));
    assert!(!suggestion.contains("Data"));
    let error = get(file, "/Data/table[3]").expect_err("one table");
    let suggestion = error.suggestion().expect("table position suggestion");
    assert!(suggestion.contains("table[1]"));
    assert!(!suggestion.contains("Data"));
    assert!(
        find(file, "/Nope")
            .expect("a miss is not an error")
            .is_none()
    );
    assert!(find(file, "/Data/table[3]").expect("a miss").is_none());
    // A coordinate inside the grid is never a miss (record 0010 §4.7).
    assert!(find(file, "/Data/AA100").expect("resolves").is_some());
}

#[test]
fn a_chart_sheet_has_no_cells() {
    let file = open_path("poi/test-data/spreadsheet/chart_sheet.xlsx");
    let sheets: Vec<String> = get(&file, "/")
        .expect("root")
        .children()
        .map(|child| path_of(&child.expect("a sheet")))
        .collect();
    assert_eq!(sheets, ["/sheet[1]", "/sheet[2]", "/sheet[3]", "/sheet[4]"]);
    let chart = get(&file, "/sheet[3]").expect("the chart sheet");
    assert_eq!(path_of(&chart), "/sheet[3]");
    assert_eq!(chart.children().count(), 0);
    for path in ["/Chart1/A1", "/Chart1/row[1]"] {
        let error = get(&file, path).expect_err("no grid");
        assert_eq!(error.code(), Code::UnsupportedType, "{path}: {error}");
    }
}

fn views(file: &crate::format::xlsx::Xlsx, paths: &[&str], options: ViewOptions) -> String {
    let nodes: Vec<Node<'_>> = paths
        .iter()
        .map(|path| get(file, path).unwrap_or_else(|error| panic!("{path}: {error}")))
        .collect();
    let views: Vec<String> = nodes.iter().map(|node| json(node, options)).collect();
    format!("[\n{}\n]", views.join(",\n"))
}

/// Golden node views in OfficeCLI's `DocumentNode` shape, identical on both
/// owners.
#[test]
fn node_views_match_officecli_shapes() {
    let cases: &[(&str, &[&str], ViewOptions)] = &[
        ("workbook", &["/"], ViewOptions::default()),
        ("sheet", &["/Data", "/Notes"], ViewOptions::default()),
        (
            "cells",
            &[
                "/Data/A1",
                "/Data/B2",
                "/Data/C3",
                "/Data/D2",
                "/Data/D3",
                "/Data/D4",
                "/Data/B4",
                "/Data/B5",
                "/Data/D5",
                "/Data/B6",
                "/Data/C6",
                "/Data/D6",
                "/Data/E1",
                "/Data/F2",
                "/Data/E2",
                "/Data/Z99",
                "/Ärger/A1",
            ],
            ViewOptions::default(),
        ),
        (
            "rows",
            &[
                "/Data/row[1]",
                "/Data/row[2]",
                "/Data/row[7]",
                "/Data/row[50]",
            ],
            ViewOptions::default(),
        ),
        (
            "columns",
            &["/Data/col[A]", "/Data/col[C]", "/Data/col[Z]"],
            ViewOptions::default(),
        ),
        (
            "range",
            &["/Data/A2:B3", "/Data/D1:E2"],
            ViewOptions::default(),
        ),
        (
            "range_dense",
            &["/Data/D1:E2"],
            ViewOptions::default().dense(16),
        ),
        ("table", &["/Data/table[1]"], ViewOptions::default()),
        (
            "named_ranges",
            &["/namedrange[1]", "/namedrange[2]", "/namedrange[3]"],
            ViewOptions::default(),
        ),
    ];
    let files = both(&kitchen());
    for (name, paths, options) in cases {
        let source = views(&files[0], paths, *options);
        let eager = views(&files[1], paths, *options);
        assert_eq!(source, eager, "{name}: the owners disagree");
        check_golden(name, &source);
    }
}

/// Real fixtures: the ListObject of `ExcelTables.xlsx` and its sheet.
#[test]
fn fixture_views_match_officecli_shapes() {
    let file = open_path("ooxml/xlsx/ExcelTables.xlsx");
    check_golden(
        "fixture_excel_tables",
        &views(
            &file,
            &[
                "/",
                "/ExcelTable",
                "/ExcelTable/table[1]",
                "/ExcelTable/C2",
                "/namedrange[2]",
            ],
            ViewOptions::default(),
        ),
    );
    let file = open_path("ooxml/xlsx/table-sample.xlsx");
    check_golden(
        "fixture_table_sample",
        &views(
            &file,
            &["/Tabelle1/table[1]", "/Tabelle1/G5"],
            ViewOptions::default(),
        ),
    );
}

#[test]
fn dense_ranges_are_bounded() {
    let file = &both(&kitchen())[0];
    let range = get(file, "/Data/A1:Z100").expect("a range");
    let mut sink = Vec::new();
    let error = range
        .view_with(ViewOptions::default().dense(100))
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect_err("2,600 cells exceed the budget");
    assert!(matches!(
        error,
        crate::node::SerializeError::View(ref error) if error.code() == Code::OutputTooLarge
    ));
    // The sparse listing has only the stored records.
    assert_eq!(range.children().count(), 25);
}

#[test]
fn typed_accessors_say_where_a_node_sits() {
    let file = &both(&kitchen())[0];
    let cell = get(file, "/Data/C3").expect("a cell");
    let xlsx = cell.as_xlsx().expect("an XLSX node");
    assert_eq!(xlsx.sheet_name().expect("sheet name"), Some("Data"));
    assert_eq!((xlsx.kind(), xlsx.cell()), ("cell", Some((2, 2))));
    assert_eq!(xlsx.number(), None);
    let row = get(file, "Data!7:7").expect("a row");
    assert_eq!(row.as_xlsx().and_then(|row| row.number()), Some(7));
    let column = get(file, "Data!AB:AB").expect("a column");
    assert_eq!(
        column.as_xlsx().and_then(|column| column.number()),
        Some(28)
    );
    let root = get(file, "/").expect("the root");
    assert_eq!(
        root.as_xlsx()
            .expect("XLSX root")
            .sheet_name()
            .expect("root sheet name"),
        None
    );
}
