//! Queries and header-row predicates on workbooks (record 0010 §5).

use super::{SML, SheetPart, both, kitchen, query, query_paths, workbook};
use crate::Code;
use crate::format::xlsx::Nav;
use crate::node::Node;
use crate::query::walk::{PlanRef, Walk};
use crate::query::{Plan, Query};

#[test]
fn cell_queries_follow_officecli_semantics() {
    let cases: &[(&str, &[&str])] = &[
        ("cell[type=Boolean]", &["/sheet[1]/D3"]),
        ("cell[type=error]", &["/sheet[1]/D4"]),
        (
            "cell[type!=Number][type!=SharedString]",
            &["/sheet[1]/D3", "/sheet[1]/D4", "/sheet[1]/D5"],
        ),
        (
            "B",
            &[
                "/sheet[1]/B1",
                "/sheet[1]/B2",
                "/sheet[1]/B3",
                "/sheet[1]/B4",
                "/sheet[1]/B5",
                "/sheet[1]/B6",
                "/sheet[2]/B2",
                "/sheet[2]/B3",
                "/sheet[2]/B4",
                "/sheet[3]/B2",
            ],
        ),
        (
            "Data!A1:B2",
            &[
                "/sheet[1]/A1",
                "/sheet[1]/B1",
                "/sheet[1]/A2",
                "/sheet[1]/B2",
            ],
        ),
        (
            "cell[value>1000]",
            &[
                "/sheet[1]/B2",
                "/sheet[1]/B5",
                "/sheet[1]/B6",
                "/sheet[1]/D6",
            ],
        ),
        ("cell[value=closed]", &["/sheet[1]/C3", "/sheet[1]/C5"]),
        ("cell:contains(ORTH)", &["/sheet[1]/A4"]),
        (
            "cell:empty",
            &["/sheet[1]/D2", "/sheet[1]/F2", "/sheet[1]/C6"],
        ),
        (
            "cell[formula]",
            &[
                "/sheet[1]/D5",
                "/sheet[1]/B6",
                "/sheet[1]/C6",
                "/sheet[1]/D6",
            ],
        ),
        (
            "cell:has(formula)",
            &[
                "/sheet[1]/D5",
                "/sheet[1]/B6",
                "/sheet[1]/C6",
                "/sheet[1]/D6",
            ],
        ),
        ("cell[evaluated=false]", &["/sheet[1]/C6"]),
        ("cell[arrayformula=true]", &["/sheet[1]/D6"]),
        (
            "cell[bold=true]",
            &["/sheet[1]/A1", "/sheet[1]/B1", "/sheet[1]/C1"],
        ),
        (
            "cell[font.color=#FF0000]",
            &["/sheet[1]/B2", "/sheet[1]/D2"],
        ),
        ("cell[fill=yellow]", &["/sheet[1]/B2", "/sheet[1]/D2"]),
        ("cell[font.size>=14pt]", &["/sheet[1]/B2", "/sheet[1]/D2"]),
        ("cell[mergeAnchor=true]", &["/sheet[1]/E1"]),
        ("cell[merge=E1:F2]", &["/sheet[1]/E1", "/sheet[1]/F2"]),
        (
            "Data!cell[type=SharedString]:contains(e)",
            &[
                "/sheet[1]/A1",
                "/sheet[1]/B1",
                "/sheet[1]/E1",
                "/sheet[1]/A2",
                "/sheet[1]/C2",
                "/sheet[1]/A3",
                "/sheet[1]/C3",
                "/sheet[1]/C4",
                "/sheet[1]/C5",
                "/sheet[1]/A7",
            ],
        ),
        (
            "/Data/row[2]/cell",
            &[
                "/sheet[1]/A2",
                "/sheet[1]/B2",
                "/sheet[1]/C2",
                "/sheet[1]/D2",
                "/sheet[1]/F2",
            ],
        ),
        (
            "/Data/col[D]/cell[type=Number]",
            &["/sheet[1]/D2", "/sheet[1]/D6"],
        ),
        (
            "/Data/A5:D6/cell[formula]",
            &[
                "/sheet[1]/D5",
                "/sheet[1]/B6",
                "/sheet[1]/C6",
                "/sheet[1]/D6",
            ],
        ),
        ("'My Data'!cell", &["/sheet[3]/A1", "/sheet[3]/B2"]),
    ];
    for file in both(&kitchen()) {
        for (selector, expected) in cases {
            assert_eq!(&query_paths(&file, selector), expected, "{selector}");
        }
    }
}

#[test]
fn element_queries_reach_every_kind() {
    let cases: &[(&str, &[&str])] = &[
        ("row[hidden=true]", &["/sheet[1]/row[7]"]),
        ("row[outlineLevel=1]", &["/sheet[1]/row[7]"]),
        ("row[height>15pt]", &["/sheet[1]/row[1]"]),
        ("row[@height>15]", &["/sheet[1]/row[1]"]),
        ("row[3]", &["/sheet[1]/row[3]", "/sheet[2]/row[3]"]),
        ("row:contains(South)", &["/sheet[1]/row[5]"]),
        ("col[hidden=true]", &["/sheet[1]/col[C]"]),
        ("col[width>10]", &["/sheet[1]/col[A]"]),
        ("col[C]", &["/sheet[1]/col[C]"]),
        ("column", &["/sheet[1]/col[A]", "/sheet[1]/col[C]"]),
        ("table", &["/sheet[1]/table[1]"]),
        ("listobject[name=sales]", &["/sheet[1]/table[1]"]),
        ("table[columns~=Status]", &["/sheet[1]/table[1]"]),
        (
            "namedrange",
            &["/namedrange[1]", "/namedrange[2]", "/namedrange[3]"],
        ),
        (
            "namedrange[scope=workbook]",
            &["/namedrange[2]", "/namedrange[3]"],
        ),
        ("definedname[volatile=true]", &["/namedrange[3]"]),
        ("sheet[hidden=true]", &["/sheet[2]"]),
        ("sheet:contains(data)", &["/sheet[1]", "/sheet[3]"]),
        (
            "row[3], cell[type=Boolean]",
            &["/sheet[1]/row[3]", "/sheet[1]/D3", "/sheet[2]/row[3]"],
        ),
        (
            "cell[type=Boolean], row[3]",
            &["/sheet[1]/row[3]", "/sheet[1]/D3", "/sheet[2]/row[3]"],
        ),
        ("sheet > table", &["/sheet[1]/table[1]"]),
        ("sheet[hidden=true] cell[value>80]", &["/sheet[2]/C3"]),
    ];
    for file in both(&kitchen()) {
        for (selector, expected) in cases {
            assert_eq!(&query_paths(&file, selector), expected, "{selector}");
        }
    }
}

#[test]
fn positions_and_limits_stop_the_walk() {
    for file in both(&kitchen()) {
        assert_eq!(query_paths(&file, "cell[2]"), ["/sheet[1]/B1"]);
        assert_eq!(query_paths(&file, "cell[last()]"), ["/sheet[4]/A1"]);
        assert_eq!(query_paths(&file, "B[value>500][2]"), ["/sheet[1]/B3"]);
        let query = Query::parse("cell").expect("parses");
        let plan = Plan::compile(&query, litchi_core::FileFormat::Xlsx).expect("compiles");
        let nav = Nav::new(&file);
        let mut walk = Walk::new(&nav, PlanRef::Borrowed(&plan))
            .expect("runs")
            .limit(3);
        let mut paths = Vec::new();
        for node in walk.by_ref() {
            paths.push(Node::Xlsx(node.expect("reads")).path().to_string());
        }
        assert_eq!(paths, ["/sheet[1]/A1", "/sheet[1]/B1", "/sheet[1]/C1"]);
        // A limit of one stops inside the first row: the walk has produced
        // the sheet and the first cell only.
        let first = Walk::new(&nav, PlanRef::Borrowed(&plan)).expect("runs");
        assert_eq!(
            first
                .first()
                .expect("reads")
                .map(|node| Node::Xlsx(node).path().to_string()),
            Some("/sheet[1]/A1".to_owned())
        );
    }
}

#[test]
fn compile_time_findings_are_diagnostics_or_errors() {
    let compile = |selector: &str| {
        let query = Query::parse(selector).expect("parses");
        Plan::compile(&query, litchi_core::FileFormat::Xlsx).map(|plan| {
            plan.diagnostics()
                .iter()
                .map(|diagnostic| diagnostic.kind().as_str().to_owned())
                .collect::<Vec<_>>()
        })
    };
    assert_eq!(
        compile("cell[bolld=true]").expect("compiles"),
        ["unknown_key"]
    );
    assert_eq!(
        compile("cell[value>high]").expect("compiles"),
        ["non_numeric"]
    );
    assert_eq!(
        compile("cell[numFmtId=164]")
            .expect_err("ordinary queries reject native IDs")
            .code(),
        Code::InvalidSelector
    );
    for (selector, code) in [
        ("sparkline", Code::UnsupportedCapability),
        ("pivottable[1]", Code::UnsupportedCapability),
        ("cell[link=x]", Code::UnsupportedCapability),
        ("cell[computedValue=1]", Code::UnsupportedCapability),
        ("paragraph", Code::UnsupportedPath),
        ("row[Sales>1 and hidden=true]", Code::InvalidSelector),
        ("row[5>1]", Code::InvalidSelector),
        ("row[@Region=x]", Code::InvalidSelector),
    ] {
        assert_eq!(
            compile(selector).expect_err(selector).code(),
            code,
            "{selector}"
        );
    }
}

#[test]
fn header_rows_bind_to_the_table_owning_the_columns() {
    let cases: &[(&str, &[&str])] = &[
        ("row[Sales>1000]", &["/sheet[1]/row[2]", "/sheet[1]/row[5]"]),
        ("row[sales>1000]", &["/sheet[1]/row[2]", "/sheet[1]/row[5]"]),
        ("row[Region=east]", &["/sheet[1]/row[3]"]),
        (
            "row[Region=West or Status=closed]",
            &["/sheet[1]/row[2]", "/sheet[1]/row[3]", "/sheet[1]/row[5]"],
        ),
        ("row[col.Sales>=800 and Status=open]", &["/sheet[1]/row[2]"]),
        ("row[B>1000]", &["/sheet[1]/row[2]", "/sheet[1]/row[5]"]),
        (
            "row[not(Status=open)]",
            &["/sheet[1]/row[3]", "/sheet[1]/row[5]"],
        ),
        (
            "Data!row[Sales<1000]",
            &["/sheet[1]/row[3]", "/sheet[1]/row[4]"],
        ),
        // No table owns `Score`: the header row of the block on `Notes` is
        // detected.
        ("row[Score>80]", &["/sheet[2]/row[3]"]),
        ("Notes!row[Name=bob]", &["/sheet[2]/row[4]"]),
    ];
    for file in both(&kitchen()) {
        for (selector, expected) in cases {
            assert_eq!(&query_paths(&file, selector), expected, "{selector}");
        }
    }
}

#[test]
fn header_rows_report_missing_and_ambiguous_columns() {
    let file = &both(&kitchen())[0];
    let error = query(file, "row[Revenue>5]").expect_err("no such column");
    assert_eq!(error.code(), Code::NotFound, "{error}");
    let suggestion = error.suggestion().expect("lists columns");
    assert!(
        suggestion.contains('A') && suggestion.contains('B'),
        "{suggestion}"
    );
    assert!(!suggestion.contains("Region"));
    assert!(!suggestion.contains("Status"));
    assert!(!format!("{error:?}").contains("Revenue"));
    let error = query(file, "'My Data'!row[Sales>5]").expect_err("no table on the sheet");
    assert_eq!(error.code(), Code::NotFound, "{error}");
    assert!(!format!("{error:?}").contains("My Data"));

    // Two sheets each own a table with an `Amount` column, and one has a
    // column named like a row property.
    let table = |name: &str, columns: &[&str]| {
        let columns: String = columns
            .iter()
            .enumerate()
            .map(|(index, column)| format!(r#"<tableColumn id="{}" name="{column}"/>"#, index + 1))
            .collect();
        format!(
            r#"<table xmlns="{SML}" id="1" name="{name}" displayName="{name}" ref="A1:B3"><tableColumns count="2">{columns}</tableColumns></table>"#
        )
    };
    let sheet = |name: &'static str, header: &str, table: String| SheetPart {
        name,
        state: None,
        xml: format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c><c r="B1" t="inlineStr"><is><t>{header}</t></is></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>7</v></c><c r="B3"><v>0</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#
        ),
        table: Some(table),
    };
    let bytes = workbook(
        &[
            sheet("One", "Hidden", table("First", &["Amount", "Hidden"])),
            sheet("Two", "Other", table("Second", &["Amount", "Other"])),
        ],
        "",
        &[],
        &super::styles(),
    );
    for file in both(&bytes) {
        let error = query(&file, "row[Amount>1]").expect_err("two tables");
        assert_eq!(error.code(), Code::InvalidSelector, "{error}");
        assert!(error.message().contains("2 tables") && error.message().contains("2 worksheets"));
        let rendered = format!("{error:?}");
        assert!(!rendered.contains("Amount"));
        assert!(!rendered.contains("First"));
        assert!(!rendered.contains("Second"));
        assert_eq!(
            query_paths(&file, "Two!row[Amount>6]"),
            ["/sheet[2]/row[3]"]
        );
        let error = query(&file, "row[hidden=true]").expect_err("a column is named Hidden");
        assert_eq!(error.code(), Code::InvalidSelector, "{error}");
        assert!(!format!("{error:?}").contains("Hidden"));
        assert_eq!(
            query_paths(&file, "row[col.Hidden=1]"),
            ["/sheet[1]/row[2]"]
        );
        assert!(query_paths(&file, "row[@hidden=true]").is_empty());
    }
}
