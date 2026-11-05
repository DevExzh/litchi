//! DOCX charts, extended chart types and the chart `set` coverage of record
//! 0283, through the unified API: round trips on every host, read-back of
//! formatting keys, family changes, data replacement, frame geometry, axis
//! ranges and series colors.
#![cfg(all(feature = "pptx", feature = "xlsx", feature = "docx"))]

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

fn ok<T>(result: litchi_ops::Result<T>) -> T {
    result.unwrap_or_else(|error| panic!("{error} ({:?})", error.property()))
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

fn document() -> File {
    litchi_ops::create(FileFormat::Docx)
        .unwrap()
        .add(
            "/body",
            NewNode::new("paragraph").prop("text", "Sales by quarter"),
        )
        .unwrap()
        .file
        .add("/body", NewNode::new("paragraph").prop("text", ""))
        .unwrap()
        .file
}

fn deck() -> File {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
}

fn workbook() -> File {
    litchi_ops::create(FileFormat::Xlsx).unwrap()
}

fn column() -> NewNode<'static> {
    NewNode::new("chart")
        .prop("chartType", "column")
        .prop("title", "Sales")
        .prop("categories", "Q1,Q2,Q3,Q4")
        .prop("data", "East:120,135,148,162;West:95,108,115,128")
        .prop("gapWidth", "80")
        .prop("catTitle", "Quarter")
        .prop("axisTitle", "Units")
}

#[test]
fn docx_charts_round_trip_through_every_path() {
    let file = document();
    let added = file
        .add(
            "/body/p[2]",
            column().prop("width", "16cm").prop("height", "9cm"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[2]/chart[1]");
    let file = added.file;
    let extended = file
        .add(
            "/body/p[2]",
            NewNode::new("chart")
                .prop("chartType", "waterfall")
                .prop("title", "Cash")
                .prop("data", "Start:100,Q1:-30,Q2:50,End:120")
                .prop("increaseColor", "00AA00")
                .prop("decreaseColor", "C00000")
                .prop("totalColor", "4472C4"),
        )
        .unwrap();
    assert_eq!(extended.path.to_string(), "/body/p[2]/chart[2]");
    let file = extended.file;
    for eager in [false, true] {
        let file = reopen(&file, eager);
        for path in ["/body/p[2]/chart[1]", "/chart[1]", "/body/chart[1]"] {
            assert_eq!(prop(&file, path, "chartType").as_deref(), Some("column"));
            assert_eq!(prop(&file, path, "title").as_deref(), Some("Sales"));
        }
        let chart = "/chart[1]";
        assert_eq!(
            prop(&file, chart, "categories").as_deref(),
            Some("Q1,Q2,Q3,Q4")
        );
        assert_eq!(prop(&file, chart, "seriesCount").as_deref(), Some("2"));
        assert_eq!(prop(&file, chart, "gapWidth").as_deref(), Some("80"));
        assert_eq!(prop(&file, chart, "catTitle").as_deref(), Some("Quarter"));
        assert_eq!(prop(&file, chart, "width").as_deref(), Some("16cm"));
        assert_eq!(prop(&file, chart, "name").as_deref(), Some("Sales"));
        let cash = "/chart[2]";
        assert_eq!(prop(&file, cash, "chartType").as_deref(), Some("waterfall"));
        assert_eq!(
            prop(&file, cash, "categories").as_deref(),
            Some("Start,Q1,Q2,End")
        );
        assert_eq!(
            prop(&file, cash, "increaseColor").as_deref(),
            Some("00AA00")
        );
        assert_eq!(file.query("chart").unwrap().count(), 2);
        assert!(file.get("/chart[3]").is_err());
    }

    // Set: titles, a new family, new data, the frame, a series color and a
    // formatting key; the embedded workbook follows.
    let edited = file
        .set(
            "/chart[1]",
            props! {
                "title" => "Revenue",
                "chartType" => "bar",
                "data" => "North:1,2,3,4",
                "width" => "10cm",
                "gapWidth" => "120",
                "series1.color" => "C00000",
            },
        )
        .unwrap();
    for eager in [false, true] {
        let file = reopen(&edited, eager);
        let chart = "/chart[1]";
        assert_eq!(prop(&file, chart, "chartType").as_deref(), Some("bar"));
        assert_eq!(prop(&file, chart, "title").as_deref(), Some("Revenue"));
        assert_eq!(prop(&file, chart, "seriesCount").as_deref(), Some("1"));
        assert_eq!(prop(&file, chart, "gapWidth").as_deref(), Some("120"));
        assert_eq!(prop(&file, chart, "colors").as_deref(), Some("C00000"));
        assert_eq!(prop(&file, chart, "width").as_deref(), Some("10cm"));
        assert_eq!(
            prop(&file, chart, "categories").as_deref(),
            Some("Q1,Q2,Q3,Q4")
        );
    }
    // The extended chart is rebuilt on every set; families do not cross.
    let retitled = edited
        .set(
            "/chart[2]",
            props! { "title" => "Cash flow", "chartType" => "funnel" },
        )
        .unwrap();
    let file = reopen(&retitled, true);
    assert_eq!(
        prop(&file, "/chart[2]", "chartType").as_deref(),
        Some("funnel")
    );
    assert_eq!(
        prop(&file, "/chart[2]", "title").as_deref(),
        Some("Cash flow")
    );
    let before = retitled.to_bytes().unwrap();
    let error = retitled
        .set("/chart[2]", props! { "chartType" => "column" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(retitled.to_bytes().unwrap(), before);
    // An exact no-op keeps every byte.
    let same = retitled
        .set("/chart[1]", props! { "title" => "Revenue" })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), before);

    // Remove drops the drawing; the paragraph keeps its text.
    let removed = retitled.remove("/chart[1]").unwrap();
    let file = reopen(&removed, true);
    assert_eq!(file.query("chart").unwrap().count(), 1);
    assert_eq!(
        prop(&file, "/chart[1]", "chartType").as_deref(),
        Some("funnel")
    );
}

#[test]
fn docx_chart_adds_are_deterministic_and_refusals_change_nothing() {
    let file = document();
    let once = file.add("/body/p[1]", column()).unwrap().file;
    let twice = file.add("/body/p[1]", column()).unwrap().file;
    assert_eq!(once.to_bytes().unwrap(), twice.to_bytes().unwrap());
    let before = file.to_bytes().unwrap();
    for (node, code) in [
        (
            NewNode::new("chart").prop("title", "x"),
            Code::MissingProperty,
        ),
        (
            NewNode::new("chart")
                .prop("data", "A:1")
                .prop("width", "0cm"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("chart")
                .prop("data", "A:1")
                .prop("chartType", "nope"),
            Code::InvalidValue,
        ),
    ] {
        assert_eq!(file.add("/body/p[1]", node).unwrap_err().code(), code);
    }
    assert!(file.add("/body", column()).is_err());
    assert_eq!(file.to_bytes().unwrap(), before);
}

const EXTENDED: [&str; 7] = [
    "waterfall",
    "funnel",
    "treemap",
    "sunburst",
    "histogram",
    "pareto",
    "boxWhisker",
];

fn extended_chart(kind: &'static str) -> NewNode<'static> {
    let node = NewNode::new("chart")
        .prop("chartType", kind)
        .prop("title", kind)
        .prop("categories", "A,B,C,D,E");
    match kind {
        "histogram" | "boxWhisker" => node
            .prop("data", "First:1,2,2,3,9;Second:4,5,5,6,1")
            .prop("legend", "bottom"),
        _ => node.prop("data", "Values:50,-20,30,15,75"),
    }
}

#[test]
fn extended_chart_types_round_trip_on_every_host() {
    for kind in EXTENDED {
        for (file, parent, path) in [
            (deck(), "/slide[1]", "/slide[1]/chart[1]"),
            (workbook(), "/Sheet1", "/Sheet1/chart[1]"),
            (document(), "/body/p[2]", "/body/p[2]/chart[1]"),
        ] {
            let added = file.add(parent, extended_chart(kind)).unwrap();
            assert!(added.path.to_string().ends_with("/chart[1]"), "{kind}");
            for eager in [false, true] {
                let file = reopen(&added.file, eager);
                assert_eq!(prop(&file, path, "chartType").as_deref(), Some(kind));
                assert_eq!(prop(&file, path, "title").as_deref(), Some(kind));
            }
            let edited = added
                .file
                .set(path, props! { "title" => "Renamed", "legend" => "top" })
                .unwrap();
            let file = reopen(&edited, true);
            assert_eq!(
                prop(&file, path, "title").as_deref(),
                Some("Renamed"),
                "{kind}"
            );
            assert_eq!(
                prop(&file, path, "legend").as_deref(),
                Some("top"),
                "{kind}"
            );
        }
    }
}

#[test]
fn histogram_and_box_keys_read_back() {
    let file = workbook()
        .add(
            "/Sheet1",
            NewNode::new("chart")
                .prop("chartType", "histogram")
                .prop("data", "Scores:1,2,2,3,9,12")
                .prop("binCount", "4")
                .prop("intervalClosed", "l")
                .prop("overflowBin", "10"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("chart")
                .prop("chartType", "boxWhisker")
                .prop("data", "A:1,2,3;B:4,5,6")
                .prop("quartileMethod", "inclusive"),
        )
        .unwrap()
        .file
        .add(
            "/Sheet1",
            NewNode::new("chart")
                .prop("chartType", "treemap")
                .prop("categories", "A,B,C")
                .prop("data", "Size:3,2,1")
                .prop("parentLabelLayout", "banner"),
        )
        .unwrap()
        .file;
    let file = reopen(&file, true);
    assert_eq!(
        prop(&file, "/Sheet1/chart[1]", "binCount").as_deref(),
        Some("4")
    );
    assert_eq!(
        prop(&file, "/Sheet1/chart[1]", "intervalClosed").as_deref(),
        Some("l")
    );
    assert_eq!(
        prop(&file, "/Sheet1/chart[2]", "quartileMethod").as_deref(),
        Some("inclusive")
    );
    assert_eq!(
        prop(&file, "/Sheet1/chart[3]", "parentLabelLayout").as_deref(),
        Some("banner")
    );
    let edited = file
        .set("/Sheet1/chart[1]", props! { "binSize" => "2.5" })
        .unwrap();
    assert_eq!(
        prop(&reopen(&edited, true), "/Sheet1/chart[1]", "binSize").as_deref(),
        Some("2.5")
    );
}

#[test]
fn chart_sets_cover_type_data_geometry_axes_and_colors() {
    for (file, parent, path, x) in [
        (deck(), "/slide[1]", "/slide[1]/chart[1]", "2in"),
        (workbook(), "/Sheet1", "/Sheet1/chart[1]", "3"),
    ] {
        let file = file.add(parent, column()).unwrap().file;
        let edited = ok(file.set(
            path,
            props! {
                "chartType" => "line",
                "data" => "North:5,6,7,8;South:1,2,3,4",
                "x" => x,
                "axisMax" => "50",
            },
        ));
        let edited = ok(edited.set(
            &format!("{path}/axis[@role=value]"),
            props! { "min" => "-5" },
        ));
        let edited = ok(edited.set(&format!("{path}/series[2]"), props! { "color" => "70AD47" }));
        let edited = ok(edited.set(path, props! { "axisTitle" => "Count" }));
        for eager in [false, true] {
            let file = reopen(&edited, eager);
            assert_eq!(prop(&file, path, "chartType").as_deref(), Some("line"));
            assert_eq!(
                prop(&file, &format!("{path}/series[1]"), "name").as_deref(),
                Some("North")
            );
            assert_eq!(
                prop(&file, &format!("{path}/series[2]"), "color").as_deref(),
                Some("70AD47")
            );
            let axis = format!("{path}/axis[@role=value]");
            assert_eq!(prop(&file, &axis, "max").as_deref(), Some("50"));
            assert_eq!(prop(&file, &axis, "min").as_deref(), Some("-5"));
            assert_eq!(prop(&file, &axis, "title").as_deref(), Some("Count"));
            assert_eq!(prop(&file, path, "axisMin").as_deref(), Some("-5"));
            assert_eq!(
                prop(&file, path, "gapWidth").as_deref(),
                None,
                "a line chart has no gap"
            );
            let expected_x = if path.starts_with("/slide") {
                "5.08cm"
            } else {
                "3"
            };
            assert_eq!(prop(&file, path, "x").as_deref(), Some(expected_x));
        }
    }
}
