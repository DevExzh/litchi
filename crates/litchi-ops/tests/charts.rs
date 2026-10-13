//! Typed chart CRUD through the unified API (record 0270): add, get, set and
//! remove of PPTX slide charts and XLSX worksheet charts, with their series
//! and axes, round trips, no-ops, signatures, encryption, batches and the
//! capability report.
#![cfg(all(feature = "pptx", feature = "xlsx"))]

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

fn chart() -> NewNode<'static> {
    NewNode::new("chart")
        .prop("chartType", "stackedColumn")
        .prop("title", "Sales")
        .prop("legend", "right")
        .prop("categories", "Q1,Q2,Q3,Q4")
        .prop("data", "East:120,135,148,162;West:95,108,115,128")
        .prop("catTitle", "Quarter")
        .prop("axisTitle", "Units")
}

fn deck() -> File {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
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

#[test]
fn pptx_chart_add_get_set_remove_round_trip() {
    let source = deck();
    let before = source.to_bytes().unwrap();
    let added = source
        .add(
            "/slide[1]",
            chart()
                .prop("x", "0.5in")
                .prop("y", "1in")
                .prop("width", "6in")
                .prop("height", "3in"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/slide[1]/chart[1]");
    // The input snapshot is unchanged.
    assert_eq!(source.to_bytes().unwrap(), before);
    for eager in [false, true] {
        let file = reopen(&added.file, eager);
        let chart = "/slide[1]/chart[1]";
        assert_eq!(
            prop(&file, chart, "chartType").as_deref(),
            Some("stackedColumn")
        );
        assert_eq!(prop(&file, chart, "title").as_deref(), Some("Sales"));
        assert_eq!(prop(&file, chart, "legend").as_deref(), Some("right"));
        assert_eq!(
            prop(&file, chart, "categories").as_deref(),
            Some("Q1,Q2,Q3,Q4")
        );
        assert_eq!(prop(&file, chart, "catTitle").as_deref(), Some("Quarter"));
        assert_eq!(prop(&file, chart, "width").as_deref(), Some("15.24cm"));
        assert_eq!(
            prop(&file, "/slide[1]/chart[1]/series[2]", "name").as_deref(),
            Some("West")
        );
        assert_eq!(
            prop(&file, "/slide[1]/chart[1]/series[2]", "values").as_deref(),
            Some("95,108,115,128")
        );
        assert_eq!(
            prop(&file, "/slide[1]/chart[1]/axis[@role=value]", "title").as_deref(),
            Some("Units")
        );
        assert_eq!(
            file.get("/slide[1]/chart[1]").unwrap().children().count(),
            4,
            "two series and two axes"
        );

        let edited = file
            .set(
                chart,
                props! { "title" => "Revenue", "legend" => "none", "catTitle" => "Period" },
            )
            .unwrap()
            .set(
                "/slide[1]/chart[1]/series[1]",
                props! { "name" => "North", "values" => "1,2,3,4" },
            )
            .unwrap()
            .set("/slide[1]/chart[1]/axis[2]", props! { "title" => "Count" })
            .unwrap()
            .set(chart, props! { "categories" => "A,B,C,D" })
            .unwrap();
        let reread = reopen(&edited, eager);
        assert_eq!(prop(&reread, chart, "title").as_deref(), Some("Revenue"));
        assert_eq!(prop(&reread, chart, "legend").as_deref(), Some("none"));
        assert_eq!(prop(&reread, chart, "catTitle").as_deref(), Some("Period"));
        assert_eq!(
            prop(&reread, chart, "categories").as_deref(),
            Some("A,B,C,D")
        );
        assert_eq!(
            prop(&reread, "/slide[1]/chart[1]/series[1]", "values").as_deref(),
            Some("1,2,3,4")
        );
        assert_eq!(
            prop(&reread, "/slide[1]/chart[1]/series[1]", "name").as_deref(),
            Some("North")
        );
        assert_eq!(
            prop(&reread, "/slide[1]/chart[1]/axis[@role=value]", "title").as_deref(),
            Some("Count")
        );
        // A title can be cleared.
        let cleared = reread
            .set(chart, props! { "title" => litchi_ops::PropValue::Null })
            .unwrap();
        assert_eq!(prop(&cleared, chart, "title"), None);

        let removed = reread.remove(chart).unwrap();
        assert_eq!(removed.get(chart).unwrap_err().code(), Code::PathNotFound);
        // The eager owner drops the chart part and its workbook with the
        // frame; the source owner keeps unused parts for exact source replay
        // (`litchi_pptx::source` removal policy).
        if eager {
            let names = zip_names(&removed.to_bytes().unwrap());
            assert!(names.iter().all(
                |name| !name.starts_with("ppt/charts/") && !name.starts_with("ppt/embeddings/")
            ));
        }
    }
}

fn zip_names(bytes: &[u8]) -> Vec<String> {
    let archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    archive.file_names().map(str::to_owned).collect()
}

#[test]
fn pptx_chart_add_is_deterministic_and_writes_an_embedded_workbook() {
    let first = deck()
        .add("/slide[1]", chart())
        .unwrap()
        .file
        .to_bytes()
        .unwrap();
    let second = deck()
        .add("/slide[1]", chart())
        .unwrap()
        .file
        .to_bytes()
        .unwrap();
    assert_eq!(first, second);
    let names = zip_names(&first);
    assert!(names.contains(&"ppt/charts/chart1.xml".to_owned()));
    assert!(names.contains(&"ppt/charts/_rels/chart1.xml.rels".to_owned()));
    assert!(names.contains(&"ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx".to_owned()));
    // The embedded workbook is an XLSX that holds the chart data.
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&first)).unwrap();
    let mut workbook = Vec::new();
    std::io::Read::read_to_end(
        &mut archive
            .by_name("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")
            .unwrap(),
        &mut workbook,
    )
    .unwrap();
    let embedded = litchi_ops::open_bytes(workbook).unwrap();
    let text = |path: &str| {
        embedded
            .get(path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .map(std::borrow::Cow::into_owned)
    };
    assert_eq!(text("/Sheet1/B1").as_deref(), Some("East"));
    assert_eq!(text("/Sheet1/A3").as_deref(), Some("Q2"));
    assert_eq!(text("/Sheet1/C5").as_deref(), Some("128"));
}

#[test]
fn pptx_chart_noops_and_refusals_leave_the_input_unchanged() {
    let file = deck().add("/slide[1]", chart()).unwrap().file;
    let before = file.to_bytes().unwrap();
    let same = file
        .edit(|edit| {
            edit.set("/slide[1]/chart[1]", props! { "title" => "Sales" })?;
            Ok(())
        })
        .unwrap();
    assert!(same.patch.is_noop());
    assert_eq!(same.file.to_bytes().unwrap(), before);
    for (node, code) in [
        (
            NewNode::new("chart").prop("chartType", "column"),
            Code::MissingProperty,
        ),
        (
            NewNode::new("chart").prop("data", "A:1,x"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("chart")
                .prop("chartType", "warp")
                .prop("data", "A:1"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("chart")
                .prop("chartType", "waterfall")
                .prop("data", "A:1"),
            Code::UnsupportedCapability,
        ),
        (
            NewNode::new("chart")
                .prop("legend", "middle")
                .prop("data", "A:1"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("chart")
                .prop("preset", "minimal")
                .prop("data", "A:1"),
            Code::UnsupportedProperty,
        ),
    ] {
        assert_eq!(file.add("/slide[1]", node).unwrap_err().code(), code);
    }
    assert_eq!(
        file.set("/slide[1]/chart[1]", props! { "chartType" => "bar" })
            .unwrap_err()
            .code(),
        Code::UnsupportedProperty
    );
    assert_eq!(
        file.set(
            "/slide[1]/chart[1]/series[1]",
            props! { "values" => "1,nope" }
        )
        .unwrap_err()
        .code(),
        Code::InvalidValue
    );
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn pptx_chart_batches_replay_and_invert() {
    use litchi_ops::{BatchItem, BatchMode, Props};
    let source = deck();
    let before = source.to_bytes().unwrap();
    let mut add = BatchItem::new("add");
    add.parent = Some("/slide[1]".into());
    add.kind = Some("chart".into());
    add.props = Some(
        Props::new()
            .with("chartType", "line")
            .with("data", "A:1,2,3")
            .with("title", "Trend"),
    );
    let mut set = BatchItem::new("set");
    set.path = Some("/slide[1]/chart[1]".into());
    set.props = Some(Props::new().with("title", "Renamed"));
    let commit = source.batch([add, set], BatchMode::Atomic).unwrap();
    assert_eq!(
        prop(&commit.file, "/slide[1]/chart[1]", "title").as_deref(),
        Some("Renamed")
    );
    assert_eq!(
        commit
            .patch
            .inverse()
            .apply(&commit.file)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn signed_and_encrypted_packages_refuse_chart_writes() {
    let signed =
        std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let file = litchi_ops::open_bytes(signed.clone()).unwrap();
    assert_eq!(
        file.add("/slide[1]", chart()).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(file.to_bytes().unwrap(), signed);
    let File::Signed(wrapper) = file else {
        panic!("signed fixture")
    };
    let unsigned = wrapper.unsign();
    let added = unsigned.add("/slide[1]", chart()).unwrap();
    assert_eq!(
        prop(&added.file, "/slide[1]/chart[1]", "title").as_deref(),
        Some("Sales")
    );

    let signed =
        std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.xlsx").unwrap();
    let file = litchi_ops::open_bytes(signed.clone()).unwrap();
    assert_eq!(
        file.add("/sheet[1]", chart()).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(file.to_bytes().unwrap(), signed);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_packages_refuse_plaintext_chart_publication() {
    use litchi_ops::Password;
    let clear = deck().to_bytes().unwrap();
    let encrypted =
        litchi_pptx::encryption::encrypt(clear, "pw", litchi_pptx::encryption::Mode::Standard)
            .unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(encrypted).unwrap() else {
        panic!("encrypted")
    };
    let source = locked.unlock(Password::new("pw".to_owned())).unwrap();
    assert_eq!(
        source.add("/slide[1]", chart()).unwrap_err().code(),
        Code::DocumentProtected
    );
}

fn workbook() -> File {
    let mut file = litchi_ops::create(FileFormat::Xlsx).unwrap();
    for (cell, value) in [
        ("A1", "Month"),
        ("B1", "East"),
        ("C1", "West"),
        ("A2", "Jan"),
        ("A3", "Feb"),
    ] {
        file = file
            .set(
                format!("/Sheet1/{cell}").as_str(),
                props! { "value" => value },
            )
            .unwrap();
    }
    for (cell, value) in [("B2", 10), ("B3", 15), ("C2", 20), ("C3", 25)] {
        file = file
            .set(
                format!("/Sheet1/{cell}").as_str(),
                props! { "value" => value },
            )
            .unwrap();
    }
    file
}

#[test]
fn xlsx_chart_add_get_set_remove_round_trip() {
    let source = workbook();
    let inline = source
        .add(
            "/Sheet1",
            chart()
                .prop("x", 4)
                .prop("y", 0)
                .prop("width", 8)
                .prop("height", 12),
        )
        .unwrap();
    assert_eq!(inline.path.to_string(), "/sheet[1]/chart[1]");
    let ranged = inline
        .file
        .add(
            "/Sheet1",
            NewNode::new("chart")
                .prop("chartType", "bar")
                .prop("dataRange", "Sheet1!A1:C3")
                .prop("title", "Ranged"),
        )
        .unwrap()
        .file;
    for eager in [false, true] {
        let file = reopen(&ranged, eager);
        assert_eq!(
            prop(&file, "/Sheet1/chart[1]", "chartType").as_deref(),
            Some("stackedColumn")
        );
        assert_eq!(
            prop(&file, "/Sheet1/chart[1]", "seriesCount").as_deref(),
            Some("2")
        );
        assert_eq!(
            prop(&file, "/Sheet1/chart[2]", "categories").as_deref(),
            Some("Jan,Feb")
        );
        assert_eq!(
            prop(&file, "/Sheet1/chart[2]/series[2]", "name").as_deref(),
            Some("West")
        );
        assert_eq!(
            prop(&file, "/Sheet1/chart[2]/series[2]", "values").as_deref(),
            Some("20,25")
        );
        let edited = file
            .set(
                "/Sheet1/chart[2]",
                props! { "title" => "Edited", "legend" => "top" },
            )
            .unwrap()
            .set("/Sheet1/chart[2]/series[1]", props! { "name" => "North" })
            .unwrap()
            .set(
                "/Sheet1/chart[1]/axis[@role=category]",
                props! { "title" => "When" },
            )
            .unwrap();
        let reread = reopen(&edited, eager);
        assert_eq!(
            prop(&reread, "/Sheet1/chart[2]", "title").as_deref(),
            Some("Edited")
        );
        assert_eq!(
            prop(&reread, "/Sheet1/chart[2]", "legend").as_deref(),
            Some("top")
        );
        assert_eq!(
            prop(&reread, "/Sheet1/chart[2]/series[1]", "name").as_deref(),
            Some("North")
        );
        assert_eq!(
            prop(&reread, "/Sheet1/chart[1]", "catTitle").as_deref(),
            Some("When")
        );
        let removed = reread.remove("/Sheet1/chart[1]").unwrap();
        assert_eq!(
            prop(&removed, "/Sheet1/chart[1]", "title").as_deref(),
            Some("Edited")
        );
        assert_eq!(
            removed.get("/Sheet1/chart[2]").unwrap_err().code(),
            Code::PathNotFound
        );
    }
    let first = workbook()
        .add("/Sheet1", chart())
        .unwrap()
        .file
        .to_bytes()
        .unwrap();
    let second = workbook()
        .add("/Sheet1", chart())
        .unwrap()
        .file
        .to_bytes()
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn chart_capabilities_report_the_registry() {
    use litchi_ops::capabilities::Capabilities;
    for format in [FileFormat::Pptx, FileFormat::Xlsx] {
        let caps = Capabilities::of(format);
        let chart = caps.element("chart").expect("chart element");
        assert!(chart.operations().add, "{format:?}");
        assert!(chart.operations().set, "{format:?}");
        assert!(chart.operations().remove, "{format:?}");
        assert!(chart.gap().is_none(), "{format:?}");
        let data = chart.property("data").expect("data");
        assert!(data.add() && !data.set() && !data.get());
        let title = chart.property("title").expect("title");
        assert!(title.add() && title.set() && title.get());
        let chart_type = chart.property("chartType").expect("chartType");
        assert!(chart_type.add() && !chart_type.set() && chart_type.get());
        let series = caps.element("series").expect("series element");
        let values = series.property("values").expect("values");
        assert!(values.set() && values.get() && !values.add());
        let axis = caps.element("axis").expect("axis element");
        assert!(axis.property("title").expect("title").set());
    }
}
