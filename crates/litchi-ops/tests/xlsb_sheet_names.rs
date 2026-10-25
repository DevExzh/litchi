//! Unified XLSB complete-tab name writes through the public [`File`] API.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_ops::{BatchMode, Code, File, ItemResult, OpenOptions, Props, SourceMode, props};

const CHART: &str = "/sheet[1]";
const WORKSHEET: &str = "/sheet[2]";

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::chart::{Anchor, Chart};
    use litchi_xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("build chart fixture");
    let mut builder = Builder::new();
    builder
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("add chart sheet");
    builder.add_worksheet(builder::Worksheet::new("Data"));
    builder.add_worksheet(builder::Worksheet::new("Tail"));

    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build mixed XLSB fixture")
        .write_to(&mut output)
        .expect("serialize mixed XLSB fixture");
    output.into_inner()
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .expect("open XLSB owner")
}

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn batch_set(path: &'static str, name: &'static str) -> litchi_ops::BatchItem<'static> {
    let mut item = litchi_ops::BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(Props::new().with("name", name));
    item
}

#[test]
fn complete_catalog_name_writes_round_trip_for_source_and_eager_owners() {
    let bytes = fixture();

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish original mixed catalog");
        assert_eq!(
            sheet_property(&source, CHART, "name").as_deref(),
            Some("Sales Chart")
        );
        assert_eq!(
            sheet_property(&source, WORKSHEET, "name").as_deref(),
            Some("Data")
        );

        let renamed = source
            .set(CHART, props! { "name" => "Dashboard" })
            .expect("rename chart tab");
        assert_eq!(
            sheet_property(&renamed, CHART, "name").as_deref(),
            Some("Dashboard")
        );
        assert_eq!(
            sheet_property(&renamed, WORKSHEET, "name").as_deref(),
            Some("Data")
        );

        let reopened = open_owner(&renamed.to_bytes().expect("publish renamed chart"), eager);
        assert_eq!(
            sheet_property(&reopened, CHART, "name").as_deref(),
            Some("Dashboard")
        );
        assert_eq!(
            sheet_property(&reopened, WORKSHEET, "name").as_deref(),
            Some("Data")
        );
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
    }
}

#[test]
fn complete_catalog_name_and_visibility_are_one_atomic_assignment() {
    let bytes = fixture();

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish original mixed catalog");
        let updated = source
            .set(
                CHART,
                props! {
                    "name" => "Dashboard",
                    "visibility" => "hidden",
                },
            )
            .expect("rename and hide chart tab");
        assert_eq!(
            sheet_property(&updated, CHART, "name").as_deref(),
            Some("Dashboard")
        );
        assert_eq!(
            sheet_property(&updated, CHART, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(&updated, WORKSHEET, "visibility"),
            None,
            "the worksheet remains visible"
        );
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let reopened = open_owner(&updated.to_bytes().expect("publish combined update"), eager);
        assert_eq!(
            sheet_property(&reopened, CHART, "name").as_deref(),
            Some("Dashboard")
        );
        assert_eq!(
            sheet_property(&reopened, CHART, "visibility").as_deref(),
            Some("hidden")
        );
    }
}

#[test]
fn complete_catalog_duplicate_names_fail_atomically_and_roll_back_batches() {
    let bytes = fixture();

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish original mixed catalog");

        let duplicate = source
            .set(CHART, props! { "name" => "Data" })
            .expect_err("chart name must not collide with worksheet");
        assert_eq!(duplicate.code(), Code::DuplicateName);
        assert_eq!(duplicate.property(), Some("name"));
        assert_eq!(
            source.to_bytes().expect("duplicate preserves source"),
            before
        );

        let atomic = source
            .batch(
                [
                    batch_set(CHART, "Dashboard"),
                    batch_set(WORKSHEET, "Dashboard"),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must roll back the duplicate");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::DuplicateName);
        assert_eq!(
            source.to_bytes().expect("atomic rollback preserves source"),
            before
        );

        let best_effort = source
            .batch(
                [
                    batch_set(CHART, "Dashboard"),
                    batch_set(WORKSHEET, "Dashboard"),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch keeps the first rename");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::DuplicateName
        ));
        assert_eq!(
            sheet_property(best_effort.file(), CHART, "name").as_deref(),
            Some("Dashboard")
        );
        assert_eq!(
            sheet_property(best_effort.file(), WORKSHEET, "name").as_deref(),
            Some("Data")
        );
    }
}

#[test]
fn complete_catalog_name_write_reports_missing_tab_without_mutation() {
    let bytes = fixture();

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish original mixed catalog");
        let error = source
            .set("/sheet[99]", props! { "name" => "Missing" })
            .expect_err("missing catalog tab");
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(source.to_bytes().expect("missing preserves source"), before);
    }
}
