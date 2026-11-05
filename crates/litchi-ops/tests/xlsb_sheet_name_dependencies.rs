//! Dependency-aware XLSB sheet-name writes through the unified facade.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_opc::{OpcPackage, PackURI};
use litchi_ops::error::Source;
use litchi_ops::{BatchMode, Code, File, OpenOptions, Props, SourceMode, props};
use litchi_xlsb::package::PackageError;

const CHART: &str = "/sheet[1]";
const DATA: &str = "/sheet[2]";
const CHART_XML: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const APP_XML: &str = "/docProps/app.xml";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

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
        .expect("add chart tab");
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

fn xml_part(bytes: &[u8], content_type: &str) -> String {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let part = package
        .iter_parts()
        .find(|part| part.content_type() == content_type)
        .unwrap_or_else(|| panic!("missing {content_type} part"));
    String::from_utf8(
        package
            .get_part(part.partname())
            .expect("read XML part")
            .blob()
            .to_vec(),
    )
    .expect("XML part is UTF-8")
}

fn app_xml(bytes: &[u8]) -> String {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let uri = PackURI::new(APP_XML).expect("app properties URI");
    String::from_utf8(
        package
            .get_part(&uri)
            .expect("read application properties")
            .blob()
            .to_vec(),
    )
    .expect("application properties are UTF-8")
}

fn rewrite_app_xml(bytes: &[u8], replacement: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let uri = PackURI::new(APP_XML).expect("app properties URI");
    let source = String::from_utf8(
        package
            .get_part(&uri)
            .expect("read application properties")
            .blob()
            .to_vec(),
    )
    .expect("application properties are UTF-8");
    let source = source.replace("</Properties>", &format!("{replacement}</Properties>"));
    assert_ne!(
        source,
        app_xml(bytes),
        "test fixture rewrite must take effect"
    );
    package
        .get_part_mut(&uri)
        .expect("mutate application properties")
        .set_blob(source.into_bytes());
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("serialize rewritten XLSB package");
    output
}

fn unknown_reference_fixture() -> Vec<u8> {
    rewrite_app_xml(&fixture(), r#"<fmla>Data!$A$1</fmla>"#)
}

fn mce_reference_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(), litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let uri = PackURI::new(APP_XML).expect("app properties URI");
    let source = String::from_utf8(
        package
            .get_part(&uri)
            .expect("read application properties")
            .blob()
            .to_vec(),
    )
    .expect("application properties are UTF-8")
    .replace(
        " xmlns:vt=",
        &format!(" xmlns:mc=\"{MCE_NS}\" xmlns:vt="),
    )
    .replace(
        "</Properties>",
        r#"<mc:AlternateContent><mc:Choice Requires="vt"><fmla>Data!$A$1</fmla></mc:Choice></mc:AlternateContent></Properties>"#,
    );
    package
        .get_part_mut(&uri)
        .expect("mutate application properties")
        .set_blob(source.into_bytes());
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("serialize MCE XLSB package");
    output
}

fn set_batch(path: &'static str, name: &'static str) -> litchi_ops::BatchItem<'static> {
    let mut item = litchi_ops::BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(Props::new().with("name", name));
    item
}

#[test]
fn worksheet_rename_rewrites_chart_formulas_and_app_titles_for_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let renamed = source
            .set(DATA, props! { "name" => "Input" })
            .expect("rename worksheet with local dependencies");
        let output = renamed.to_bytes().expect("publish renamed workbook");
        let chart = xml_part(&output, CHART_XML);
        assert!(
            chart.contains("Input!$A$1:$A$2"),
            "chart category formula: {chart}"
        );
        assert!(
            chart.contains("Input!$B$1:$B$2"),
            "chart value formula: {chart}"
        );
        assert!(!chart.contains("Data!"), "stale chart formula: {chart}");

        let app = app_xml(&output);
        assert!(
            app.contains("<vt:lpstr>Input</vt:lpstr>"),
            "app titles: {app}"
        );
        assert!(
            !app.contains("<vt:lpstr>Data</vt:lpstr>"),
            "stale app title: {app}"
        );

        let reopened = open_owner(&output, eager);
        assert_eq!(
            reopened
                .get(DATA)
                .expect("resolve renamed worksheet")
                .view()
                .format()
                .expect("worksheet properties")
                .find(|property| property.key() == "name")
                .map(|property| property.value().to_string())
                .as_deref(),
            Some("Input")
        );
    }
}

#[test]
fn unknown_reference_carrier_refusal_is_typed_late_and_batch_atomic() {
    for eager in [false, true] {
        let source = open_owner(&unknown_reference_fixture(), eager);
        let before = source.to_bytes().expect("publish original package");
        let noop = source
            .set(DATA, props! { "name" => "Data" })
            .expect("unchanged name bypasses dependencies");
        assert_eq!(noop.to_bytes().expect("exact no-op"), before);
        let error = source
            .set(DATA, props! { "name" => "Input" })
            .expect_err("unknown formula-like carrier must refuse rename");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(matches!(
            error.leaf(),
            Source::Xlsb(PackageError::SheetNameDependency(_))
        ));
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let atomic = source
            .batch(
                [set_batch(CHART, "Dashboard"), set_batch(DATA, "Input")],
                BatchMode::Atomic,
            )
            .expect_err("atomic batch must roll back before refusal");
        assert_eq!(atomic.code(), Code::UnsupportedCapability);
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(source.to_bytes().expect("atomic source unchanged"), before);

        let best_effort = source
            .batch(
                [set_batch(CHART, "Dashboard"), set_batch(DATA, "Input")],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch isolates unsupported rename");
        assert!(matches!(
            &best_effort.value()[0],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            litchi_ops::ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::UnsupportedCapability
        ));
    }
}

#[test]
fn markup_compatibility_reference_refusal_preserves_source_and_batch_result() {
    for eager in [false, true] {
        let source = open_owner(&mce_reference_fixture(), eager);
        let before = source.to_bytes().expect("publish original package");
        let error = source
            .set(DATA, props! { "name" => "Input" })
            .expect_err("MCE reference carrier must refuse rename");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(matches!(
            error.leaf(),
            Source::Xlsb(PackageError::SheetNameDependency(_))
        ));
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let best_effort = source
            .batch(
                [set_batch(CHART, "Dashboard"), set_batch(DATA, "Input")],
                BatchMode::BestEffort,
            )
            .expect("best-effort MCE batch remains transactional");
        assert!(matches!(
            &best_effort.value()[0],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            litchi_ops::ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::UnsupportedCapability
        ));
    }
}

#[test]
fn sheet_name_scan_limits_keep_typed_sources_and_limit_codes() {
    use litchi_ooxml_common::xml::sheet_name::{Error as ScanError, Limit};
    let error = litchi_ops::Error::from(PackageError::SheetNameDependency(ScanError::Limit(
        Limit::NestingDepth { limit: 256 },
    )));
    assert_eq!(error.code(), Code::LimitExceeded);
    assert!(matches!(
        error.leaf(),
        Source::Xlsb(PackageError::SheetNameDependency(ScanError::Limit(_)))
    ));
    #[cfg(feature = "xlsx")]
    {
        let error = litchi_ops::Error::from(litchi_xlsx::Error::SheetNameDependency(
            ScanError::Limit(Limit::CapturedTextBytes {
                limit: 8 * 1024 * 1024,
            }),
        ));
        assert_eq!(error.code(), Code::LimitExceeded);
    }
}
