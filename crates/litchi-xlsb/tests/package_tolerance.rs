#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

//! Packaging repairs applied by the OPC layer are folded into the
//! workbook-scope report on every open path, and a conformant package
//! contributes nothing.

use std::io::Cursor;

use litchi_opc::SourceBackedPackage;
use litchi_opc::leniency::{Defect as OpcDefect, RelationshipDefect};
use litchi_xlsb::{Defect, Package, SourceBackedWorkbook, ToleranceReport, Workbook};

const WORKBOOK_RELS: &str = "xl/_rels/workbook.bin.rels";
const DUPLICATE: &str = r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#;

fn fixture() -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../test-data/ooxml/xlsb/Simple.xlsb",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap()
}

/// `Simple.xlsb` with a second `rId4` relationship in the workbook manifest.
fn duplicate_relationship_id() -> Vec<u8> {
    let source = fixture();
    let archive = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if name == WORKBOOK_RELS {
            let text = String::from_utf8(bytes).unwrap();
            assert!(text.contains(r#"Id="rId4""#));
            bytes = text
                .replace("</Relationships>", &format!("{DUPLICATE}</Relationships>"))
                .into_bytes();
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn assert_duplicate_recorded(report: &ToleranceReport, route: &str) {
    let packaging = report
        .iter()
        .filter(|entry| matches!(entry.defect(), Defect::Package(_)))
        .collect::<Vec<_>>();
    assert_eq!(packaging.len(), 1, "{route}: {:?}", report.defects());
    let entry = packaging[0];
    assert_eq!(
        entry.defect(),
        Defect::Package(OpcDefect::Relationship(RelationshipDefect::DuplicateId)),
        "{route}"
    );
    assert_eq!(entry.part(), Some("/xl/workbook.bin"), "{route}");
    assert_eq!(entry.item(), Some("rId4"), "{route}");
    assert_eq!(entry.value(), Some("theme/theme1.xml"), "{route}");
    assert_eq!(
        (entry.record_type(), entry.location(), entry.observed()),
        (0, 0, 0),
        "{route}"
    );
    assert!(!entry.defect().repair().is_empty(), "{route}");
    // Packaging repairs lead the report in discovery order.
    assert_eq!(report.defects()[0], *entry, "{route}");
    assert_eq!(
        report.count(OpcDefect::Relationship(RelationshipDefect::DuplicateId)),
        1,
        "{route}"
    );
}

fn assert_clean(report: &ToleranceReport, route: &str) {
    assert!(report.is_clean(), "{route}: {:?}", report.defects());
}

#[test]
fn duplicate_relationship_id_is_listed_in_the_workbook_report_on_every_route() {
    let bytes = duplicate_relationship_id();

    let eager = Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_duplicate_recorded(eager.tolerance_report(), "eager");
    let worksheet = eager.worksheet(0).unwrap();
    assert_clean(worksheet.tolerance_report(), "eager worksheet");

    let owned = Package::from_bytes(bytes.clone()).unwrap();
    assert_duplicate_recorded(owned.tolerance_report(), "owned package");
    let adopted = owned.into_workbook().unwrap();
    assert_duplicate_recorded(adopted.tolerance_report(), "owned package workbook");

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert_duplicate_recorded(&source.tolerance_report(), "source-backed");

    let materialized = SourceBackedPackage::from_reader_with_limits(
        Cursor::new(bytes),
        litchi_opc::ReadLimits::default(),
    )
    .unwrap()
    .into_opc_package()
    .unwrap();
    let materialized = Workbook::from_opc_package(materialized).unwrap();
    assert_duplicate_recorded(materialized.tolerance_report(), "materialized");
}

#[test]
fn conformant_package_contributes_no_packaging_repairs() {
    let bytes = fixture();
    assert_clean(
        Workbook::new(Cursor::new(bytes.clone()))
            .unwrap()
            .tolerance_report(),
        "eager",
    );
    assert_clean(
        Package::from_bytes(bytes.clone())
            .unwrap()
            .tolerance_report(),
        "owned package",
    );
    assert_clean(
        &SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone()))
            .unwrap()
            .tolerance_report(),
        "source-backed",
    );
    let materialized = SourceBackedPackage::from_reader_with_limits(
        Cursor::new(bytes),
        litchi_opc::ReadLimits::default(),
    )
    .unwrap()
    .into_opc_package()
    .unwrap();
    assert_clean(
        Workbook::from_opc_package(materialized)
            .unwrap()
            .tolerance_report(),
        "materialized",
    );
}
