#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Packages carrying cataloged OPC non-conformities (docs/LENIENCY_AUDIT.md
//! section 8) open by default, and every repair is recorded in the package's
//! tolerance report. A conformant control package reports clean through every
//! ingress path.

use litchi_opc::leniency::{Defect, ToleranceReport};
use litchi_opc::{OpcPackage, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;

const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

const DOCUMENT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#;

const DOCUMENT: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#;
const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

/// Build a minimal package from the given manifests.
fn package(content_types: &str, root_rels: &str, document_rels: &str) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_rels.as_bytes())
        .unwrap();
    writer
        .write_stored("word/_rels/document.xml.rels", document_rels.as_bytes())
        .unwrap();
    writer
        .write_stored("word/document.xml", DOCUMENT.as_bytes())
        .unwrap();
    writer
        .write_stored("word/styles.xml", STYLES.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// Open `bytes` through every ingress path and return each path's report,
/// asserting they agree.
fn reports(bytes: &[u8]) -> ToleranceReport {
    let eager = OpcPackage::from_bytes(bytes).expect("eager open");
    let owned = OpcPackage::from_vec(bytes.to_vec()).expect("owned open");
    let source = SourceBackedPackage::from_vec(bytes.to_vec()).expect("source-backed open");
    let materialized = source.to_opc_package().expect("materialize");

    let report = eager.tolerance_report().clone();
    assert_eq!(owned.tolerance_report(), &report, "owned ingress");
    assert_eq!(source.tolerance_report(), &report, "source-backed ingress");
    assert_eq!(
        materialized.tolerance_report(),
        &report,
        "materialized source-backed package"
    );
    report
}

/// Defect classes and parts, in discovery order.
fn classes(report: &ToleranceReport) -> Vec<(Defect, String, Option<String>)> {
    report
        .iter()
        .map(|tolerated| {
            (
                tolerated.defect(),
                tolerated.part().to_owned(),
                tolerated.item().map(str::to_owned),
            )
        })
        .collect()
}

#[test]
fn conformant_package_reports_clean_through_every_ingress() {
    let report = reports(&package(CONTENT_TYPES, ROOT_RELS, DOCUMENT_RELS));
    assert!(report.is_clean(), "{:?}", classes(&report));
    assert_eq!(report.total(), 0);
    assert_eq!(report.unrecorded(), 0);
}

#[test]
fn new_packages_report_clean() {
    assert!(OpcPackage::new().tolerance_report().is_clean());
}
