#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, sign::Report};
use litchi_sign::{Policy, Status};

const DOCX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
const XLSX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
const PPTX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx");

fn assert_valid(reports: &[Report]) {
    assert_eq!(reports.len(), 1);
    assert_eq!(reports[0].details().integrity(), Status::Valid);
    assert_eq!(reports[0].details().signature(), Status::Valid);
}

#[test]
fn digital_signature_accessors_verify_real_microsoft_packages() {
    let policy = Policy::compatible();

    let docx = OpcPackage::from_bytes(DOCX).unwrap();
    assert!(docx.is_signed());
    assert_valid(&docx.signatures_with(&policy).unwrap());

    let xlsx = OpcPackage::from_bytes(XLSX).unwrap();
    assert!(xlsx.is_signed());
    assert_valid(&xlsx.signatures_with(&policy).unwrap());

    let pptx = OpcPackage::from_bytes(PPTX).unwrap();
    assert!(pptx.is_signed());
    assert_valid(&pptx.signatures_with(&policy).unwrap());
}

#[test]
fn digital_signature_mutable_opc_access_drops_stale_signatures() {
    let policy = Policy::compatible();

    let mut docx = OpcPackage::from_bytes(DOCX).unwrap();
    docx.unsign();
    assert!(!docx.is_signed());
    assert!(docx.signatures_with(&policy).unwrap().is_empty());

    let mut xlsx = OpcPackage::from_bytes(XLSX).unwrap();
    xlsx.unsign();
    assert!(!xlsx.is_signed());
    assert!(xlsx.signatures_with(&policy).unwrap().is_empty());

    let mut pptx = OpcPackage::from_bytes(PPTX).unwrap();
    pptx.unsign();
    assert!(!pptx.is_signed());
    assert!(pptx.signatures_with(&policy).unwrap().is_empty());
}
