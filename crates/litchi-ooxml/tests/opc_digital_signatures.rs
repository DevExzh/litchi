use litchi_ooxml::{OpcPackage, Policy, Status};
use std::io::Cursor;

const DOCX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
const XLSX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
const PPTX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx");

fn assert_valid(reports: &[litchi_ooxml::opc::sign::Report]) {
    assert_eq!(reports.len(), 1);
    assert_eq!(reports[0].details().integrity(), Status::Valid);
    assert_eq!(reports[0].details().signature(), Status::Valid);
}

#[test]
fn digital_signature_accessors_verify_real_microsoft_packages() {
    let policy = Policy::compatible();

    let docx = litchi_ooxml::docx::Package::from_reader(Cursor::new(DOCX)).unwrap();
    assert!(docx.is_signed());
    assert_valid(&docx.signatures_with(&policy).unwrap());

    let xlsx = litchi_ooxml::xlsx::Workbook::new(OpcPackage::from_bytes(XLSX).unwrap()).unwrap();
    assert!(xlsx.is_signed());
    assert_valid(&xlsx.signatures_with(&policy).unwrap());

    let pptx = litchi_ooxml::pptx::Package::from_reader(Cursor::new(PPTX)).unwrap();
    assert!(pptx.is_signed());
    assert_valid(&pptx.signatures_with(&policy).unwrap());
}

#[test]
fn digital_signature_mutable_opc_access_drops_stale_signatures() {
    let policy = Policy::compatible();

    let mut docx = litchi_ooxml::docx::Package::from_reader(Cursor::new(DOCX)).unwrap();
    let _ = docx.opc_package_mut();
    assert!(!docx.is_signed());
    assert!(docx.signatures_with(&policy).unwrap().is_empty());

    let mut xlsx =
        litchi_ooxml::xlsx::Workbook::new(OpcPackage::from_bytes(XLSX).unwrap()).unwrap();
    xlsx.unsign();
    assert!(!xlsx.is_signed());
    assert!(xlsx.signatures_with(&policy).unwrap().is_empty());

    let mut pptx = litchi_ooxml::pptx::Package::from_reader(Cursor::new(PPTX)).unwrap();
    pptx.edit_opc(|_| Ok(())).unwrap();
    assert!(!pptx.is_signed());
    assert!(pptx.signatures_with(&policy).unwrap().is_empty());
}
