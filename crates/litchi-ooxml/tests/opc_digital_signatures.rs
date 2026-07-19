use litchi_ooxml::{OpcPackage, SignatureVerificationPolicy, VerificationStatus};
use std::io::Cursor;

const DOCX: &[u8] =
    include_bytes!("../../../3rdparty/poi/test-data/xmldsign/ms-office-2010-signed.docx");
const XLSX: &[u8] =
    include_bytes!("../../../3rdparty/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
const PPTX: &[u8] =
    include_bytes!("../../../3rdparty/poi/test-data/xmldsign/ms-office-2010-signed.pptx");

fn assert_valid(reports: &[litchi_ooxml::DigitalSignatureVerification]) {
    assert_eq!(reports.len(), 1);
    assert_eq!(reports[0].package_integrity, VerificationStatus::Valid);
    assert_eq!(reports[0].signature_value, VerificationStatus::Valid);
}

#[test]
fn digital_signature_accessors_verify_real_microsoft_packages() {
    let policy = SignatureVerificationPolicy::compatibility();

    let docx = litchi_ooxml::docx::Package::from_reader(Cursor::new(DOCX)).unwrap();
    assert_valid(&docx.verify_digital_signatures(&policy).unwrap());

    let xlsx = litchi_ooxml::xlsx::Workbook::new(OpcPackage::from_bytes(XLSX).unwrap()).unwrap();
    assert_valid(&xlsx.verify_digital_signatures(&policy).unwrap());

    let pptx = litchi_ooxml::pptx::Package::from_reader(Cursor::new(PPTX)).unwrap();
    assert_valid(&pptx.verify_digital_signatures(&policy).unwrap());
}

#[test]
fn digital_signature_mutable_opc_access_drops_stale_signatures() {
    let policy = SignatureVerificationPolicy::compatibility();

    let mut docx = litchi_ooxml::docx::Package::from_reader(Cursor::new(DOCX)).unwrap();
    let _ = docx.opc_package_mut();
    assert!(docx.verify_digital_signatures(&policy).unwrap().is_empty());

    let mut xlsx =
        litchi_ooxml::xlsx::Workbook::new(OpcPackage::from_bytes(XLSX).unwrap()).unwrap();
    let _ = xlsx.opc_package_mut();
    assert!(xlsx.verify_digital_signatures(&policy).unwrap().is_empty());

    let mut pptx = litchi_ooxml::pptx::Package::from_reader(Cursor::new(PPTX)).unwrap();
    let _ = pptx.opc_package_mut();
    assert!(pptx.verify_digital_signatures(&policy).unwrap().is_empty());
}
