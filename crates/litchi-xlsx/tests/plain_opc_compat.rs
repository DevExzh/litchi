#![cfg(not(feature = "encryption"))]

use litchi_opc::OpcPackage;
use litchi_xlsx::Package;

#[test]
fn plaintext_build_exposes_explicit_opc_conversion() {
    let raw: OpcPackage = Package::new().unwrap().into_plain_opc();

    assert!(raw.main_document_part().is_ok());
}
