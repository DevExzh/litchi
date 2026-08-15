#![cfg(not(feature = "encryption"))]

use litchi_opc::OpcPackage;
use litchi_xlsx::Package;

#[test]
fn plaintext_build_retains_the_legacy_opc_conversion() {
    let raw: OpcPackage = Package::create().unwrap().into();

    assert!(raw.main_document_part().is_ok());
}
