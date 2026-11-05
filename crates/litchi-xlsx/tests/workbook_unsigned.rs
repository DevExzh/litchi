#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! `Workbook::unsigned` is the explicit signature removal that authorizes
//! editing a signed workbook (ADR 0006; unified-ops record 0013).

use std::path::PathBuf;

use litchi_opc::OpcPackage;
use litchi_xlsx::{Edit, Error, Package};

/// The producer-signed fixtures: one signature by `xmlsec`, one by Office 2010.
const SIGNED: [&str; 2] = ["hello-world-signed.xlsx", "ms-office-2010-signed.xlsx"];

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/xmldsign")
            .join(name),
    )
    .expect("signed XLSX fixture")
}

#[test]
fn an_unsigned_copy_authorizes_an_edit_and_publishes_no_signature() {
    for name in SIGNED {
        let bytes = fixture(name);
        let workbook = Package::from_bytes(&bytes).unwrap();
        // An exact no-op returns the signed source unchanged (ADR 0006).
        let noop = workbook.edit().and_then(Edit::commit).unwrap();
        assert!(!noop.diagnostics().changed(), "{name}");
        assert_eq!(noop.package().to_bytes().unwrap(), bytes, "{name}: no-op");
        let mut changed = workbook.edit().unwrap();
        changed
            .sheet(0_usize)
            .unwrap()
            .expect("a first worksheet")
            .set("A1", "edited")
            .unwrap();
        assert!(
            matches!(changed.commit(), Err(Error::Signed)),
            "{name}: a signed workbook refuses a changed transaction"
        );

        let unsigned = workbook.unsign().unwrap();
        // The signed snapshot is left as it was.
        assert_eq!(workbook.to_bytes().unwrap(), bytes, "{name}");

        let mut edit = unsigned.edit().unwrap();
        edit.sheet(0_usize)
            .unwrap()
            .expect("a first worksheet")
            .set("A1", "edited")
            .unwrap();
        let output = edit.commit().unwrap().package().to_bytes().unwrap();
        let reopened =
            OpcPackage::from_bytes(&output, litchi_opc::SourceLimits::default()).unwrap();
        assert!(!reopened.is_signed(), "{name}");
        assert!(
            reopened
                .iter_parts()
                .all(|part| !part.partname().as_str().starts_with("/_xmlsignatures/")),
            "{name}"
        );
    }
}

#[test]
fn an_unsigned_workbook_comes_back_as_itself() {
    let workbook = Package::new().unwrap();
    let unsigned = workbook.unsign().unwrap();
    assert_eq!(unsigned.workbook_ref().len(), workbook.workbook_ref().len());
    assert_eq!(unsigned.to_bytes().unwrap(), workbook.to_bytes().unwrap());
    let mut edit = unsigned.edit().unwrap();
    edit.sheet(0_usize)
        .unwrap()
        .expect("a first worksheet")
        .set("A1", 1_i32)
        .unwrap();
    edit.commit().unwrap();
}
