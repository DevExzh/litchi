#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
#![cfg(feature = "sign")]

use std::path::PathBuf;

#[test]
fn ppt_facade_discovers_unsigned_signature_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut ppt =
        litchi_ppt::Package::open(root.join("test-data/ole/ppt/text-margins.ppt")).unwrap();
    assert!(ppt.signatures().unwrap().is_empty());
}
