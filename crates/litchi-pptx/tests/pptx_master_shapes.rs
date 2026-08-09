#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Package;

#[test]
fn master_and_layout_shape_inventories_are_available() {
    let package = Package::new().unwrap();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();

    assert!(!masters[0].shapes().unwrap().is_empty());
    assert!(masters[0].shapes().unwrap().placeholders().next().is_some());

    let layouts = masters[0].layouts().unwrap();
    assert!(!layouts.is_empty());
    assert!(!layouts[0].shapes().unwrap().is_empty());
    assert!(layouts[0].shapes().unwrap().placeholders().next().is_some());
}
