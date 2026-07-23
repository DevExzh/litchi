use litchi_ooxml::pptx::Package;

#[test]
fn master_and_layout_shape_inventories_are_available() {
    let package = Package::new().unwrap();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();

    assert!(!masters[0].shapes().unwrap().is_empty());

    let layouts = masters[0].slide_layouts().unwrap();
    assert!(!layouts.is_empty());
    assert!(!layouts[0].shapes().unwrap().is_empty());
}
