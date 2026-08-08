use litchi_opc::OpcPackage;
use litchi_xlsx::Package;
use litchi_xlsx::volatile_dependencies::{
    VolatileDependencies, VolatileDependenciesConformance, load_from_package, remove_from_package,
    store_in_package,
};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn fixture_value() -> VolatileDependencies {
    VolatileDependencies::parse(
        format!(
            r#"<volTypes xmlns="{MAIN_NS}"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#
        )
        .as_bytes(),
    )
    .expect("parse volatile dependencies")
}

fn package() -> OpcPackage {
    Package::create()
        .expect("create workbook package")
        .into_plain_opc()
}

#[test]
fn volatile_dependencies_survive_direct_package_publication() {
    let mut package = package();
    let value = fixture_value();
    store_in_package(
        &mut package,
        &value,
        VolatileDependenciesConformance::Strict,
    )
    .unwrap();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("volatile-dependencies.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path).unwrap();
    assert_eq!(load_from_package(&reopened).unwrap(), Some(value));
}

#[test]
fn volatile_dependencies_removal_is_idempotent() {
    let mut package = package();
    store_in_package(
        &mut package,
        &fixture_value(),
        VolatileDependenciesConformance::Strict,
    )
    .unwrap();
    assert!(remove_from_package(&mut package).unwrap());
    assert_eq!(load_from_package(&package).unwrap(), None);
    assert!(!remove_from_package(&mut package).unwrap());
}
