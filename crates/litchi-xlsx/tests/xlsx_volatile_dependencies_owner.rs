use litchi_opc::OpcPackage;
use litchi_xlsx::Package;
use litchi_xlsx::volatile_dependencies::{
    VolatileDependencies, VolatileDependenciesConformance, load_from_package,
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

fn stored(value: &VolatileDependencies) -> Package {
    let mut edit = Package::new()
        .expect("create workbook package")
        .edit_volatile_dependencies()
        .unwrap();
    edit.set(value, VolatileDependenciesConformance::Strict)
        .unwrap();
    edit.commit().unwrap().into_package()
}

#[test]
fn volatile_dependencies_survive_direct_package_publication() {
    let value = fixture_value();
    let package = stored(&value).into_plain_opc();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("volatile-dependencies.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path, litchi_opc::SourceLimits::default()).unwrap();
    assert_eq!(load_from_package(&reopened).unwrap(), Some(value));
}

#[test]
fn volatile_dependencies_removal_is_idempotent() {
    let mut edit = stored(&fixture_value())
        .edit_volatile_dependencies()
        .unwrap();
    assert!(edit.remove().unwrap());
    assert!(!edit.remove().unwrap());
    let package = edit.commit().unwrap().into_package().into_plain_opc();
    assert_eq!(load_from_package(&package).unwrap(), None);
}
