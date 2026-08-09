#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::{Error, Package};

#[test]
fn master_theme_inventory_resolves_through_the_standalone_owner() {
    let package = package_with_slide();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    let master_theme = masters[0].theme().unwrap().unwrap();

    assert!(!master_theme.name.is_empty());
    assert!(!master_theme.colors.is_empty());

    let relationship = masters[0]
        .part()
        .part()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .unwrap();
    let theme_name = relationship.target_partname().unwrap();
    let summary = litchi_pptx::shape::theme::package::load_summary(
        presentation.package(),
        theme_name.as_str(),
    )
    .unwrap();
    assert_eq!(summary.name, master_theme.name);
    assert_eq!(summary.colors.len(), master_theme.colors.len());
}

#[test]
fn master_theme_accessor_rejects_external_theme_relationships() {
    let package = package_with_external_master_theme();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();

    assert!(matches!(
        masters[0].theme(),
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));
}

fn package_with_slide() -> Package {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    Package::from_bytes(&package_bytes).unwrap()
}

fn package_with_external_master_theme() -> Package {
    let mut package = package_with_slide();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let master_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    let relationship_id = opc
        .get_part(&master_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .unwrap()
        .r_id()
        .to_string();
    let master = opc.get_part_mut(&master_name).unwrap();
    master.rels_mut().remove(&relationship_id);
    master.rels_mut().add_relationship(
        rt::THEME.to_string(),
        "https://example.invalid/theme.xml".to_string(),
        relationship_id,
        true,
    );
    Package::from_opc_package(opc).unwrap()
}
