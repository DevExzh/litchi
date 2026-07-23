use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use litchi_opc::constants::relationship_type as rt;

#[test]
fn master_theme_and_presentation_theme_inventory_resolve() {
    let package = Package::new().unwrap();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();

    let master_theme = masters[0].theme().unwrap();
    assert!(!master_theme.name.is_empty());
    assert!(!master_theme.colors.is_empty());

    let themes = presentation.get_themes().unwrap();
    assert_eq!(themes.len(), 1);
    assert_eq!(themes[0].name, master_theme.name);
}

#[test]
fn master_theme_accessor_rejects_external_theme_relationships() {
    let mut package = Package::new().unwrap();
    let master_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    let relationship_id = package
        .opc_package()
        .get_part(&master_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .unwrap()
        .r_id()
        .to_string();
    let master = package
        .opc_package_mut()
        .get_part_mut(&master_name)
        .unwrap();
    master.rels_mut().remove(&relationship_id);
    master.rels_mut().add_relationship(
        rt::THEME.to_string(),
        "https://example.invalid/theme.xml".to_string(),
        relationship_id,
        true,
    );

    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    assert!(matches!(
        masters[0].theme(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
    assert!(matches!(
        presentation.get_themes(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
}
