use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use litchi_opc::constants::relationship_type as rt;
use tempfile::NamedTempFile;

#[test]
fn master_layout_slide_and_presentation_theme_inventory_resolve() {
    let package = package_with_slide();
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();

    let master_theme = masters[0].theme().unwrap();
    assert!(!master_theme.name.is_empty());
    assert!(!master_theme.colors.is_empty());

    let slides = presentation.slides().unwrap();
    let layout = slides[0].layout().unwrap();
    assert_eq!(
        slides[0].master().unwrap().name().unwrap(),
        masters[0].name().unwrap()
    );
    assert_eq!(layout.theme().unwrap().name, master_theme.name);
    assert_eq!(slides[0].theme().unwrap().name, master_theme.name);

    let themes = presentation.get_themes().unwrap();
    assert_eq!(themes.len(), 1);
    assert_eq!(themes[0].name, master_theme.name);
}

#[test]
fn master_theme_accessor_rejects_external_theme_relationships() {
    let mut package = package_with_slide();
    let master_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    let relationship_id = package
        .opc()
        .unwrap()
        .get_part(&master_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .unwrap()
        .r_id()
        .to_string();
    package
        .edit_opc(|opc| {
            let master = opc.get_part_mut(&master_name)?;
            master.rels_mut().remove(&relationship_id);
            master.rels_mut().add_relationship(
                rt::THEME.to_string(),
                "https://example.invalid/theme.xml".to_string(),
                relationship_id,
                true,
            );
            Ok(())
        })
        .unwrap();

    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    let slides = presentation.slides().unwrap();
    assert!(matches!(
        masters[0].theme(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
    assert!(matches!(
        slides[0].layout().unwrap().theme(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
    assert!(matches!(
        slides[0].theme(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
    assert!(matches!(
        presentation.get_themes(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
}

fn package_with_slide() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();
    Package::open(output.path()).unwrap()
}
