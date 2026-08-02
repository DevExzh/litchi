use litchi_ooxml::pptx::{HandoutMaster, Package};
use litchi_opc::{BlobPart, PackURI, constants};
use tempfile::NamedTempFile;

#[test]
fn presentation_exposes_the_default_notes_graph() {
    let package = Package::new().unwrap();
    let notes = package.presentation().unwrap().notes().unwrap().unwrap();

    assert_eq!(notes.master().part(), "/ppt/notesMasters/notesMaster1.xml");
    assert_eq!(notes.master().theme().part(), "/ppt/theme/theme2.xml");
    assert!(notes.slides().is_empty());
}

#[test]
fn default_notes_graph_survives_save_and_reopen() {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    let notes = reopened.presentation().unwrap().notes().unwrap().unwrap();
    assert_eq!(notes.master().part(), "/ppt/notesMasters/notesMaster1.xml");
    assert_eq!(notes.master().theme().part(), "/ppt/theme/theme2.xml");
    assert!(notes.slides().is_empty());
}

#[test]
fn handout_theme_allocation_does_not_overwrite_an_existing_theme() {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    let default_theme_uri = PackURI::new("/ppt/theme/theme1.xml").unwrap();
    let occupied_theme_uri = PackURI::new("/ppt/theme/theme3.xml").unwrap();
    let default_theme = package
        .opc_package()
        .get_part(&default_theme_uri)
        .unwrap()
        .blob()
        .to_vec();
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        occupied_theme_uri,
        constants::content_type::OFC_THEME.to_owned(),
        default_theme,
    )));
    package
        .presentation_mut()
        .unwrap()
        .set_handout_master(HandoutMaster::new());
    package.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    let handout_uri = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();
    let handout = reopened.opc_package().get_part(&handout_uri).unwrap();
    let theme = handout
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == constants::relationship_type::THEME)
        .expect("handout theme relationship");
    assert_eq!(theme.target_ref(), "../theme/theme4.xml");
    assert!(
        reopened
            .opc_package()
            .contains_part(&PackURI::new("/ppt/theme/theme4.xml").unwrap())
    );
}
