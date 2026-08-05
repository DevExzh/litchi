use litchi_opc::constants::relationship_type::{HYPERLINK, SLIDE};
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use litchi_pptx::Error;
use litchi_pptx::actions::{Jump, Kind, Limits, Target, Trigger, load_slide_action_settings};

const LOCAL_ACTIONS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/actions/basic_actions.xml");

#[test]
fn slide_action_owner_reports_local_action_settings() {
    let package = package_with_local_actions();
    let settings = load_actions(&package);

    assert_eq!(settings.len(), 7);

    assert_eq!(settings[0].slide_index(), 0);
    assert_eq!(settings[0].action_index(), 0);
    assert_eq!(settings[0].trigger(), Trigger::Click);
    assert_eq!(
        settings[0].kind(),
        Kind::Presentation {
            start_slide_index: 2
        }
    );
    assert_eq!(settings[0].relationship_id(), Some("rIdExternal"));
    assert_eq!(settings[0].tooltip(), Some("Open presentation"));
    assert_eq!(settings[0].target_frame(), Some("_blank"));
    assert!(matches!(
        settings[0].target(),
        Some(Target::External {
            target,
            relationship_type,
        }) if target == "https://example.invalid/other.pptx" && relationship_type == HYPERLINK
    ));

    assert_eq!(settings[1].trigger(), Trigger::Hover);
    assert_eq!(settings[1].kind(), Kind::SlideShowJump(Jump::NextSlide));
    assert_eq!(settings[1].target(), None);

    assert_eq!(settings[2].kind(), Kind::SlideJump);
    assert!(matches!(
        settings[2].target(),
        Some(Target::Internal {
            part_name,
            relationship_type,
        }) if part_name.as_str() == "/ppt/slides/slide2.xml" && relationship_type == SLIDE
    ));
    assert_eq!(settings[3].kind(), Kind::Macro);
    assert_eq!(
        settings[3].action(),
        Some("ppaction://macro?name=Module1.Run")
    );
    assert_eq!(settings[4].kind(), Kind::Program);
    assert_eq!(settings[5].kind(), Kind::Media);
    assert_eq!(settings[5].relationship_id(), None);
    assert_eq!(settings[6].kind(), Kind::Unknown);
    assert_eq!(settings[6].action(), Some("urn:producer:unrecognized"));
}

#[test]
fn slide_action_owner_rejects_missing_action_relationships() {
    let fragment = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="101" name="Broken"><a:hlinkClick r:id="rIdMissing" action="ppaction://hlinkfile"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>"#;
    let package = package_with_actions(fragment, false);
    let slide = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    let mut limits = Limits::default();

    assert!(matches!(
        load_slide_action_settings(&package, 0, slide, &mut limits),
        Err(Error::Relationship(message)) if message.contains("rIdMissing")
    ));
}

fn load_actions(package: &OpcPackage) -> Vec<litchi_pptx::actions::Setting> {
    let slide = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    load_slide_action_settings(package, 0, slide, &mut Limits::default()).unwrap()
}

fn package_with_local_actions() -> OpcPackage {
    package_with_actions(std::str::from_utf8(LOCAL_ACTIONS).unwrap(), true)
}

fn package_with_actions(fragment: &str, add_relationships: bool) -> OpcPackage {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_vec(bytes).unwrap();
    install_actions(&mut package, fragment, add_relationships);
    package
}

fn install_actions(package: &mut OpcPackage, fragment: &str, add_relationships: bool) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replacen("</p:spTree>", &format!("{fragment}</p:spTree>"), 1);
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());

    if add_relationships {
        slide.rels_mut().add_relationship(
            HYPERLINK.to_string(),
            "https://example.invalid/other.pptx".to_string(),
            "rIdExternal".to_string(),
            true,
        );
        slide.rels_mut().add_relationship(
            SLIDE.to_string(),
            "slide2.xml".to_string(),
            "rIdSlide".to_string(),
            false,
        );
    }
}
