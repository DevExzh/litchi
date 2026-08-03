use litchi_ooxml::pptx::{
    Package, PptxActionKind, PptxActionTarget, PptxActionTrigger, PptxSlideShowJump,
};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::relationship_type::{HYPERLINK, SLIDE};
use tempfile::NamedTempFile;

const LOCAL_ACTIONS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/actions/basic_actions.xml");

#[test]
fn package_inventory_reports_local_action_settings() {
    let package = package_with_local_actions();

    let settings = package.action_settings().unwrap();
    assert_eq!(settings.len(), 7);

    assert_eq!(settings[0].slide_index(), 0);
    assert_eq!(settings[0].action_index(), 0);
    assert_eq!(settings[0].trigger(), PptxActionTrigger::Click);
    assert_eq!(
        settings[0].kind(),
        PptxActionKind::Presentation {
            start_slide_index: 2
        }
    );
    assert_eq!(settings[0].relationship_id(), Some("rIdExternal"));
    assert_eq!(settings[0].tooltip(), Some("Open presentation"));
    assert_eq!(settings[0].target_frame(), Some("_blank"));
    assert!(matches!(
        settings[0].target(),
        Some(PptxActionTarget::External {
            target,
            relationship_type,
        }) if target == "https://example.invalid/other.pptx" && relationship_type == HYPERLINK
    ));

    assert_eq!(settings[1].trigger(), PptxActionTrigger::Hover);
    assert_eq!(
        settings[1].kind(),
        PptxActionKind::SlideShowJump(PptxSlideShowJump::NextSlide)
    );
    assert_eq!(settings[1].target(), None);

    assert_eq!(settings[2].kind(), PptxActionKind::SlideJump);
    assert!(matches!(
        settings[2].target(),
        Some(PptxActionTarget::Internal {
            part_name,
            relationship_type,
        }) if part_name.as_str() == "/ppt/slides/slide2.xml" && relationship_type == SLIDE
    ));
    assert_eq!(settings[3].kind(), PptxActionKind::Macro);
    assert_eq!(
        settings[3].action(),
        Some("ppaction://macro?name=Module1.Run")
    );
    assert_eq!(settings[4].kind(), PptxActionKind::Program);
    assert_eq!(settings[5].kind(), PptxActionKind::Media);
    assert_eq!(settings[5].relationship_id(), None);
    assert_eq!(settings[6].kind(), PptxActionKind::Unknown);
    assert_eq!(settings[6].action(), Some("urn:producer:unrecognized"));

    assert_eq!(
        package.presentation().unwrap().action_settings().unwrap(),
        settings
    );
}

#[test]
fn package_inventory_rejects_missing_action_relationships() {
    let fragment = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="101" name="Broken"><a:hlinkClick r:id="rIdMissing" action="ppaction://hlinkfile"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>"#;
    let package = package_with_actions(fragment, false);

    assert!(matches!(
        package.action_settings(),
        Err(OoxmlError::InvalidRelationship(message))
            if message.contains("rIdMissing")
    ));
}

fn package_with_local_actions() -> Package {
    package_with_actions(std::str::from_utf8(LOCAL_ACTIONS).unwrap(), true)
}

fn package_with_actions(fragment: &str, add_relationships: bool) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    install_actions(&mut package, fragment, add_relationships);
    package
}

fn install_actions(package: &mut Package, fragment: &str, add_relationships: bool) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name).unwrap();
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
            Ok(())
        })
        .unwrap();
}
