#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "bounded fixture assertions panic on failure by design"
)]

use super::{Limits, apply_patch, load_snapshot};
use crate::Error;
use litchi_drawingml::ink::actions::ActionSelector;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const ACTION: &str = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction";

fn fixture() -> (OpcPackage, PackURI) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let target_name = PackURI::new("/ppt/actions.xml").unwrap();
    let slide_xml = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p14="{P14}" xmlns:ia="{ACTION}"><p:spTree><mc:AlternateContent><mc:Choice Requires="p14 ia"><p:contentPart r:id="rIdAction"/></mc:Choice><mc:Fallback><p:pic/></mc:Fallback></mc:AlternateContent></p:spTree></p:sld>"#
    );
    let target_xml = format!(
        r#"<ia:actions xmlns:ia="{ACTION}" lengthUnit="cm" timeUnit="ms"><ia:action type="add" startTime="0"/></ia:actions>"#
    );
    let mut package = OpcPackage::new();
    let mut slide = BlobPart::new(
        slide_name.clone(),
        litchi_opc::constants::content_type::PML_SLIDE.to_owned(),
        slide_xml.into_bytes(),
    );
    slide
        .rels_mut()
        .try_add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "../actions.xml".to_owned(),
            "rIdAction".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(slide)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            target_name,
            "text/xml".to_owned(),
            target_xml.into_bytes(),
        )))
        .unwrap();
    (package, slide_name)
}

#[test]
fn reads_exact_mce_and_shared_profile() {
    let (package, slide_name) = fixture();
    let slide = package.get_part(&slide_name).unwrap();
    let snapshot = load_snapshot(&package, 0, slide, &Limits::default()).unwrap();
    assert_eq!(snapshot.anchors().len(), 1);
    let anchor = &snapshot.anchors()[0];
    assert_eq!(anchor.semantic_ordinal(), 0);
    assert_eq!(anchor.source_ordinal(), 0);
    assert_eq!(anchor.relationship_id(), "rIdAction");
    assert_eq!(anchor.content_type(), "text/xml");
    assert_eq!(anchor.profile().actions().count(), 1);
    assert!(
        std::str::from_utf8(anchor.owner_xml())
            .unwrap()
            .contains("Fallback")
    );
}

#[test]
fn shared_edit_replaces_only_target_and_inverse_restores_bytes() {
    let (mut package, slide_name) = fixture();
    let source = {
        let slide = package.get_part(&slide_name).unwrap();
        load_snapshot(&package, 0, slide, &Limits::default()).unwrap()
    };
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_start_time(ActionSelector::ordinal(0), "1")?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    apply_patch(&mut package, commit.patch()).unwrap();
    let changed = {
        let slide = package.get_part(&slide_name).unwrap();
        load_snapshot(&package, 0, slide, &Limits::default()).unwrap()
    };
    assert_eq!(
        changed.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .start_time(),
        "1"
    );
    assert_eq!(
        changed.anchors()[0].fallback_xml(),
        source.anchors()[0].fallback_xml()
    );
    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    let restored = {
        let slide = package.get_part(&slide_name).unwrap();
        load_snapshot(&package, 0, slide, &Limits::default()).unwrap()
    };
    assert!(restored.same_source(&source));
}

#[test]
fn rejects_wrong_target_type_and_stale_context() {
    let (mut package, slide_name) = fixture();
    let source = {
        let slide = package.get_part(&slide_name).unwrap();
        load_snapshot(&package, 0, slide, &Limits::default()).unwrap()
    };
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.replace_profile_bytes(
        selector,
        format!(r#"<ia:actions xmlns:ia="{ACTION}" lengthUnit="cm" timeUnit="ms"/>"#),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/actions.xml").unwrap())
        .unwrap()
        .set_content_type("application/inkml+xml".to_owned())
        .unwrap();
    assert!(matches!(
        apply_patch(&mut package, commit.patch()),
        Err(Error::ContentType { .. })
    ));
}
