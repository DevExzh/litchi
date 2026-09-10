#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test fixtures and assertions panic on failure by design"
)]

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::{Error, Package};

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P1710_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2017/10/main";
const URI: &str = "{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}";

fn package_with_properties(xml: &[u8]) -> Package {
    let mut authored = Package::new().expect("new package");
    let bytes = authored.to_bytes().expect("serialize package");
    let mut opc = OpcPackage::from_bytes(&bytes).expect("reopen package");
    let properties = PackURI::new("/ppt/presProps.xml").expect("properties part URI");
    opc.get_part_mut(&properties)
        .expect("properties part")
        .set_blob(xml.to_vec());
    let presentation = PackURI::new("/ppt/presentation.xml").expect("presentation URI");
    let presentation_part = opc.get_part_mut(&presentation).expect("presentation part");
    if !presentation_part
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == rt::PRES_PROPS)
    {
        presentation_part.rels_mut().add_relationship(
            rt::PRES_PROPS.to_owned(),
            "presProps.xml".to_owned(),
            "rIdReadonlyRecommended".to_owned(),
            false,
        );
    }
    Package::from_opc_package(opc).expect("open package")
}

#[test]
fn package_creation_round_trip_and_owner_retention() {
    let mut package = Package::new().expect("new package");
    assert_eq!(package.readonly_recommended().unwrap(), None);
    assert_eq!(package.put_readonly_recommended(true).unwrap(), None);
    assert_eq!(package.read_only_recommended().unwrap(), Some(true));
    let bytes = package.to_bytes().expect("serialize changed package");
    let mut reopened = Package::from_bytes(&bytes).expect("reopen changed package");
    assert_eq!(reopened.readonly_recommended().unwrap(), Some(true));
    assert_eq!(reopened.remove_read_only_recommended().unwrap(), Some(true));
    assert_eq!(reopened.readonly_recommended().unwrap(), None);
    assert_eq!(reopened.remove_readonly_recommended().unwrap(), None);

    let properties = PackURI::new("/ppt/presProps.xml").unwrap();
    assert!(reopened.opc().unwrap().get_part(&properties).is_ok());
}

#[test]
fn package_snapshot_patch_is_source_checked_and_reversible() {
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{P1710_NS}"><p:extLst><!-- keep --><p:ext uri="{URI}"><p1710:readonlyRecommended val="1"/></p:ext><p:ext uri="urn:other"><o:data xmlns:o="urn:other"/></p:ext></p:extLst></p:presentationPr>"#
    );
    let mut package = package_with_properties(xml.as_bytes());
    let snapshot = package
        .readonly_recommended_snapshot()
        .unwrap()
        .expect("owner snapshot");
    let original = snapshot.source_xml().to_vec();
    let mut edit = snapshot.edit();
    edit.set(Some(false)).unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    package
        .apply_read_only_recommended_commit(commit)
        .expect("publish commit");
    assert_eq!(package.readonly_recommended().unwrap(), Some(false));
    package
        .apply_readonly_recommended_patch(&inverse)
        .expect("publish inverse");
    assert_eq!(package.readonly_recommended().unwrap(), Some(true));
    let restored = package
        .readonly_recommended_snapshot()
        .unwrap()
        .expect("restored owner snapshot");
    assert_eq!(restored.source_xml(), original.as_slice());
}

#[test]
fn exact_signed_noop_is_allowed_but_changed_edit_requires_explicit_policy() {
    let mut package = Package::new().expect("new package");
    assert_eq!(package.put_readonly_recommended(true).unwrap(), None);
    package
        .edit_opc(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .expect("install inert signature origin");
    assert!(package.opc().unwrap().is_signed());

    assert_eq!(package.put_readonly_recommended(true).unwrap(), Some(true));
    assert!(package.opc().unwrap().is_signed());
    let error = package.put_readonly_recommended(false).unwrap_err();
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(package.opc().unwrap().is_signed());
}

#[test]
fn duplicate_presentation_properties_parts_are_rejected() {
    let mut authored = Package::new().expect("new package");
    let bytes = authored.to_bytes().expect("serialize package");
    let mut opc = OpcPackage::from_bytes(&bytes).expect("reopen package");
    let orphan = PackURI::new("/ppt/presProps-orphan.xml").expect("orphan URI");
    opc.try_add_part(Box::new(BlobPart::new(
        orphan,
        litchi_opc::constants::content_type::PML_PRES_PROPS.to_owned(),
        br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
            .to_vec(),
    )))
    .expect("add orphan owner");
    let package = Package::from_opc_package(opc).expect("open package");
    assert!(package.readonly_recommended().is_err());
}
